"""Advanced operation tools, on Apple Events.

Mirrors ``ppt_com/advanced_ops.py``. Same function names, same signatures, same
returned shapes; what differs is the walk through PowerPoint's object model.
The icon search is pure Python on both platforms and is not touched here.

Five things about this group of tools are worth knowing before reading on.

**Tags do not exist.** The whole dictionary contains one ``tag``, on
``command bar control``, and it has nothing to do with the ``Tags`` collection
Windows hangs off a shape, a slide and a presentation. There is no near miss to
degrade to, so both tag tools refuse.

**Shapes cannot be selected.** The dictionary has an ``unselect`` command and no
``select``, and the only ``shape range`` PowerPoint hands out is the one already
selected in a window. So ``ppt_select_shapes`` refuses on the same grounds
``ppt_group_shapes`` and ``ppt_merge_shapes`` do, and borrows layout.py's
wording so all of them agree. Reading the existing selection is a different
question and it does work, so ``ppt_get_selection`` is implemented rather than
refused.

**A shape can be exported even though a slide cannot.** There is no ``export``
command anywhere in the dictionary, which is why export.py renders slides from
the deck's PDF. But ``save as picture`` is declared on ``shape``, and
MACOS_PORT section 5.3 measured it writing a 28 KB PNG in 0.15 s. It takes no
size, so ``width`` and ``height`` come back as a warning rather than being
silently dropped, and it writes nothing at all when handed a path outside the
sandbox, so the file is staged inside PowerPoint's container and moved out here.

**Animation cannot be copied.** There is no ``pickup animation`` and no
``apply animation``; the ``pick up`` and ``apply`` commands in the dictionary
are the formatting pair, not the animation one. The only route left is reading
the source shape's ``animation settings`` and writing them onto the target,
which is the write MACOS_PORT section 5.2 measured turning every effect on the
slide into a plain appear and dropping exit animations outright. So this
refuses, and it refuses before touching PowerPoint.

**A group's members cannot be read.** groups.py measured a group of two text
boxes answering 0 for its ``shapes`` and -1728 for ``shapes[1]``. The font walk
still recurses, so a PowerPoint that starts answering needs no code change, but
it says in ``warnings`` that grouped text was probably not reached.
"""

import json
import logging
import os
import shutil
import subprocess
import tempfile
import urllib.error
import urllib.request

from appscript import k
from appscript.reference import CommandError

from backend.mac_ae import (
    EXPORT_STAGING_DIR,
    count,
    count_of,
    is_missing,
    ppt,
    shapes_of,
    target_window,
)
from backend.mac_enums import (
    MsoAutoShapeType,
    MsoPictureColorType,
    PpSelectionType,
    to_keyword,
)
from backend.unsupported import refusal as _refusal
from ppt_com.constants import (
    ICON_PACKAGE_BASE,
    PICTURE_COLOR_TYPE_MAP,
    PICTURE_COLOR_TYPE_NAMES,
    VIEW_TYPE_MAP,
    VIEW_TYPE_NAMES,
    msoShapeRectangle,
)
from ppt_mac.export import _staging_path
from ppt_mac.layout import _NO_SHAPE_RANGE
from ppt_mac.shapes import (
    _WIN_AUTO_SHAPE_TYPE,
    _WIN_SHAPE_TYPE,
    _apply_line_visibility,
    _get_shape,
    _slide,
    _win_constant,
)
from utils.color import hex_to_rgb_list, rgb_list_to_hex
from utils.navigation import goto_slide

from ppt_mac.shapes import place_in_zorder

logger = logging.getLogger(__name__)

# The two shape types Windows accepts as a picture, as macOS names them.
_PICTURE_TYPES = (k.shape_type_picture, k.shape_type_linked_picture)

# msoShapeRoundedRectangle, the one auto shape whose adjustment is a corner
# radius. Named here because ppt_com/constants.py does not export it.
_ROUNDED_RECTANGLE = 5

# The theme palette, in the order Windows numbers it. themes.py measured all
# twelve reading and writing through the slide master's theme colour scheme,
# and the accents land at 5 to 10, which is what `_resolve_color` needs.
_THEME_COLOR_INDEX = {
    "dark1": 1, "light1": 2, "dark2": 3, "light2": 4,
    "accent1": 5, "accent2": 6, "accent3": 7, "accent4": 8,
    "accent5": 9, "accent6": 10, "hyperlink": 11,
    "followed_hyperlink": 12,
}

# Windows numbers its shape export formats through SHAPE_FORMAT_MAP and macOS
# names them, and the numbers do not agree, so this is hand written rather than
# generated. scripts/gen_mac_enums.py pairs by name and these two enumerations
# have no names in common at all, Windows spelling them `ppShapeFormatPNG` and
# macOS `save as PNG file`. The suffix travels with the keyword because the
# staged file has to carry the extension PowerPoint expects.
_SHAPE_FORMATS = {
    0: (k.save_as_GIF_file, ".gif"),
    1: (k.save_as_JPG_file, ".jpg"),
    2: (k.save_as_PNG_file, ".png"),
    3: (k.save_as_BMP_file, ".bmp"),
}

# The two Windows shape formats macOS has no word for. `MsoPictureType` offers
# GIF, JPG, PNG, BMP and PDF and nothing vector beyond that.
_SHAPE_FORMAT_NAMES_UNSUPPORTED = {4: "wmf", 5: "emf"}

# macOS view types, by the keyword PowerPoint answers with, giving the Windows
# constant and the name this project's tools speak.
#
# Carried here rather than looked up through VIEW_TYPE_MAP because that map and
# the generated PpViewType table disagree, and reading one through the other
# would be wrong in two places. VIEW_TYPE_MAP calls 1 "normal" where the Windows
# constant at 1 is ppViewSlide, and calls 10 "reading" where the constant at 10
# is ppViewPrintPreview on both platforms. The input vocabulary is kept exactly
# as Windows has it so the same call works on either machine, and this table is
# what the read back is named from.
#
# EPPViewType numbers every one of its eleven members the way Windows numbers
# PpViewType, so 2 and 3 are paired on their position even though the generator
# could not pair them on their names, macOS calling them `master view` and
# `page view` where Windows says ppViewSlideMaster and ppViewNotesPage.
_VIEW_TYPES = {
    k.slide_view: (1, "normal"),
    k.master_view: (2, "slide_master"),
    k.page_view: (3, "notes_page"),
    k.handout_master_view: (4, "handout_master"),
    k.notes_master_view: (5, "notes_master"),
    k.outline_view: (6, "outline"),
    k.slide_sorter_view: (7, "slide_sorter"),
    k.title_master_view: (8, "title_master"),
    k.normal_view: (9, "normal_view"),
    k.print_preview: (10, "reading"),
    k.thumbnail_view: (11, "thumbnails"),
}

_VIEW_KEYWORD_BY_WINDOWS = {number: word for word, (number, _) in _VIEW_TYPES.items()}

# The macOS enumerator for each Windows selection constant, reversed, so a
# selection reads back as the same number a Windows client would see. The
# picture colour type is reversed for the same reason.
_WIN_SELECTION_TYPE = {word: number for number, word in PpSelectionType.items()}
_WIN_PICTURE_COLOR_TYPE = {
    word: number for number, word in MsoPictureColorType.items()
}

# How deep the font walk follows groups. Groups do not answer for their members
# here, so this only guards against a PowerPoint that starts to.
_GROUP_DEPTH = 4


def _staged_file(suffix: str) -> str:
    """Create an empty file inside PowerPoint's container and return its path.

    Anything this process downloads and then names to PowerPoint goes here.
    Handing it a path it has no grant for blocks for tens of seconds and then
    kills the application, and its own container is the one place it is always
    allowed, so a temp file in the system temp directory is not good enough.
    POSIX paths only; an HFS colon path is read as a literal filename.
    """
    os.makedirs(EXPORT_STAGING_DIR, exist_ok=True)
    handle, path = tempfile.mkstemp(
        prefix="ppt_mcp_media_", suffix=suffix, dir=EXPORT_STAGING_DIR
    )
    os.close(handle)
    return path


def _is_picture(shape) -> bool:
    """True when a shape is one of the two kinds Windows calls a picture."""
    try:
        return shape.shape_type() in _PICTURE_TYPES
    except CommandError:
        return False


def _require_picture(shape, tool_name: str):
    """Raise the message Windows raises when a shape is not a picture."""
    if _is_picture(shape):
        return
    try:
        type_word = shape.shape_type()
    except CommandError:
        type_word = None
    raise ValueError(
        f"Shape '{shape.name()}' is not a picture "
        f"(type={_win_constant(_WIN_SHAPE_TYPE, type_word)}). "
        f"{tool_name} only works on picture shapes."
    )


def _resolve_color(pres, color_str):
    """Resolve a colour name or hex string to '#RRGGBB'.

    Theme names go through the slide master's theme colour scheme, which
    themes.py measured reading and writing all twelve colours. When a deck
    answers ``missing value`` for it there is no palette to read, and saying so
    beats handing back a colour nobody asked for.
    """
    if color_str.startswith("#"):
        return color_str

    index = _THEME_COLOR_INDEX.get(color_str.lower())
    if index is None:
        raise ValueError(
            f"Unknown color '{color_str}'. Use '#RRGGBB' or theme name: "
            f"{list(_THEME_COLOR_INDEX.keys())}"
        )

    try:
        rgb = pres.slide_master.theme.theme_color_scheme.theme_colors[index].RGB()
    except CommandError as exc:
        raise ValueError(
            f"This deck's theme palette could not be read, so the theme colour "
            f"'{color_str}' cannot be resolved. Give a '#RRGGBB' value instead."
        ) from exc
    hex_value = rgb_list_to_hex(rgb)
    if hex_value is None:
        raise ValueError(
            f"This deck answered with no colour for '{color_str}', which "
            "PowerPoint for Mac does on a deck whose theme palette it will not "
            "hand over. Give a '#RRGGBB' value instead."
        )
    return hex_value


def _text_shapes(container, depth: int = 0):
    """Yield every shape under a container that carries a text frame.

    Groups are followed, but a group here reports no members at all, so the
    recursion is expected to end at the top level and the callers say so in
    their warnings rather than reporting a number they cannot stand behind.
    """
    for shape in shapes_of(container):
        try:
            has_text = bool(shape.has_text_frame())
        except CommandError:
            has_text = False
        if has_text:
            yield shape
        if depth < _GROUP_DEPTH:
            yield from _text_shapes(shape, depth + 1)


def _grouped_text_warning() -> str:
    """The one sentence every font walk owes the caller."""
    return (
        "Text inside grouped shapes was probably not reached. PowerPoint for "
        "Mac reports no members for a group over Apple Events, so a group is "
        "walked but answers empty. Ungroup with ppt_ungroup_shapes to include "
        "its text."
    )


# ---------------------------------------------------------------------------
# Tags
# ---------------------------------------------------------------------------
def _set_tag_impl(slide_index, shape_name_or_index, tag_name, tag_value, target_type):
    """Refuse, because PowerPoint for Mac has no tags to set.

    No PowerPoint is touched, not even to look the target up, because there is
    nothing that could be written whatever the target turns out to be.
    """
    return _refusal(
        "ppt_set_tag",
        "PowerPoint for Mac's Apple Event dictionary has no Tags collection on "
        "a shape, a slide or a presentation. The only `tag` in it belongs to a "
        "command bar control and is unrelated, so there is nowhere to put a "
        "name and value pair that PowerPoint would keep with the file.",
        [
            "Carry the value in the shape's name with ppt_update_shape",
            "ppt_set_slide_notes",
        ],
    )


def _get_tags_impl(slide_index, shape_name_or_index, target_type):
    """Refuse, for the reason ``_set_tag_impl`` gives."""
    return _refusal(
        "ppt_get_tags",
        "PowerPoint for Mac's Apple Event dictionary has no Tags collection on "
        "a shape, a slide or a presentation, so there is nothing to read. A "
        "deck that carries tags written on Windows still has them in the file; "
        "they are simply not reachable from here.",
        ["ppt_list_shapes", "ppt_get_slide_notes"],
    )


# ---------------------------------------------------------------------------
# Fonts
# ---------------------------------------------------------------------------
def _replace_font_impl(original_font, replacement_font):
    """Swap one font for another by walking the deck.

    Windows calls ``Presentation.Fonts.Replace``. There is no ``replace``
    command in the whole dictionary, so this reads each shape's font and writes
    the replacement where it matches. Two things that does not reach are named
    in ``warnings`` rather than left for the caller to discover. A shape whose
    text mixes fonts answers ``missing value`` for ``font name`` and is skipped,
    and masters and layouts are not walked, so placeholder text that has never
    been overridden keeps the theme font.
    """
    ppt._get_app_impl()
    pres = ppt._get_pres_impl()

    wanted = original_font.strip().lower()
    shapes_updated = 0
    mixed = 0
    refused = 0
    for index in range(1, count(pres.slides) + 1):
        slide = pres.slides[index]
        for shape in _text_shapes(slide):
            font = shape.text_frame.text_range.font
            try:
                latin = font.font_name()
                east_asian = font.east_asian_name()
            except CommandError:
                continue
            if is_missing(latin) and is_missing(east_asian):
                mixed += 1
                continue
            wrote_latin = False
            wrote_east_asian = False
            if not is_missing(latin) and str(latin).strip().lower() == wanted:
                font.font_name.set(replacement_font)
                wrote_latin = True
            if (
                not is_missing(east_asian)
                and str(east_asian).strip().lower() == wanted
            ):
                font.east_asian_name.set(replacement_font)
                wrote_east_asian = True
            # Counted from what the font reads back as, not from a write that
            # did not raise. The read is the same one two lines above, so the
            # route is known to work, and this number is the whole answer the
            # tool gives.
            if wrote_latin or wrote_east_asian:
                if _font_reads_back(font, replacement_font,
                                    wrote_latin, wrote_east_asian):
                    shapes_updated += 1
                else:
                    refused += 1

    warnings = [
        "PowerPoint for Mac has no font replace command, so this walked the "
        "deck shape by shape. Masters and layouts were not walked, so "
        "placeholder text still using the theme font is unchanged; use "
        "ppt_set_default_fonts for that.",
        _grouped_text_warning(),
    ]
    if mixed:
        warnings.append(
            f"{mixed} shape(s) mix more than one font in their text and were "
            "skipped, because PowerPoint answers with no font name at all for "
            "those rather than with a list."
        )
    if refused:
        warnings.append(
            f"{refused} shape(s) were written and still read back with their "
            "old font, so PowerPoint declined those without saying so. They "
            "are not counted in shapes_updated."
        )

    return {
        "success": True,
        "original_font": original_font,
        "replacement_font": replacement_font,
        "shapes_updated": shapes_updated,
        "warnings": warnings,
    }


def _font_reads_back(font, expected, check_latin: bool, check_east_asian: bool) -> bool:
    """Say whether a font write is actually there.

    One Apple Event per face written. Both faces are checked when both were
    written, because PowerPoint taking the Latin one says nothing about the
    East Asian one, and a deck set only through `font name` renders Japanese in
    the theme font.
    """
    try:
        if check_latin and str(font.font_name()) != expected:
            return False
        if check_east_asian and str(font.east_asian_name()) != expected:
            return False
    except CommandError:
        return False
    return True


def _list_fonts_impl():
    """List the fonts the deck uses, by walking it.

    ``presentation`` declares a ``font`` element and it is a trap. Asking the
    presentation how many fonts it has answers 2 on a deck using one font,
    asking the collection for its contents answers nothing, and ``fonts[1]``
    answers -1728. The count is real and the elements behind it are not
    reachable, which is the same shape of defect as everywhere else in this
    port.

    So the deck is walked instead, the way ``_replace_font_impl`` walks it,
    which answers the question the tool is actually asking. The theme fonts are
    included because placeholder text that has never been overridden is set in
    them, and a list that left them out would be missing the fonts most of a
    deck is written in.
    """
    ppt._get_app_impl()
    pres = ppt._get_pres_impl()

    found = []

    def remember(value):
        if is_missing(value) or not value:
            return
        name = str(value).strip()
        if name and name not in found:
            found.append(name)

    scheme = pres.slide_master.theme.theme_font_scheme
    for group in (scheme.major_theme_fonts, scheme.minor_theme_fonts):
        for slot in (1, 3):
            try:
                remember(group[slot].name())
            except CommandError:
                continue

    for index in range(1, count(pres.slides) + 1):
        for shape in _text_shapes(pres.slides[index]):
            font = shape.text_frame.text_range.font
            for attribute in ("font_name", "east_asian_name"):
                try:
                    remember(getattr(font, attribute)())
                except CommandError:
                    continue

    return {
        "success": True,
        "fonts_count": len(found),
        "fonts": sorted(found),
        "warnings": [
            "PowerPoint for Mac will not enumerate a presentation's fonts, so "
            "this walked the theme and every shape with text instead. A shape "
            "whose text mixes fonts answers nothing for its font name and "
            "contributes none, and masters and layouts beyond the theme are "
            "not walked."
        ],
    }


def _set_default_fonts_impl(latin, east_asian, apply_to_existing):
    """Set the theme fonts, and optionally rewrite the text already there.

    Windows walks ``Designs(n).SlideMaster.Theme``. ``presentation.designs``
    counts zero on an ordinary deck here and asking ``designs[1]`` for anything
    answers -1728, which themes.py recorded, so this goes straight to the slide
    master's theme font scheme. Latin sits at position 1 and East Asian at
    position 3 in both the major and the minor scheme, the same positions
    Windows addresses as ``MajorFont(1)`` and ``MajorFont(3)``.
    """
    if not latin and not east_asian:
        raise ValueError("At least one of 'latin' or 'east_asian' must be provided")

    ppt._get_app_impl()
    pres = ppt._get_pres_impl()

    warnings = []
    theme_updated = False
    try:
        scheme = pres.slide_master.theme.theme_font_scheme
        for family in (scheme.major_theme_fonts, scheme.minor_theme_fonts):
            if latin:
                family[1].name.set(latin)
            if east_asian:
                family[3].name.set(east_asian)
        # Nothing is trusted because it did not raise, and every slot written
        # is read back rather than one standing in for the others. That means
        # both families as well as both slots, because the major scheme taking
        # a face says nothing about the minor scheme, and the minor is the one
        # body text is set in. PowerPoint accepting the Latin face and quietly
        # dropping the East Asian one is the failure worth catching, and a deck
        # set only through `font name` renders Japanese in the theme font.
        theme_updated = True
        families = (
            ("major", scheme.major_theme_fonts),
            ("minor", scheme.minor_theme_fonts),
        )
        for slot, expected in ((1, latin), (3, east_asian)):
            if not expected:
                continue
            for family_name, family in families:
                probe = family[slot].name()
                if is_missing(probe) or str(probe) != expected:
                    theme_updated = False
                    warnings.append(
                        "PowerPoint reported no error but the "
                        f"{family_name} theme font reads back as {probe!r} "
                        f"rather than {expected!r}, so that part of the theme "
                        "was left alone. Text already on the slides was still "
                        "updated when apply_to_existing was set."
                    )
    except CommandError as exc:
        logger.warning("Could not update the theme fonts: %s", exc)
        warnings.append(
            f"The theme font scheme could not be written ({exc}). New text will "
            "keep the fonts it had."
        )

    slides_processed = 0
    shapes_updated = 0
    refused = 0
    if apply_to_existing:
        for index in range(1, count(pres.slides) + 1):
            slide = pres.slides[index]
            slides_processed += 1
            for shape in _text_shapes(slide):
                font = shape.text_frame.text_range.font
                try:
                    if latin:
                        font.font_name.set(latin)
                    if east_asian:
                        font.east_asian_name.set(east_asian)
                except CommandError:
                    continue
                # Counted from the read back, the same way ppt_replace_font
                # counts, so a shape PowerPoint declined is not in the total.
                latin_ok = not latin or _font_reads_back(font, latin, True, False)
                east_ok = (
                    not east_asian
                    or _font_reads_back(font, east_asian, False, True)
                )
                if latin_ok and east_ok:
                    shapes_updated += 1
                else:
                    refused += 1
        if slides_processed:
            # Only when there was something to walk. A deck with no slides yet,
            # which is where a caller sets the default fonts, was being told
            # that text inside groups had probably been missed, having walked
            # no shapes at all.
            warnings.append(_grouped_text_warning())
        if refused:
            warnings.append(
                f"{refused} shape(s) were written and still read back with "
                "their old font, so PowerPoint declined those without saying "
                "so. They are not counted in shapes_updated."
            )

    result = {"success": True, "theme_updated": theme_updated}
    if latin:
        result["latin"] = latin
    if east_asian:
        result["east_asian"] = east_asian
    if apply_to_existing:
        result["slides_processed"] = slides_processed
        result["shapes_updated"] = shapes_updated
    if warnings:
        result["warnings"] = warnings
    return result


# ---------------------------------------------------------------------------
# Picture crop
# ---------------------------------------------------------------------------
def _crop_picture_impl(slide_index, shape_name_or_index, crop_left, crop_right,
                       crop_top, crop_bottom, crop_shape, crop_fit, crop_anchor,
                       corner_radius_pt):
    from ppt_com.shapes import SHAPE_NAME_MAP

    # Resolve crop_shape first so a bad name aborts before anything is written,
    # which is the same ordering the Windows code keeps and for the same reason.
    auto_shape_int = None
    if crop_shape is not None:
        if isinstance(crop_shape, str):
            key = crop_shape.strip().lower()
            if key.isdigit():
                auto_shape_int = int(key)
            elif key not in SHAPE_NAME_MAP:
                raise ValueError(
                    f"Unknown crop_shape '{crop_shape}'. "
                    f"Available names: {', '.join(sorted(SHAPE_NAME_MAP.keys()))}"
                )
            else:
                auto_shape_int = SHAPE_NAME_MAP[key]
        else:
            auto_shape_int = int(crop_shape)

    if crop_fit is not None:
        fit_key = crop_fit.strip().lower()
        if fit_key not in ("square", "1:1"):
            raise ValueError(
                f"Unknown crop_fit '{crop_fit}'. Supported values: 'square', '1:1'"
            )

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, shape_name_or_index)
    _require_picture(shape, "ppt_crop_picture")

    pic_fmt = shape.picture_format

    if crop_fit is not None:
        anchor = crop_anchor if crop_anchor is not None else 0.5

        pic_fmt.crop_left.set(0)
        pic_fmt.crop_right.set(0)
        pic_fmt.crop_top.set(0)
        pic_fmt.crop_bottom.set(0)

        old_lock = shape.lock_aspect_ratio()
        shape.lock_aspect_ratio.set(False)

        cur_w = shape.width()
        cur_h = shape.height()

        # `scale width` and `scale height` take the same three arguments as
        # Windows's ScaleWidth and ScaleHeight, and a factor of 1 measured
        # against the original size is what puts the shape back to the image's
        # own dimensions so the aspect ratio can be read off it.
        shape.scale_width(
            factor=1.0,
            relative_to_original_size=True,
            scale=k.scale_from_top_left,
        )
        shape.scale_height(
            factor=1.0,
            relative_to_original_size=True,
            scale=k.scale_from_top_left,
        )
        orig_w = shape.width()
        orig_h = shape.height()

        shape.width.set(cur_w)
        shape.height.set(cur_h)

        if orig_w > orig_h:
            excess = orig_w - orig_h
            pic_fmt.crop_left.set(excess * anchor)
            pic_fmt.crop_right.set(excess * (1.0 - anchor))
        elif orig_h > orig_w:
            excess = orig_h - orig_w
            pic_fmt.crop_top.set(excess * anchor)
            pic_fmt.crop_bottom.set(excess * (1.0 - anchor))

        min_dim = min(cur_w, cur_h)
        shape.width.set(min_dim)
        shape.height.set(min_dim)

        shape.lock_aspect_ratio.set(bool(old_lock))
    else:
        if crop_left is not None:
            pic_fmt.crop_left.set(crop_left)
        if crop_right is not None:
            pic_fmt.crop_right.set(crop_right)
        if crop_top is not None:
            pic_fmt.crop_top.set(crop_top)
        if crop_bottom is not None:
            pic_fmt.crop_bottom.set(crop_bottom)

    if auto_shape_int is not None:
        try:
            shape.auto_shape_type.set(
                to_keyword(MsoAutoShapeType, auto_shape_int, "auto shape type")
            )
        except CommandError as exc:
            raise ValueError(
                f"Invalid crop_shape value {auto_shape_int}: PowerPoint for Mac "
                f"refused it. Apple Event error: {exc}"
            ) from exc

    if corner_radius_pt is not None:
        try:
            current_type = _win_constant(
                _WIN_AUTO_SHAPE_TYPE, shape.auto_shape_type()
            )
        except CommandError:
            current_type = None
        if current_type == _ROUNDED_RECTANGLE:
            short_side = min(shape.width(), shape.height())
            adj_value = min(0.5, corner_radius_pt / short_side) if short_side > 0 else 0
            # `adjustment 1`'s value is `adjustment_value`, never `value`.
            shape.adjustments[1].adjustment_value.set(adj_value)

    try:
        auto_shape_val = _win_constant(
            _WIN_AUTO_SHAPE_TYPE, shape.auto_shape_type()
        )
    except CommandError:
        auto_shape_val = None

    return {
        "success": True,
        "shape_name": shape.name(),
        "width": round(shape.width(), 2),
        "height": round(shape.height(), 2),
        "crop_left": round(pic_fmt.crop_left(), 2),
        "crop_right": round(pic_fmt.crop_right(), 2),
        "crop_top": round(pic_fmt.crop_top(), 2),
        "crop_bottom": round(pic_fmt.crop_bottom(), 2),
        "crop_shape": auto_shape_val,
    }


# ---------------------------------------------------------------------------
# Picture format
# ---------------------------------------------------------------------------
def _set_picture_format_impl(slide_index, shape_name_or_index, brightness,
                             contrast, color_type, transparent_color,
                             transparent_background):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, shape_name_or_index)
    _require_picture(shape, "ppt_set_picture_format")

    pf = shape.picture_format

    if brightness is not None:
        pf.brightness.set(brightness)
    if contrast is not None:
        pf.contrast.set(contrast)
    if color_type is not None:
        pf.color_type.set(
            to_keyword(
                MsoPictureColorType,
                PICTURE_COLOR_TYPE_MAP[color_type],
                "picture colour type",
            )
        )
    if transparent_color is not None:
        # Apple Events take three numbers where COM takes one packed one.
        pf.transparency_color.set(hex_to_rgb_list(transparent_color))
        if transparent_background is not False:
            pf.transparent_background.set(True)
    if transparent_background is not None:
        pf.transparent_background.set(bool(transparent_background))

    cur_color_type = _win_constant(_WIN_PICTURE_COLOR_TYPE, pf.color_type())

    return {
        "success": True,
        "shape_name": shape.name(),
        "brightness": pf.brightness(),
        "contrast": pf.contrast(),
        "color_type": cur_color_type,
        "color_type_name": PICTURE_COLOR_TYPE_NAMES.get(cur_color_type, "unknown"),
        "transparent_background": bool(pf.transparent_background()),
        "transparent_color_hex": rgb_list_to_hex(pf.transparency_color()),
    }


# ---------------------------------------------------------------------------
# Shape export
# ---------------------------------------------------------------------------
def _export_shape_impl(slide_index, shape_name_or_index, file_path, format_type, width, height):
    """Write one shape to an image file, through PowerPoint's own container.

    ``save as picture`` writes nothing at all when handed a path the sandbox
    does not cover, and says nothing about it, so the file is written inside
    the container and moved out afterwards by this process, which is not
    sandboxed.
    """
    if isinstance(format_type, str):
        from ppt_com.constants import SHAPE_FORMAT_MAP

        fmt_key = format_type.strip().lower()
        if fmt_key not in SHAPE_FORMAT_MAP:
            raise ValueError(
                f"Unknown format '{format_type}'. "
                f"Valid values: {list(SHAPE_FORMAT_MAP.keys())}"
            )
        format_type = SHAPE_FORMAT_MAP[fmt_key]

    if format_type in _SHAPE_FORMAT_NAMES_UNSUPPORTED:
        name = _SHAPE_FORMAT_NAMES_UNSUPPORTED[format_type]
        return _refusal(
            "ppt_export_shape",
            f"PowerPoint for Mac cannot write {name}. Its `save as picture` "
            "command offers gif, jpg, png and bmp and nothing vector beyond "
            "that, so there is no format to fall back to that would still be "
            "a vector file.",
            ["Call ppt_export_shape with format='png'"],
            error=f"ppt_export_shape cannot write {name} on macOS",
        )
    if format_type not in _SHAPE_FORMATS:
        raise ValueError(
            f"Unknown format {format_type!r}. "
            f"Valid values: {sorted(_SHAPE_FORMATS)}"
        )

    picture_type, suffix = _SHAPE_FORMATS[format_type]

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, shape_name_or_index)
    shape_name = shape.name()

    abs_path = os.path.abspath(os.path.expanduser(file_path))
    out_dir = os.path.dirname(abs_path)
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    staged = _staging_path(suffix)
    try:
        shape.save_as_picture(picture_type=picture_type, file_name=staged)
        # Nothing is trusted because it did not raise. An export that writes
        # nothing is the exact failure MACOS_PORT section 5 records for this
        # command, and the file on disk is the only evidence worth acting on.
        if not os.path.exists(staged) or os.path.getsize(staged) == 0:
            return _refusal(
                "ppt_export_shape",
                "PowerPoint reported success but wrote no file, which is what "
                "it does when it declines an export without raising. Nothing "
                "was written to the path asked for. The tool itself works; "
                "this one call did not land.",
                ["ppt_export_images", "ppt_export_pdf"],
                error="ppt_export_shape wrote no file",
            )
        shutil.move(staged, abs_path)
    finally:
        if os.path.exists(staged):
            os.remove(staged)

    result = {
        "success": True,
        "shape_name": shape_name,
        "file_path": abs_path,
    }
    if width is not None or height is not None:
        result["warnings"] = [
            "PowerPoint for Mac's `save as picture` takes no size, so width "
            "and height were ignored and the shape was written at its own "
            "size. Resize the shape with ppt_update_shape before exporting, or "
            "use ppt_export_images, which renders at any size asked for."
        ]
    return result


# ---------------------------------------------------------------------------
# Slide hidden
# ---------------------------------------------------------------------------
def _set_slide_hidden_impl(slide_index, hidden):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)

    # macOS takes a real boolean here, not msoTrue and msoFalse.
    slide.slide_show_transition.hidden.set(bool(hidden))

    landed = slide.slide_show_transition.hidden()
    if is_missing(landed) or bool(landed) != bool(hidden):
        return _refusal(
            "ppt_set_slide_hidden",
            "PowerPoint reported no error but the slide's hidden flag reads "
            f"back as {landed!r}, so the write was a silent no-op. The tool "
            "itself works; this one call did not land.",
            error="ppt_set_slide_hidden did not change the slide",
        )

    return {
        "success": True,
        "slide_index": slide_index,
        "hidden": hidden,
    }


# ---------------------------------------------------------------------------
# Selection
# ---------------------------------------------------------------------------
def _select_shapes_impl(slide_index, shape_names):
    """Refuse, because nothing can be put into a selection from a script.

    No PowerPoint is touched. The dictionary settles this before any deck is
    involved, so there is no reason to move the view first.
    """
    return _refusal(
        "ppt_select_shapes",
        _NO_SHAPE_RANGE,
        [
            "Act on each shape by name, which every formatting tool takes",
            "ppt_batch_apply_formatting",
        ],
    )


def _get_selection_impl():
    """Read what is selected in the target deck's window.

    Windows asks ``ActiveWindow``. That is the wrong window here, because
    ``active window`` raises whenever PowerPoint's start gallery is in front and
    it can belong to a different deck than the one being worked on, which is why
    ``goto_slide`` addresses the presentation's own window too.

    The ranges PowerPoint hands back are references, and a reference PowerPoint
    hands back is not trusted, so they are counted through their container and
    then indexed rather than asked for their elements.
    """
    ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    # A deck can outlive its window, and there is nothing to read a selection
    # out of when it has. `target_window` says that in a sentence rather than
    # letting -1728 stand in for it.
    selection = target_window(pres).selection

    sel_word = selection.selection_type()
    sel_type = _WIN_SELECTION_TYPE.get(sel_word)

    result = {
        "success": True,
        "type": sel_type,
    }

    if sel_word == k.selection_type_none:
        result["type_name"] = "none"
    elif sel_word == k.selection_type_slides:
        result["type_name"] = "slides"
        indices = []
        try:
            total = count_of(selection.slide_range, k.slide)
            for i in range(1, int(total) + 1):
                indices.append(selection.slide_range.slides[i].slide_index())
        except CommandError as exc:
            logger.warning("Could not read the selected slides: %s", exc)
            result["warnings"] = [
                "PowerPoint would not say which slides are selected "
                f"({exc}); only the selection type could be read."
            ]
        result["slide_indices"] = indices
    elif sel_word == k.selection_type_shapes:
        result["type_name"] = "shapes"
        names = []
        try:
            total = count_of(selection.shape_range, k.shape)
            for i in range(1, int(total) + 1):
                names.append(selection.shape_range.shapes[i].name())
        except CommandError as exc:
            logger.warning("Could not read the selected shapes: %s", exc)
            result["warnings"] = [
                "PowerPoint would not say which shapes are selected "
                f"({exc}); only the selection type could be read."
            ]
        result["shape_names"] = names
        result["count"] = len(names)
    elif sel_word == k.selection_type_text:
        result["type_name"] = "text"
        content = selection.text_range.content()
        result["text"] = "" if is_missing(content) else content

    return result


# ---------------------------------------------------------------------------
# View
# ---------------------------------------------------------------------------
def _set_view_impl(view_type, zoom):
    """Set the window's view type and zoom.

    The writable ``view type`` is on the window, not on its ``view``, where the
    dictionary marks it read only. The zoom is the other way round.
    """
    if view_type is not None:
        vt_key = view_type.strip().lower().replace(" ", "_").replace("-", "_")
        if vt_key not in VIEW_TYPE_MAP:
            raise ValueError(
                f"Unknown view_type '{view_type}'. "
                f"Use one of: {', '.join(VIEW_TYPE_MAP.keys())}"
            )
        wanted = VIEW_TYPE_MAP[vt_key]
        keyword = _VIEW_KEYWORD_BY_WINDOWS.get(wanted)
        if keyword is None:
            return _refusal(
                "ppt_set_view",
                f"PowerPoint for Mac has no view matching '{view_type}'. Its "
                "EPPViewType offers slide, master, page, handout master, notes "
                "master, outline, slide sorter, title master, normal, print "
                "preview and thumbnail views.",
                ["Call ppt_set_view with view_type='normal'"],
                error=f"ppt_set_view cannot show '{view_type}' on macOS",
            )

    ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    # The view and the zoom both live on the window, so a deck without one has
    # nothing this tool can set or report.
    window = target_window(pres)

    warnings = []
    if view_type is not None:
        window.view_type.set(keyword)
        # Nothing is trusted because it did not raise. PowerPoint declines some
        # views depending on what the deck holds and says nothing about it.
        if window.view_type() != keyword:
            warnings.append(
                f"PowerPoint did not switch to '{view_type}' and reported no "
                "error. The view below is the one it is actually in."
            )

    if zoom is not None:
        window.view.zoom.set(zoom)

    current_word = window.view_type()
    current_view_type, current_name = _VIEW_TYPES.get(current_word, (None, None))
    current_zoom = window.view.zoom()

    # The zoom below is read back, so it only needs saying when it disagrees.
    # PowerPoint holds a view to its own range and reports no error for
    # clamping, and some views will not zoom at all.
    if zoom is not None and (is_missing(current_zoom) or current_zoom != zoom):
        warnings.append(
            f"A zoom of {zoom} was asked for and the window holds "
            f"{None if is_missing(current_zoom) else current_zoom}. PowerPoint "
            "keeps a view inside its own range and says nothing about doing it."
        )

    result = {
        "success": True,
        "view_type": current_name or VIEW_TYPE_NAMES.get(
            current_view_type, f"Unknown({current_view_type})"
        ),
        "view_type_id": current_view_type,
        "zoom": None if is_missing(current_zoom) else current_zoom,
    }
    if warnings:
        result["warnings"] = warnings
    return result


# ---------------------------------------------------------------------------
# Copy animation
# ---------------------------------------------------------------------------
def _copy_animation_impl(slide_index, source_shape, target_shape):
    """Refuse, because the only route to it costs the rest of the slide.

    No PowerPoint is touched, so a refused call does not even move the view.
    """
    return _refusal(
        "ppt_copy_animation",
        "PowerPoint for Mac has no pickup animation and no apply animation "
        "command; the `pick up` and `apply` pair in its dictionary copies "
        "formatting, not animation. The only remaining route is to read the "
        "source shape's animation settings and write them onto the target, and "
        "writing anything through animation settings rewrites the whole "
        "slide, turning every effect on it into a plain appear and dropping "
        "exit animations. That is measured behaviour, recorded in MACOS_PORT "
        "section 5.2, so it is not attempted.",
        [
            "Add the same effect to the target with ppt_add_animation",
            "ppt_list_animations",
        ],
    )


# ---------------------------------------------------------------------------
# Pictures from the network
# ---------------------------------------------------------------------------
def _place_picture(app, slide, path, left, top):
    """Put a picture on a slide and hand back a reference that resolves.

    ``make``'s own return value is not used. A new shape lands at the end of
    the z order, so counting before and after and taking the last one is both
    the check that something arrived and the way to reach it.
    """
    before = count(slide.shapes)
    app.make(
        new=k.picture,
        # At the slide, not at its shapes. `slide.shapes.end` raises -1708.
        at=slide.end,
        with_properties={
            k.file_name: path,
            k.left_position: left,
            k.top: top,
        },
    )
    shapes_now = shapes_of(slide)
    if len(shapes_now) != before + 1:
        raise RuntimeError(
            "PowerPoint reported success but the slide gained no shape, which "
            "is the silent no-op recorded in MACOS_PORT section 5."
        )
    picture = shapes_now[-1]
    if picture.shape_type() not in _PICTURE_TYPES:
        # Clear up before reporting. The empty box PowerPoint leaves behind is
        # 25 by 25 and unnamed, so a caller who retries collects one of these
        # per attempt and has to find them all by hand afterwards.
        removed = True
        try:
            picture.delete()
        except CommandError:
            removed = False
            logger.debug("Could not remove the empty shape left at %s", path)
        raise RuntimeError(
            "PowerPoint reported success but left a plain autoshape on the "
            "slide rather than a picture, which is what it does when it cannot "
            "read the file it was given."
            + ("" if removed else " The empty shape is still on the slide.")
        )
    return picture


def _fit_picture(picture, left, top, width, height) -> None:
    """Scale a picture into a box, keeping its proportions and centring it."""
    picture.lock_aspect_ratio.set(True)
    scale = min(width / picture.width(), height / picture.height())
    new_w = picture.width() * scale
    new_h = picture.height() * scale
    picture.width.set(new_w)
    picture.left_position.set(left + (width - new_w) / 2)
    picture.top.set(top + (height - new_h) / 2)


def _add_picture_from_url_impl(slide_index, url, left, top, width, height, svg_color, fit,
                               zorder="front"):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)

    resp = urllib.request.urlopen(url)
    content_type = resp.headers.get("Content-Type", "")
    is_svg = url.lower().endswith(".svg") or "svg" in content_type

    # The download lands inside PowerPoint's container rather than in the
    # system temp directory. A path it has no grant for blocks for tens of
    # seconds and then kills the application.
    if is_svg:
        svg_text = resp.read().decode("utf-8")
        if svg_color:
            svg_text = svg_text.replace("currentColor", svg_color)
        tmp_path = _staged_file(".svg")
        with open(tmp_path, "w", encoding="utf-8") as handle:
            handle.write(svg_text)
    else:
        data = resp.read()
        suffix = os.path.splitext(url.split("?")[0])[-1] or ".png"
        tmp_path = _staged_file(suffix)
        with open(tmp_path, "wb") as handle:
            handle.write(data)

    try:
        pic = _place_picture(app, slide, tmp_path, left, top)

        if fit and width is not None and height is not None:
            _fit_picture(pic, left, top, width, height)
        elif width is not None and height is not None:
            pic.lock_aspect_ratio.set(False)
            pic.width.set(width)
            pic.height.set(height)
        elif width is not None:
            pic.lock_aspect_ratio.set(True)
            pic.width.set(width)
        elif height is not None:
            pic.lock_aspect_ratio.set(True)
            pic.height.set(height)

        # Read before the move; the reference is stale afterwards.
        name = pic.name()
        size = (round(pic.width(), 2), round(pic.height(), 2))
        placed = place_in_zorder(slide, pic, zorder)
        return {
            "success": True,
            "shape_name": name,
            "shape_index": placed.get("z_position", pic.z_order_position()),
            "width": size[0],
            "height": size[1],
            "source_url": url,
            **placed,
        }
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


# PowerPoint for Mac cannot read an SVG. Handing it one is not an error it
# reports: it says the picture was made and leaves an empty 25 by 25 box on the
# slide instead. So the file it is handed is always a PNG, rendered here.
#
# `sips` does the rendering, which arrived with macOS 13. It honours `viewBox`
# and rasterises at whatever size is asked for rather than scaling the 48 by 48
# the file declares, keeps the alpha channel, and keeps the fill colour that was
# substituted in. All of that was measured rather than assumed, so a machine
# where it is not true should say so instead of leaving an empty box behind.
_SVG_RENDER_SCALE = 4       # pixels per point, so the icon survives a zoom
_SVG_RENDER_MIN = 256
_SVG_RENDER_MAX = 2048

_sips_svg_support = None


def _sips_renders_svg() -> bool:
    """Ask this machine, once, whether ``sips`` can read an SVG."""
    global _sips_svg_support
    if _sips_svg_support is not None:
        return _sips_svg_support

    probe_dir = tempfile.mkdtemp(prefix="ppt_mcp_svg_probe_")
    svg_path = os.path.join(probe_dir, "probe.svg")
    png_path = os.path.join(probe_dir, "probe.png")
    try:
        with open(svg_path, "w", encoding="utf-8") as handle:
            handle.write(
                '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10">'
                '<rect width="10" height="10" fill="#000000"/></svg>'
            )
        _sips_svg_support = _sips_to_png(svg_path, png_path, 32)
    except Exception:  # noqa: BLE001 - a probe that fails is a "no"
        _sips_svg_support = False
    finally:
        shutil.rmtree(probe_dir, ignore_errors=True)
    return _sips_svg_support


def _sips_to_png(svg_path: str, png_path: str, pixels: int) -> bool:
    """Rasterise an SVG to a PNG of ``pixels`` on its longest side."""
    try:
        subprocess.run(
            ["sips", "-s", "format", "png", svg_path,
             "--out", png_path, "-Z", str(pixels)],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            timeout=30, check=False,
        )
    except (OSError, subprocess.TimeoutExpired):
        return False
    return os.path.exists(png_path) and os.path.getsize(png_path) > 0


def _add_svg_icon_impl(slide_index, icon_name, left, top, width, height, color, style, filled,
                       zorder="front"):
    if not _sips_renders_svg():
        return _refusal(
            "ppt_add_svg_icon",
            "PowerPoint for Mac cannot read an SVG, so the file has to be "
            "rasterised first, and `sips` on this machine will not read one "
            "either. SVG support in `sips` arrived with macOS 13.",
            alternatives=[
                "Convert the icon to PNG yourself and use ppt_add_picture",
                "ppt_add_shape for a plain geometric marker",
            ],
        )

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)

    hex_color = _resolve_color(pres, color)

    base = ICON_PACKAGE_BASE
    file_name = f"{icon_name}-fill" if filled else icon_name
    svg_url = f"{base}/{style}/{file_name}.svg"

    try:
        resp = urllib.request.urlopen(svg_url)
    except urllib.error.HTTPError as e:
        if e.code == 404:
            raise ValueError(
                f"Icon '{icon_name}' not found (style='{style}', filled={filled}). "
                "The icon set this reads from is a pinned npm package, and it "
                "holds fewer names than the Google Fonts site lists, so a name "
                "that exists there can still be missing here. Search again with "
                "ppt_search_icons, which only offers names this package has. "
                f"URL: {svg_url}"
            ) from None
        raise
    svg_text = resp.read().decode("utf-8")

    svg_text = svg_text.replace("currentColor", hex_color)
    if f'fill="{hex_color}"' not in svg_text:
        svg_text = svg_text.replace("<svg ", f'<svg fill="{hex_color}" ', 1)

    pixels = int(max(width or 0, height or 0, 0) * _SVG_RENDER_SCALE)
    pixels = max(_SVG_RENDER_MIN, min(pixels, _SVG_RENDER_MAX))

    # The SVG itself never reaches PowerPoint, so it can live in the ordinary
    # temp directory. Only the PNG has to be staged inside the container.
    svg_dir = tempfile.mkdtemp(prefix="ppt_mcp_svg_")
    svg_path = os.path.join(svg_dir, "icon.svg")
    png_path = _staged_file(".png")
    try:
        with open(svg_path, "w", encoding="utf-8") as handle:
            handle.write(svg_text)

        if not _sips_to_png(svg_path, png_path, pixels):
            raise RuntimeError(
                f"'{icon_name}' was downloaded but could not be rendered to a "
                "PNG, so nothing was put on the slide. The slide is unchanged."
            )

        pic = _place_picture(app, slide, png_path, left, top)
        _fit_picture(pic, left, top, width, height)

        name = pic.name()
        size = (round(pic.width(), 2), round(pic.height(), 2))
        placed = place_in_zorder(slide, pic, zorder)
        return {
            "success": True,
            "shape_name": name,
            "shape_index": placed.get("z_position", pic.z_order_position()),
            "width": size[0],
            "height": size[1],
            "icon_name": icon_name,
            **placed,
            "source_url": svg_url,
        }
    finally:
        shutil.rmtree(svg_dir, ignore_errors=True)
        if os.path.exists(png_path):
            os.remove(png_path)


# ---------------------------------------------------------------------------
# Default shape style
# ---------------------------------------------------------------------------
# No read route at all. `set shapes default properties` is a command with no
# matching property anywhere in the dictionary, so nothing says what the
# default style now is, and a new shape is the only way to find out.
_DEFAULT_STYLE_WARNING = (
    "PowerPoint for Mac takes the default shape style as one command and "
    "answers nothing about what it now holds, and no property reads it back. "
    "So this reports the command was sent rather than a measurement. Add a "
    "shape with ppt_add_shape to see what new shapes look like now."
)


def _set_default_shape_style_from_shape_impl(slide_index, shape_name_or_index):
    """Capture one shape's whole style as the default for new shapes.

    Returns an encoded string rather than a dict, because the tool hands what
    this returns straight back to the caller without encoding it. The Windows
    version does the same and the two have to agree.
    """
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    total = count(pres.slides)
    if slide_index > total:
        raise ValueError(f"slide_index {slide_index} out of range (1-{total})")
    goto_slide(app, slide_index)
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, shape_name_or_index)
    shape.set_shapes_default_properties()
    return json.dumps({
        "success": True,
        "source_shape": shape.name(),
        "warnings": [_DEFAULT_STYLE_WARNING],
    })


def _set_default_shape_style_impl(
    fill_type, fill_color,
    line_visible, line_color, line_weight,
    font_name, font_size, font_bold, font_italic, font_color,
):
    """Set the default style for new shapes through a throwaway template shape.

    No ``goto_slide`` here. This mode names no slide, and yanking the user's
    view to slide 1 for a tool that is not about slide 1 would be a surprise.
    Encoded rather than returned as a dict, for the reason above.
    """
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    if count(pres.slides) == 0:
        raise ValueError("Presentation has no slides. Add a slide first.")

    fill_color_hex = _resolve_color(pres, fill_color) if fill_color else None
    line_color_hex = _resolve_color(pres, line_color) if line_color else None
    font_color_hex = _resolve_color(pres, font_color) if font_color else None

    # A one point template shape, far enough off the canvas to be invisible.
    slide = pres.slides[1]
    before = count(slide.shapes)
    app.make(
        new=k.shape,
        at=slide.end,
        with_properties={
            k.auto_shape_type: to_keyword(
                MsoAutoShapeType, msoShapeRectangle, "shape type"
            ),
            k.left_position: -10000,
            k.top: -10000,
            k.width: 1,
            k.height: 1,
        },
    )
    shapes_now = shapes_of(slide)
    if len(shapes_now) != before + 1:
        return json.dumps(_refusal(
            "ppt_set_default_shape_style",
            "PowerPoint reported success but the template shape never arrived "
            "on the slide, which is the silent no-op recorded in MACOS_PORT "
            "section 5. Nothing was changed. The tool itself works; this one "
            "call did not land.",
            error="ppt_set_default_shape_style could not make a template shape",
        ))
    shape = shapes_now[-1]

    try:
        if fill_type == "none":
            shape.fill_format.visible.set(False)
        elif fill_color_hex is not None:
            shape.fill_format.visible.set(True)
            shape.fill_format.solid()
            shape.fill_format.fore_color.set(hex_to_rgb_list(fill_color_hex))

        if line_visible is not None:
            # `line format` has no `visible` here, so this is weight and
            # transparency standing in for it, the same way ppt_set_line does.
            _apply_line_visibility(shape.line_format, line_visible)
        if line_color_hex is not None:
            shape.line_format.fore_color.set(hex_to_rgb_list(line_color_hex))
        if line_weight is not None:
            shape.line_format.line_weight.set(line_weight)

        font = shape.text_frame.text_range.font
        if font_name is not None:
            font.font_name.set(font_name)
            font.east_asian_name.set(font_name)
        if font_size is not None:
            font.font_size.set(font_size)
        if font_bold is not None:
            font.bold.set(bool(font_bold))
        if font_italic is not None:
            font.italic.set(bool(font_italic))
        if font_color_hex is not None:
            font.font_color.set(hex_to_rgb_list(font_color_hex))

        shape.set_shapes_default_properties()
    finally:
        shape.delete()

    return json.dumps({"success": True, "warnings": [_DEFAULT_STYLE_WARNING]})


# ---------------------------------------------------------------------------
# Lock aspect ratio
# ---------------------------------------------------------------------------
def _lock_aspect_ratio_impl(slide_index, shape_name_or_index, locked):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    shape.lock_aspect_ratio.set(bool(locked))

    landed = shape.lock_aspect_ratio()
    if is_missing(landed) or bool(landed) != bool(locked):
        return _refusal(
            "ppt_lock_aspect_ratio",
            "PowerPoint reported no error but the shape's aspect ratio lock "
            f"reads back as {landed!r}, so the write was a silent no-op. The "
            "tool itself works; this one call did not land.",
            error="ppt_lock_aspect_ratio did not change the shape",
        )

    return {
        "success": True,
        "shape_name": shape.name(),
        "locked": locked,
    }
