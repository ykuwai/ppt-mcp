"""Shape operations, on Apple Events.

Mirrors ``ppt_com/shapes.py``. Same function names, same signatures, same
returned shapes; what differs is the walk through PowerPoint's object model.

Three habits run through this module and are worth reading once.

**Insertion goes at the slide, not at its shapes.** ``at=slide.end`` works and
``at=slide.shapes.end`` raises -1708, which is the least helpful error in the
whole dictionary because it names neither.

**Nothing is trusted because it did not raise.** PowerPoint answers a creation
it silently refused with an empty 25 by 25 autoshape and no error at all, so
every ``make`` here is checked against what was asked for. See MACOS_PORT.md
section 5.

**Reads are cheaper in bulk.** The property record of every shape on a slide
costs one Apple Event, so anything that walks all the shapes asks once rather
than per shape.
"""

import logging
import os
from typing import Optional, Union

from appscript import k
from appscript.reference import CommandError

from backend.mac_ae import count, elements, is_missing, osascript, ppt, raw
from backend.mac_enums import (
    MsoAutoShapeType,
    MsoFillType,
    MsoGradientStyle,
    MsoLineDashStyle,
    MsoShapeType,
    MsoVerticalAnchor,
    MsoZOrderCmd,
    PpParagraphAlignment,
    to_keyword,
)
from ppt_com.constants import GRADIENT_STYLE_MAP, SHAPE_TYPE_NAMES, msoGroup
from utils.color import hex_to_rgb_list, rgb_list_to_hex
from utils.navigation import goto_slide
from utils.redraw import FrozenRedraw

logger = logging.getLogger(__name__)

# The enumerator tables run Windows constant to macOS keyword, which is the
# direction every write needs. Reads need the way back, so each one used for
# reporting is inverted once here rather than searched every time.
_WIN_SHAPE_TYPE = {word: number for number, word in MsoShapeType.items()}
_WIN_AUTO_SHAPE_TYPE = {word: number for number, word in MsoAutoShapeType.items()}
_WIN_FILL_TYPE = {word: number for number, word in MsoFillType.items()}
_WIN_DASH_STYLE = {word: number for number, word in MsoLineDashStyle.items()}

# The four character codes for the two properties appscript cannot reach by
# name. Both names are taken by AppleScript's own built-in vocabulary, which
# has a different code for each, so appscript renames PowerPoint's out of the
# way and the plain spelling raises AttributeError.
_ROTATION = b"ShRt"
_DASH_STYLE = b"LFds"

# The size PowerPoint leaves behind when it decided not to make what it was
# asked for and said nothing about it.
_SILENT_FAILURE_SIZE = 25

_ALIGN_VALUES = {"left": 1, "center": 2, "right": 3, "justify": 4}

_VERTICAL_ANCHOR_VALUES = {
    "top": 1,       # msoAnchorTop
    "middle": 3,    # msoAnchorMiddle
    "bottom": 4,    # msoAnchorBottom
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _keyword_name(value) -> Optional[str]:
    """Return an appscript keyword's own name, as it reads in the dictionary."""
    if is_missing(value):
        return None
    return str(value).replace("k.", "").replace("_", " ")


def _win_constant(table: dict, word, default=None):
    """Translate a macOS enumerator back to the Windows constant it pairs with.

    A word with no pair is not an error. macOS has shape types Windows never
    named, and reporting None for the number beats reporting a near miss.
    """
    if is_missing(word):
        return default
    return table.get(word, default)


def _slide(pres, slide_index: int):
    """Return a slide reference, checking the index first.

    An out of range element reference does not fail where it is built, it fails
    somewhere later with -1728 and no mention of the index, so the range is
    checked here where the number is still in hand.
    """
    total = count(pres.slides)
    if slide_index < 1 or slide_index > total:
        raise ValueError(
            f"Slide index {slide_index} is out of range. "
            f"The presentation has {total} slides (1-based)."
        )
    return pres.slides[slide_index]


def _shape_names(slide) -> list:
    """Return the name of every shape on a slide, in one Apple Event."""
    return [
        "" if is_missing(name) else name
        for name in elements(slide.shapes.name)
    ]


def _shape_index(
    slide,
    name_or_index: Union[str, int, None] = None,
    shape_name: Optional[str] = None,
    shape_index: Optional[int] = None,
) -> int:
    """Resolve a shape identifier to its 1-based position on the slide.

    The index rather than the reference, because deleting and duplicating both
    need a number to address the shape from AppleScript.
    """
    if shape_name is not None:
        identifier = shape_name
    elif shape_index is not None:
        identifier = shape_index
    elif name_or_index is not None:
        identifier = name_or_index
    else:
        raise ValueError("Either shape_name or shape_index must be provided.")

    names = _shape_names(slide)

    if isinstance(identifier, int):
        if identifier < 1 or identifier > len(names):
            raise ValueError(
                f"Shape index {identifier} is out of range. "
                f"Slide has {len(names)} shapes (1-based)."
            )
        return identifier

    for position, name in enumerate(names, start=1):
        if name == identifier:
            return position
    raise ValueError(f"Shape '{identifier}' not found on this slide.")


def _get_shape(
    slide,
    name_or_index: Union[str, int, None] = None,
    shape_name: Optional[str] = None,
    shape_index: Optional[int] = None,
):
    """Find a shape on a slide by name or 1-based index."""
    return slide.shapes[_shape_index(slide, name_or_index, shape_name, shape_index)]


def _from_record(record: dict, key, shape, attribute: str):
    """Read one property out of a bulk record, falling back to one event.

    The record for a whole slide arrives in a single Apple Event, but
    PowerPoint leaves a key out rather than answering ``missing value`` for a
    property a particular shape does not carry. A missing key read straight out
    of the dict would come back as None and, for a boolean, quietly read as
    false, so anything absent is asked for directly instead.
    """
    if key in record:
        value = record[key]
        return None if is_missing(value) else value
    try:
        value = getattr(shape, attribute)()
    except (CommandError, AttributeError):
        return None
    return None if is_missing(value) else value


def _verify_created(shape, width, height, expected_type=None, what="shape"):
    """Check that PowerPoint made what it was asked for.

    It reports success and does nothing often enough that this is policy rather
    than caution. The signature of a silent refusal is an empty 25 by 25
    autoshape, and for anything built from a file it is a shape of the wrong
    type. Both come back with no error attached.
    """
    if shape is None or is_missing(shape):
        raise RuntimeError(
            f"PowerPoint answered the request for a {what} with nothing at all, "
            "so there is no way to tell what, if anything, it put on the slide."
        )
    if width is not None and height is not None:
        asked_for_the_default_size = (
            round(width) == _SILENT_FAILURE_SIZE
            and round(height) == _SILENT_FAILURE_SIZE
        )
        landed = (
            round(shape.width()) == _SILENT_FAILURE_SIZE
            and round(shape.height()) == _SILENT_FAILURE_SIZE
        )
        if landed and not asked_for_the_default_size:
            raise RuntimeError(
                f"PowerPoint reported success but left an empty 25 by 25 {what} "
                "on the slide, which is what it does when it declines a request "
                "without raising. Check the arguments and try again."
            )
    if expected_type is not None:
        actual_type = shape.shape_type()
        if actual_type != expected_type:
            raise RuntimeError(
                f"PowerPoint reported success but made a "
                f"'{_keyword_name(actual_type)}' rather than a "
                f"'{_keyword_name(expected_type)}'. It does this instead of "
                "raising when it cannot use what it was given."
            )


def _apply_font(font, font_name, font_size, bold, italic, font_color) -> None:
    """Apply the inline font settings shared by shapes and text boxes."""
    if font_name is not None:
        font.font_name.set(font_name)
        # East Asian characters take their own face on both platforms, and a
        # deck set only through `font name` renders Japanese in the theme font.
        font.east_asian_name.set(font_name)
    if font_size is not None:
        font.font_size.set(font_size)
    if bold is not None:
        # macOS takes a real boolean here, not msoTrue/msoFalse.
        font.bold.set(bool(bold))
    if italic is not None:
        font.italic.set(bool(italic))
    if font_color is not None:
        font.font_color.set(hex_to_rgb_list(font_color))


def _apply_alignment(text_range, align) -> None:
    """Apply a paragraph alignment given by its user facing name."""
    align_val = _ALIGN_VALUES.get(align.lower())
    if align_val is None:
        raise ValueError(
            f"Invalid align '{align}'. Must be one of: {sorted(_ALIGN_VALUES)}"
        )
    text_range.paragraph_format.alignment.set(
        to_keyword(PpParagraphAlignment, align_val, "alignment")
    )


def _resolve_image_path(file_path: str) -> str:
    """Return an absolute POSIX path, having checked the file is really there.

    Two macOS traps meet here. PowerPoint reports success for a picture whose
    file does not exist and leaves an empty placeholder behind, and an HFS
    colon path is taken as a literal filename rather than as a path.
    """
    absolute = os.path.abspath(os.path.expanduser(file_path))
    if not os.path.isfile(absolute):
        raise ValueError(f"Image file not found: {absolute}")
    return absolute


# ---------------------------------------------------------------------------
# Apple Event implementation functions (run on the worker thread via ppt.execute)
# ---------------------------------------------------------------------------
def _add_shape_impl(
    slide_index, shape_type_int, left, top, width, height, text,
    font_name, font_size, bold, italic, font_color, align,
    fill_color, fill_type, fill_color2, fill_gradient_style, fill_transparency,
    line_visible, line_color, line_weight,
    corner_radius, corner_radius_pt,
):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    # FrozenRedraw already degrades to a no-op where win32 is missing, so this
    # block does nothing on macOS. It is kept so the two platforms read the
    # same, and because the flicker it fixes barely appears here: the whole
    # call is one round of Apple Events rather than a visible sequence.
    with FrozenRedraw():
        goto_slide(app, slide_index)
        slide = _slide(pres, slide_index)
        shape = app.make(
            new=k.shape,
            # At the slide, not at its shapes. `slide.shapes.end` raises -1708.
            at=slide.end,
            with_properties={
                k.auto_shape_type: to_keyword(
                    MsoAutoShapeType, shape_type_int, "shape type"
                ),
                k.left_position: left,
                k.top: top,
                k.width: width,
                k.height: height,
            },
        )
        _verify_created(shape, width, height)
        return _apply_shape_attrs(
            shape, text, font_name, font_size, bold, italic, font_color, align,
            fill_color, fill_type, fill_color2, fill_gradient_style, fill_transparency,
            line_visible, line_color, line_weight, corner_radius, corner_radius_pt,
            width, height,
        )


def _apply_shape_attrs(
    shape, text, font_name, font_size, bold, italic, font_color, align,
    fill_color, fill_type, fill_color2, fill_gradient_style, fill_transparency,
    line_visible, line_color, line_weight, corner_radius, corner_radius_pt,
    width, height,
):
    if text:
        text = text.replace("\n", "\r")  # \n -> paragraph break (Enter)
        # \v (vertical tab) -> line break (Shift+Enter) — passed through as-is
        text_range = shape.text_frame.text_range
        text_range.content.set(text)

        if font_name is not None or font_size is not None or bold is not None \
                or italic is not None or font_color is not None:
            # The font hangs off the text range, not off the text frame.
            # Reaching for it any other way answers -10006.
            _apply_font(
                text_range.font, font_name, font_size, bold, italic, font_color
            )

        if align is not None:
            _apply_alignment(text_range, align)

    # Inline fill — avoids a follow-up ppt_set_fill call
    _VALID_FILL_TYPES = {"solid", "none", "gradient"}
    if fill_type is not None and fill_type not in _VALID_FILL_TYPES:
        raise ValueError(
            f"Invalid fill_type '{fill_type}'. Must be one of: {sorted(_VALID_FILL_TYPES)}"
        )
    if fill_color is not None or fill_type is not None or fill_transparency is not None:
        effective_type = fill_type or ("solid" if fill_color is not None else None)
        fill = shape.fill_format
        if effective_type == "none":
            fill.visible.set(False)
        elif effective_type == "gradient":
            gstyle = GRADIENT_STYLE_MAP.get(fill_gradient_style or "horizontal", 1)
            fill.two_color_gradient(
                style=to_keyword(MsoGradientStyle, gstyle, "gradient style"),
                variant=1,
            )
            if fill_color is not None:
                fill.fore_color.set(hex_to_rgb_list(fill_color))
            if fill_color2 is not None:
                fill.back_color.set(hex_to_rgb_list(fill_color2))
        elif effective_type == "solid":
            # `solid` is a command on macOS, the same as Fill.Solid() is a
            # method on Windows.
            fill.solid()
            if fill_color is not None:
                fill.fore_color.set(hex_to_rgb_list(fill_color))
        if fill_transparency is not None and effective_type != "none":
            fill.transparency.set(fill_transparency)

    # Inline line/border — avoids a follow-up ppt_set_line call
    if line_visible is not None:
        _apply_line_visibility(shape.line_format, line_visible)
    if line_color is not None:
        shape.line_format.fore_color.set(hex_to_rgb_list(line_color))
    if line_weight is not None:
        shape.line_format.line_weight.set(line_weight)

    # Corner radius for rounded rectangles
    if corner_radius is not None or corner_radius_pt is not None:
        from ppt_com.shapes import SHAPE_NAME_MAP  # local: ppt_com imports us

        try:
            rounded = to_keyword(
                MsoAutoShapeType, SHAPE_NAME_MAP["rounded_rectangle"], "shape type"
            )
            if shape.auto_shape_type() == rounded:
                if corner_radius_pt is not None:
                    # Absolute: convert points to the adjustment ratio, clamp to 0.5
                    short_side = min(width, height)
                    adj_value = min(0.5, corner_radius_pt / short_side)
                else:
                    # Ratio: map user-facing 0.0–1.0 to the adjustment's 0.0–0.5
                    adj_value = corner_radius * 0.5
                # An adjustment carries its number in `adjustment value`. Its
                # `value` is a different thing and setting that does nothing.
                shape.adjustments[1].adjustment_value.set(adj_value)
        except Exception:
            logger.warning("Failed to set corner_radius on shape '%s'", shape.name())

    return {
        "success": True,
        "shape_name": shape.name(),
        "shape_index": shape.z_order_position(),
        "shape_type": _win_constant(
            _WIN_AUTO_SHAPE_TYPE,
            shape.auto_shape_type(),
        ),
    }


def _apply_line_visibility(line, visible: bool) -> None:
    """Show or hide a border without a `visible` property to set.

    ``line format`` has no ``visible`` on macOS, so this is a stand-in and not
    the same thing. Weight 0 and full transparency both read as no border to
    the eye, but PowerPoint's own "no line" flag stays where it was, so a later
    weight or colour will bring the border back.

    A shape with no line format to speak of, a picture or a placeholder, answers
    with an error rather than ignoring the request. That is worth a warning and
    not worth throwing away a shape that has already been created and styled.
    """
    try:
        if visible:
            line.transparency.set(0.0)
            # A border hidden by the branch below has weight 0, and turning it
            # back on has to give it something to draw.
            if not line.line_weight():
                line.line_weight.set(1.0)
        else:
            line.line_weight.set(0.0)
            line.transparency.set(1.0)
    except CommandError as exc:
        logger.warning("This shape has no border to show or hide: %s", exc)


def _add_textbox_impl(
    slide_index, left, top, width, height, text,
    font_name, font_size, bold, italic, font_color, align,
    vertical_anchor,
):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    # A text box is its own class here rather than an orientation argument to a
    # shape, so the Windows Orientation parameter has nowhere to go; horizontal
    # is what PowerPoint makes by default anyway.
    textbox = app.make(
        new=k.text_box,
        at=slide.end,
        with_properties={
            k.left_position: left,
            k.top: top,
            k.width: width,
            k.height: height,
        },
    )
    # The type is checked as well as the size, because a `make` PowerPoint did
    # not want to honour comes back as a plain autoshape and says nothing.
    _verify_created(
        textbox, width, height,
        expected_type=k.shape_type_text_box, what="text box",
    )
    if text:
        text = text.replace("\n", "\r")  # \n -> paragraph break (Enter)
        # \v (vertical tab) -> line break (Shift+Enter) — passed through as-is
        textbox.text_frame.text_range.content.set(text)

    # Inline font — avoids a follow-up ppt_format_text call
    if any(x is not None for x in [font_name, font_size, bold, italic, font_color]):
        _apply_font(
            textbox.text_frame.text_range.font,
            font_name, font_size, bold, italic, font_color,
        )

    # Inline alignment — avoids a follow-up ppt_set_paragraph_format call
    if align is not None:
        _apply_alignment(textbox.text_frame.text_range, align)

    # Inline vertical anchor — avoids a follow-up ppt_set_textframe call
    if vertical_anchor is not None:
        anchor_val = _VERTICAL_ANCHOR_VALUES.get(vertical_anchor.lower())
        if anchor_val is None:
            raise ValueError(
                f"Invalid vertical_anchor '{vertical_anchor}'. "
                f"Must be one of: {sorted(_VERTICAL_ANCHOR_VALUES)}"
            )
        textbox.text_frame.vertical_anchor.set(
            to_keyword(MsoVerticalAnchor, anchor_val, "vertical anchor")
        )

    return {
        "success": True,
        "shape_name": textbox.name(),
        "shape_index": textbox.z_order_position(),
    }


def _add_picture_impl(slide_index, file_path, left, top, width, height):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    absolute = _resolve_image_path(file_path)
    # Insert at natural size first to obtain the true aspect ratio, the same
    # reason Windows passes -1 for both dimensions.
    pic = app.make(
        new=k.picture,
        at=slide.end,
        with_properties={
            k.file_name: absolute,
            k.left_position: left,
            k.top: top,
        },
    )
    # A file PowerPoint could not read leaves a plain autoshape behind and says
    # nothing, so the type is what proves the picture actually arrived.
    _verify_created(
        pic, None, None, expected_type=k.shape_type_picture, what="picture"
    )
    if width is not None and height is not None:
        # Both specified: user intentionally overrides aspect ratio.
        pic.lock_aspect_ratio.set(False)
        pic.width.set(width)
        pic.height.set(height)
    elif width is not None:
        pic.lock_aspect_ratio.set(True)
        pic.width.set(width)
    elif height is not None:
        pic.lock_aspect_ratio.set(True)
        pic.height.set(height)
    return {
        "success": True,
        "shape_name": pic.name(),
        "shape_index": pic.z_order_position(),
        "width": round(pic.width(), 2),
        "height": round(pic.height(), 2),
    }


def _add_line_impl(slide_index, begin_x, begin_y, end_x, end_y):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    line = app.make(
        new=k.line_shape,
        at=slide.end,
        with_properties={
            k.begin_line_X: begin_x,
            k.begin_line_Y: begin_y,
            k.end_line_X: end_x,
            k.end_line_Y: end_y,
        },
    )
    _verify_created(
        line, None, None, expected_type=k.shape_type_line, what="line"
    )
    return {
        "success": True,
        "shape_name": line.name(),
        "shape_index": line.z_order_position(),
    }


def _list_shapes_impl(slide_index):
    ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)

    # Every property of every shape on the slide, in one Apple Event. Walking
    # them one at a time is correct but costs an event per property per shape.
    total = count(slide.shapes)
    records = elements(slide.shapes.properties)
    if len(records) != total:
        # PowerPoint did not answer with one record per shape. Rather than
        # report a short slide, fall back to asking each shape directly.
        records = [{} for _ in range(total)]

    shapes = []
    for i, record in enumerate(records, start=1):
        shape = slide.shapes[i]
        if not isinstance(record, dict):
            record = {}

        type_word = _from_record(record, k.shape_type, shape, "shape_type")
        type_int = _win_constant(_WIN_SHAPE_TYPE, type_word)

        has_text = bool(_from_record(record, k.has_text_frame, shape, "has_text_frame"))
        text_preview = ""
        if has_text:
            try:
                if shape.text_frame.has_text():
                    full_text = shape.text_frame.text_range.content()
                    if not is_missing(full_text):
                        text_preview = full_text[:50] + (
                            "..." if len(full_text) > 50 else ""
                        )
            except CommandError:
                pass

        shapes.append({
            "index": i,
            "name": _from_record(record, k.name, shape, "name"),
            # macOS exposes no shape id, only stacking order, and a made-up id
            # would be worse than an honest null.
            "id": None,
            "type": type_int,
            "type_name": (
                SHAPE_TYPE_NAMES.get(type_int)
                or _keyword_name(type_word)
                or f"Unknown({type_int})"
            ),
            "left": round(_from_record(record, k.left_position, shape, "left_position") or 0.0, 2),
            "top": round(_from_record(record, k.top, shape, "top") or 0.0, 2),
            "width": round(_from_record(record, k.width, shape, "width") or 0.0, 2),
            "height": round(_from_record(record, k.height, shape, "height") or 0.0, 2),
            "has_text": has_text,
            "text_preview": text_preview,
        })
    return {
        "slide_index": slide_index,
        "shapes_count": total,
        "shapes": shapes,
    }


def _get_shape_info_impl(slide_index, shape_name, shape_index):
    ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, None, shape_name=shape_name, shape_index=shape_index)

    type_word = shape.shape_type()
    type_int = _win_constant(_WIN_SHAPE_TYPE, type_word)
    name = shape.name()
    # `rotation` is one of the two names AppleScript's own vocabulary has
    # already claimed, so it is reached by its four character code.
    try:
        rotation = raw(shape, _ROTATION).get()
    except CommandError:
        rotation = None

    info = {
        "name": name,
        # No shape id on macOS; see _list_shapes_impl.
        "id": None,
        "type": type_int,
        "type_name": (
            SHAPE_TYPE_NAMES.get(type_int)
            or _keyword_name(type_word)
            or f"Unknown({type_int})"
        ),
        "left": round(shape.left_position(), 2),
        "top": round(shape.top(), 2),
        "width": round(shape.width(), 2),
        "height": round(shape.height(), 2),
        "rotation": None if is_missing(rotation) else round(rotation, 2),
        "z_order": shape.z_order_position(),
        "is_group": type_int == msoGroup,
        "has_animation": None,
        "aspect_ratio_locked": False,
        "text": None,
        "fill": None,
        "line": None,
    }

    # No animation check on macOS. Reading a slide's animation timeline is not
    # slow or unsupported, it takes PowerPoint down: touching
    # `timeline.main_sequence.effects` kills the application with -609, on a
    # slide with no animations at all, reproducibly. Reporting null here is the
    # only honest answer, and it costs less than losing the user's deck.
    info["has_animation"] = None
    info["has_animation_note"] = (
        "not read on macOS, because reading a slide's animation timeline "
        "crashes PowerPoint"
    )

    # Aspect ratio lock
    try:
        info["aspect_ratio_locked"] = bool(shape.lock_aspect_ratio())
    except Exception:
        pass

    # Text content
    try:
        if shape.has_text_frame() and shape.text_frame.has_text():
            content = shape.text_frame.text_range.content()
            info["text"] = None if is_missing(content) else content
    except Exception:
        pass

    # Fill info
    try:
        fill = shape.fill_format
        info["fill"] = {
            "type": _win_constant(_WIN_FILL_TYPE, fill.fill_format_type()),
            "visible": bool(fill.visible()),
        }
        try:
            color_hex = rgb_list_to_hex(fill.fore_color())
            if color_hex is not None:
                info["fill"]["color_hex"] = color_hex
        except Exception:
            pass
        try:
            info["fill"]["transparency"] = round(fill.transparency(), 2)
        except Exception:
            pass
    except Exception:
        pass

    # Line info
    try:
        line = shape.line_format
        info["line"] = {
            # `line format` carries no `visible` on macOS. Weight and
            # transparency can be read, but neither is PowerPoint's own no-line
            # flag, so guessing from them would answer a question nobody asked.
            "visible": None,
        }
        try:
            info["line"]["weight"] = round(line.line_weight(), 2)
        except Exception:
            pass
        try:
            color_hex = rgb_list_to_hex(line.fore_color())
            if color_hex is not None:
                info["line"]["color_hex"] = color_hex
        except Exception:
            pass
        try:
            # The other name AppleScript has claimed; reached by code.
            info["line"]["dash_style"] = _win_constant(
                _WIN_DASH_STYLE, raw(line, _DASH_STYLE).get()
            )
        except Exception:
            pass
    except Exception:
        pass

    # Connector info. The connector face of a shape is an element of it here
    # rather than a property, and it only resolves for a real connector.
    try:
        cf = shape.connectors[1].connector_format
        conn_info = {}
        try:
            if cf.begin_connected():
                conn_info["begin_connected_shape"] = cf.begin_connected_shape.name()
                conn_info["begin_connection_site"] = cf.begin_connection_site()
        except Exception:
            pass
        try:
            if cf.end_connected():
                conn_info["end_connected_shape"] = cf.end_connected_shape.name()
                conn_info["end_connection_site"] = cf.end_connection_site()
        except Exception:
            pass
        if conn_info:
            info["connector_format"] = conn_info
    except Exception:
        pass

    # Adjustment handles, every value in one Apple Event.
    try:
        values = elements(shape.adjustments.adjustment_value)
        if values:
            adj_dict = {
                i: round(value, 4)
                for i, value in enumerate(values, start=1)
                if not is_missing(value)
            }
            info["adjustments"] = adj_dict
            info["adjustments_count"] = len(adj_dict)
            # Include semantic labels when available
            try:
                from ppt_com.shapes import ADJUSTMENT_LABELS  # local: cycle

                labels = ADJUSTMENT_LABELS.get(
                    _win_constant(_WIN_AUTO_SHAPE_TYPE, shape.auto_shape_type())
                )
                if labels:
                    info["adjustment_labels"] = labels
            except Exception:
                pass
    except Exception:
        pass

    return info


def _update_shape_impl(slide_index, shape_name, shape_index, left, top, width, height, rotation, name, adjustments):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, None, shape_name=shape_name, shape_index=shape_index)

    if left is not None:
        shape.left_position.set(left)
    if top is not None:
        shape.top.set(top)
    if width is not None:
        shape.width.set(width)
    if height is not None:
        shape.height.set(height)
    if rotation is not None:
        raw(shape, _ROTATION).set(rotation)
    if name is not None:
        shape.name.set(name)

    # Apply adjustment handle values.
    adj_count = 0
    if adjustments:
        adj_count = count(shape.adjustments)
        if not adj_count:
            raise ValueError(
                f"Shape '{shape.name()}' does not support adjustment handles"
            )
        for idx, value in adjustments.items():
            idx = int(idx)  # ensure int even if str comes through deserialization
            if idx < 1 or idx > adj_count:
                raise ValueError(
                    f"Adjustment index {idx} out of range for shape "
                    f"'{shape.name()}' (has {adj_count} adjustment(s))"
                )
            shape.adjustments[idx].adjustment_value.set(value)

    # Read every number back out of PowerPoint rather than echoing what was
    # asked for, which is also the check that the writes above landed.
    result = {
        "success": True,
        "shape_name": shape.name(),
        "left": round(shape.left_position(), 2),
        "top": round(shape.top(), 2),
        "width": round(shape.width(), 2),
        "height": round(shape.height(), 2),
    }

    # Include current adjustment values in response when adjustments were set.
    if adjustments:
        values = elements(shape.adjustments.adjustment_value)
        result["adjustments"] = {
            i: round(value, 4)
            for i, value in enumerate(values, start=1)
            if not is_missing(value)
        }
        # Include semantic labels when available
        try:
            from ppt_com.shapes import ADJUSTMENT_LABELS  # local: cycle

            labels = ADJUSTMENT_LABELS.get(
                _win_constant(_WIN_AUTO_SHAPE_TYPE, shape.auto_shape_type())
            )
            if labels:
                result["adjustment_labels"] = labels
        except Exception:
            pass

    return result


def _delete_shape_impl(slide_index, shape_name, shape_index):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    index = _shape_index(slide, None, shape_name=shape_name, shape_index=shape_index)
    shape = slide.shapes[index]
    deleted_name = shape.name()
    before = count(slide.shapes)
    # PowerPoint's dictionary declares no Standard Suite commands, yet it
    # answers `make` and `delete` all the same. Since the dictionary is not
    # promising anything here, the count is what says the shape really went.
    shape.delete()
    if count(slide.shapes) != before - 1:
        raise RuntimeError(
            f"PowerPoint reported success but shape '{deleted_name}' is still "
            f"on slide {slide_index}. The deletion did not happen."
        )
    return {"success": True, "deleted": deleted_name}


def _duplicate_shape_impl(slide_index, shape_name, shape_index):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    index = _shape_index(slide, None, shape_name=shape_name, shape_index=shape_index)
    shape = slide.shapes[index]
    left = shape.left_position()
    top = shape.top()

    before = _shape_names(slide)

    # appscript's own `duplicate` sends an event code PowerPoint does not
    # answer and fails inside the bridge with "unpack requires a buffer of 4
    # bytes". The AppleScript spelling works, so it is what runs here.
    pres_index = _presentation_index(app, pres)
    if pres_index is None:
        raise RuntimeError(
            "Could not work out which open presentation to duplicate the shape "
            "in. Set the target with ppt_set_target_presentation and try again."
        )
    osascript(
        'tell application "Microsoft PowerPoint" to duplicate '
        f"shape {index} of slide {slide_index} of presentation {pres_index}"
    )

    after = _shape_names(slide)
    if len(after) != len(before) + 1:
        raise RuntimeError(
            "PowerPoint reported success but the slide still has "
            f"{len(after)} shapes, so nothing was duplicated."
        )

    # Which one is new, rather than an assumption that it landed on top. A
    # duplicate often carries the same name as its original, so the names are
    # matched off one at a time and whatever is left over is the new shape.
    remaining = list(before)
    new_index = None
    for position, shape_name_after in enumerate(after, start=1):
        if shape_name_after in remaining:
            remaining.remove(shape_name_after)
        else:
            new_index = position
            break
    if new_index is None:
        # The slide grew by one shape and yet every name was already there,
        # which the matching above cannot explain. Better to say so than to
        # guess at a shape and move it.
        raise RuntimeError(
            "A shape was duplicated but could not be told apart from the "
            f"originals on slide {slide_index}, so it was left where it landed."
        )

    new_shape = slide.shapes[new_index]
    new_shape.left_position.set(left + 20)
    new_shape.top.set(top + 20)
    return {
        "success": True,
        "new_shape_name": new_shape.name(),
        "new_shape_index": new_shape.z_order_position(),
    }


def _presentation_index(app, pres) -> Optional[int]:
    """Return the 1-based index of a presentation among the open ones.

    AppleScript addresses a presentation by index or by name, and the index is
    the one that cannot be confused by two files sharing a basename, which is
    the case the session target exists to defend against.
    """
    try:
        full_name = pres.full_name()
    except CommandError:
        return None
    for index, candidate in enumerate(elements(app.presentations), start=1):
        try:
            if candidate.full_name() == full_name:
                return index
        except CommandError:
            continue
    return None


def _set_zorder_impl(slide_index, shape_name, shape_index, z_order_cmd):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, None, shape_name=shape_name, shape_index=shape_index)
    shape.z_order(
        z_order_position=to_keyword(MsoZOrderCmd, z_order_cmd, "z order command")
    )
    return {
        "success": True,
        "shape_name": shape.name(),
        "new_z_position": shape.z_order_position(),
    }
