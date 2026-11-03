"""Reading a slide for the layout checks, on Apple Events.

Mirrors the reading half of ``ppt_com/layout_check.py``: the same function
names and signatures, and the same plain dicts coming back, so every check
after the reading is shared.

The write tools run this after every write, so it is built around how few
events a slide can be read in. Asking an object for ``properties`` answers
every property it has in one event, and that is the whole technique:

- ``slide.shapes.properties`` is one event for every shape on the slide, and
  carries geometry, type, visibility, the connector flag, ``rotation`` (under
  the name ``rotation_``, appscript's way round the clash with AppleScript's
  own word) and a placeholder's type.
- A fill, a text frame and a text range are one event each the same way.
- Reaching through every shape at once does not work. ``fill format of every
  shape`` answers ``missing value`` for each of them, so those reads stay one
  per shape.

So a shape costs one event and a shape with text two, or about seven for
the shape a write tool touched, which is read in full. Reading the same things
property by property cost about twenty five a shape.

Two things differ from Windows besides the spelling. ``line format`` has no
``visible`` here, so ``line_visible`` stays None and the shared code decides
from the shape's kind, as ``ppt_get_shape_info`` already declines to guess. And
there is no run element, so the colour, size and weight of every character are
read as columns and folded into the distinct styles the checks need, as
``ppt_get_text`` rebuilds runs.
"""

import logging

from appscript import k

from backend.mac_ae import (
    elements,
    is_missing,
    ppt,
    shapes_of,
    slide_at as _slide,
    windows_constant as _windows_constant,
)
from backend.mac_enums import (
    MsoTextOrientation,
    MsoVerticalAnchor,
    PpParagraphAlignment,
    PpPlaceholderType,
)
from ppt_com.constants import (
    msoFillGradient, msoFillSolid, msoGroup, msoLinkedPicture, msoPicture, msoPlaceholder,
)
from utils.color import rgb_list_to_hex
from utils.navigation import goto_slide

logger = logging.getLogger(__name__)


def _value(value):
    return None if is_missing(value) else value


def _record(ref) -> dict:
    """Every property of one object, in one event, or an empty dict."""
    try:
        record = ref.properties.get()
    except Exception:
        logger.debug("properties refused", exc_info=True)
        return {}
    return record if isinstance(record, dict) else {}


def _fill_facts(shape):
    from ppt_mac.shapes import _WIN_FILL_TYPE, _win_constant

    record = _record(shape.fill_format)
    if not record:
        return None
    visible = _value(record.get(k.visible))
    transparency = _value(record.get(k.transparency))
    colour = _value(record.get(k.fore_color))
    kind = _win_constant(_WIN_FILL_TYPE, record.get(k.fill_format_type))
    facts = {
        "visible": None if visible is None else bool(visible),
        "solid": kind == msoFillSolid,
        "color_hex": rgb_list_to_hex(colour) if isinstance(colour, list) else None,
        "transparency": None if transparency is None else round(transparency, 2),
    }
    if kind == msoFillGradient:
        facts["gradient"] = _gradient_facts(shape.fill_format)
    return facts


def _gradient_facts(fill):
    """The macOS half of ppt_com.layout_check._gradient_facts.

    The stops come in one event. This dictionary has no angle for a gradient,
    so it is None here, and the checks judge the text against the whole of it.
    """
    records = [r for r in elements(fill.gradient_stops.properties) if isinstance(r, dict)]
    stops = []
    for record in records:
        colour = _value(record.get(k.color))
        position = _value(record.get(k.position))
        transparency = _value(record.get(k.transparency)) or 0.0
        if not isinstance(colour, list) or position is None:
            return None
        stops.append([position, rgb_list_to_hex(colour), transparency])
    if not stops:
        return None
    # Asked for in percent, answered in percent or as a fraction.
    if max(stop[0] for stop in stops) > 1.0 + 1e-6:
        stops = [[p / 100.0, c, t] for p, c, t in stops]
    if max(stop[2] for stop in stops) > 1.0 + 1e-6:
        stops = [[p, c, t / 100.0] for p, c, t in stops]
    return {"angle": None, "stops": [[round(p, 4), c, round(t, 2)] for p, c, t in stops]}


def _frame_state(frame: dict) -> dict:
    """The text frame settings, in the words ppt_set_textframe takes.

    The same answer as ppt_mac.shapes._text_frame_state, read out of one
    record rather than asked for one property at a time.
    """
    from ppt_com.text import AUTO_SIZE_NAMES, ORIENTATION_NAMES, VERTICAL_ANCHOR_NAMES
    from ppt_mac.text import _AUTO_SIZE

    auto_size = _windows_constant(_AUTO_SIZE, frame.get(k.auto_size))
    wrap = _value(frame.get(k.word_wrap))
    margins = {
        side: _value(frame.get(key))
        for side, key in (("left", k.margin_left), ("right", k.margin_right),
                          ("top", k.margin_top), ("bottom", k.margin_bottom))
    }
    return {
        "autofit": AUTO_SIZE_NAMES.get(auto_size),
        "word_wrap": None if wrap is None else bool(wrap),
        "vertical_anchor": VERTICAL_ANCHOR_NAMES.get(
            _windows_constant(MsoVerticalAnchor, frame.get(k.vertical_anchor))),
        "orientation": ORIENTATION_NAMES.get(
            _windows_constant(MsoTextOrientation, frame.get(k.orientation))),
        "margins": None if any(v is None for v in margins.values()) else {
            side: round(value, 2) for side, value in margins.items()
        },
    }


def _theme_flag(index):
    """Whether a colour is one of the theme's, from its theme index.

    A theme colour is the template's choice; text the caller coloured itself
    reads ``no theme color``.
    """
    index = _value(index)
    if index is None or isinstance(index, (list, str, int, float)):
        return None
    return index not in (k.no_theme_color, k.theme_color_unset)


def _run_styles(text_range) -> list:
    """The distinct colour, size and weight the visible characters use.

    Most text is set in one style, and the range's own font record answers
    that in one event: a property the characters disagree on reads back as
    ``missing value``, so a record with all three is the whole answer. Mixed
    text is read as columns, six events however long it is, and when
    PowerPoint will not answer per character either, nothing is guessed.
    """
    font = _record(text_range.font)
    colour = _value(font.get(k.font_color))
    size = _value(font.get(k.font_size))
    bold = _value(font.get(k.bold))
    name = _value(font.get(k.font_name))
    if isinstance(colour, list) and isinstance(size, (int, float)) and bold is not None:
        return [{"color_hex": rgb_list_to_hex(colour), "size": round(size, 1),
                 "bold": bold is True, "font": name if isinstance(name, str) else None,
                 "theme": _theme_flag(font.get(k.font_color_theme_index))}]

    try:
        chars = text_range.characters
        texts = chars.content.get()
        font = chars.font
        columns = [font.font_color.get(), font.font_size.get(), font.bold.get(),
                   font.font_color_theme_index.get(), font.font_name.get()]
        if not isinstance(texts, list) or any(
            not isinstance(column, list) or len(column) != len(texts)
            for column in columns
        ):
            raise ValueError("not one value per character")
        rows = zip(texts, *columns, strict=True)
    except Exception:
        logger.debug("Bulk character read refused", exc_info=True)
        return []

    styles = set()
    for char, colour, size, bold, index, name in rows:
        if not isinstance(char, str) or not char.strip():
            continue
        colour, size, name = _value(colour), _value(size), _value(name)
        styles.add((
            rgb_list_to_hex(colour) if isinstance(colour, list) else None,
            round(size, 1) if isinstance(size, (int, float)) else None,
            _value(bold) is True,
            _theme_flag(index),
            name if isinstance(name, str) else None,
        ))
    return [
        {"color_hex": colour, "size": size, "bold": bold, "font": name, "theme": theme}
        for colour, size, bold, theme, name in sorted(
            styles, key=lambda s: (s[0] or "", s[1] or 0, s[4] or ""))
    ]


def _text_facts(shape, full):
    """The macOS half of ppt_com.layout_check._text_facts.

    The text range's record carries the content and where it is drawn, so a
    shape that is read only for where its text sits costs one event. In full,
    the frame's record and the character columns are added.
    """
    from ppt_com.layout_check import _CONTENT_KEPT, has_right_to_left

    frame = shape.text_frame
    facts = {
        "chars": 0, "content": "",
        "left": None, "top": None, "width": None, "height": None,
        "alignment": None, "runs": [],
        "margins": None, "word_wrap": None, "autofit": None,
        "orientation": None, "vertical_anchor": None,
    }
    tr = frame.text_range
    words = _record(tr)
    content = _value(words.get(k.content)) or ""
    facts["chars"] = len(content.strip()) if isinstance(content, str) else 0
    facts["content"] = content[:_CONTENT_KEPT] if isinstance(content, str) else ""
    if full:
        facts.update(_frame_state(_record(frame)))
    if not facts["chars"]:
        return facts

    for key, name in (("left", k.left_bounds), ("top", k.top_bounds),
                      ("width", k.bounds_width), ("height", k.bounds_height)):
        value = _value(words.get(name))
        if isinstance(value, (int, float)):
            facts[key] = round(value, 2)
    # Alignment only places text that PowerPoint would not place itself.
    if facts["left"] is None or facts["top"] is None:
        try:
            facts["alignment"] = _windows_constant(
                PpParagraphAlignment, tr.paragraphs[1].paragraph_format.alignment()
            )
        except Exception:
            pass
    if full:
        facts["runs"] = _run_styles(tr)
        caps = _value(_record(tr.font).get(k.caps_type))
        facts["all_caps"] = {k.all_caps: True, k.no_caps: False}.get(caps)
        if has_right_to_left(content):
            facts["directions"] = _directions(tr, content)
    return facts


def _directions(tr, content):
    """The macOS half of ppt_com.layout_check._directions, in one event.

    A paragraph set right to left is known. One that is not may still be
    drawn right to left, since PowerPoint draws text tagged as Hebrew or
    Arabic from the right, and the dictionary has no way to read a text's
    language: it is None, unknown, rather than a guess. A text of one
    paragraph answers its one value rather than a list.
    """
    from ppt_com.layout_check import _paragraphs

    names = {k.right_to_left: "rtl"}
    try:
        values = tr.paragraphs.paragraph_format.text_direction.get()
    except Exception:
        logger.debug("Cannot read the paragraphs' direction", exc_info=True)
        return None
    if not isinstance(values, list):
        if len(_paragraphs(content)) != 1:
            return None
        values = [values]
    return [names.get(value) for value in values]


def _geometry_facts(shape, record, z):
    """Where a shape is and what kind it is, out of its bulk record."""
    from ppt_mac.shapes import _WIN_SHAPE_TYPE, _from_record, _win_constant

    type_int = _win_constant(
        _WIN_SHAPE_TYPE, _from_record(record, k.shape_type, shape, "shape_type")
    )
    rotation = _value(record.get(k.rotation_))
    facts = {
        "name": _from_record(record, k.name, shape, "name"),
        "z": z,
        "type": type_int,
        "left": round(_from_record(record, k.left_position, shape, "left_position") or 0.0, 2),
        "top": round(_from_record(record, k.top, shape, "top") or 0.0, 2),
        "width": round(_from_record(record, k.width, shape, "width") or 0.0, 2),
        "height": round(_from_record(record, k.height, shape, "height") or 0.0, 2),
        "rotation": round(rotation, 2) if isinstance(rotation, (int, float)) else None,
        "visible": True,
        "connector": bool(_value(record.get(k.is_connector))),
        "placeholder_type": None,
        "fill": None,
        "line_visible": None,
        "text": None,
    }
    visible = _from_record(record, k.visible, shape, "visible")
    if visible is not None:
        facts["visible"] = bool(visible)
    if type_int == msoPlaceholder:
        facts["placeholder_type"] = _windows_constant(
            PpPlaceholderType, record.get(k.placeholder_type))
    if type_int in (msoPicture, msoLinkedPicture):
        locked = _value(record.get(k.lock_aspect_ratio))
        if locked is not None:
            facts["lock_aspect"] = bool(locked)
    # Kept for _read_appearance, which has no record; dropped before return.
    facts["_has_text_frame"] = bool(
        _from_record(record, k.has_text_frame, shape, "has_text_frame"))
    return facts


def _read_appearance(facts, shape, full):
    """Add what the shape draws: its fill, and its text."""
    if facts["type"] != msoGroup:
        try:
            facts["fill"] = _fill_facts(shape)
        except Exception:
            logger.debug("Cannot read the fill of '%s'", facts["name"], exc_info=True)
    if facts.get("_has_text_frame"):
        try:
            facts["text"] = _text_facts(shape, full)
        except Exception:
            logger.debug("Cannot read the text of '%s'", facts["name"], exc_info=True)
    return facts


def _background(slide, follows_master):
    """The macOS half of ppt_com.layout_check._background.

    A background is a shape here, so its colour is that shape's fill.
    """
    from ppt_mac.shapes import _WIN_FILL_TYPE, _win_constant

    try:
        source = slide
        if follows_master:
            source = slide.custom_layout
            if source.follow_master_background():
                source = slide.slide_master
        fill = _record(source.background.fill_format)
        kind = _win_constant(_WIN_FILL_TYPE, fill.get(k.fill_format_type))
        if kind == msoFillSolid:
            colour = _value(fill.get(k.fore_color))
            return (rgb_list_to_hex(colour) if isinstance(colour, list) else None), None
        if kind == msoFillGradient:
            return None, _gradient_facts(source.background.fill_format)
    except Exception:
        logger.debug("Cannot read the slide background", exc_info=True)
    return None, None


def _theme_colours(slide):
    """The macOS half of ppt_com.layout_check._theme_colours.

    The colours come in one event, and the reader asks once per master.
    """
    names = ("dark1", "light1", "dark2", "light2")
    try:
        scheme = slide.slide_master.theme.theme_color_scheme
        values = scheme.theme_colors.RGB.get()
        if not isinstance(values, list) or len(values) < 4:
            values = [scheme.theme_colors[i].RGB() for i in range(1, 5)]
        return {name: rgb_list_to_hex(value) for name, value in zip(names, values)}
    except Exception:
        logger.debug("Cannot read the theme's colours", exc_info=True)
        return None


def _template_facts(slide, shows, cache, touched=None):
    """The macOS half of ppt_com.layout_check._template_facts.

    Only the master's graphics can be read here: PowerPoint for Mac answers no
    shapes for a custom layout. A layout that draws a design of its own, a
    title or section layout usually, hides the master's graphics and is not
    seen, which the checks allow for by trusting the template's own colours on
    its placeholders.

    Most slides of a deck share one master, so a graphic's fill and text are
    read once per call: ``cache`` holds them by name and position. A write
    reads them only for the graphics within reach of the shapes it touched,
    ``touched``, as on Windows.
    """
    from ppt_com.layout_check import within_reach
    from ppt_com.layout_check import _CONTENT_PLACEHOLDERS, _TEMPLATE_Z

    if shows is False:
        return [], []
    try:
        if not slide.custom_layout.display_master_shapes():
            return [], []
        master = slide.slide_master
        refs = shapes_of(master)
        records = elements(master.shapes.properties)
    except Exception:
        logger.debug("Cannot read the master's graphics", exc_info=True)
        return [], []
    if len(records) != len(refs):
        records = [{} for _ in refs]
    graphics, area = [], []
    for ref, record in zip(refs, records, strict=True):
        record = record if isinstance(record, dict) else {}
        facts = _geometry_facts(ref, record, _TEMPLATE_Z + len(graphics))
        if facts["type"] == msoPlaceholder:
            if facts["placeholder_type"] in _CONTENT_PLACEHOLDERS:
                area.append([facts["left"], facts["top"], round(facts["left"] + facts["width"], 2),
                             round(facts["top"] + facts["height"], 2)])
            continue
        key = (facts["name"], facts["left"], facts["top"], facts["width"], facts["height"])
        if key not in cache:
            if touched is None or within_reach(facts, touched):
                _read_appearance(facts, ref, False)
            facts.pop("_has_text_frame", None)
            facts["source"] = "master"
            cache[key] = facts
        graphics.append({**cache[key], "z": _TEMPLATE_Z + len(graphics)})
    return graphics, area


def _layout_name(word):
    """The slide's layout, in the words of the dictionary.

    A custom layout has no name on macOS, so this is the built-in layout the
    slide reports, "title only" rather than "slide layout title only".
    """
    word = _value(word)
    if word is None:
        return None
    name = str(word).replace("k.", "").replace("_", " ")
    return name[len("slide layout "):] if name.startswith("slide layout ") else name


def _slide_facts_impl(slide_indices, navigate, detail, template_cache=None):
    from ppt_com.layout_check import design_key, needs_background, read_shapes
    from ppt_mac.layout import _slide_dimensions
    from ppt_mac.slides import _slide_count

    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide_w, slide_h = _slide_dimensions(pres)
    facts = {"slide_width": slide_w, "slide_height": slide_h, "slides": []}
    # Carried between calls only for a whole read, never a write's partial one.
    if detail is not None or template_cache is None:
        template_cache = {}
    themes = template_cache.setdefault("theme colours", {})
    total = _slide_count(pres)
    for slide_index in slide_indices:
        slide = _slide(pres, slide_index)
        if navigate:
            goto_slide(app, slide_index)
        refs = shapes_of(slide)
        # Every property of every shape in one event. A short answer means
        # PowerPoint did not give one record per shape, and each shape is
        # then asked directly instead.
        try:
            records = elements(slide.shapes.properties)
        except Exception:
            records = []
        if len(records) != len(refs):
            records = [{} for _ in refs]
        own = _record(slide)
        follows = _value(own.get(k.follow_master_background))
        shows = _value(own.get(k.display_master_shapes))
        records = [r if isinstance(r, dict) else {} for r in records]
        shapes = read_shapes(
            refs,
            lambda ref, z, records=records: _geometry_facts(ref, records[z - 1], z),
            _read_appearance,
            detail,
        )
        for shape in shapes:
            shape.pop("_has_text_frame", None)
        template, area = _template_facts(
            slide, None if shows is None else bool(shows), template_cache,
            None if detail is None else [s for s in shapes if s["name"] in detail])
        drawn = _saved_layout(pres, slide_index, template_cache) if detail is None else None
        inner = _saved_groups(slide_index, template_cache) if detail is None else []
        if drawn and shows is not False:
            from ppt_com.layout_check import _TEMPLATE_Z

            for graphic in drawn["graphics"]:
                template.append({**graphic, "z": _TEMPLATE_Z + len(template),
                                 "source": "layout"})
        background_hex, background_gradient = (
            _background(slide, True if follows is None else bool(follows))
            if needs_background(shapes, detail) else (None, None))
        design = design_key(area)
        if design not in themes:
            themes[design] = _theme_colours(slide)
        try:
            hidden = bool(slide.slide_show_transition.hidden())
        except Exception:
            hidden = False
        facts["slides"].append({
            "slide_index": slide_index,
            "hidden": hidden,
            "animations": _animations(slide),
            # A Morph moves shapes onto this slide, or off it onto the next.
            "morph": _morph(slide) or (
                slide_index < total and _morph(_slide(pres, slide_index + 1))),
            "layout_title": (drawn or {}).get("title"),
            # From the saved file, like the right-to-left fonts it is held to.
            "theme_fonts": (drawn or {}).get("theme_fonts"),
            "inner": inner,
            "design": design,
            "layout": _layout_name(own.get(k.layout)),
            "background_hex": background_hex,
            "background_gradient": background_gradient,
            "theme_colors": themes[design],
            "follows_background": None if follows is None else bool(follows),
            "template": template,
            "template_area": area,
            "shapes": shapes,
        })
    return facts


def _animations(slide):
    """The slide's main sequence, as ``on_screen`` takes it.

    PowerPoint for Mac does not say what starts an effect, so the trigger is
    None and each effect counts as its own click; see ``on_screen``. None of
    it rather than part, as on Windows.
    """
    from backend.mac_enums import MsoAnimEffect
    from ppt_com.animation import _get_animation_category
    from ppt_mac.animation import _effect_count, _effect_shape_name, _main_sequence, _read

    effects = []
    try:
        sequence = _main_sequence(slide)
        for i in range(1, _effect_count(sequence) + 1):
            effect = sequence.effects[i]
            name = _effect_shape_name(effect)
            word = _read(effect.animation_effect_type)
            if name is None or word is None:
                return []
            effects.append({
                "shape": name,
                "category": _get_animation_category(
                    _windows_constant(MsoAnimEffect, word), bool(_read(effect.exit_animation))),
                "trigger": None,
            })
    except Exception:
        logger.debug("Cannot read the animations of a slide", exc_info=True)
        return []
    return effects


def _morph(slide) -> bool:
    """Whether the slide is entered with a Morph transition.

    PowerPoint for Mac answers ``entry_effect_morph_by_object`` (or by word,
    or by character), which its PpEntryEffect table here does not list.
    """
    try:
        word = slide.slide_show_transition.entry_effect()
    except Exception:
        return False
    return getattr(word, "name", "").startswith("entry_effect_morph")


def _package_facts_impl(slide_indices, charts=False):
    """The macOS half of ppt_com.layout_check._package_facts_impl.

    PowerPoint for Mac has no way to save a copy aside: saving to another
    path moves the open deck to it. So the deck's own file is read, and only
    while it holds what is on screen. Otherwise the pictures go unjudged, and
    with ``charts`` each chart is read the way ppt_get_chart_data reads it,
    through the clipboard, which is put back afterwards.
    """
    from ppt_com.layout_check import package_facts
    from ppt_mac.presentation import _local_path

    pres = ppt._get_pres_impl()
    try:
        path = _local_path(pres.full_name()) if pres.saved() else None
    except Exception:
        logger.debug("Cannot tell whether the deck is saved", exc_info=True)
        path = None
    if path and path.lower().endswith((".pptx", ".pptm", ".potx", ".ppsx")):
        return package_facts(path, slide_indices)
    if not charts:
        return {}
    return {"charts": _charts_by_clipboard(pres, slide_indices)}


def _charts_by_clipboard(pres, slide_indices):
    from ppt_com.constants import msoChart
    from ppt_mac.charts import _get_chart_data_impl
    from ppt_mac.shapes import _WIN_SHAPE_TYPE, _win_constant

    found = {}
    for index in slide_indices:
        data = {}
        for shape in shapes_of(_slide(pres, index)):
            try:
                if _win_constant(_WIN_SHAPE_TYPE, shape.shape_type()) != msoChart:
                    continue
                read = _get_chart_data_impl(index, shape.name())
            except Exception:
                logger.debug("Cannot read a chart's data", exc_info=True)
                continue
            if read.get("success"):
                data[read["shape_name"]] = {"categories": read["categories"],
                                            "series": read["series"]}
        found[index] = data
    return found


_FONTS = {}


def _installed_fonts_impl():
    """The macOS half of ppt_com.layout_check._installed_fonts_impl.

    The system's fonts, the Office fonts PowerPoint for Mac carries inside
    its own bundle, where the system's list does not see them, and the
    Office cloud fonts already downloaded.
    """
    if "installed" in _FONTS:
        return _FONTS["installed"]
    import os

    from AppKit import NSFontManager

    from ppt_com.layout_check import FontNames
    from utils.fonts import names_in_folder

    names = set()
    manager = NSFontManager.sharedFontManager()
    for family in manager.availableFontFamilies() or []:
        names.add(str(family))
        for member in manager.availableMembersOfFontFamily_(family) or []:
            names.add(f"{family} {member[1]}")
    bundle = "/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts"
    names |= names_in_folder(bundle)
    cloud = os.path.expanduser(
        "~/Library/Group Containers/UBF8T346G9.Office/FontCache/4/CloudFonts")
    if os.path.isdir(cloud):
        for family in os.listdir(cloud):
            names.add(family)
            names |= names_in_folder(os.path.join(cloud, family))
    # The system's list names a font once; the files hold the names a deck
    # made elsewhere asks for it by: "ヒラギノ角ゴシック W3" and "Hiragino Kaku
    # Gothic ProN W3" for what it lists as "Hiragino Sans".
    _FONTS["installed"] = FontNames(
        names, ("/System/Library/Fonts", "/Library/Fonts", os.path.expanduser("~/Library/Fonts")))
    return _FONTS["installed"]



def _saved_layout(pres, slide_index, cache):
    """What the slide's layout draws, from the deck's own file.

    PowerPoint for Mac answers no shapes for a custom layout, so a layout's
    own graphics, a band or a rule a title or section layout draws, and
    where it puts its title, are read from the saved file, while it holds
    what is on screen. A deck never saved, or changed since, answers None.
    """
    from ppt_mac.presentation import _local_path
    from utils.package import layout_facts

    if "saved file" not in cache:
        path = None
        try:
            if pres.saved():
                path = _local_path(pres.full_name())
        except Exception:
            logger.debug("Cannot tell whether the deck is saved", exc_info=True)
        if path and not path.lower().endswith((".pptx", ".pptm", ".potx", ".ppsx")):
            path = None
        cache["saved file"] = path
    path = cache["saved file"]
    if not path:
        return None
    try:
        return layout_facts(path, [slide_index]).get(slide_index)
    except Exception:
        logger.debug("Cannot read the layouts from %s", path, exc_info=True)
        return None


def _saved_groups(slide_index, cache):
    """The text inside the slide's groups, from the deck's own file.

    A group answers no members over Apple Events, so what is inside it is
    read from the saved file, while it holds what is on screen: its words,
    fonts and paragraph directions, enough for the checks on those. Where
    PowerPoint draws them is not in the file, so whether they fit is not
    known here. ``_saved_layout`` has already decided whether the file can
    be used.
    """
    from utils.package import group_texts

    path = cache.get("saved file")
    if not path:
        return []
    try:
        return group_texts(path, [slide_index]).get(slide_index, [])
    except Exception:
        logger.debug("Cannot read the groups from %s", path, exc_info=True)
        return []


def _photo_samples_impl(slide_index, names):
    """The macOS half of ppt_com.layout_check._photo_samples_impl.

    ``save as picture`` writes a picture's own image file, whatever format it
    is asked for, and only inside PowerPoint's container. Quartz reads any of
    them, turned the way PowerPoint shows it. Being the whole image, it lines
    up with what is drawn only when the picture is not cropped, so a cropped
    one is left unjudged.
    """
    import os

    from ppt_mac.export import _staging_path

    slide = _slide(ppt._get_pres_impl(), slide_index)
    by_name = {}
    for ref in shapes_of(slide):
        try:
            by_name[ref.name()] = ref
        except Exception:
            continue
    found = {}
    for name in names:
        shape = by_name.get(name)
        if shape is None:
            continue
        crop = _record(shape.picture_format)
        if any(abs(_value(crop.get(side)) or 0) > 0.5
               for side in (k.crop_left, k.crop_right, k.crop_top, k.crop_bottom)):
            continue
        staged = _staging_path(".img")
        try:
            shape.save_as_picture(picture_type=k.save_as_PNG_file, file_name=staged)
            with open(staged, "rb") as image:
                sampled = _sample_image(image.read(), shape.width() / max(shape.height(), 1))
            if sampled:
                # A flipped or half-turned picture comes out as the slide
                # shows it (checked live), so the sample needs no turning.
                found[name] = sampled
        except Exception:
            logger.debug("Cannot export '%s'", name, exc_info=True)
        finally:
            if os.path.exists(staged):
                os.remove(staged)
    return found


def _sample_image(data, aspect, columns=96):
    """A grid of an image file's colours, as utils.bitmap.sample_bmp gives it,
    drawn at the proportions the picture is shown at."""
    import Quartz
    from Foundation import NSData

    source = Quartz.CGImageSourceCreateWithData(NSData.dataWithBytes_length_(data, len(data)), None)
    if source is None:
        return None
    image = Quartz.CGImageSourceCreateThumbnailAtIndex(source, 0, {
        Quartz.kCGImageSourceCreateThumbnailFromImageAlways: True,
        Quartz.kCGImageSourceCreateThumbnailWithTransform: True,
        Quartz.kCGImageSourceThumbnailMaxPixelSize: 1024,
    })
    if image is None:
        return None
    rows = max(1, round(columns / max(aspect, 0.01)))
    space = Quartz.CGColorSpaceCreateWithName(Quartz.kCGColorSpaceSRGB)
    context = Quartz.CGBitmapContextCreate(None, columns, rows, 8, columns * 4, space,
                                           Quartz.kCGImageAlphaPremultipliedLast)
    Quartz.CGContextSetInterpolationQuality(context, Quartz.kCGInterpolationNone)
    Quartz.CGContextDrawImage(context, Quartz.CGRectMake(0, 0, columns, rows), image)
    pixels = Quartz.CGBitmapContextGetData(context)
    raw = bytes(pixels.as_buffer(columns * rows * 4))
    return {"columns": columns, "rows": rows, "colours": _colours(raw)}


def _colours(raw):
    """Premultiplied RGBA bytes as colours, None where the image is see-through:
    what shows there is whatever is under the picture, not black."""
    colours = []
    for i in range(0, len(raw), 4):
        alpha = raw[i + 3]
        if alpha < 128:
            colours.append(None)
            continue
        colours.append("#" + "".join(f"{min(255, round(raw[i + c] * 255 / alpha)):02X}"
                                     for c in range(3)))
    return colours

