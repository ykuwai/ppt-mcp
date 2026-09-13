"""Placeholder, design and layout tools, on Apple Events.

Mirrors ``ppt_com/placeholders.py``. Same function names, same signatures,
same returned shapes; what differs is the walk through PowerPoint's object
model.

This is the area of the port that carries over most cleanly. ``place holder``
is two words on macOS and ``placeholder type`` reads back correctly, so
resolving a title or a body placeholder works exactly as it does on Windows.

One thing genuinely does not carry over. ``design``, ``master`` and
``custom layout`` declare no ``name`` property in PowerPoint's dictionary, so
layouts cannot be addressed by name here. Each listing says so in the payload
and every entry keeps its index, which is the handle that does work.
"""

import logging

from appscript.reference import CommandError

from backend.mac_ae import (
    count,
    elements,
    is_missing,
    positional,
    ppt,
    raw,
    shapes_of,
    windows_constant as _windows_constant,
)
from backend.mac_enums import (
    MsoShapeType,
    PpParagraphAlignment,
    PpPlaceholderType,
    to_keyword,
)
from utils.color import rgb_list_to_hex
from utils.navigation import goto_slide
from ppt_com.constants import (
    PLACEHOLDER_TYPE_NAMES,
    msoPlaceholder,
    ppPlaceholderTitle, ppPlaceholderBody, ppPlaceholderCenterTitle,
    ppPlaceholderSubtitle,
)

logger = logging.getLogger(__name__)

# Keys are MsoShapeType values, which is what `contained type` reports once it
# has been mapped back from its macOS enumerator. Same table as the COM module.
_CONTAINED_TYPE_NAMES = {
    1: "AutoShape",
    3: "Chart",
    11: "LinkedPicture",
    13: "Picture",
    14: "Placeholder",
    16: "Media",
    19: "Table",
    24: "SmartArt",
}

# What every listing says once, rather than leaving a row of nulls that reads
# as "these layouts happen to be unnamed".
_NO_NAMES_NOTE = (
    "PowerPoint for Mac exposes no name on designs, masters or custom layouts "
    "to Apple Events, so names may come back null. Address them by index."
)

# The property code behind every `name` in PowerPoint's dictionary. The classes
# above do not declare one, but the code is worth a try because Office objects
# often answer it anyway, and a failed try costs one event and reports null.
_NAME_CODE = b'pnam'


def _name_or_none(ref):
    """Read an object's name, or None when its class does not carry one."""
    try:
        value = raw(ref, _NAME_CODE).get()
    except Exception:
        return None
    return None if is_missing(value) else value


def _placeholders(container):
    """Return the placeholder shapes of a slide, master or custom layout.

    ``place holder`` is an element of ``shape`` in the dictionary rather than
    of ``slide``, yet asking a slide for its placeholders resolves anyway. That
    was measured on a slide and not on a custom layout, so a layout that
    refuses falls back to filtering its shapes by type, which reaches the same
    objects the long way round.
    """
    found = positional(container.place_holders)
    if found:
        return found

    placeholder_type = to_keyword(MsoShapeType, msoPlaceholder, "shape type")
    filtered = []
    for shape in shapes_of(container):
        try:
            if shape.shape_type() == placeholder_type:
                filtered.append(shape)
        except CommandError:
            continue
    return filtered


def _placeholder_type_value(placeholder):
    """Return the Windows PpPlaceholderType int for a placeholder, or None."""
    try:
        return _windows_constant(PpPlaceholderType, placeholder.placeholder_type())
    except CommandError:
        return None


def _find_placeholder_by_type(slide, placeholder_type: int):
    """Find the first placeholder of a given type on a slide.

    Returns a placeholder reference, or None if not found.
    """
    for ph in _placeholders(slide):
        if _placeholder_type_value(ph) == placeholder_type:
            return ph
    return None


def _resolve_placeholder(slide, placeholder_index=None, placeholder_type=None):
    """Resolve a placeholder by index (int) or type name (str) or type int."""
    # Imported lazily. ppt_com/placeholders.py imports this module at the bottom
    # of its own file, so importing it back at module scope would let an
    # "import ppt_mac.placeholders first" ordering run that swap block against
    # a module that has defined nothing yet, and the swap would silently not
    # happen. By call time both modules are fully loaded.
    from ppt_com.placeholders import PLACEHOLDER_TYPE_MAP

    if placeholder_index is not None:
        phs = _placeholders(slide)
        if placeholder_index < 1 or placeholder_index > len(phs):
            raise ValueError(
                f"Placeholder index {placeholder_index} out of range "
                f"(1-{len(phs)})"
            )
        return phs[placeholder_index - 1]
    elif placeholder_type is not None:
        # Convert string type to int if needed
        if isinstance(placeholder_type, str):
            type_int = PLACEHOLDER_TYPE_MAP.get(placeholder_type.lower())
            if type_int is None:
                raise ValueError(
                    f"Unknown placeholder type '{placeholder_type}'. "
                    f"Valid types: {list(PLACEHOLDER_TYPE_MAP.keys())}"
                )
        else:
            type_int = placeholder_type

        ph = _find_placeholder_by_type(slide, type_int)

        # Fallback: Title -> CenterTitle, Subtitle -> Body
        if ph is None:
            if type_int == ppPlaceholderTitle:
                ph = _find_placeholder_by_type(slide, ppPlaceholderCenterTitle)
            elif type_int == ppPlaceholderSubtitle:
                ph = _find_placeholder_by_type(slide, ppPlaceholderBody)

        if ph is None:
            type_name = PLACEHOLDER_TYPE_NAMES.get(type_int, str(type_int))
            raise ValueError(
                f"No placeholder of type '{type_name}' found on slide"
            )
        return ph
    else:
        raise ValueError("Must specify either placeholder_index or placeholder_type")


def _text_or_none(text_frame):
    """Read a text frame's content, treating `missing value` as no text."""
    content = text_frame.text_range.content()
    return None if is_missing(content) else content


# ---------------------------------------------------------------------------
# Apple Event implementation functions
# ---------------------------------------------------------------------------
def _list_placeholders_impl(slide_index: int) -> dict:
    ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]
    phs = _placeholders(slide)

    placeholders = []
    for i, ph in enumerate(phs, start=1):
        ph_type = _placeholder_type_value(ph)

        info = {
            "index": i,
            "name": ph.name(),
            "type": ph_type,
            "type_name": PLACEHOLDER_TYPE_NAMES.get(ph_type, f"Unknown({ph_type})"),
            "has_text": False,
            "text_preview": None,
            "left": round(ph.left_position(), 1),
            "top": round(ph.top(), 1),
            "width": round(ph.width(), 1),
            "height": round(ph.height(), 1),
        }

        if ph.has_text_frame():
            tf = ph.text_frame
            has_text = bool(tf.has_text())
            info["has_text"] = has_text
            if has_text:
                text = _text_or_none(tf) or ""
                info["text_preview"] = text[:100] + ("..." if len(text) > 100 else "")

        placeholders.append(info)

    return {
        "status": "success",
        "slide_index": slide_index,
        "placeholder_count": len(phs),
        "placeholders": placeholders,
    }


def _get_placeholder_impl(slide_index, placeholder_index, placeholder_type) -> dict:
    ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]

    ph = _resolve_placeholder(slide, placeholder_index, placeholder_type)
    ph_type = _placeholder_type_value(ph)

    contained = None
    try:
        contained = _windows_constant(
            MsoShapeType, ph.place_holder_format.contained_type()
        )
    except CommandError:
        pass

    info = {
        # Windows reports the type here too rather than the ordinal. Kept as
        # it is so both platforms answer the same shape.
        "index": ph_type,
        "name": ph.name(),
        "type": ph_type,
        "type_name": PLACEHOLDER_TYPE_NAMES.get(ph_type, f"Unknown({ph_type})"),
        "contained_type": contained,
        "contained_type_name": _CONTAINED_TYPE_NAMES.get(
            contained, f"Unknown({contained})"
        ),
        "left": round(ph.left_position(), 1),
        "top": round(ph.top(), 1),
        "width": round(ph.width(), 1),
        "height": round(ph.height(), 1),
        # `rotation` is both a property and an enumerator in the dictionary and
        # appscript resolves the collision toward the enumerator, so the
        # property is only reachable by its raw code.
        "rotation": None,
        "has_text_frame": bool(ph.has_text_frame()),
    }

    try:
        rotation = raw(ph, b'ShRt').get()
        info["rotation"] = None if is_missing(rotation) else rotation
    except CommandError:
        pass

    if info["has_text_frame"]:
        tf = ph.text_frame
        tr = tf.text_range
        has_text = bool(tf.has_text())
        info["text"] = _text_or_none(tf) if has_text else None
        info["paragraph_count"] = count(tr.paragraphs)
        if has_text:
            # Font formatting has to be reached through `font of text range`.
            # Any other route answers -10006.
            font = tr.font
            info["font_name"] = font.font_name()
            info["font_name_far_east"] = font.east_asian_name()
            try:
                size = font.font_size()
                info["font_size"] = None if is_missing(size) else size
            except Exception:
                info["font_size"] = None
            info["alignment"] = _windows_constant(
                PpParagraphAlignment, tr.paragraph_format.alignment()
            )

    return {
        "status": "success",
        "placeholder": info,
    }


def _set_placeholder_text_impl(slide_index, placeholder_index, placeholder_type, text) -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]

    ph = _resolve_placeholder(slide, placeholder_index, placeholder_type)

    if not ph.has_text_frame():
        raise ValueError(
            f"Placeholder '{ph.name()}' does not have a text frame "
            f"(type={_placeholder_type_value(ph)})"
        )

    text = text.replace('\n', '\r')  # CR is the paragraph break here too
    tr = ph.text_frame.text_range
    tr.content.set(text)

    # Nothing is trusted because it did not raise. Read the content back and
    # make sure the write landed.
    written = tr.content()
    if is_missing(written) or (text and not written):
        raise RuntimeError(
            f"PowerPoint reported no error but placeholder '{ph.name()}' is "
            "still empty. The text was not written."
        )

    return {
        "status": "success",
        "slide_index": slide_index,
        "placeholder_name": ph.name(),
        "placeholder_type": _placeholder_type_value(ph),
        "text_length": tr.text_length(),
    }


def _list_layouts_impl(design_index) -> dict:
    ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    design = pres.designs[design_index]
    master = design.slide_master
    layouts_col = elements(master.custom_layouts)

    layouts = []
    for i, layout in enumerate(layouts_col, start=1):
        phs = _placeholders(layout)

        ph_list = []
        for ph in phs:
            ph_type = _placeholder_type_value(ph)
            ph_list.append({
                "type": ph_type,
                "type_name": PLACEHOLDER_TYPE_NAMES.get(
                    ph_type, f"Unknown({ph_type})"
                ),
                "name": ph.name(),
            })

        layouts.append({
            "index": i,
            "name": _name_or_none(layout),
            "placeholder_count": len(phs),
            "placeholders": ph_list,
        })

    return {
        "status": "success",
        "design_name": _name_or_none(design),
        "layout_count": len(layouts_col),
        "layouts": layouts,
        "note": _NO_NAMES_NOTE,
    }


def _list_designs_impl(include_layouts: bool = False) -> dict:
    ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    designs = elements(pres.designs)
    result = []
    for d, design in enumerate(designs, start=1):
        master = design.slide_master
        layouts = elements(master.custom_layouts)
        entry = {
            "design_index": d,
            "design_name": _name_or_none(design),
            "layout_count": len(layouts),
        }
        if include_layouts:
            entry["layout_names"] = [_name_or_none(layout) for layout in layouts]
        result.append(entry)
    return {
        "status": "success",
        "design_count": len(designs),
        "designs": result,
        "note": _NO_NAMES_NOTE,
    }


def _get_slide_master_info_impl(design_index) -> dict:
    ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    design = pres.designs[design_index]
    master = design.slide_master

    # Theme colours. MACOS_PORT records `theme color scheme` reading back as
    # `missing value` on a slide master, so this is an attempt rather than a
    # promise. Whatever actually comes back is reported, and nothing is
    # invented when it does not.
    colors = []
    try:
        theme_color_names = [
            "Dark1", "Light1", "Dark2", "Light2",
            "Accent1", "Accent2", "Accent3", "Accent4",
            "Accent5", "Accent6", "Hyperlink", "FollowedHyperlink",
        ]
        for i, theme_color in enumerate(
            elements(master.theme.theme_color_scheme.theme_colors)[:12], start=1
        ):
            hex_color = rgb_list_to_hex(theme_color.RGB())
            if hex_color is None:
                continue
            colors.append({
                "index": i,
                "name": theme_color_names[i - 1] if i <= 12 else f"Color{i}",
                "color_hex": hex_color,
            })
    except Exception:
        logger.debug("Theme colour scheme unavailable on this master",
                     exc_info=True)

    return {
        "status": "success",
        "design_name": _name_or_none(design),
        "master_name": _name_or_none(master),
        "layout_count": count(master.custom_layouts),
        # `has title master` lives on the presentation here, not on the design
        # the way COM has it.
        "has_title_master": bool(pres.has_title_master()),
        "theme_colors": colors,
        "note": _NO_NAMES_NOTE,
    }
