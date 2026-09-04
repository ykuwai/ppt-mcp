"""Fill, line, and shadow tools, on Apple Events.

Mirrors ``ppt_com/formatting.py``. Same function names, same signatures, same
returned shapes.

One thing does not carry over. ``line format`` has no ``visible`` property on
macOS, so ``ppt_set_line`` with ``visible`` set reaches for weight and
transparency instead. That is a stand-in and it is described where it happens.
"""

import logging

from appscript import k
from appscript.reference import CommandError

from backend.mac_ae import is_missing, ppt, raw
from backend.mac_enums import MsoGradientStyle, MsoLineDashStyle, to_keyword
from ppt_com.constants import msoGradientHorizontal
from ppt_mac.shapes import (
    _DASH_STYLE,
    _line_visibility_warning,
    _apply_line_visibility,
    _get_shape,
    _slide,
)
from utils.color import hex_to_rgb_list, rgb_list_to_hex
from utils.navigation import goto_slide

# ``GRADIENT_STYLE_MAP`` and ``DASH_STYLE_MAP`` live in ``ppt_com.formatting``,
# which imports this module at its own bottom, so they are fetched inside the
# functions that need them. A top level import would work in one direction and
# quietly hand back a half built module in the other.

logger = logging.getLogger(__name__)

_VALID_FILL_TYPES = ("solid", "gradient", "none")

# What each `fill format type` enumerator is called in this server's own
# vocabulary, so a fill read back can be reported in the same words the caller
# asked in. macOS names six and the tool only writes three, and the other three
# are here because a shape can already be wearing one.
_FILL_TYPE_NAMES = {
    k.fill_solid: "solid",
    k.fill_gradient: "gradient",
    k.fill_picture: "picture",
    k.fill_patterned: "patterned",
    k.fill_textured: "textured",
    k.fill_background: "background",
}


def _measure_fill(fill, color_asked: bool) -> tuple:
    """Read a fill back, and say what PowerPoint is actually holding.

    Two properties decide whether the call did what was asked. ``visible`` is
    the whole of what "none" means here, since there is no Fill.Background(),
    and ``fill format type`` is the only word for the difference between solid
    and gradient. The fore colour is read as well when one was asked for,
    because for a solid fill the colour is the request. Nothing else is read;
    each one is an Apple Event and this is the busiest styling tool here.

    Every read is guarded on its own. A shape carrying no fill at all answers
    with an error rather than a value, the same way ``ppt_get_shape_info``
    finds it.
    """
    visible = None
    try:
        visible = bool(fill.visible())
    except (CommandError, AttributeError):
        pass

    measured = None
    try:
        word = fill.fill_format_type()
        if not is_missing(word):
            measured = _FILL_TYPE_NAMES.get(word)
    except (CommandError, AttributeError):
        pass

    color_hex = None
    if color_asked:
        try:
            color_hex = rgb_list_to_hex(fill.fore_color())
        except (CommandError, AttributeError, TypeError, ValueError):
            pass

    return measured, visible, color_hex


def _fill_verdict(asked: str, measured, visible):
    """Turn what was read back into the one word the caller gets.

    ``none`` is a visibility here rather than a type, so the two readings
    answer different halves of the question and neither one alone is the
    answer. None means PowerPoint would not say.
    """
    if asked == "none":
        if visible is None:
            return None
        return "none" if visible is False else (measured or "unknown")
    if visible is False:
        return "none"
    return measured


# ---------------------------------------------------------------------------
# Apple Event implementation functions
# ---------------------------------------------------------------------------
def _set_fill_impl(slide_index, shape_name_or_index, fill_type,
                    color, gradient_color1, gradient_color2, gradient_style,
                    transparency) -> dict:
    from ppt_com.formatting import GRADIENT_STYLE_MAP

    # Before the first Apple Event. Checked at the point of use, a misspelling
    # cost the caller a jump to a slide they were not looking at.
    if fill_type not in _VALID_FILL_TYPES:
        raise ValueError(
            f"Invalid fill_type '{fill_type}'. Valid values: 'solid', 'gradient', 'none'"
        )

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    fill = shape.fill_format

    if fill_type == "none":
        # There is no Fill.Background() here. The fill format carries its own
        # visibility flag, which is what "no fill" means in the interface.
        fill.visible.set(False)
    elif fill_type == "solid":
        # Visible first, because "none" is this flag rather than a fill type
        # and a shape left invisible by an earlier call would take the colour
        # and still show nothing. ppt_set_table_cell and the default shape
        # style both set it the same way for the same reason.
        fill.visible.set(True)
        fill.solid()
        if color is not None:
            fill.fore_color.set(hex_to_rgb_list(color))
    else:  # gradient
        fill.visible.set(True)
        style_val = GRADIENT_STYLE_MAP.get(gradient_style, msoGradientHorizontal)
        fill.two_color_gradient(
            style=to_keyword(MsoGradientStyle, style_val, "gradient style"),
            variant=1,
        )
        if gradient_color1 is not None:
            fill.fore_color.set(hex_to_rgb_list(gradient_color1))
        if gradient_color2 is not None:
            fill.back_color.set(hex_to_rgb_list(gradient_color2))

    if transparency is not None and fill_type != "none":
        fill.transparency.set(transparency)

    # Read back rather than echoed. A fill PowerPoint declined comes back
    # unchanged and without an error, so the answer is what it is holding now.
    asked_color = {"solid": color, "gradient": gradient_color1}.get(fill_type)
    measured, visible, color_hex = _measure_fill(fill, asked_color is not None)
    verdict = _fill_verdict(fill_type, measured, visible)

    result = {
        "status": "success",
        "shape_name": shape.name(),
        "fill_type": verdict or fill_type,
    }
    if visible is not None:
        result["fill_visible"] = visible
    if color_hex is not None:
        result["color_hex"] = color_hex

    warnings = []
    if verdict is not None and verdict != fill_type:
        warnings.append(
            f"A {fill_type} fill was asked for and the shape reads back as "
            f"{verdict}, so PowerPoint did not take it."
        )
    if verdict is None:
        warnings.append(
            "PowerPoint would not say what fill this shape carries, so the "
            "fill type above is what was asked for rather than what was "
            "measured."
        )
    if asked_color is not None and color_hex is None:
        warnings.append(
            "The fill colour could not be read back, so it is not reported. "
            "The write itself raised nothing."
        )
    if warnings:
        result["warnings"] = warnings
    return result


def _set_line_impl(slide_index, shape_name_or_index,
                    color, weight, dash_style, visible, transparency) -> dict:
    from ppt_com.formatting import DASH_STYLE_MAP

    # Before the first write and before the view moves. This used to be checked
    # after the visibility, the colour and the weight had already been applied,
    # so a misspelled dash style handed the caller an error and a changed shape.
    dash_val = None
    if dash_style is not None:
        dash_val = DASH_STYLE_MAP.get(dash_style)
        if dash_val is None:
            raise ValueError(
                f"Invalid dash_style '{dash_style}'. "
                f"Valid values: {list(DASH_STYLE_MAP.keys())}"
            )

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    line = shape.line_format

    warnings = []

    # Order matters and follows the Windows one, so an explicit weight or
    # transparency in the same call overrides the visibility stand-in below.
    visibility_failed = None
    if visible is not None:
        warnings.append(_line_visibility_warning(visible))
        visibility_failed = _apply_line_visibility(line, visible)

    if visibility_failed and color is None and weight is None \
            and dash_val is None and transparency is None:
        # Visibility was the whole request and it did not happen, so there is
        # nothing here to call a success.
        return {
            "error": "ppt_set_line could not change this shape's border",
            "reason": visibility_failed,
            "platform": "macOS",
            "shape_name": shape.name(),
        }
    if visibility_failed:
        warnings.append(visibility_failed)

    if color is not None:
        line.fore_color.set(hex_to_rgb_list(color))

    if weight is not None:
        line.line_weight.set(weight)

    if dash_val is not None:
        # `dash style` is a name AppleScript's own vocabulary already owns, so
        # appscript cannot reach PowerPoint's property by it. The code can.
        raw(line, _DASH_STYLE).set(
            to_keyword(MsoLineDashStyle, dash_val, "dash style")
        )

    if transparency is not None:
        line.transparency.set(transparency)

    result = {
        "status": "success",
        "shape_name": shape.name(),
    }

    # Read back only what was asked for. PowerPoint clamps a weight to its own
    # minimum and reports no error for doing it, so the number it holds is
    # worth more than the number it was handed, and a property nobody named is
    # not worth an Apple Event.
    if weight is not None or visible is not None:
        try:
            result["weight"] = round(line.line_weight(), 2)
        except (CommandError, AttributeError, TypeError):
            pass
    if transparency is not None or visible is not None:
        try:
            result["transparency"] = round(line.transparency(), 2)
        except (CommandError, AttributeError, TypeError):
            pass
    if color is not None:
        try:
            measured_color = rgb_list_to_hex(line.fore_color())
            if measured_color is not None:
                result["color_hex"] = measured_color
        except (CommandError, AttributeError, TypeError, ValueError):
            pass

    if warnings:
        result["warnings"] = warnings
    return result


def _set_shadow_impl(slide_index, shape_name_or_index,
                      visible, blur, offset_x, offset_y, color, transparency) -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    shadow = shape.shadow_format

    shadow.visible.set(bool(visible))

    if visible:
        if blur is not None:
            shadow.blur.set(blur)
        if offset_x is not None:
            # The dictionary spells these with a capital X and Y, which
            # appscript keeps.
            shadow.offset_X.set(offset_x)
        if offset_y is not None:
            shadow.offset_Y.set(offset_y)
        if color is not None:
            shadow.fore_color.set(hex_to_rgb_list(color))
        if transparency is not None:
            shadow.transparency.set(transparency)

    # No read route. Nothing in PowerPoint for Mac's dictionary reads a shadow
    # back, so unlike the fill and the line above, this one cannot be measured
    # and says so rather than letting the echo pass for evidence.
    warnings = [
        "PowerPoint for Mac answers nothing about a shape's shadow, so "
        "shadow_visible is what was asked for rather than a measurement, and "
        "the shadow may be untouched."
    ]
    ignored = [
        name for name, value in (
            ("blur", blur), ("offset_x", offset_x), ("offset_y", offset_y),
            ("color", color), ("transparency", transparency),
        )
        if value is not None
    ]
    if not visible and ignored:
        warnings.append(
            f"{', '.join(ignored)} was not written, because a hidden shadow "
            "has nothing to style. Set visible to true in the same call to "
            "apply it."
        )

    return {
        "status": "success",
        "shape_name": shape.name(),
        "shadow_visible": visible,
        "warnings": warnings,
    }
