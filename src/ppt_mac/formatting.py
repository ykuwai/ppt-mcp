"""Fill, line, and shadow tools, on Apple Events.

Mirrors ``ppt_com/formatting.py``. Same function names, same signatures, same
returned shapes.

One thing does not carry over. ``line format`` has no ``visible`` property on
macOS, so ``ppt_set_line`` with ``visible`` set reaches for weight and
transparency instead. That is a stand-in and it is described where it happens.
"""

import logging

from backend.mac_ae import ppt, raw
from backend.mac_enums import MsoGradientStyle, MsoLineDashStyle, to_keyword
from ppt_com.constants import msoGradientHorizontal
from ppt_mac.shapes import (
    _DASH_STYLE,
    _apply_line_visibility,
    _get_shape,
    _slide,
)
from utils.color import hex_to_rgb_list
from utils.navigation import goto_slide

# ``GRADIENT_STYLE_MAP`` and ``DASH_STYLE_MAP`` live in ``ppt_com.formatting``,
# which imports this module at its own bottom, so they are fetched inside the
# functions that need them. A top level import would work in one direction and
# quietly hand back a half built module in the other.

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Apple Event implementation functions
# ---------------------------------------------------------------------------
def _set_fill_impl(slide_index, shape_name_or_index, fill_type,
                    color, gradient_color1, gradient_color2, gradient_style,
                    transparency) -> dict:
    from ppt_com.formatting import GRADIENT_STYLE_MAP

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
        fill.solid()
        if color is not None:
            fill.fore_color.set(hex_to_rgb_list(color))
    elif fill_type == "gradient":
        style_val = GRADIENT_STYLE_MAP.get(gradient_style, msoGradientHorizontal)
        fill.two_color_gradient(
            style=to_keyword(MsoGradientStyle, style_val, "gradient style"),
            variant=1,
        )
        if gradient_color1 is not None:
            fill.fore_color.set(hex_to_rgb_list(gradient_color1))
        if gradient_color2 is not None:
            fill.back_color.set(hex_to_rgb_list(gradient_color2))
    else:
        raise ValueError(
            f"Invalid fill_type '{fill_type}'. Valid values: 'solid', 'gradient', 'none'"
        )

    if transparency is not None and fill_type != "none":
        fill.transparency.set(transparency)

    return {
        "status": "success",
        "shape_name": shape.name(),
        "fill_type": fill_type,
    }


def _set_line_impl(slide_index, shape_name_or_index,
                    color, weight, dash_style, visible, transparency) -> dict:
    from ppt_com.formatting import DASH_STYLE_MAP

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    line = shape.line_format

    # Order matters and follows the Windows one, so an explicit weight or
    # transparency in the same call overrides the visibility stand-in below.
    if visible is not None:
        _apply_line_visibility(line, visible)

    if color is not None:
        line.fore_color.set(hex_to_rgb_list(color))

    if weight is not None:
        line.line_weight.set(weight)

    if dash_style is not None:
        dash_val = DASH_STYLE_MAP.get(dash_style)
        if dash_val is None:
            raise ValueError(
                f"Invalid dash_style '{dash_style}'. "
                f"Valid values: {list(DASH_STYLE_MAP.keys())}"
            )
        # `dash style` is a name AppleScript's own vocabulary already owns, so
        # appscript cannot reach PowerPoint's property by it. The code can.
        raw(line, _DASH_STYLE).set(
            to_keyword(MsoLineDashStyle, dash_val, "dash style")
        )

    if transparency is not None:
        line.transparency.set(transparency)

    return {
        "status": "success",
        "shape_name": shape.name(),
    }


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

    return {
        "status": "success",
        "shape_name": shape.name(),
        "shadow_visible": visible,
    }
