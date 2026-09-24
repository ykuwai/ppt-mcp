"""Fill, line, and shadow effect tools for PowerPoint COM automation."""

import json
import logging
from typing import Literal, Optional, Union

from pydantic import BaseModel, Field, ConfigDict, field_validator

from utils.offload import run_offloaded
from backend import ppt
from utils.navigation import goto_slide
from utils.color import hex_to_int
from ppt_com.constants import (
    msoTrue, msoFalse,
    msoFillSolid, msoFillGradient,
    msoGradientHorizontal, msoGradientVertical,
    msoGradientDiagonalUp, msoGradientDiagonalDown,
    msoGradientFromCorner, msoGradientFromCenter,
    DASH_STYLE_DESCRIPTION,
    check_dash_style, dash_style_value,
)
from ppt_com.shape_lookup import resolve_shape as _get_shape
from ppt_com.effects import (
    TARGET_FIELD_DESCRIPTION, effect_of, nothing_drawn_warning,
    will_not_draw,
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Constant maps
# ---------------------------------------------------------------------------
GRADIENT_STYLE_MAP = {
    "horizontal": msoGradientHorizontal,
    "vertical": msoGradientVertical,
    "diagonal_up": msoGradientDiagonalUp,
    "diagonal_down": msoGradientDiagonalDown,
    "from_corner": msoGradientFromCorner,
    "from_center": msoGradientFromCenter,
}

# DASH_STYLE_MAP lives in ppt_com.constants and is shared by every line tool.


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class SetFillInput(BaseModel):
    """Input for setting shape fill."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    slide_index: int = Field(..., description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    fill_type: str = Field(
        ..., description="'solid', 'gradient', or 'none'"
    )
    color: Optional[str] = Field(
        default=None, description="Fill color as '#RRGGBB' hex (for solid fill)"
    )
    gradient_color1: Optional[str] = Field(
        default=None, description="Gradient start color as '#RRGGBB'"
    )
    gradient_color2: Optional[str] = Field(
        default=None, description="Gradient end color as '#RRGGBB'"
    )
    gradient_style: Optional[str] = Field(
        default=None,
        description="'horizontal', 'vertical', 'diagonal_up', 'diagonal_down', 'from_corner', or 'from_center'"
    )
    transparency: Optional[float] = Field(
        default=None, description="Transparency 0.0 (opaque) to 1.0 (fully transparent)"
    )


class SetLineInput(BaseModel):
    """Input for setting shape border/line."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    slide_index: int = Field(..., description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    color: Optional[str] = Field(default=None, description="Line color as '#RRGGBB'")
    weight: Optional[float] = Field(default=None, description="Line weight in points")
    dash_style: Optional[str] = Field(
        default=None, description=DASH_STYLE_DESCRIPTION
    )
    visible: Optional[bool] = Field(default=None, description="Line visible on/off")
    transparency: Optional[float] = Field(
        default=None, description="Transparency 0.0 (opaque) to 1.0 (fully transparent)"
    )

    @field_validator("dash_style")
    @classmethod
    def _dash_style_known(cls, v):
        return check_dash_style(v)


class SetShadowInput(BaseModel):
    """Input for setting shadow effect."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    slide_index: int = Field(..., description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    visible: bool = Field(..., description="Shadow visible on/off")
    blur: Optional[float] = Field(default=None, description="Shadow blur radius in points")
    offset_x: Optional[float] = Field(default=None, description="Shadow horizontal offset in points")
    offset_y: Optional[float] = Field(default=None, description="Shadow vertical offset in points")
    color: Optional[str] = Field(default=None, description="Shadow color as '#RRGGBB'")
    transparency: Optional[float] = Field(
        default=None, description="Transparency 0.0 (opaque) to 1.0 (fully transparent)"
    )
    target: Literal["shape", "text"] = Field(
        default="shape", description=TARGET_FIELD_DESCRIPTION
    )


# ---------------------------------------------------------------------------
# COM implementation functions
# ---------------------------------------------------------------------------
def _set_fill_impl(slide_index, shape_name_or_index, fill_type,
                    color, gradient_color1, gradient_color2, gradient_style,
                    transparency) -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    fill = shape.Fill

    if fill_type == "none":
        fill.Visible = msoFalse
    elif fill_type == "solid":
        fill.Solid()
        if color is not None:
            fill.ForeColor.RGB = hex_to_int(color)
    elif fill_type == "gradient":
        style_val = GRADIENT_STYLE_MAP.get(gradient_style, msoGradientHorizontal)
        fill.TwoColorGradient(Style=style_val, Variant=1)
        if gradient_color1 is not None:
            fill.ForeColor.RGB = hex_to_int(gradient_color1)
        if gradient_color2 is not None:
            fill.BackColor.RGB = hex_to_int(gradient_color2)
    else:
        raise ValueError(
            f"Invalid fill_type '{fill_type}'. Valid values: 'solid', 'gradient', 'none'"
        )

    if transparency is not None and fill_type != "none":
        fill.Transparency = transparency

    return {
        "status": "success",
        "shape_name": shape.Name,
        "fill_type": fill_type,
    }


def _set_line_impl(slide_index, shape_name_or_index,
                    color, weight, dash_style, visible, transparency) -> dict:
    # Resolved before the view moves or anything is written, so an unknown
    # name changes nothing.
    dash_val = dash_style_value(dash_style) if dash_style is not None else None

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    line = shape.Line

    if visible is not None:
        line.Visible = msoTrue if visible else msoFalse

    if color is not None:
        line.ForeColor.RGB = hex_to_int(color)

    if weight is not None:
        line.Weight = weight

    if dash_val is not None:
        line.DashStyle = dash_val

    if transparency is not None:
        line.Transparency = transparency

    return {
        "status": "success",
        "shape_name": shape.Name,
    }


def _set_shadow_impl(slide_index, shape_name_or_index,
                      visible, blur, offset_x, offset_y, color, transparency,
                      target="shape") -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    shadow = effect_of(shape, target, "Shadow")

    shadow.Visible = msoTrue if visible else msoFalse

    if visible:
        if blur is not None:
            shadow.Blur = blur
        if offset_x is not None:
            shadow.OffsetX = offset_x
        if offset_y is not None:
            shadow.OffsetY = offset_y
        if color is not None:
            shadow.ForeColor.RGB = hex_to_int(color)
        if transparency is not None:
            shadow.Transparency = transparency

    result = {
        "status": "success",
        "shape_name": shape.Name,
        "target": target,
        "shadow_visible": visible,
    }
    if target == "shape" and visible and will_not_draw(shape):
        result["warnings"] = [nothing_drawn_warning(shape, "shadow")]
    return result


# ---------------------------------------------------------------------------
# MCP tool functions
# ---------------------------------------------------------------------------
def set_fill(params: SetFillInput) -> str:
    """Set shape fill (solid, gradient, or none)."""
    try:
        result = ppt.execute(
            _set_fill_impl,
            params.slide_index, params.shape_name_or_index, params.fill_type,
            params.color, params.gradient_color1, params.gradient_color2,
            params.gradient_style, params.transparency,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def set_line(params: SetLineInput) -> str:
    """Set shape border/line properties."""
    try:
        result = ppt.execute(
            _set_line_impl,
            params.slide_index, params.shape_name_or_index,
            params.color, params.weight, params.dash_style,
            params.visible, params.transparency,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def set_shadow(params: SetShadowInput) -> str:
    """Set shadow effect on a shape."""
    try:
        result = ppt.execute(
            _set_shadow_impl,
            params.slide_index, params.shape_name_or_index,
            params.visible, params.blur, params.offset_x, params.offset_y,
            params.color, params.transparency, params.target,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register all formatting tools with the MCP server."""

    @mcp.tool(
        name="ppt_set_fill",
        annotations={
            "title": "Set Shape Fill",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_ppt_set_fill(params: SetFillInput) -> str:
        """Set the fill of a shape.

        Supports solid color, two-color gradient, or no fill.
        For solid fills, provide a color hex. For gradients, provide
        gradient_color1, gradient_color2, and gradient_style.
        """
        return await run_offloaded(set_fill, params)

    @mcp.tool(
        name="ppt_set_line",
        annotations={
            "title": "Set Shape Line/Border",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_ppt_set_line(params: SetLineInput) -> str:
        """Set the border/line of a shape.

        Configure color, weight, dash style, visibility, and transparency.
        Line cap (round, flat or square ends) is not in the PowerPoint object
        model and cannot be set here. To match a line that has one, use
        ppt_copy_formatting from that line.
        """
        return await run_offloaded(set_line, params)

    @mcp.tool(
        name="ppt_set_shadow",
        annotations={
            "title": "Set Shape Shadow",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_ppt_set_shadow(params: SetShadowInput) -> str:
        """Set a shadow on a shape, or on its text.

        Configure blur, offset, color and transparency. visible=false
        removes it.

        target='text' shadows the glyphs rather than the shape. A shape
        shadow on a text box with no fill and no line draws nothing, and the
        call says so in a warning rather than reporting a plain success.
        """
        return await run_offloaded(set_shadow, params)


# ---------------------------------------------------------------------------
# macOS
# ---------------------------------------------------------------------------
# The implementations above walk COM. Their Apple Event counterparts have the
# same names and signatures, so on macOS they simply take their place; nothing
# else in this module changes.
from backend import IS_MACOS, use_mac_impls  # noqa: E402

if IS_MACOS:  # pragma: no cover - platform specific
    from ppt_mac import formatting as _mac_formatting

    use_mac_impls(globals(), _mac_formatting)
