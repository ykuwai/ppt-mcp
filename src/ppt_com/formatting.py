"""Fill, line, and shadow effect tools for PowerPoint COM automation."""

import json
import logging
from typing import Optional, Union

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from utils.navigation import goto_slide
from utils.color import hex_to_int
from ppt_com.constants import (
    msoTrue, msoFalse,
    msoFillSolid, msoFillGradient,
    msoGradientHorizontal, msoGradientVertical,
    msoGradientDiagonalUp, msoGradientDiagonalDown,
    msoGradientFromCorner, msoGradientFromCenter,
    msoLineSolid, msoLineRoundDot, msoLineDash,
    msoLineDashDot, msoLineLongDash,
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Helper: reuse _get_shape from text module
# ---------------------------------------------------------------------------
def _get_shape(slide, name_or_index):
    """Find shape by name (str) or index (int)."""
    if isinstance(name_or_index, int):
        if name_or_index < 1 or name_or_index > slide.Shapes.Count:
            raise ValueError(
                f"Shape index {name_or_index} out of range "
                f"(1-{slide.Shapes.Count})"
            )
        return slide.Shapes(name_or_index)
    else:
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            if shape.Name == name_or_index:
                return shape
        raise ValueError(f"Shape '{name_or_index}' not found on slide")


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

DASH_STYLE_MAP = {
    "solid": msoLineSolid,
    "round_dot": msoLineRoundDot,
    "dash": msoLineDash,
    "dash_dot": msoLineDashDot,
    "long_dash": msoLineLongDash,
}


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class SetFillInput(BaseModel):
    """Input for setting shape fill."""
    model_config = ConfigDict(str_strip_whitespace=True)

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
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    color: Optional[str] = Field(default=None, description="Line color as '#RRGGBB'")
    weight: Optional[float] = Field(default=None, description="Line weight in points")
    dash_style: Optional[str] = Field(
        default=None,
        description="'solid', 'round_dot', 'dash', 'dash_dot', or 'long_dash'"
    )
    visible: Optional[bool] = Field(default=None, description="Line visible on/off")
    transparency: Optional[float] = Field(
        default=None, description="Transparency 0.0 (opaque) to 1.0 (fully transparent)"
    )


class SetShadowInput(BaseModel):
    """Input for setting shadow effect."""
    model_config = ConfigDict(str_strip_whitespace=True)

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
        fill.Background()
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

    if transparency is not None:
        fill.Transparency = transparency

    return {
        "status": "success",
        "shape_name": shape.Name,
        "fill_type": fill_type,
    }


def _set_line_impl(slide_index, shape_name_or_index,
                    color, weight, dash_style, visible, transparency) -> dict:
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

    if dash_style is not None:
        dash_val = DASH_STYLE_MAP.get(dash_style)
        if dash_val is None:
            raise ValueError(
                f"Invalid dash_style '{dash_style}'. "
                f"Valid values: {list(DASH_STYLE_MAP.keys())}"
            )
        line.DashStyle = dash_val

    if transparency is not None:
        line.Transparency = transparency

    return {
        "status": "success",
        "shape_name": shape.Name,
    }


def _set_shadow_impl(slide_index, shape_name_or_index,
                      visible, blur, offset_x, offset_y, color, transparency) -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    shadow = shape.Shadow

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

    return {
        "status": "success",
        "shape_name": shape.Name,
        "shadow_visible": visible,
    }


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
            params.color, params.transparency,
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
        return set_fill(params)

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
        """
        return set_line(params)

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
        """Set shadow effect on a shape.

        Configure blur, offset, color, and transparency.
        Set visible=false to remove the shadow.
        """
        return set_shadow(params)
