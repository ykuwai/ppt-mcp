"""Visual effect tools (glow, reflection, soft edge) for PowerPoint COM automation.

Handles glow, reflection, and soft edge effects on shapes.
"""

import json
import logging
from typing import Optional, Union

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from utils.color import hex_to_int

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Helper: find a shape by name or index
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
# Pydantic input models
# ---------------------------------------------------------------------------
class SetGlowInput(BaseModel):
    """Input for setting glow effect on a shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (string) or 1-based index (int)"
    )
    radius: float = Field(
        ..., ge=0, description="Glow radius in points (0 to remove glow)"
    )
    color: Optional[str] = Field(
        default=None, description="Glow color as '#RRGGBB' hex"
    )
    transparency: Optional[float] = Field(
        default=None, ge=0, le=1,
        description="Transparency 0.0 (opaque) to 1.0 (fully transparent)"
    )


class SetReflectionInput(BaseModel):
    """Input for setting reflection effect on a shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (string) or 1-based index (int)"
    )
    reflection_type: Optional[int] = Field(
        default=None, ge=0, le=9,
        description="MsoReflectionType: 0=none, 1-9=presets"
    )
    blur: Optional[float] = Field(
        default=None, ge=0, description="Reflection blur radius in points"
    )
    offset: Optional[float] = Field(
        default=None, ge=0, description="Reflection offset in points"
    )
    size: Optional[float] = Field(
        default=None, ge=0, le=100,
        description="Reflection size as percentage (0-100)"
    )
    transparency: Optional[float] = Field(
        default=None, ge=0, le=1,
        description="Transparency 0.0 (opaque) to 1.0 (fully transparent)"
    )


class SetSoftEdgeInput(BaseModel):
    """Input for setting soft edge effect on a shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (string) or 1-based index (int)"
    )
    radius: float = Field(
        ..., ge=0, description="Soft edge radius in points (0 to remove)"
    )


# ---------------------------------------------------------------------------
# COM implementation functions
# ---------------------------------------------------------------------------
def _set_glow_impl(slide_index, shape_name_or_index, radius,
                    color, transparency) -> dict:
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    glow = shape.Glow
    glow.Radius = radius

    if color is not None:
        glow.Color.RGB = hex_to_int(color)

    if transparency is not None:
        glow.Transparency = transparency

    return {
        "status": "success",
        "shape_name": shape.Name,
        "glow_radius": radius,
    }


def _set_reflection_impl(slide_index, shape_name_or_index, reflection_type,
                          blur, offset, size, transparency) -> dict:
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    reflection = shape.Reflection

    if reflection_type is not None:
        reflection.Type = reflection_type

    if blur is not None:
        reflection.Blur = blur

    if offset is not None:
        reflection.Offset = offset

    if size is not None:
        reflection.Size = size

    if transparency is not None:
        reflection.Transparency = transparency

    return {
        "status": "success",
        "shape_name": shape.Name,
    }


def _set_soft_edge_impl(slide_index, shape_name_or_index, radius) -> dict:
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    shape.SoftEdge.Radius = radius

    return {
        "status": "success",
        "shape_name": shape.Name,
        "soft_edge_radius": radius,
    }


# ---------------------------------------------------------------------------
# MCP tool functions
# ---------------------------------------------------------------------------
def set_glow(params: SetGlowInput) -> str:
    """Set glow effect on a shape."""
    try:
        result = ppt.execute(
            _set_glow_impl,
            params.slide_index, params.shape_name_or_index, params.radius,
            params.color, params.transparency,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def set_reflection(params: SetReflectionInput) -> str:
    """Set reflection effect on a shape."""
    try:
        result = ppt.execute(
            _set_reflection_impl,
            params.slide_index, params.shape_name_or_index,
            params.reflection_type, params.blur, params.offset,
            params.size, params.transparency,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def set_soft_edge(params: SetSoftEdgeInput) -> str:
    """Set soft edge effect on a shape."""
    try:
        result = ppt.execute(
            _set_soft_edge_impl,
            params.slide_index, params.shape_name_or_index, params.radius,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register all visual effect tools with the MCP server."""

    @mcp.tool(
        name="ppt_set_glow",
        annotations={
            "title": "Set Shape Glow",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_ppt_set_glow(params: SetGlowInput) -> str:
        """Set glow effect on a shape.

        Configure radius, color, and transparency.
        Set radius=0 to remove the glow effect.
        """
        return set_glow(params)

    @mcp.tool(
        name="ppt_set_reflection",
        annotations={
            "title": "Set Shape Reflection",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_ppt_set_reflection(params: SetReflectionInput) -> str:
        """Set reflection effect on a shape.

        Configure reflection type (0=none, 1-9=presets), blur, offset,
        size, and transparency.
        """
        return set_reflection(params)

    @mcp.tool(
        name="ppt_set_soft_edge",
        annotations={
            "title": "Set Shape Soft Edge",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_ppt_set_soft_edge(params: SetSoftEdgeInput) -> str:
        """Set soft edge effect on a shape.

        Configure the soft edge radius in points.
        Set radius=0 to remove the soft edge effect.
        """
        return set_soft_edge(params)
