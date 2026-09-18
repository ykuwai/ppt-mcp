"""Visual effect tools (glow, reflection, soft edge) for PowerPoint COM automation.

Handles glow, reflection, and soft edge effects on shapes.
"""

import json
import logging
from typing import Literal, Optional, Union

from pydantic import BaseModel, Field, ConfigDict

from utils.offload import run_offloaded
from backend import ppt
from utils.navigation import goto_slide
from utils.color import hex_to_int
from ppt_com.shape_lookup import resolve_shape as _get_shape

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Shape effect or text effect
# ---------------------------------------------------------------------------
EFFECT_TARGETS = ("shape", "text")

TARGET_FIELD_DESCRIPTION = (
    "Which effect to set. 'shape' is the outline of the shape itself. 'text' "
    "is the halo around the glyphs, which is what lifts a caption off an "
    "illustration behind it, and the only one that draws anything on a text "
    "box with no fill and no line."
)


def effect_of(shape, target, name):
    """The shape's effect object, or the one on its text.

    A shape glow is drawn around the shape's fill and line. A text box has
    neither by default, so setting one there succeeds, reads back, and draws
    nothing. The effect such a box actually wants is on the font.
    """
    if target != "text":
        return getattr(shape, name)

    try:
        has_frame = shape.HasTextFrame
    except Exception:
        has_frame = False
    if not has_frame:
        raise ValueError(
            f"Shape '{shape.Name}' has no text frame, so it has no text "
            f"{name.lower()}. Use target='shape'."
        )
    return getattr(shape.TextFrame2.TextRange.Font, name)


def will_not_draw(shape):
    """True when a shape effect on this shape has nothing to be drawn around.

    Only asked of a shape that holds text, because a line or a picture with no
    fill is perfectly ordinary and its effect draws on what is there. A text
    box with no fill and no line is the case worth warning about, and it is
    the default a text box is created with.
    """
    try:
        if not shape.HasTextFrame:
            return False
        return not shape.Fill.Visible and not shape.Line.Visible
    except Exception:
        return False


def nothing_drawn_warning(shape, name):
    return (
        f"The {name} was set on shape '{shape.Name}', which has no fill and "
        f"no line, so there is nothing for it to be drawn around and the "
        f"slide will not change. For the halo around the text itself, use "
        f"target='text'."
    )


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class SetGlowInput(BaseModel):
    """Input for setting glow effect on a shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
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
    target: Literal["shape", "text"] = Field(
        default="shape", description=TARGET_FIELD_DESCRIPTION
    )


class SetReflectionInput(BaseModel):
    """Input for setting reflection effect on a shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
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
    target: Literal["shape", "text"] = Field(
        default="shape", description=TARGET_FIELD_DESCRIPTION
    )


class SetSoftEdgeInput(BaseModel):
    """Input for setting soft edge effect on a shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    radius: float = Field(
        ..., ge=0, description="Soft edge radius in points (0 to remove)"
    )


# ---------------------------------------------------------------------------
# COM implementation functions
# ---------------------------------------------------------------------------
def _set_glow_impl(slide_index, shape_name_or_index, radius,
                    color, transparency, target="shape") -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    glow = effect_of(shape, target, "Glow")
    glow.Radius = radius

    if color is not None:
        glow.Color.RGB = hex_to_int(color)

    if transparency is not None:
        glow.Transparency = transparency

    result = {
        "status": "success",
        "shape_name": shape.Name,
        "target": target,
        "glow_radius": radius,
    }
    if target == "shape" and radius and will_not_draw(shape):
        result["warnings"] = [nothing_drawn_warning(shape, "glow")]
    return result


def _set_reflection_impl(slide_index, shape_name_or_index, reflection_type,
                          blur, offset, size, transparency, target="shape") -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    reflection = effect_of(shape, target, "Reflection")

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

    result = {
        "status": "success",
        "shape_name": shape.Name,
        "target": target,
    }
    if target == "shape" and reflection_type and will_not_draw(shape):
        result["warnings"] = [nothing_drawn_warning(shape, "reflection")]
    return result


def _set_soft_edge_impl(slide_index, shape_name_or_index, radius) -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
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
            params.color, params.transparency, params.target,
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
            params.size, params.transparency, params.target,
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
        """Set a glow on a shape, or on its text.

        Configure radius, color and transparency. radius=0 removes it.

        **target='shape'** glows the shape's fill and line. A text box has
        neither by default, so a shape glow on one is set, reads back, and
        draws nothing at all.

        **target='text'** glows the glyphs. That is the white halo that lifts
        a caption off an illustration behind it, and on a text box it is
        almost always the one wanted. ppt_get_shape_info reports both.
        """
        return await run_offloaded(set_glow, params)

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
        """Set a reflection on a shape, or on its text.

        target='text' reflects the glyphs rather than the shape, and a shape
        reflection on a text box with no fill and no line draws nothing.


        Configure reflection type (0=none, 1-9=presets), blur, offset,
        size, and transparency.
        """
        return await run_offloaded(set_reflection, params)

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
        return await run_offloaded(set_soft_edge, params)


# ---------------------------------------------------------------------------
# macOS
# ---------------------------------------------------------------------------
# The implementations above walk COM. Their Apple Event counterparts have the
# same names and signatures, so on macOS they simply take their place; nothing
# else in this module changes.
from backend import IS_MACOS, use_mac_impls  # noqa: E402

if IS_MACOS:  # pragma: no cover - platform specific
    from ppt_mac import effects as _mac_effects

    use_mac_impls(globals(), _mac_effects)
