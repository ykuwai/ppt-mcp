"""Batch formatting operations for multiple shapes at once."""

import json
import logging
from typing import Any, List, Optional, Union

from pydantic import BaseModel, Field

from utils.com_wrapper import ppt

# Import impl functions from existing modules
from ppt_com.formatting import (
    _set_fill_impl, _set_line_impl, _set_shadow_impl,
    SetFillInput, SetLineInput, SetShadowInput,
)
from ppt_com.effects import (
    _set_glow_impl, _set_reflection_impl, _set_soft_edge_impl,
    SetGlowInput, SetReflectionInput, SetSoftEdgeInput,
)
from ppt_com.text import (
    _format_text_impl,
    FormatTextInput,
)

logger = logging.getLogger(__name__)


SUPPORTED_OPERATIONS = [
    "set_fill", "set_line", "set_shadow",
    "set_glow", "set_reflection", "set_soft_edge",
    "format_text",
]


# ---------------------------------------------------------------------------
# Helper
# ---------------------------------------------------------------------------

def _get_shape(slide, name_or_index):
    """Find shape by name (str) or 1-based index (int)."""
    if isinstance(name_or_index, int):
        if name_or_index < 1 or name_or_index > slide.Shapes.Count:
            raise ValueError(
                f"Shape index {name_or_index} out of range "
                f"(1-{slide.Shapes.Count})"
            )
        return slide.Shapes(name_or_index)
    for i in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes(i)
        if shape.Name == name_or_index:
            return shape
    raise ValueError(f"Shape '{name_or_index}' not found on slide")


# ---------------------------------------------------------------------------
# Input models
# ---------------------------------------------------------------------------

class BatchOperation(BaseModel):
    """A single formatting operation to apply."""
    tool: str = Field(
        ...,
        description=(
            "Operation name: set_fill, set_line, set_shadow, "
            "set_glow, set_reflection, set_soft_edge, or format_text"
        ),
    )
    params: dict = Field(
        default_factory=dict,
        description="Operation-specific parameters (without slide_index or shape_name_or_index)",
    )


class BatchApplyFormattingInput(BaseModel):
    """Input for batch applying formatting to multiple shapes."""
    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shapes: List[Union[str, int]] = Field(
        ...,
        min_length=1,
        description="List of shape names (str) or 1-based indices (int)",
    )
    operations: List[BatchOperation] = Field(
        ...,
        min_length=1,
        description="List of formatting operations to apply to each shape",
    )


# ---------------------------------------------------------------------------
# Dispatch
# ---------------------------------------------------------------------------

def _dispatch_op(slide_index, shape_name_or_index, tool_name, params):
    """Validate params and call the appropriate impl function."""
    if tool_name == "set_fill":
        m = SetFillInput(
            slide_index=slide_index,
            shape_name_or_index=shape_name_or_index,
            **params,
        )
        return _set_fill_impl(
            slide_index, shape_name_or_index,
            m.fill_type, m.color, m.gradient_color1, m.gradient_color2,
            m.gradient_style, m.transparency,
        )

    elif tool_name == "set_line":
        m = SetLineInput(
            slide_index=slide_index,
            shape_name_or_index=shape_name_or_index,
            **params,
        )
        return _set_line_impl(
            slide_index, shape_name_or_index,
            m.color, m.weight, m.dash_style, m.visible, m.transparency,
        )

    elif tool_name == "set_shadow":
        m = SetShadowInput(
            slide_index=slide_index,
            shape_name_or_index=shape_name_or_index,
            **params,
        )
        return _set_shadow_impl(
            slide_index, shape_name_or_index,
            m.visible, m.blur, m.offset_x, m.offset_y, m.color,
            m.transparency,
        )

    elif tool_name == "set_glow":
        m = SetGlowInput(
            slide_index=slide_index,
            shape_name_or_index=shape_name_or_index,
            **params,
        )
        return _set_glow_impl(
            slide_index, shape_name_or_index,
            m.radius, m.color, m.transparency,
        )

    elif tool_name == "set_reflection":
        m = SetReflectionInput(
            slide_index=slide_index,
            shape_name_or_index=shape_name_or_index,
            **params,
        )
        return _set_reflection_impl(
            slide_index, shape_name_or_index,
            m.reflection_type, m.blur, m.offset, m.size, m.transparency,
        )

    elif tool_name == "set_soft_edge":
        m = SetSoftEdgeInput(
            slide_index=slide_index,
            shape_name_or_index=shape_name_or_index,
            **params,
        )
        return _set_soft_edge_impl(
            slide_index, shape_name_or_index,
            m.radius,
        )

    elif tool_name == "format_text":
        m = FormatTextInput(
            slide_index=slide_index,
            shape_name_or_index=shape_name_or_index,
            **params,
        )
        return _format_text_impl(
            slide_index, shape_name_or_index,
            m.font_name, m.font_size, m.bold, m.italic, m.underline,
            m.color, m.font_color_theme,
        )

    else:
        supported = ", ".join(SUPPORTED_OPERATIONS)
        raise ValueError(
            f"Unsupported operation: '{tool_name}'. Supported: {supported}"
        )


# ---------------------------------------------------------------------------
# Batch implementation (runs on STA thread)
# ---------------------------------------------------------------------------

def _batch_apply_impl(slide_index, shapes, operations):
    """Apply multiple formatting operations to multiple shapes."""
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)

    results = []
    for shape_id in shapes:
        # Verify shape exists first
        try:
            _get_shape(slide, shape_id)
        except Exception as e:
            results.append({
                "shape": str(shape_id),
                "error": str(e),
                "operations": [],
            })
            continue

        shape_results = []
        for op in operations:
            try:
                _dispatch_op(slide_index, shape_id, op["tool"], op.get("params", {}))
                shape_results.append({"tool": op["tool"], "status": "success"})
            except Exception as e:
                shape_results.append({
                    "tool": op["tool"],
                    "status": "error",
                    "error": str(e),
                })

        results.append({"shape": str(shape_id), "operations": shape_results})

    return {"results": results}


# ---------------------------------------------------------------------------
# Tool function
# ---------------------------------------------------------------------------

def batch_apply_formatting(params: BatchApplyFormattingInput) -> str:
    """Apply formatting operations to multiple shapes at once.

    Applies one or more formatting operations (set_fill, set_line,
    set_shadow, set_glow, set_reflection, set_soft_edge, format_text)
    to multiple shapes in a single call.

    Each operation's params should NOT include slide_index or
    shape_name_or_index — these are provided at the top level.

    If a shape is not found or an operation fails, the error is recorded
    and processing continues with the remaining shapes/operations.

    Args:
        params: Slide index, list of shape identifiers, and operations.

    Returns:
        JSON with per-shape, per-operation results.
    """
    try:
        # Serialize operations to dicts for COM thread
        ops = [{"tool": op.tool, "params": op.params} for op in params.operations]
        result = ppt.execute(
            _batch_apply_impl,
            params.slide_index,
            list(params.shapes),
            ops,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Batch apply failed: {str(e)}"})


# ---------------------------------------------------------------------------
# Registration
# ---------------------------------------------------------------------------

def register_tools(mcp):
    @mcp.tool(
        name="ppt_batch_apply_formatting",
        annotations={"readOnlyHint": False},
    )
    async def tool_batch_apply_formatting(params: BatchApplyFormattingInput) -> str:
        return batch_apply_formatting(params)
