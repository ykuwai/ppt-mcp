"""Batch formatting operations for multiple shapes at once."""

import json
import logging
from typing import List, Literal, Union, get_args

from pydantic import BaseModel, Field, ConfigDict

from utils.offload import run_offloaded
from backend import ppt

# The input models are imported by name because they are platform neutral. The
# impl functions are reached through their module instead, and looked up at call
# time, because on macOS they are swapped for their Apple Event counterparts
# after this module has already been imported. Binding the names here would
# quietly keep the COM versions.
from ppt_com import effects as _effects
from ppt_com import formatting as _formatting
from ppt_com import text as _text
from ppt_com.effects import SetGlowInput, SetReflectionInput, SetSoftEdgeInput
from ppt_com.formatting import SetFillInput, SetLineInput, SetShadowInput
from ppt_com.text import FormatTextInput

logger = logging.getLogger(__name__)


OperationName = Literal[
    "set_fill", "set_line", "set_shadow",
    "set_glow", "set_reflection", "set_soft_edge",
    "format_text",
]

SUPPORTED_OPERATIONS = list(get_args(OperationName))


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
    model_config = ConfigDict(str_strip_whitespace=True)

    tool: OperationName = Field(
        ...,
        description="Formatting operation to apply.",
    )
    params: dict = Field(
        default_factory=dict,
        description="Operation-specific parameters (without slide_index or shape_name_or_index)",
    )


class BatchApplyFormattingInput(BaseModel):
    """Apply one or more formatting operations to multiple shapes in a single call.

    Use this to bulk-style shapes efficiently — e.g., set fill + remove borders on
    4 shapes with 1 call instead of 8 separate ppt_set_fill / ppt_set_line calls.
    Supported operations: set_fill, set_line, set_shadow, set_glow, set_reflection,
    set_soft_edge, format_text.
    """
    model_config = ConfigDict(str_strip_whitespace=True)

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

# Where one idea goes by two names across these tools. The value is the name
# the operation models use; the key is what a caller arrives with, having just
# used `ppt_add_shape` or `ppt_add_textbox`.
_SIBLING_NAMES = {
    "font_color": "color",
    "line_visible": "visible",
    "line_color": "color",
    "line_weight": "weight",
    "fill_color": "color",
    "fill_transparency": "transparency",
}


def _checked(model_cls, tool_name, params, **fixed):
    """Build an operation's input model, refusing arguments it does not have.

    Pydantic drops unknown keys by default, so a batch operation carrying a
    misspelled or borrowed argument used to apply nothing and report success.
    That happened for real: `format_text` was given `font_color`, which is what
    `ppt_add_shape` and `ppt_add_textbox` call it, while this tool's own name
    for it is `color`. The text stayed the colour it was and the result said
    `"status": "success"`, and the only way to notice was to look at the slide.

    A near miss is named, because the argument that gets passed here is almost
    always the right idea under a sibling tool's name.
    """
    known = set(model_cls.model_fields)
    unknown = [key for key in params if key not in known]
    if unknown:
        import difflib

        parts = []
        for key in unknown:
            # The splits this server actually has, where spelling is no guide.
            # `font_color` is what the shape and textbox tools call what this
            # one calls `color`, and difflib answers `font_color_theme` for it,
            # which is a different thing entirely.
            suggestion = _SIBLING_NAMES.get(key)
            if suggestion not in known:
                close = difflib.get_close_matches(key, known, n=1, cutoff=0.6)
                suggestion = close[0] if close else None
            parts.append(
                f"{key!r}" + (f" (did you mean {suggestion!r}?)" if suggestion else "")
            )
        accepted = ", ".join(sorted(known - set(fixed)))
        raise ValueError(
            f"{tool_name} does not take " + ", ".join(parts)
            + f". Nothing was applied. It takes: {accepted}."
        )
    return model_cls(**fixed, **params)


def _dispatch_op(slide_index, shape_name_or_index, tool_name, params):
    """Validate params and call the appropriate impl function."""
    if tool_name == "set_fill":
        m = _checked(
            SetFillInput, tool_name, params,
            slide_index=slide_index,
            shape_name_or_index=shape_name_or_index,
        )
        return _formatting._set_fill_impl(
            slide_index, shape_name_or_index,
            m.fill_type, m.color, m.gradient_color1, m.gradient_color2,
            m.gradient_style, m.transparency,
        )

    elif tool_name == "set_line":
        m = _checked(
            SetLineInput, tool_name, params,
            slide_index=slide_index,
            shape_name_or_index=shape_name_or_index,
        )
        return _formatting._set_line_impl(
            slide_index, shape_name_or_index,
            m.color, m.weight, m.dash_style, m.visible, m.transparency,
        )

    elif tool_name == "set_shadow":
        m = _checked(
            SetShadowInput, tool_name, params,
            slide_index=slide_index,
            shape_name_or_index=shape_name_or_index,
        )
        return _formatting._set_shadow_impl(
            slide_index, shape_name_or_index,
            m.visible, m.blur, m.offset_x, m.offset_y, m.color,
            m.transparency,
        )

    elif tool_name == "set_glow":
        m = _checked(
            SetGlowInput, tool_name, params,
            slide_index=slide_index,
            shape_name_or_index=shape_name_or_index,
        )
        return _effects._set_glow_impl(
            slide_index, shape_name_or_index,
            m.radius, m.color, m.transparency,
        )

    elif tool_name == "set_reflection":
        m = _checked(
            SetReflectionInput, tool_name, params,
            slide_index=slide_index,
            shape_name_or_index=shape_name_or_index,
        )
        return _effects._set_reflection_impl(
            slide_index, shape_name_or_index,
            m.reflection_type, m.blur, m.offset, m.size, m.transparency,
        )

    elif tool_name == "set_soft_edge":
        m = _checked(
            SetSoftEdgeInput, tool_name, params,
            slide_index=slide_index,
            shape_name_or_index=shape_name_or_index,
        )
        return _effects._set_soft_edge_impl(
            slide_index, shape_name_or_index,
            m.radius,
        )

    elif tool_name == "format_text":
        m = _checked(
            FormatTextInput, tool_name, params,
            slide_index=slide_index,
            shape_name_or_index=shape_name_or_index,
        )
        return _text._format_text_impl(
            slide_index, shape_name_or_index,
            m.font_name, m.font_name_fareast,
            m.font_size, m.bold, m.italic, m.underline,
            m.color, m.font_color_theme, m.highlight_color,
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
        return await run_offloaded(batch_apply_formatting, params)


# ---------------------------------------------------------------------------
# macOS
# ---------------------------------------------------------------------------
# The implementation above walks COM. Its Apple Event counterpart has the same
# name and signature, so on macOS it simply takes its place; nothing else in
# this module changes, and `_dispatch_op` above is already reaching the Apple
# Event versions of the tools it calls.
from backend import IS_MACOS, use_mac_impls  # noqa: E402

if IS_MACOS:  # pragma: no cover - platform specific
    from ppt_mac import batch_apply as _mac_batch_apply

    use_mac_impls(globals(), _mac_batch_apply)
