"""Editing operations for PowerPoint COM automation.

Handles undo, redo, shape copy across slides, format painting,
undo entry management, and MSO command execution.
"""

import json
import logging
from typing import Union

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from ppt_com.constants import msoTrue, msoFalse

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class UndoInput(BaseModel):
    """Input for undo operations."""
    model_config = ConfigDict(str_strip_whitespace=True)

    times: int = Field(
        default=1, ge=1,
        description="Number of undo operations to perform",
    )


class RedoInput(BaseModel):
    """Input for redo operations."""
    model_config = ConfigDict(str_strip_whitespace=True)

    times: int = Field(
        default=1, ge=1,
        description="Number of redo operations to perform",
    )


class CopyShapeToSlideInput(BaseModel):
    """Input for copying a shape to another slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    src_slide_index: int = Field(..., ge=1, description="1-based source slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int) on the source slide. Prefer name — indices shift when shapes are added/removed"
    )
    dst_slide_index: int = Field(..., ge=1, description="1-based destination slide index")


class CopyFormattingInput(BaseModel):
    """Input for copying formatting from one shape to others."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    source_shape: Union[str, int] = Field(
        ..., description="Source shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    target_shapes: list[Union[str, int]] = Field(
        ..., min_length=1,
        description="List of target shape names (string) or 1-based indices (int)",
    )


class ExecuteMsoInput(BaseModel):
    """Input for executing an MSO command."""
    model_config = ConfigDict(str_strip_whitespace=True)

    command_name: str = Field(
        ..., description="MSO command name to execute"
    )
    check_enabled: bool = Field(
        default=True,
        description="If true, check whether the command is enabled before executing",
    )


# ---------------------------------------------------------------------------
# Helper: find a shape by name or index
# ---------------------------------------------------------------------------
def _get_shape(slide, name_or_index: Union[str, int]):
    """Find a shape on a slide by name (str) or 1-based index (int).

    Args:
        slide: Slide COM object
        name_or_index: Shape name (str) or 1-based index (int)

    Returns:
        Shape COM object

    Raises:
        ValueError: If shape not found or index out of range
    """
    if isinstance(name_or_index, int):
        if name_or_index < 1 or name_or_index > slide.Shapes.Count:
            raise ValueError(
                f"Shape index {name_or_index} out of range "
                f"(1-{slide.Shapes.Count})"
            )
        return slide.Shapes(name_or_index)
    else:
        for i in range(1, slide.Shapes.Count + 1):
            if slide.Shapes(i).Name == name_or_index:
                return slide.Shapes(i)
        raise ValueError(f"Shape '{name_or_index}' not found on slide")


# ---------------------------------------------------------------------------
# COM implementation functions (run on COM thread via ppt.execute)
# ---------------------------------------------------------------------------
def _undo_impl(times):
    app = ppt._get_app_impl()
    count = 0
    for _ in range(times):
        if not app.CommandBars.GetEnabledMso("Undo"):
            break
        app.CommandBars.ExecuteMso("Undo")
        count += 1
    return {"success": True, "actions_undone": count}


def _redo_impl(times):
    app = ppt._get_app_impl()
    count = 0
    for _ in range(times):
        if not app.CommandBars.GetEnabledMso("Redo"):
            break
        app.CommandBars.ExecuteMso("Redo")
        count += 1
    return {"success": True, "actions_redone": count}


def _copy_shape_to_slide_impl(src_slide_index, shape_name_or_index, dst_slide_index):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()

    src_slide = pres.Slides(src_slide_index)
    shape = _get_shape(src_slide, shape_name_or_index)

    dst_slide = pres.Slides(dst_slide_index)

    shape.Copy()
    pasted = dst_slide.Shapes.Paste()
    new_shape = pasted(1)

    return {
        "success": True,
        "new_shape_name": new_shape.Name,
        "destination_slide": dst_slide_index,
    }


def _copy_formatting_impl(slide_index, source_shape, target_shapes):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)

    src = _get_shape(slide, source_shape)
    src.PickUp()

    applied_to = []
    for target_id in target_shapes:
        target = _get_shape(slide, target_id)
        target.Apply()
        applied_to.append(target.Name)

    return {
        "success": True,
        "source": src.Name,
        "applied_to": applied_to,
    }


def _start_undo_entry_impl():
    app = ppt._get_app_impl()
    app.StartNewUndoEntry()
    return {"success": True}


def _execute_mso_impl(command_name, check_enabled):
    app = ppt._get_app_impl()

    if check_enabled:
        enabled = app.CommandBars.GetEnabledMso(command_name)
        if not enabled:
            return {
                "error": f"Command '{command_name}' is not currently enabled/available"
            }

    app.CommandBars.ExecuteMso(command_name)
    return {"success": True, "command": command_name}


# ---------------------------------------------------------------------------
# MCP tool functions (return JSON strings)
# ---------------------------------------------------------------------------
def undo(params: UndoInput) -> str:
    """Undo recent actions in PowerPoint.

    Args:
        params: Number of undo operations to perform.

    Returns:
        JSON with success status and number of actions undone.
    """
    try:
        result = ppt.execute(_undo_impl, params.times)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to undo: {str(e)}"})


def redo(params: RedoInput) -> str:
    """Redo recently undone actions in PowerPoint.

    Args:
        params: Number of redo operations to perform.

    Returns:
        JSON with success status and number of actions redone.
    """
    try:
        result = ppt.execute(_redo_impl, params.times)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to redo: {str(e)}"})


def copy_shape_to_slide(params: CopyShapeToSlideInput) -> str:
    """Copy a shape from one slide to another.

    Args:
        params: Source slide, shape identifier, and destination slide.

    Returns:
        JSON with new shape name and destination slide index.
    """
    try:
        result = ppt.execute(
            _copy_shape_to_slide_impl,
            params.src_slide_index, params.shape_name_or_index,
            params.dst_slide_index,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to copy shape to slide: {str(e)}"})


def copy_formatting(params: CopyFormattingInput) -> str:
    """Copy formatting from one shape to other shapes.

    Args:
        params: Slide index, source shape, and list of target shapes.

    Returns:
        JSON with source name and list of shapes formatting was applied to.
    """
    try:
        result = ppt.execute(
            _copy_formatting_impl,
            params.slide_index, params.source_shape, params.target_shapes,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to copy formatting: {str(e)}"})


def start_undo_entry() -> str:
    """Start a new undo entry in PowerPoint.

    Returns:
        JSON with success status.
    """
    try:
        result = ppt.execute(_start_undo_entry_impl)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to start undo entry: {str(e)}"})


def execute_mso(params: ExecuteMsoInput) -> str:
    """Execute an MSO command in PowerPoint.

    Args:
        params: Command name and whether to check if enabled first.

    Returns:
        JSON with success status and command name.
    """
    try:
        result = ppt.execute(
            _execute_mso_impl,
            params.command_name, params.check_enabled,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to execute MSO command: {str(e)}"})


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register all edit operation tools with the MCP server."""

    @mcp.tool(
        name="ppt_undo",
        annotations={
            "title": "Undo",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_undo(params: UndoInput) -> str:
        """Undo recent actions in PowerPoint.

        Performs one or more undo operations. Stops early if no more
        actions can be undone. Returns the number of actions actually undone.
        """
        return undo(params)

    @mcp.tool(
        name="ppt_redo",
        annotations={
            "title": "Redo",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_redo(params: RedoInput) -> str:
        """Redo recently undone actions in PowerPoint.

        Performs one or more redo operations. Stops early if no more
        actions can be redone. Returns the number of actions actually redone.
        """
        return redo(params)

    @mcp.tool(
        name="ppt_copy_shape_to_slide",
        annotations={
            "title": "Copy Shape to Slide",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_copy_shape_to_slide(params: CopyShapeToSlideInput) -> str:
        """Copy a shape from one slide to another slide.

        Copies the shape via the clipboard and pastes it onto the destination
        slide. The original shape remains on the source slide.
        Identify the shape by name (string) or 1-based index (int).
        """
        return copy_shape_to_slide(params)

    @mcp.tool(
        name="ppt_copy_formatting",
        annotations={
            "title": "Copy Formatting",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_copy_formatting(params: CopyFormattingInput) -> str:
        """Copy formatting from a source shape to one or more target shapes.

        Uses PickUp/Apply to transfer fill, line, shadow, and other visual
        formatting from the source shape to each target shape on the same slide.
        Identify shapes by name (string) or 1-based index (int).
        """
        return copy_formatting(params)

    @mcp.tool(
        name="ppt_start_undo_entry",
        annotations={
            "title": "Start Undo Entry",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_start_undo_entry() -> str:
        """Start a new undo entry in PowerPoint.

        Call this before a batch of operations so that all subsequent
        changes can be undone with a single Ctrl+Z. Useful for grouping
        multiple tool calls into one undoable action.
        """
        return start_undo_entry()

    @mcp.tool(
        name="ppt_execute_mso",
        annotations={
            "title": "Execute MSO Command",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_execute_mso(params: ExecuteMsoInput) -> str:
        """Execute a built-in PowerPoint MSO command by name.

        Runs any command from the Office command bar system. By default,
        checks if the command is enabled before executing.

        Common commands: SelectAll, Copy, Cut, Paste, AnimationPreview,
        SlideShowFromBeginning, SlideShowFromCurrent, Bold, Italic,
        Underline, AlignLeft, AlignCenter, AlignRight, GroupObjects,
        UngroupObjects, BringToFront, SendToBack.

        Set check_enabled=false to skip the enabled check (may raise
        a COM error if the command is not available).
        """
        return execute_mso(params)
