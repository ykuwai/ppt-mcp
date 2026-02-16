"""Shape grouping and ungrouping tools for PowerPoint COM automation.

Handles grouping multiple shapes, ungrouping group shapes,
and inspecting the items within a group.
"""

import json
import logging
from typing import Union

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from ppt_com.constants import msoGroup, SHAPE_TYPE_NAMES

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class GroupShapesInput(BaseModel):
    """Input for grouping shapes on a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_names: list[str] = Field(
        ..., min_length=2,
        description="List of shape names to group (minimum 2)",
    )


class UngroupShapesInput(BaseModel):
    """Input for ungrouping a group shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Group shape name (string) or 1-based index (int)"
    )


class GetGroupItemsInput(BaseModel):
    """Input for getting items within a group shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Group shape name (string) or 1-based index (int)"
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
def _group_shapes_impl(slide_index, shape_names):
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    # Validate all shape names exist before grouping
    for name in shape_names:
        found = False
        for i in range(1, slide.Shapes.Count + 1):
            if slide.Shapes(i).Name == name:
                found = True
                break
        if not found:
            raise ValueError(f"Shape '{name}' not found on slide {slide_index}")

    shape_range = slide.Shapes.Range(shape_names)
    group = shape_range.Group()

    return {
        "success": True,
        "group_name": group.Name,
        "shape_index": group.ZOrderPosition,
    }


def _ungroup_shapes_impl(slide_index, shape_name_or_index):
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    if shape.Type != msoGroup:
        raise ValueError(
            f"Shape '{shape.Name}' is not a group (type={shape.Type}). "
            f"Only group shapes (type={msoGroup}) can be ungrouped."
        )

    ungrouped = shape.Ungroup()
    names = []
    for i in range(1, ungrouped.Count + 1):
        names.append(ungrouped(i).Name)

    return {
        "success": True,
        "ungrouped_count": ungrouped.Count,
        "shape_names": names,
    }


def _get_group_items_impl(slide_index, shape_name_or_index):
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    if shape.Type != msoGroup:
        raise ValueError(
            f"Shape '{shape.Name}' is not a group (type={shape.Type}). "
            f"Only group shapes (type={msoGroup}) can be inspected."
        )

    items = []
    for i in range(1, shape.GroupItems.Count + 1):
        item = shape.GroupItems(i)
        type_val = item.Type
        items.append({
            "name": item.Name,
            "type": type_val,
            "type_name": SHAPE_TYPE_NAMES.get(type_val, f"Unknown({type_val})"),
            "left": round(item.Left, 2),
            "top": round(item.Top, 2),
            "width": round(item.Width, 2),
            "height": round(item.Height, 2),
        })

    return {
        "success": True,
        "group_name": shape.Name,
        "items": items,
    }


# ---------------------------------------------------------------------------
# MCP tool functions (async wrappers that delegate to COM thread)
# ---------------------------------------------------------------------------
def group_shapes(params: GroupShapesInput) -> str:
    """Group multiple shapes into a single group shape.

    Args:
        params: Slide index and list of shape names to group.

    Returns:
        JSON with group name and shape index.
    """
    try:
        result = ppt.execute(
            _group_shapes_impl,
            params.slide_index, params.shape_names,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to group shapes: {str(e)}"})


def ungroup_shapes(params: UngroupShapesInput) -> str:
    """Ungroup a group shape into its individual shapes.

    Args:
        params: Slide index and group shape identifier.

    Returns:
        JSON with ungrouped count and shape names.
    """
    try:
        result = ppt.execute(
            _ungroup_shapes_impl,
            params.slide_index, params.shape_name_or_index,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to ungroup shapes: {str(e)}"})


def get_group_items(params: GetGroupItemsInput) -> str:
    """Get information about all items within a group shape.

    Args:
        params: Slide index and group shape identifier.

    Returns:
        JSON with group name and list of item details.
    """
    try:
        result = ppt.execute(
            _get_group_items_impl,
            params.slide_index, params.shape_name_or_index,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to get group items: {str(e)}"})


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register all group tools with the MCP server."""

    @mcp.tool(
        name="ppt_group_shapes",
        annotations={
            "title": "Group Shapes",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_group_shapes(params: GroupShapesInput) -> str:
        """Group multiple shapes into a single group.

        Provide at least 2 shape names to combine into a group.
        The individual shapes are replaced by a single group shape.
        """
        return group_shapes(params)

    @mcp.tool(
        name="ppt_ungroup_shapes",
        annotations={
            "title": "Ungroup Shapes",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_ungroup_shapes(params: UngroupShapesInput) -> str:
        """Ungroup a group shape into its individual shapes.

        The group shape is removed and its child shapes become
        independent shapes on the slide.
        """
        return ungroup_shapes(params)

    @mcp.tool(
        name="ppt_get_group_items",
        annotations={
            "title": "Get Group Items",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_get_group_items(params: GetGroupItemsInput) -> str:
        """Get information about all items within a group shape.

        Returns name, type, position, and size for each item in the group.
        Identify the group by shape name or 1-based shape index.
        """
        return get_group_items(params)
