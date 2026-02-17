"""SmartArt operations for PowerPoint COM automation.

Handles creating SmartArt graphics, modifying nodes (set text, add, delete),
and listing available SmartArt layouts.
"""

import json
import logging
from typing import Optional, Union

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from utils.navigation import goto_slide
from ppt_com.constants import msoSmartArt, SHAPE_TYPE_NAMES

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class AddSmartArtInput(BaseModel):
    """Input for adding a SmartArt graphic to a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    layout_name: Optional[str] = Field(
        default=None,
        description="Partial or full SmartArt layout name to search for (case-insensitive match)",
    )
    layout_index: Optional[int] = Field(
        default=None, ge=1,
        description="1-based SmartArt layout index (used if layout_name is not provided, defaults to 1)",
    )
    left: float = Field(default=50.0, description="Left position in points")
    top: float = Field(default=50.0, description="Top position in points")
    width: float = Field(default=600.0, description="Width in points")
    height: float = Field(default=400.0, description="Height in points")
    node_texts: Optional[list[str]] = Field(
        default=None,
        description="List of text strings to populate SmartArt nodes in order",
    )


class ModifySmartArtInput(BaseModel):
    """Input for modifying a SmartArt node."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="SmartArt shape name (string) or 1-based index (int)"
    )
    action: str = Field(
        ..., description="Action to perform: 'set_text', 'add_node', or 'delete_node'"
    )
    node_index: Optional[int] = Field(
        default=None, ge=1,
        description="1-based node index in AllNodes. Required for 'set_text' and 'delete_node'. For 'add_node', specifies the node to add after (omit to append).",
    )
    text: Optional[str] = Field(
        default=None,
        description="Text to set on the node. Required for 'set_text', optional for 'add_node'.",
    )


# ---------------------------------------------------------------------------
# Helper: find a shape on a slide
# ---------------------------------------------------------------------------
def _get_shape(slide, name_or_index: Union[str, int]):
    """Find a shape on a slide by name or 1-based index.

    Args:
        slide: Slide COM object
        name_or_index: Shape name (str) or 1-based index (int)

    Returns:
        Shape COM object

    Raises:
        ValueError: If shape not found
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
def _add_smartart_impl(slide_index, layout_name, layout_index, left, top, width, height, node_texts):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    # Resolve the SmartArt layout
    if layout_name:
        layout = None
        for j in range(1, app.SmartArtLayouts.Count + 1):
            if layout_name.lower() in app.SmartArtLayouts(j).Name.lower():
                layout = app.SmartArtLayouts(j)
                break
        if not layout:
            raise ValueError(f"SmartArt layout '{layout_name}' not found")
    else:
        idx = layout_index if layout_index else 1
        layout = app.SmartArtLayouts(idx)

    resolved_layout_name = layout.Name
    shape = slide.Shapes.AddSmartArt(layout, left, top, width, height)

    # Set node texts if provided
    if node_texts:
        smart_art = shape.SmartArt
        for i, text in enumerate(node_texts):
            if i < smart_art.AllNodes.Count:
                # CRITICAL: Use TextFrame2 NOT TextFrame for SmartArt
                smart_art.AllNodes(i + 1).TextFrame2.TextRange.Text = text
            else:
                # Add new node
                try:
                    node = smart_art.AllNodes(smart_art.AllNodes.Count).AddNode()
                    node.TextFrame2.TextRange.Text = text
                except Exception:
                    break  # Some layouts have fixed node counts

    node_count = shape.SmartArt.AllNodes.Count

    return {
        "success": True,
        "shape_name": shape.Name,
        "node_count": node_count,
        "layout_name": resolved_layout_name,
    }


def _modify_smartart_impl(slide_index, shape_name_or_index, action, node_index, text):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    if shape.Type != msoSmartArt:
        raise ValueError(
            f"Shape '{shape.Name}' is not a SmartArt graphic "
            f"(type={SHAPE_TYPE_NAMES.get(shape.Type, shape.Type)})"
        )

    smart_art = shape.SmartArt

    if action == "set_text":
        if node_index is None:
            raise ValueError("node_index is required for 'set_text' action")
        if text is None:
            raise ValueError("text is required for 'set_text' action")
        node = smart_art.AllNodes(node_index)
        # CRITICAL: Use TextFrame2 NOT TextFrame for SmartArt
        node.TextFrame2.TextRange.Text = text

    elif action == "add_node":
        if node_index and node_index <= smart_art.AllNodes.Count:
            new_node = smart_art.AllNodes(node_index).AddNode()
        else:
            new_node = smart_art.AllNodes(smart_art.AllNodes.Count).AddNode()
        if text:
            new_node.TextFrame2.TextRange.Text = text

    elif action == "delete_node":
        if node_index is None:
            raise ValueError("node_index is required for 'delete_node' action")
        smart_art.AllNodes(node_index).Delete()

    else:
        raise ValueError(
            f"Unknown action '{action}'. Use: 'set_text', 'add_node', or 'delete_node'"
        )

    total_nodes = smart_art.AllNodes.Count

    return {
        "success": True,
        "action": action,
        "total_nodes": total_nodes,
    }


def _list_smartart_layouts_impl():
    app = ppt._get_app_impl()
    layouts = []
    total = app.SmartArtLayouts.Count
    count = min(total, 50)
    for i in range(1, count + 1):
        layout = app.SmartArtLayouts(i)
        layouts.append({
            "index": i,
            "name": layout.Name,
            "description": layout.Description,
        })

    return {
        "success": True,
        "total_count": total,
        "returned_count": count,
        "layouts": layouts,
    }


# ---------------------------------------------------------------------------
# MCP tool functions (sync wrappers that delegate to COM thread)
# ---------------------------------------------------------------------------
def add_smartart(params: AddSmartArtInput) -> str:
    """Add a SmartArt graphic to a slide.

    Args:
        params: SmartArt parameters including layout, position, size, and node texts.

    Returns:
        JSON with shape name, node count, and layout name.
    """
    try:
        result = ppt.execute(
            _add_smartart_impl,
            params.slide_index, params.layout_name, params.layout_index,
            params.left, params.top, params.width, params.height,
            params.node_texts,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to add SmartArt: {str(e)}"})


def modify_smartart(params: ModifySmartArtInput) -> str:
    """Modify a SmartArt node (set text, add node, or delete node).

    Args:
        params: SmartArt shape identifier, action, and optional node index/text.

    Returns:
        JSON confirming the action and updated node count.
    """
    try:
        result = ppt.execute(
            _modify_smartart_impl,
            params.slide_index, params.shape_name_or_index,
            params.action, params.node_index, params.text,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to modify SmartArt: {str(e)}"})


def list_smartart_layouts() -> str:
    """List available SmartArt layouts.

    Returns:
        JSON with layout index, name, and description for each layout (first 50).
    """
    try:
        result = ppt.execute(_list_smartart_layouts_impl)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to list SmartArt layouts: {str(e)}"})


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register all SmartArt tools with the MCP server."""

    @mcp.tool(
        name="ppt_add_smartart",
        annotations={
            "title": "Add SmartArt",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_add_smartart(params: AddSmartArtInput) -> str:
        """Add a SmartArt graphic to a slide.

        Creates a SmartArt with the specified layout and optionally populates
        node text. Find layouts using ppt_list_smartart_layouts.
        All positions and sizes are in points (72 points = 1 inch).
        """
        return add_smartart(params)

    @mcp.tool(
        name="ppt_modify_smartart",
        annotations={
            "title": "Modify SmartArt",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_modify_smartart(params: ModifySmartArtInput) -> str:
        """Modify a SmartArt graphic node.

        Actions: 'set_text' (update node text), 'add_node' (add a new node),
        'delete_node' (remove a node). Uses 1-based node indexing from AllNodes.
        SmartArt nodes use TextFrame2, not TextFrame.
        """
        return modify_smartart(params)

    @mcp.tool(
        name="ppt_list_smartart_layouts",
        annotations={
            "title": "List SmartArt Layouts",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_list_smartart_layouts() -> str:
        """List available SmartArt layouts.

        Returns the first 50 SmartArt layouts with their index, name,
        and description. Use the layout name or index with ppt_add_smartart.
        """
        return list_smartart_layouts()
