"""Hyperlink operations for PowerPoint COM automation.

Handles adding, listing, and removing hyperlinks on shapes.
Supports click and mouseover actions with URL, file, mailto, and slide links.
"""

import json
import logging
from typing import Optional, Union

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from utils.navigation import goto_slide
from ppt_com.constants import (
    ppActionNone, ppActionHyperlink,
    ppMouseClick, ppMouseOver,
)

logger = logging.getLogger(__name__)

ACTION_ON_MAP: dict[str, int] = {
    "click": ppMouseClick,
    "mouseover": ppMouseOver,
}


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class AddHyperlinkInput(BaseModel):
    """Input for adding a hyperlink to a shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    address: str = Field(
        ..., description="Hyperlink URL, file path, or mailto: address"
    )
    sub_address: Optional[str] = Field(
        default=None,
        description="Sub-address for slide links (e.g. '3,,' to link to slide 3)",
    )
    screen_tip: Optional[str] = Field(
        default=None, description="Tooltip text shown on hover"
    )
    action_on: str = Field(
        default="click",
        description="Trigger action: 'click' or 'mouseover'",
    )


class GetHyperlinksInput(BaseModel):
    """Input for listing hyperlinks on a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")


class RemoveHyperlinkInput(BaseModel):
    """Input for removing a hyperlink from a shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    action_on: str = Field(
        default="click",
        description="Which action to remove: 'click' or 'mouseover'",
    )


# ---------------------------------------------------------------------------
# Helper: find a shape by name or 1-based index
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

    # String name lookup
    for i in range(1, slide.Shapes.Count + 1):
        if slide.Shapes(i).Name == name_or_index:
            return slide.Shapes(i)
    raise ValueError(f"Shape '{name_or_index}' not found on slide")


# ---------------------------------------------------------------------------
# COM implementation functions (run on COM thread via ppt.execute)
# ---------------------------------------------------------------------------
def _add_hyperlink_impl(slide_index, shape_name_or_index, address, sub_address, screen_tip, action_on):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    action_key = action_on.strip().lower()
    if action_key not in ACTION_ON_MAP:
        raise ValueError(
            f"Unknown action_on '{action_on}'. Use: {', '.join(ACTION_ON_MAP.keys())}"
        )
    action_idx = ACTION_ON_MAP[action_key]

    # CRITICAL: Must set Action = ppActionHyperlink BEFORE setting Hyperlink.Address
    action_setting = shape.ActionSettings(action_idx)
    action_setting.Action = ppActionHyperlink
    action_setting.Hyperlink.Address = address
    if sub_address is not None:
        action_setting.Hyperlink.SubAddress = sub_address
    if screen_tip is not None:
        action_setting.Hyperlink.ScreenTip = screen_tip

    return {
        "success": True,
        "shape_name": shape.Name,
        "address": address,
        "sub_address": sub_address,
        "action_on": action_key,
    }


def _get_hyperlinks_impl(slide_index):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)

    hyperlinks = []
    for i in range(1, slide.Hyperlinks.Count + 1):
        hl = slide.Hyperlinks(i)
        hyperlinks.append({
            "index": i,
            "address": hl.Address,
            "sub_address": hl.SubAddress,
            "type": hl.Type,
        })

    return {
        "success": True,
        "slide_index": slide_index,
        "hyperlinks_count": slide.Hyperlinks.Count,
        "hyperlinks": hyperlinks,
    }


def _remove_hyperlink_impl(slide_index, shape_name_or_index, action_on):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    action_key = action_on.strip().lower()
    if action_key not in ACTION_ON_MAP:
        raise ValueError(
            f"Unknown action_on '{action_on}'. Use: {', '.join(ACTION_ON_MAP.keys())}"
        )
    action_idx = ACTION_ON_MAP[action_key]

    shape.ActionSettings(action_idx).Action = ppActionNone

    return {
        "success": True,
        "shape_name": shape.Name,
        "action_on": action_key,
    }


# ---------------------------------------------------------------------------
# MCP tool functions (async wrappers that delegate to COM thread)
# ---------------------------------------------------------------------------
def add_hyperlink(params: AddHyperlinkInput) -> str:
    """Add a hyperlink to a shape.

    Args:
        params: Hyperlink parameters including shape, address, and action trigger.

    Returns:
        JSON with shape name and hyperlink address.
    """
    try:
        result = ppt.execute(
            _add_hyperlink_impl,
            params.slide_index, params.shape_name_or_index,
            params.address, params.sub_address, params.screen_tip,
            params.action_on,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to add hyperlink: {str(e)}"})


def get_hyperlinks(params: GetHyperlinksInput) -> str:
    """Get all hyperlinks on a slide.

    Args:
        params: Slide index to list hyperlinks from.

    Returns:
        JSON with hyperlinks count and list of hyperlink details.
    """
    try:
        result = ppt.execute(
            _get_hyperlinks_impl,
            params.slide_index,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to get hyperlinks: {str(e)}"})


def remove_hyperlink(params: RemoveHyperlinkInput) -> str:
    """Remove a hyperlink from a shape.

    Args:
        params: Shape identifier and which action trigger to remove.

    Returns:
        JSON confirming the hyperlink removal.
    """
    try:
        result = ppt.execute(
            _remove_hyperlink_impl,
            params.slide_index, params.shape_name_or_index,
            params.action_on,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to remove hyperlink: {str(e)}"})


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register all hyperlink tools with the MCP server."""

    @mcp.tool(
        name="ppt_add_hyperlink",
        annotations={
            "title": "Add Hyperlink",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_add_hyperlink(params: AddHyperlinkInput) -> str:
        """Add a hyperlink to a shape.

        Supports URLs, file paths, mailto: links, and slide links (via sub_address).
        Set action_on to 'click' (default) or 'mouseover' for the trigger type.
        For slide links, use sub_address like '3,,' to link to slide 3.
        """
        return add_hyperlink(params)

    @mcp.tool(
        name="ppt_get_hyperlinks",
        annotations={
            "title": "Get Hyperlinks",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_get_hyperlinks(params: GetHyperlinksInput) -> str:
        """Get all hyperlinks on a slide.

        Returns address, sub-address, and type for each hyperlink found on the slide.
        """
        return get_hyperlinks(params)

    @mcp.tool(
        name="ppt_remove_hyperlink",
        annotations={
            "title": "Remove Hyperlink",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_remove_hyperlink(params: RemoveHyperlinkInput) -> str:
        """Remove a hyperlink from a shape.

        Clears the click or mouseover action on the specified shape.
        Set action_on to 'click' (default) or 'mouseover' to choose which to remove.
        """
        return remove_hyperlink(params)
