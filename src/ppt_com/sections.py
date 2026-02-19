"""Section operations for PowerPoint COM automation.

Handles adding, listing, and managing presentation sections.
Sections group slides into logical units for organization.
"""

import json
import logging
from typing import Optional

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class AddSectionInput(BaseModel):
    """Input for adding a section to the presentation."""
    model_config = ConfigDict(str_strip_whitespace=True)

    name: str = Field(..., description="Section name")
    slide_index: int = Field(
        ..., ge=1, description="1-based index of the first slide in the section"
    )


class ManageSectionInput(BaseModel):
    """Input for managing (rename, move, delete) a section."""
    model_config = ConfigDict(str_strip_whitespace=True)

    section_index: int = Field(..., ge=1, description="1-based section index")
    action: str = Field(
        ..., description="Action to perform: 'rename', 'move', or 'delete'"
    )
    new_name: Optional[str] = Field(
        default=None, description="New name for the section (required for 'rename')"
    )
    move_to_index: Optional[int] = Field(
        default=None, ge=1,
        description="Target position to move the section to (required for 'move')",
    )


# ---------------------------------------------------------------------------
# COM implementation functions (run on COM thread via ppt.execute)
# ---------------------------------------------------------------------------
def _add_section_impl(name, slide_index):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    sp = pres.SectionProperties

    section_index = sp.AddSection(slide_index, name)

    return {
        "success": True,
        "section_index": section_index,
        "name": name,
        "slide_index": slide_index,
    }


def _list_sections_impl():
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    sp = pres.SectionProperties

    sections = []
    for i in range(1, sp.Count + 1):
        sections.append({
            "index": i,
            "name": sp.Name(i),
            "first_slide": sp.FirstSlide(i),
            "slides_count": sp.SlidesCount(i),
        })

    return {
        "success": True,
        "sections_count": sp.Count,
        "sections": sections,
    }


def _manage_section_impl(section_index, action, new_name, move_to_index):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    sp = pres.SectionProperties

    action_key = action.strip().lower()

    if action_key == "rename":
        if new_name is None:
            raise ValueError("new_name is required for 'rename' action")
        sp.Rename(section_index, new_name)
        return {
            "success": True,
            "action": "rename",
            "section_index": section_index,
            "new_name": new_name,
        }
    elif action_key == "move":
        if move_to_index is None:
            raise ValueError("move_to_index is required for 'move' action")
        sp.Move(section_index, move_to_index)
        return {
            "success": True,
            "action": "move",
            "section_index": section_index,
            "moved_to": move_to_index,
        }
    elif action_key == "delete":
        section_name = sp.Name(section_index)
        sp.Delete(section_index, False)  # False = don't delete slides
        return {
            "success": True,
            "action": "delete",
            "deleted_section": section_name,
        }
    else:
        raise ValueError(
            f"Unknown action '{action}'. Use: 'rename', 'move', or 'delete'"
        )


# ---------------------------------------------------------------------------
# MCP tool functions (async wrappers that delegate to COM thread)
# ---------------------------------------------------------------------------
def add_section(params: AddSectionInput) -> str:
    """Add a section to the presentation.

    Args:
        params: Section name and the first slide index.

    Returns:
        JSON with section index and name.
    """
    try:
        result = ppt.execute(
            _add_section_impl,
            params.name, params.slide_index,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to add section: {str(e)}"})


def list_sections() -> str:
    """List all sections in the active presentation.

    Returns:
        JSON with sections count and list of section details.
    """
    try:
        result = ppt.execute(_list_sections_impl)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to list sections: {str(e)}"})


def manage_section(params: ManageSectionInput) -> str:
    """Manage a section: rename, move, or delete.

    Args:
        params: Section index, action, and action-specific parameters.

    Returns:
        JSON confirming the action performed.
    """
    try:
        result = ppt.execute(
            _manage_section_impl,
            params.section_index, params.action,
            params.new_name, params.move_to_index,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to manage section: {str(e)}"})


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register all section tools with the MCP server."""

    @mcp.tool(
        name="ppt_add_section",
        annotations={
            "title": "Add Section",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_add_section(params: AddSectionInput) -> str:
        """Add a section to the presentation.

        Creates a new section starting at the specified slide.
        Sections group slides for organizational purposes.
        """
        return add_section(params)

    @mcp.tool(
        name="ppt_list_sections",
        annotations={
            "title": "List Sections",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_list_sections() -> str:
        """List all sections in the active presentation.

        Returns section name, first slide index, and slide count for each section.
        """
        return list_sections()

    @mcp.tool(
        name="ppt_manage_section",
        annotations={
            "title": "Manage Section",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_manage_section(params: ManageSectionInput) -> str:
        """Manage a section: rename, move, or delete.

        Actions:
        - 'rename': Change section name (requires new_name).
        - 'move': Move section to a new position (requires move_to_index).
        - 'delete': Remove the section without deleting its slides.
        """
        return manage_section(params)
