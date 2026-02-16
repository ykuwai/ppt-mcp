"""Slide-level operations for PowerPoint COM automation.

Add, delete, duplicate, move, list, and query slides. Manage speaker notes.
"""

import json
import logging
from typing import Optional

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from ppt_com.constants import ppLayoutBlank

logger = logging.getLogger(__name__)

# Friendly layout name -> PpSlideLayout constant
LAYOUT_NAME_MAP = {
    "title": 1,
    "text": 2,
    "two_column_text": 3,
    "table": 4,
    "title_only": 11,
    "blank": 12,
    "section_header": 33,
    "comparison": 34,
    "content_with_caption": 35,
    "picture_with_caption": 36,
}


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class AddSlideInput(BaseModel):
    """Input for adding a new slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    position: Optional[int] = Field(
        default=None,
        description=(
            "1-based position for the new slide. "
            "If omitted, the slide is added at the end."
        ),
    )
    layout: Optional[int] = Field(
        default=None,
        description=(
            "PpSlideLayout constant: 1=Title, 2=Text, 11=TitleOnly, 12=Blank, "
            "33=SectionHeader, 34=Comparison, 35=ContentWithCaption, "
            "36=PictureWithCaption. Ignored if layout_name is provided."
        ),
    )
    layout_name: Optional[str] = Field(
        default=None,
        description=(
            "Layout name to look up from the slide master's custom layouts. "
            "Alternatively use a friendly name: 'title', 'text', 'blank', "
            "'title_only', 'section_header', 'comparison', "
            "'content_with_caption', 'picture_with_caption'. "
            "Overrides the layout parameter."
        ),
    )


class DeleteSlideInput(BaseModel):
    """Input for deleting a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(
        ...,
        description="1-based index of the slide to delete.",
    )


class DuplicateSlideInput(BaseModel):
    """Input for duplicating a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(
        ...,
        description="1-based index of the slide to duplicate.",
    )


class MoveSlideInput(BaseModel):
    """Input for moving a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(
        ...,
        description="Current 1-based index of the slide to move.",
    )
    new_position: int = Field(
        ...,
        description="Target 1-based position to move the slide to.",
    )


class ListSlidesInput(BaseModel):
    """Input for listing slides (no required params)."""
    model_config = ConfigDict(str_strip_whitespace=True)

    presentation_index: Optional[int] = Field(
        default=None,
        description=(
            "1-based index of the presentation to list slides from. "
            "If omitted, uses the active presentation."
        ),
    )


class GetSlideInfoInput(BaseModel):
    """Input for getting detailed slide info."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(
        ...,
        description="1-based index of the slide to query.",
    )


class SetSlideNotesInput(BaseModel):
    """Input for setting speaker notes."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(
        ...,
        description="1-based index of the slide.",
    )
    notes_text: str = Field(
        ...,
        description="The speaker notes text to set.",
    )


class GetSlideNotesInput(BaseModel):
    """Input for getting speaker notes."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(
        ...,
        description="1-based index of the slide.",
    )


# ---------------------------------------------------------------------------
# Helper to resolve a presentation
# ---------------------------------------------------------------------------
def _resolve_presentation(app, presentation_index: Optional[int] = None):
    """Return a Presentation COM object by index, or ActivePresentation if None."""
    if presentation_index is not None:
        count = app.Presentations.Count
        if presentation_index < 1 or presentation_index > count:
            raise ValueError(
                f"Presentation index {presentation_index} out of range (1-{count})"
            )
        return app.Presentations(presentation_index)
    if app.Presentations.Count == 0:
        raise RuntimeError(
            "No presentation is open. "
            "Use ppt_create_presentation or ppt_open_presentation first."
        )
    return app.ActivePresentation


# ---------------------------------------------------------------------------
# Implementation functions (run on COM thread via ppt.execute)
# ---------------------------------------------------------------------------
def _add_slide_impl(
    position: Optional[int],
    layout: Optional[int],
    layout_name: Optional[str],
) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(app)

    if position is None:
        position = pres.Slides.Count + 1

    if position < 1 or position > pres.Slides.Count + 1:
        raise ValueError(
            f"Position {position} out of range (1-{pres.Slides.Count + 1})"
        )

    if layout_name:
        # Check friendly name map first
        friendly_key = layout_name.lower().strip().replace(" ", "_")
        if friendly_key in LAYOUT_NAME_MAP:
            slide = pres.Slides.Add(Index=position, Layout=LAYOUT_NAME_MAP[friendly_key])
        else:
            # Search custom layouts by exact name
            master = pres.SlideMaster
            custom_layout = None
            for i in range(1, master.CustomLayouts.Count + 1):
                if master.CustomLayouts(i).Name == layout_name:
                    custom_layout = master.CustomLayouts(i)
                    break
            if custom_layout is None:
                available = []
                for i in range(1, master.CustomLayouts.Count + 1):
                    available.append(master.CustomLayouts(i).Name)
                raise ValueError(
                    f"Layout '{layout_name}' not found. "
                    f"Available custom layouts: {available}"
                )
            slide = pres.Slides.AddSlide(Index=position, pCustomLayout=custom_layout)
    else:
        layout_val = layout if layout is not None else ppLayoutBlank
        slide = pres.Slides.Add(Index=position, Layout=layout_val)

    return {
        "success": True,
        "slide_index": slide.SlideIndex,
        "slide_id": slide.SlideID,
        "layout": slide.Layout,
    }


def _delete_slide_impl(slide_index: int) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(app)

    if slide_index < 1 or slide_index > pres.Slides.Count:
        raise ValueError(
            f"Slide index {slide_index} out of range (1-{pres.Slides.Count})"
        )

    pres.Slides(slide_index).Delete()
    return {
        "success": True,
        "deleted_index": slide_index,
        "remaining_count": pres.Slides.Count,
    }


def _duplicate_slide_impl(slide_index: int) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(app)

    if slide_index < 1 or slide_index > pres.Slides.Count:
        raise ValueError(
            f"Slide index {slide_index} out of range (1-{pres.Slides.Count})"
        )

    dup_range = pres.Slides(slide_index).Duplicate()
    new_slide = dup_range(1)
    return {
        "success": True,
        "new_slide_index": new_slide.SlideIndex,
        "new_slide_id": new_slide.SlideID,
    }


def _move_slide_impl(slide_index: int, new_position: int) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(app)

    count = pres.Slides.Count
    if slide_index < 1 or slide_index > count:
        raise ValueError(
            f"Slide index {slide_index} out of range (1-{count})"
        )
    if new_position < 1 or new_position > count:
        raise ValueError(
            f"New position {new_position} out of range (1-{count})"
        )

    pres.Slides(slide_index).MoveTo(toPos=new_position)
    return {
        "success": True,
        "moved_from": slide_index,
        "moved_to": new_position,
    }


def _list_slides_impl(presentation_index: Optional[int]) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(app, presentation_index)

    slides = []
    for i in range(1, pres.Slides.Count + 1):
        slide = pres.Slides(i)

        layout_name = ""
        try:
            layout_name = slide.CustomLayout.Name
        except Exception:
            pass

        has_notes = False
        try:
            notes_text = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
            has_notes = len(notes_text.strip()) > 0
        except Exception:
            pass

        slides.append({
            "index": slide.SlideIndex,
            "slide_id": slide.SlideID,
            "name": slide.Name,
            "layout": slide.Layout,
            "layout_name": layout_name,
            "hidden": bool(slide.SlideShowTransition.Hidden),
            "shapes_count": slide.Shapes.Count,
            "has_notes": has_notes,
        })

    return {"slides_count": pres.Slides.Count, "slides": slides}


def _get_slide_info_impl(slide_index: int) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(app)

    if slide_index < 1 or slide_index > pres.Slides.Count:
        raise ValueError(
            f"Slide index {slide_index} out of range (1-{pres.Slides.Count})"
        )

    slide = pres.Slides(slide_index)
    trans = slide.SlideShowTransition

    layout_name = ""
    try:
        layout_name = slide.CustomLayout.Name
    except Exception:
        pass

    title_text = ""
    has_title = bool(slide.Shapes.HasTitle)
    if has_title:
        try:
            title_text = slide.Shapes.Title.TextFrame.TextRange.Text
        except Exception:
            pass

    notes_text = ""
    try:
        notes_text = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
    except Exception:
        pass

    design_name = ""
    try:
        design_name = slide.Design.Name
    except Exception:
        pass

    return {
        "index": slide.SlideIndex,
        "slide_id": slide.SlideID,
        "slide_number": slide.SlideNumber,
        "name": slide.Name,
        "layout": slide.Layout,
        "layout_name": layout_name,
        "hidden": bool(trans.Hidden),
        "shapes_count": slide.Shapes.Count,
        "has_title": has_title,
        "title_text": title_text,
        "notes_text": notes_text,
        "follow_master_background": bool(slide.FollowMasterBackground),
        "transition_effect": trans.EntryEffect,
        "advance_on_click": bool(trans.AdvanceOnClick),
        "advance_on_time": bool(trans.AdvanceOnTime),
        "advance_time": trans.AdvanceTime,
        "design_name": design_name,
    }


def _set_slide_notes_impl(slide_index: int, notes_text: str) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(app)

    if slide_index < 1 or slide_index > pres.Slides.Count:
        raise ValueError(
            f"Slide index {slide_index} out of range (1-{pres.Slides.Count})"
        )

    slide = pres.Slides(slide_index)
    slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = notes_text
    return {"success": True}


def _get_slide_notes_impl(slide_index: int) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(app)

    if slide_index < 1 or slide_index > pres.Slides.Count:
        raise ValueError(
            f"Slide index {slide_index} out of range (1-{pres.Slides.Count})"
        )

    slide = pres.Slides(slide_index)
    try:
        notes_text = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
    except Exception:
        notes_text = ""

    return {"slide_index": slide_index, "notes_text": notes_text}


# ---------------------------------------------------------------------------
# MCP tool functions (return JSON strings)
# ---------------------------------------------------------------------------
def add_slide(params: AddSlideInput) -> str:
    """Add a new slide to the active presentation."""
    try:
        result = ppt.execute(
            _add_slide_impl, params.position, params.layout, params.layout_name
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def delete_slide(params: DeleteSlideInput) -> str:
    """Delete a slide by index."""
    try:
        result = ppt.execute(_delete_slide_impl, params.slide_index)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def duplicate_slide(params: DuplicateSlideInput) -> str:
    """Duplicate a slide."""
    try:
        result = ppt.execute(_duplicate_slide_impl, params.slide_index)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def move_slide(params: MoveSlideInput) -> str:
    """Move a slide to a new position."""
    try:
        result = ppt.execute(
            _move_slide_impl, params.slide_index, params.new_position
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def list_slides(params: ListSlidesInput) -> str:
    """List all slides in the active presentation."""
    try:
        result = ppt.execute(_list_slides_impl, params.presentation_index)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def get_slide_info(params: GetSlideInfoInput) -> str:
    """Get detailed info about a specific slide."""
    try:
        result = ppt.execute(_get_slide_info_impl, params.slide_index)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def set_slide_notes(params: SetSlideNotesInput) -> str:
    """Set speaker notes for a slide."""
    try:
        result = ppt.execute(
            _set_slide_notes_impl, params.slide_index, params.notes_text
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def get_slide_notes(params: GetSlideNotesInput) -> str:
    """Get speaker notes for a slide."""
    try:
        result = ppt.execute(_get_slide_notes_impl, params.slide_index)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


class GotoSlideInput(BaseModel):
    """Input for navigating to a specific slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(
        ...,
        description="1-based index of the slide to navigate to.",
        ge=1,
    )


def _goto_slide_impl(slide_index: int) -> dict:
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    if slide_index < 1 or slide_index > pres.Slides.Count:
        raise ValueError(
            f"Slide index {slide_index} out of range (1-{pres.Slides.Count})"
        )
    app.ActiveWindow.View.GotoSlide(slide_index)
    return {
        "success": True,
        "active_slide_index": slide_index,
    }


def goto_slide(params: GotoSlideInput) -> str:
    """Navigate the active window to display a specific slide."""
    try:
        result = ppt.execute(_goto_slide_impl, params.slide_index)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register all slide tools with the MCP server."""

    @mcp.tool(
        name="ppt_add_slide",
        annotations={
            "title": "Add Slide",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_add_slide(params: AddSlideInput) -> str:
        """Add a new slide to the active presentation.

        Specify a layout by name (e.g. 'blank', 'title', 'title_only') or
        by PpSlideLayout integer constant. You can also provide a custom
        layout_name to match a layout from the slide master.
        Position is 1-based; omit to append at the end.
        """
        return add_slide(params)

    @mcp.tool(
        name="ppt_delete_slide",
        annotations={
            "title": "Delete Slide",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_delete_slide(params: DeleteSlideInput) -> str:
        """Delete a slide by its 1-based index.

        Remaining slides re-index automatically after deletion.
        When deleting multiple slides, delete from highest index first
        to avoid index shifting.
        """
        return delete_slide(params)

    @mcp.tool(
        name="ppt_duplicate_slide",
        annotations={
            "title": "Duplicate Slide",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_duplicate_slide(params: DuplicateSlideInput) -> str:
        """Duplicate a slide. The copy is inserted immediately after the original.

        Returns the new slide's index and ID.
        """
        return duplicate_slide(params)

    @mcp.tool(
        name="ppt_move_slide",
        annotations={
            "title": "Move Slide",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_move_slide(params: MoveSlideInput) -> str:
        """Move a slide to a new position within the presentation.

        Both slide_index and new_position are 1-based.
        """
        return move_slide(params)

    @mcp.tool(
        name="ppt_list_slides",
        annotations={
            "title": "List Slides",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_list_slides(params: ListSlidesInput) -> str:
        """List all slides in the active or specified presentation.

        Returns each slide's index, ID, name, layout, hidden status,
        shape count, and whether it has speaker notes.
        """
        return list_slides(params)

    @mcp.tool(
        name="ppt_get_slide_info",
        annotations={
            "title": "Get Slide Info",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_get_slide_info(params: GetSlideInfoInput) -> str:
        """Get detailed information about a specific slide.

        Returns layout, shapes count, title text, speaker notes,
        transition settings, background info, and design name.
        """
        return get_slide_info(params)

    @mcp.tool(
        name="ppt_set_slide_notes",
        annotations={
            "title": "Set Slide Notes",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_set_slide_notes(params: SetSlideNotesInput) -> str:
        """Set the speaker notes text for a slide.

        Replaces any existing notes with the provided text.
        """
        return set_slide_notes(params)

    @mcp.tool(
        name="ppt_get_slide_notes",
        annotations={
            "title": "Get Slide Notes",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_get_slide_notes(params: GetSlideNotesInput) -> str:
        """Get the speaker notes text for a slide.

        Returns the notes text. If no notes exist, returns an empty string.
        """
        return get_slide_notes(params)

    @mcp.tool(
        name="ppt_goto_slide",
        annotations={
            "title": "Go To Slide",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_goto_slide(params: GotoSlideInput) -> str:
        """Navigate the active window to display a specific slide.

        Changes which slide is shown in the PowerPoint editor.
        Useful for jumping to a slide you want to view or edit.
        """
        return goto_slide(params)
