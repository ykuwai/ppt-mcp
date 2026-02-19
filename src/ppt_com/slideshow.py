"""SlideShow control tools for PowerPoint COM automation.

Start, stop, navigate, and query slide show state.
"""

import json
import logging
import time
from typing import Optional

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from ppt_com.constants import (
    msoTrue,
    msoFalse,
    ppShowTypeSpeaker,
    ppShowTypeWindow,
    ppShowTypeKiosk,
    ppShowSlideRange,
    ppSlideShowRunning,
    SLIDESHOW_STATE_NAMES,
    SHOW_TYPE_NAMES,
)

logger = logging.getLogger(__name__)

SHOW_TYPE_MAP = {
    "speaker": ppShowTypeSpeaker,
    "window": ppShowTypeWindow,
    "kiosk": ppShowTypeKiosk,
}


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class SlideShowStartInput(BaseModel):
    """Input for starting a slide show."""
    model_config = ConfigDict(str_strip_whitespace=True)

    start_slide: Optional[int] = Field(
        default=None,
        description="1-based starting slide index. Default: 1.",
    )
    end_slide: Optional[int] = Field(
        default=None,
        description="1-based ending slide index. Default: last slide.",
    )
    loop: Optional[bool] = Field(
        default=None,
        description="Loop the slide show continuously.",
    )
    show_type: Optional[str] = Field(
        default=None,
        description="Show type: 'speaker' (fullscreen), 'window', or 'kiosk'. Default: 'speaker'.",
    )


class SlideShowGotoInput(BaseModel):
    """Input for navigating to a specific slide in the slide show."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(
        ...,
        description="1-based slide index to navigate to.",
        ge=1,
    )


# ---------------------------------------------------------------------------
# Implementation functions (run on COM thread via ppt.execute)
# ---------------------------------------------------------------------------
def _slideshow_start_impl(
    start_slide: Optional[int],
    end_slide: Optional[int],
    loop: Optional[bool],
    show_type: Optional[str],
) -> dict:
    app = ppt._get_app_impl()
    if app.Presentations.Count == 0:
        raise RuntimeError(
            "No presentation is open. "
            "Use ppt_create_presentation or ppt_open_presentation first."
        )
    pres = ppt._get_pres_impl()

    if pres.Slides.Count == 0:
        raise RuntimeError("Presentation has no slides.")

    settings = pres.SlideShowSettings

    # Show type
    if show_type is not None:
        type_key = show_type.lower().strip()
        if type_key not in SHOW_TYPE_MAP:
            raise ValueError(
                f"Unknown show_type '{show_type}'. Supported: {list(SHOW_TYPE_MAP.keys())}"
            )
        settings.ShowType = SHOW_TYPE_MAP[type_key]
    else:
        settings.ShowType = ppShowTypeSpeaker

    # Slide range
    actual_start = start_slide if start_slide is not None else 1
    actual_end = end_slide if end_slide is not None else pres.Slides.Count

    if actual_start < 1 or actual_start > pres.Slides.Count:
        raise ValueError(
            f"start_slide {actual_start} out of range (1-{pres.Slides.Count})"
        )
    if actual_end < actual_start or actual_end > pres.Slides.Count:
        raise ValueError(
            f"end_slide {actual_end} out of range ({actual_start}-{pres.Slides.Count})"
        )

    settings.RangeType = ppShowSlideRange
    settings.StartingSlide = actual_start
    settings.EndingSlide = actual_end

    # Loop
    if loop is not None:
        settings.LoopUntilStopped = msoTrue if loop else msoFalse

    # Start the show
    ssw = settings.Run()
    time.sleep(0.5)

    view = ssw.View
    return {
        "success": True,
        "show_type": SHOW_TYPE_NAMES.get(settings.ShowType, "unknown"),
        "current_slide": view.CurrentShowPosition,
        "total_slides": pres.Slides.Count,
        "start_slide": actual_start,
        "end_slide": actual_end,
    }


def _slideshow_stop_impl() -> dict:
    app = ppt._get_app_impl()
    if app.SlideShowWindows.Count == 0:
        return {"success": True, "message": "No slide show was running."}

    app.SlideShowWindows(1).View.Exit()
    return {"success": True, "message": "Slide show ended."}


def _slideshow_next_impl() -> dict:
    app = ppt._get_app_impl()
    if app.SlideShowWindows.Count == 0:
        raise RuntimeError("No slide show is running.")

    view = app.SlideShowWindows(1).View
    view.Next()
    return {
        "success": True,
        "current_slide": view.CurrentShowPosition,
        "state": SLIDESHOW_STATE_NAMES.get(view.State, f"unknown({view.State})"),
    }


def _slideshow_previous_impl() -> dict:
    app = ppt._get_app_impl()
    if app.SlideShowWindows.Count == 0:
        raise RuntimeError("No slide show is running.")

    view = app.SlideShowWindows(1).View
    view.Previous()
    return {
        "success": True,
        "current_slide": view.CurrentShowPosition,
        "state": SLIDESHOW_STATE_NAMES.get(view.State, f"unknown({view.State})"),
    }


def _slideshow_goto_impl(slide_index: int) -> dict:
    app = ppt._get_app_impl()
    if app.SlideShowWindows.Count == 0:
        raise RuntimeError("No slide show is running.")

    view = app.SlideShowWindows(1).View
    view.GotoSlide(slide_index)
    return {
        "success": True,
        "current_slide": view.CurrentShowPosition,
        "state": SLIDESHOW_STATE_NAMES.get(view.State, f"unknown({view.State})"),
    }


def _slideshow_get_status_impl() -> dict:
    app = ppt._get_app_impl()

    if app.SlideShowWindows.Count == 0:
        return {"running": False}

    ssw = app.SlideShowWindows(1)
    view = ssw.View

    return {
        "running": True,
        "current_slide": view.CurrentShowPosition,
        "state": view.State,
        "state_name": SLIDESHOW_STATE_NAMES.get(view.State, f"unknown({view.State})"),
        "pointer_type": view.PointerType,
    }


# ---------------------------------------------------------------------------
# MCP tool functions (return JSON strings)
# ---------------------------------------------------------------------------
def slideshow_start(params: SlideShowStartInput) -> str:
    """Start a slide show presentation."""
    try:
        result = ppt.execute(
            _slideshow_start_impl,
            params.start_slide,
            params.end_slide,
            params.loop,
            params.show_type,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def slideshow_stop() -> str:
    """Stop the running slide show."""
    try:
        result = ppt.execute(_slideshow_stop_impl)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def slideshow_next() -> str:
    """Navigate to the next slide in the running slide show."""
    try:
        result = ppt.execute(_slideshow_next_impl)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def slideshow_previous() -> str:
    """Navigate to the previous slide in the running slide show."""
    try:
        result = ppt.execute(_slideshow_previous_impl)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def slideshow_goto(params: SlideShowGotoInput) -> str:
    """Go to a specific slide in the running slide show."""
    try:
        result = ppt.execute(_slideshow_goto_impl, params.slide_index)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def slideshow_get_status() -> str:
    """Get the current state of the running slide show."""
    try:
        result = ppt.execute(_slideshow_get_status_impl)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register all slideshow tools with the MCP server."""

    @mcp.tool(
        name="ppt_slideshow_start",
        annotations={
            "title": "Start Slide Show",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_slideshow_start(params: SlideShowStartInput) -> str:
        """Start a slide show presentation.

        Optionally configure start/end slides, looping, and show type
        ('speaker' for fullscreen, 'window' for windowed, 'kiosk' for kiosk mode).
        """
        return slideshow_start(params)

    @mcp.tool(
        name="ppt_slideshow_stop",
        annotations={
            "title": "Stop Slide Show",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_slideshow_stop() -> str:
        """End the currently running slide show.

        If no slide show is running, returns success with an informational message.
        """
        return slideshow_stop()

    @mcp.tool(
        name="ppt_slideshow_next",
        annotations={
            "title": "Slide Show Next",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_slideshow_next() -> str:
        """Advance to the next slide in the running slide show.

        Returns the current slide position and show state.
        """
        return slideshow_next()

    @mcp.tool(
        name="ppt_slideshow_previous",
        annotations={
            "title": "Slide Show Previous",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_slideshow_previous() -> str:
        """Go back to the previous slide in the running slide show.

        Returns the current slide position and show state.
        """
        return slideshow_previous()

    @mcp.tool(
        name="ppt_slideshow_goto",
        annotations={
            "title": "Slide Show Go To Slide",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_slideshow_goto(params: SlideShowGotoInput) -> str:
        """Navigate to a specific slide in the running slide show.

        Jumps directly to the slide at the given 1-based index.
        """
        return slideshow_goto(params)

    @mcp.tool(
        name="ppt_slideshow_get_status",
        annotations={
            "title": "Get Slide Show Status",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_slideshow_get_status() -> str:
        """Get the current state of the running slide show.

        Returns whether a show is running, current slide, state
        (running/paused/black_screen/white_screen/done), and pointer type.
        """
        return slideshow_get_status()
