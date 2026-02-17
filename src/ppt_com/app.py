"""Application-level connection and info tools for PowerPoint COM automation."""

import json
import logging
from typing import Optional

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt, handle_com_error
from ppt_com.constants import WINDOW_STATE_NAMES, ppSelectionNone, ppSelectionSlides, ppSelectionShapes, ppSelectionText

logger = logging.getLogger(__name__)

SELECTION_TYPE_NAMES = {
    ppSelectionNone: "none",
    ppSelectionSlides: "slides",
    ppSelectionShapes: "shapes",
    ppSelectionText: "text",
}


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class ConnectInput(BaseModel):
    """Input for connecting to PowerPoint."""
    model_config = ConfigDict(str_strip_whitespace=True)

    visible: Optional[bool] = Field(
        default=None,
        description=(
            "If true, PowerPoint window is visible. If false, headless mode "
            "(PowerPoint runs in background). If null, keep current state."
        ),
    )


class SetWindowStateInput(BaseModel):
    """Input for setting PowerPoint window state."""
    model_config = ConfigDict(str_strip_whitespace=True)

    window_state: str = Field(
        default="maximized",
        description=(
            "Window state: 'normal' (restored), 'minimized', or 'maximized'"
        ),
    )


# ---------------------------------------------------------------------------
# Implementation functions (run on COM thread via ppt.execute)
# ---------------------------------------------------------------------------
def _connect_impl(visible: Optional[bool]) -> dict:
    app = ppt._connect_impl(visible)
    return {
        "success": True,
        "name": app.Name,
        "version": app.Version,
        "visible": bool(app.Visible),
        "presentations_count": app.Presentations.Count,
    }


def _get_app_info_impl() -> dict:
    app = ppt._get_app_impl()
    info = {
        "name": app.Name,
        "version": app.Version,
        "visible": bool(app.Visible),
        "window_state": WINDOW_STATE_NAMES.get(app.WindowState, "unknown"),
        "presentations_count": app.Presentations.Count,
        "active_presentation": None,
    }
    if app.Presentations.Count > 0:
        try:
            info["active_presentation"] = app.ActivePresentation.Name
        except Exception:
            pass
    return info


def _get_active_window_info_impl() -> dict:
    app = ppt._get_app_impl()
    if app.Windows.Count == 0:
        return {"error": "No windows are open in PowerPoint."}

    win = app.ActiveWindow
    result = {
        "caption": win.Caption,
        "view_type": win.ViewType,
        "active_slide_index": None,
        "selection_type": "none",
        "selected_shapes": [],
        "selected_text": None,
    }

    try:
        result["active_slide_index"] = win.View.Slide.SlideIndex
    except Exception:
        pass

    try:
        sel = win.Selection
        sel_type = sel.Type
        result["selection_type"] = SELECTION_TYPE_NAMES.get(sel_type, "unknown")

        if sel_type == ppSelectionShapes:
            shapes = []
            for i in range(1, sel.ShapeRange.Count + 1):
                s = sel.ShapeRange(i)
                shapes.append({"name": s.Name, "type": s.Type, "id": s.Id})
            result["selected_shapes"] = shapes

        elif sel_type == ppSelectionText:
            result["selected_text"] = sel.TextRange.Text
    except Exception:
        pass

    return result


def _list_presentations_impl() -> dict:
    app = ppt._get_app_impl()
    presentations = []
    for i in range(1, app.Presentations.Count + 1):
        p = app.Presentations(i)
        presentations.append({
            "index": i,
            "name": p.Name,
            "full_name": p.FullName,
            "path": p.Path,
            "slides_count": p.Slides.Count,
            "read_only": bool(p.ReadOnly),
            "saved": bool(p.Saved),
        })
    return {"presentations": presentations, "count": len(presentations)}


def _set_window_state_impl(window_state: str) -> dict:
    """Set the PowerPoint application window state."""
    STATE_MAP = {
        "normal": 1,    # ppWindowNormal
        "minimized": 2, # ppWindowMinimized
        "maximized": 3, # ppWindowMaximized
    }

    state_value = STATE_MAP.get(window_state.lower())
    if state_value is None:
        return {
            "error": f"Invalid window_state: {window_state}. Use 'normal', 'minimized', or 'maximized'."
        }

    app = ppt._get_app_impl()
    app.WindowState = state_value

    return {
        "success": True,
        "window_state": window_state.lower(),
    }


# ---------------------------------------------------------------------------
# MCP tool functions (async wrappers that delegate to COM thread)
# ---------------------------------------------------------------------------
def connect_to_powerpoint(params: ConnectInput) -> str:
    """Connect to a running PowerPoint instance or launch a new one.

    Attempts to connect to an already-running PowerPoint via GetActiveObject.
    If no instance is found, launches a new one via Dispatch.

    Set visible=false for headless mode where PowerPoint runs in the background
    without showing a window (useful for automated slide generation).

    Args:
        params (ConnectInput): Connection parameters:
            - visible (Optional[bool]): Window visibility. true=visible, false=headless, null=keep current

    Returns:
        str: JSON with connection status, PowerPoint version, and presentation count
    """
    try:
        result = ppt.execute(_connect_impl, params.visible)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to connect to PowerPoint: {str(e)}"})


def get_app_info() -> str:
    """Get information about the connected PowerPoint application.

    Returns the PowerPoint version, visibility, window state, number of
    open presentations, and the name of the active presentation.

    Returns:
        str: JSON with application info including version, window state, and active presentation
    """
    try:
        result = ppt.execute(_get_app_info_impl)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to get app info: {str(e)}"})


def get_active_window_info() -> str:
    """Get detailed info about the active PowerPoint window and current selection.

    Returns the window caption, view type, current slide index, and details
    about what is currently selected (nothing, slides, shapes, or text).

    Returns:
        str: JSON with window caption, active slide index, selection type, and selected items
    """
    try:
        result = ppt.execute(_get_active_window_info_impl)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to get window info: {str(e)}"})


def list_presentations() -> str:
    """List all currently open presentations in PowerPoint.

    Returns name, path, slide count, read-only status, and save status
    for each open presentation.

    Returns:
        str: JSON array of presentation objects with their properties
    """
    try:
        result = ppt.execute(_list_presentations_impl)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to list presentations: {str(e)}"})


def set_window_state(params: SetWindowStateInput) -> str:
    """Set the PowerPoint application window state.

    Controls whether the PowerPoint window is maximized, minimized, or
    restored to normal size. This affects the main PowerPoint application
    window, not individual presentation windows.

    Args:
        params (SetWindowStateInput): Window state parameters:
            - window_state (str): Target state - 'normal', 'minimized', or 'maximized'

    Returns:
        str: JSON with success status and the applied window state
    """
    try:
        result = ppt.execute(_set_window_state_impl, params.window_state)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to set window state: {str(e)}"})
