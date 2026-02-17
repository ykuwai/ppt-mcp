"""Presentation-level operations for PowerPoint COM automation.

Create, open, save, close, and query PowerPoint presentations.
"""

import glob as glob_mod
import json
import logging
import os
import winreg
from typing import Optional

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from ppt_com.constants import (
    msoTrue,
    msoFalse,
    ppSaveAsOpenXMLPresentation,
    ppSaveAsPDF,
    ppSaveAsPNG,
    ppSaveAsJPG,
    ppSaveAsDefault,
)
from utils.units import (
    SLIDE_WIDTH_16_9,
    SLIDE_HEIGHT_16_9,
    SLIDE_WIDTH_4_3,
    SLIDE_HEIGHT_4_3,
)

logger = logging.getLogger(__name__)

# Mapping of format aliases to PpSaveAsFileType constants
SAVE_FORMAT_MAP = {
    "pptx": ppSaveAsOpenXMLPresentation,
    "pdf": ppSaveAsPDF,
    "png": ppSaveAsPNG,
    "jpg": ppSaveAsJPG,
    "default": ppSaveAsDefault,
}

# Mapping of preset names to (width, height) in points
SLIDE_SIZE_PRESETS = {
    "16:9": (SLIDE_WIDTH_16_9, SLIDE_HEIGHT_16_9),
    "4:3": (SLIDE_WIDTH_4_3, SLIDE_HEIGHT_4_3),
}


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class CreatePresentationInput(BaseModel):
    """Input for creating a new presentation."""
    model_config = ConfigDict(str_strip_whitespace=True)

    template_path: Optional[str] = Field(
        default=None,
        description=(
            "Absolute path to a PowerPoint template file (.potx, .potm, .pot). "
            "When provided, creates a new presentation based on this template, "
            "preserving all slides, layouts, themes, and custom data. "
            "Size parameters (preset, slide_width, slide_height) are ignored "
            "when a template is used."
        ),
    )
    slide_width: Optional[float] = Field(
        default=None,
        description=(
            "Slide width in points (72 points = 1 inch). "
            "Ignored if preset or template_path is provided."
        ),
    )
    slide_height: Optional[float] = Field(
        default=None,
        description=(
            "Slide height in points (72 points = 1 inch). "
            "Ignored if preset or template_path is provided."
        ),
    )
    preset: Optional[str] = Field(
        default=None,
        description=(
            "Slide size preset. Supported values: '16:9' (widescreen, 960x540 pt), "
            "'4:3' (standard, 720x540 pt). Overrides slide_width/slide_height. "
            "Ignored if template_path is provided."
        ),
    )


class OpenPresentationInput(BaseModel):
    """Input for opening an existing presentation."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_path: str = Field(
        ...,
        description="Full path to the presentation file (.pptx, .pptm, .ppt, .potx, etc.).",
    )
    read_only: bool = Field(
        default=False,
        description="If true, open in read-only mode.",
    )
    with_window: bool = Field(
        default=True,
        description="If true, open with a visible window. Set false for background processing.",
    )


class SavePresentationInput(BaseModel):
    """Input for saving the active presentation."""
    model_config = ConfigDict(str_strip_whitespace=True)

    presentation_index: Optional[int] = Field(
        default=None,
        description=(
            "1-based index of the presentation to save. "
            "If omitted, saves the active presentation."
        ),
    )
    presentation_name: Optional[str] = Field(
        default=None,
        description=(
            "Name of the presentation (e.g. 'MySlides.pptx'). "
            "Alternative to presentation_index. If omitted, uses the active presentation."
        ),
    )


class SavePresentationAsInput(BaseModel):
    """Input for SaveAs operation."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_path: str = Field(
        ...,
        description="Target file path for saving.",
    )
    format: Optional[str] = Field(
        default=None,
        description=(
            "Output format: 'pptx', 'pdf', 'png', 'jpg', or 'default'. "
            "If omitted, PowerPoint infers from file extension."
        ),
    )
    presentation_index: Optional[int] = Field(
        default=None,
        description=(
            "1-based index of the presentation to save. "
            "If omitted, saves the active presentation."
        ),
    )
    presentation_name: Optional[str] = Field(
        default=None,
        description=(
            "Name of the presentation (e.g. 'MySlides.pptx'). "
            "Alternative to presentation_index. If omitted, uses the active presentation."
        ),
    )


class ClosePresentationInput(BaseModel):
    """Input for closing a presentation."""
    model_config = ConfigDict(str_strip_whitespace=True)

    save_changes: bool = Field(
        default=False,
        description="If true, save the presentation before closing. If false, discard unsaved changes.",
    )
    presentation_index: Optional[int] = Field(
        default=None,
        description=(
            "1-based index of the presentation to close. "
            "If omitted, closes the active presentation."
        ),
    )
    presentation_name: Optional[str] = Field(
        default=None,
        description=(
            "Name of the presentation (e.g. 'MySlides.pptx'). "
            "Alternative to presentation_index. If omitted, uses the active presentation."
        ),
    )


class GetPresentationInfoInput(BaseModel):
    """Input for getting presentation info."""
    model_config = ConfigDict(str_strip_whitespace=True)

    presentation_index: Optional[int] = Field(
        default=None,
        description=(
            "1-based index of the presentation to query. "
            "If omitted, uses the active presentation."
        ),
    )
    presentation_name: Optional[str] = Field(
        default=None,
        description=(
            "Name of the presentation (e.g. 'MySlides.pptx'). "
            "Alternative to presentation_index. If omitted, uses the active presentation."
        ),
    )


class ListTemplatesInput(BaseModel):
    """Input for listing available PowerPoint templates."""
    model_config = ConfigDict(str_strip_whitespace=True)

    templates_dir: Optional[str] = Field(
        default=None,
        description=(
            "Directory to scan for template files (.potx, .potm). "
            "If omitted, auto-detects the personal templates folder."
        ),
    )


# ---------------------------------------------------------------------------
# Helper to resolve a presentation by index or active
# ---------------------------------------------------------------------------
def _resolve_presentation(
    app,
    presentation_index: Optional[int] = None,
    presentation_name: Optional[str] = None,
):
    """Return a Presentation COM object by index, name, or ActivePresentation.

    Args:
        app: PowerPoint Application COM object.
        presentation_index: 1-based index of the presentation.
        presentation_name: File name of the presentation (e.g. 'MySlides.pptx').

    Raises:
        ValueError: If both parameters are provided, or if the name matches
            zero or multiple presentations.
        RuntimeError: If no presentations are open.
    """
    if presentation_index is not None and presentation_name is not None:
        raise ValueError(
            "Specify either presentation_index or presentation_name, not both"
        )

    if presentation_index is not None:
        count = app.Presentations.Count
        if presentation_index < 1 or presentation_index > count:
            raise ValueError(
                f"Presentation index {presentation_index} out of range (1-{count})"
            )
        return app.Presentations(presentation_index)

    if presentation_name is not None:
        count = app.Presentations.Count
        if count == 0:
            raise RuntimeError(
                "No presentation is open. "
                "Use ppt_create_presentation or ppt_open_presentation first."
            )
        matches = []
        available = []
        for i in range(1, count + 1):
            pres = app.Presentations(i)
            name = pres.Name
            available.append(f"  [{i}] {name}")
            if name == presentation_name:
                matches.append((i, pres))
        if len(matches) == 1:
            return matches[0][1]
        if len(matches) > 1:
            match_list = ", ".join(
                f"[{idx}] {presentation_name}" for idx, _ in matches
            )
            raise ValueError(
                f"Multiple presentations match name '{presentation_name}': "
                f"{match_list}. Use presentation_index to disambiguate."
            )
        raise ValueError(
            f"No presentation named '{presentation_name}'. "
            f"Available presentations:\n" + "\n".join(available)
        )

    if app.Presentations.Count == 0:
        raise RuntimeError(
            "No presentation is open. "
            "Use ppt_create_presentation or ppt_open_presentation first."
        )
    return app.ActivePresentation


# ---------------------------------------------------------------------------
# Implementation functions (run on COM thread via ppt.execute)
# ---------------------------------------------------------------------------
def _create_presentation_impl(
    template_path: Optional[str],
    slide_width: Optional[float],
    slide_height: Optional[float],
    preset: Optional[str],
) -> dict:
    app = ppt._get_app_impl()

    if template_path:
        # Create from template using Open + Untitled=msoTrue
        abs_path = os.path.abspath(template_path)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"Template not found: {abs_path}")

        # Positional args: FileName, ReadOnly, Untitled, WithWindow
        pres = app.Presentations.Open(abs_path, 0, -1, -1)
    else:
        # Create blank presentation (existing behavior)
        pres = app.Presentations.Add()

        # Apply preset if specified
        if preset:
            preset_key = preset.strip()
            if preset_key not in SLIDE_SIZE_PRESETS:
                raise ValueError(
                    f"Unknown preset '{preset}'. Supported: {list(SLIDE_SIZE_PRESETS.keys())}"
                )
            w, h = SLIDE_SIZE_PRESETS[preset_key]
            pres.PageSetup.SlideWidth = w
            pres.PageSetup.SlideHeight = h
        elif slide_width is not None and slide_height is not None:
            pres.PageSetup.SlideWidth = slide_width
            pres.PageSetup.SlideHeight = slide_height

    template_name = ""
    try:
        template_name = pres.TemplateName
    except Exception:
        pass

    # Find the 1-based index of the newly created presentation
    pres_index = None
    for i in range(1, app.Presentations.Count + 1):
        if app.Presentations(i).Name == pres.Name:
            pres_index = i
            break

    return {
        "success": True,
        "presentation_index": pres_index,
        "name": pres.Name,
        "slides_count": pres.Slides.Count,
        "slide_width": pres.PageSetup.SlideWidth,
        "slide_height": pres.PageSetup.SlideHeight,
        "template_name": template_name,
    }


def _open_presentation_impl(
    file_path: str,
    read_only: bool,
    with_window: bool,
) -> dict:
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    app = ppt._get_app_impl()
    pres = app.Presentations.Open(
        FileName=file_path,
        ReadOnly=msoTrue if read_only else msoFalse,
        Untitled=msoFalse,
        WithWindow=msoTrue if with_window else msoFalse,
    )

    # Find the 1-based index of the opened presentation
    pres_index = None
    for i in range(1, app.Presentations.Count + 1):
        if app.Presentations(i).Name == pres.Name:
            pres_index = i
            break

    return {
        "success": True,
        "presentation_index": pres_index,
        "name": pres.Name,
        "full_name": pres.FullName,
        "slides_count": pres.Slides.Count,
        "read_only": bool(pres.ReadOnly),
    }


def _save_presentation_impl(
    presentation_index: Optional[int],
    presentation_name: Optional[str],
) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(
        app, presentation_index=presentation_index, presentation_name=presentation_name
    )
    pres.Save()
    return {
        "success": True,
        "name": pres.Name,
        "saved": bool(pres.Saved),
    }


def _save_presentation_as_impl(
    file_path: str,
    format: Optional[str],
    presentation_index: Optional[int],
    presentation_name: Optional[str],
) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(
        app, presentation_index=presentation_index, presentation_name=presentation_name
    )

    kwargs = {"FileName": file_path}
    if format:
        fmt_key = format.lower().strip()
        if fmt_key not in SAVE_FORMAT_MAP:
            raise ValueError(
                f"Unknown format '{format}'. Supported: {list(SAVE_FORMAT_MAP.keys())}"
            )
        kwargs["FileFormat"] = SAVE_FORMAT_MAP[fmt_key]

    pres.SaveAs(**kwargs)
    return {
        "success": True,
        "name": pres.Name,
        "full_name": pres.FullName,
    }


def _close_presentation_impl(
    save_changes: bool,
    presentation_index: Optional[int],
    presentation_name: Optional[str],
) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(
        app, presentation_index=presentation_index, presentation_name=presentation_name
    )
    name = pres.Name

    if save_changes:
        pres.Save()
    else:
        # Suppress "save changes?" dialog
        pres.Saved = True

    pres.Close()
    return {"success": True, "closed": name}


def _get_presentation_info_impl(
    presentation_index: Optional[int],
    presentation_name: Optional[str],
) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(
        app, presentation_index=presentation_index, presentation_name=presentation_name
    )
    page = pres.PageSetup

    template_name = ""
    try:
        template_name = pres.TemplateName
    except Exception:
        pass

    return {
        "name": pres.Name,
        "full_name": pres.FullName,
        "path": pres.Path,
        "slides_count": pres.Slides.Count,
        "read_only": bool(pres.ReadOnly),
        "saved": bool(pres.Saved),
        "slide_width": page.SlideWidth,
        "slide_height": page.SlideHeight,
        "slide_width_inches": round(page.SlideWidth / 72.0, 3),
        "slide_height_inches": round(page.SlideHeight / 72.0, 3),
        "first_slide_number": page.FirstSlideNumber,
        "template_name": template_name,
    }


def _get_default_templates_dir() -> Optional[str]:
    """Resolve the default personal templates folder.

    Priority:
    1. Registry: HKCU\\Software\\Microsoft\\Office\\16.0\\PowerPoint\\Options → PersonalTemplates
    2. Fallback: check common paths and return the first that exists
    3. Return None if no directory is found
    """
    # 1. Try registry
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Office\16.0\PowerPoint\Options",
        )
        path, _ = winreg.QueryValueEx(key, "PersonalTemplates")
        winreg.CloseKey(key)
        if path and os.path.isdir(path):
            return path
    except (FileNotFoundError, OSError):
        pass

    # 2. Fallback: check common paths in order
    user_profile = os.environ.get("USERPROFILE", "")
    candidates = [
        os.path.join(user_profile, "OneDrive", "ドキュメント", "Office のカスタム テンプレート"),
        os.path.join(user_profile, "Documents", "Custom Office Templates"),
        os.path.join(user_profile, "Documents", "Office のカスタム テンプレート"),
    ]
    for candidate in candidates:
        if os.path.isdir(candidate):
            return candidate

    # 3. None found
    return None


def _list_templates_impl(templates_dir: Optional[str]) -> dict:
    """List PowerPoint template files (.potx, .potm) in a directory.

    This function does NOT require COM access.
    """
    if templates_dir is None:
        templates_dir = _get_default_templates_dir()

    if templates_dir is None:
        return {
            "templates_dir": None,
            "count": 0,
            "templates": [],
            "error": "Could not find a templates directory. Specify templates_dir explicitly.",
        }

    if not os.path.isdir(templates_dir):
        return {
            "templates_dir": templates_dir,
            "count": 0,
            "templates": [],
            "error": f"Directory not found: {templates_dir}",
        }

    templates = []
    for ext in ("*.potx", "*.potm"):
        pattern = os.path.join(templates_dir, ext)
        for filepath in glob_mod.glob(pattern):
            templates.append({
                "name": os.path.splitext(os.path.basename(filepath))[0],
                "file_name": os.path.basename(filepath),
                "file_path": os.path.abspath(filepath),
            })

    templates.sort(key=lambda t: t["name"])
    return {
        "templates_dir": templates_dir,
        "count": len(templates),
        "templates": templates,
    }


# ---------------------------------------------------------------------------
# MCP tool functions (return JSON strings)
# ---------------------------------------------------------------------------
def create_presentation(params: CreatePresentationInput) -> str:
    """Create a new presentation, optionally from a template."""
    try:
        result = ppt.execute(
            _create_presentation_impl,
            params.template_path,
            params.slide_width,
            params.slide_height,
            params.preset,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def open_presentation(params: OpenPresentationInput) -> str:
    """Open an existing presentation file."""
    try:
        result = ppt.execute(
            _open_presentation_impl,
            params.file_path,
            params.read_only,
            params.with_window,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def save_presentation(params: SavePresentationInput) -> str:
    """Save the active or specified presentation."""
    try:
        result = ppt.execute(
            _save_presentation_impl,
            params.presentation_index,
            params.presentation_name,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def save_presentation_as(params: SavePresentationAsInput) -> str:
    """Save a presentation with a new name and/or format."""
    try:
        result = ppt.execute(
            _save_presentation_as_impl,
            params.file_path,
            params.format,
            params.presentation_index,
            params.presentation_name,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def close_presentation(params: ClosePresentationInput) -> str:
    """Close a presentation, optionally saving first."""
    try:
        result = ppt.execute(
            _close_presentation_impl,
            params.save_changes,
            params.presentation_index,
            params.presentation_name,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def get_presentation_info(params: GetPresentationInfoInput) -> str:
    """Get detailed info about a presentation."""
    try:
        result = ppt.execute(
            _get_presentation_info_impl,
            params.presentation_index,
            params.presentation_name,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def list_templates(params: ListTemplatesInput) -> str:
    """List available PowerPoint template files."""
    try:
        result = _list_templates_impl(params.templates_dir)
        return json.dumps(result, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register all presentation tools with the MCP server."""

    @mcp.tool(
        name="ppt_create_presentation",
        annotations={
            "title": "Create Presentation",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_create_presentation(params: CreatePresentationInput) -> str:
        """Create a new PowerPoint presentation.

        Optionally create from a template file (.potx, .potm) by providing
        template_path. This preserves all slides, layouts, themes, and
        custom data from the template (like double-clicking a .potx file).

        For blank presentations, set slide dimensions via a preset ('16:9'
        or '4:3') or explicit width/height in points (72 pt = 1 inch).
        Size parameters are ignored when a template is used.
        """
        return create_presentation(params)

    @mcp.tool(
        name="ppt_open_presentation",
        annotations={
            "title": "Open Presentation",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": True,
        },
    )
    async def tool_open_presentation(params: OpenPresentationInput) -> str:
        """Open an existing PowerPoint file.

        Supports .pptx, .pptm, .ppt, .potx, .ppsx and other PowerPoint formats.
        Set read_only=true to prevent accidental edits.
        Set with_window=false for background processing.
        """
        return open_presentation(params)

    @mcp.tool(
        name="ppt_save_presentation",
        annotations={
            "title": "Save Presentation",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": True,
        },
    )
    async def tool_save_presentation(params: SavePresentationInput) -> str:
        """Save the active or specified presentation to its current file.

        This overwrites the existing file. Use ppt_save_presentation_as to
        save to a new location or format.
        """
        return save_presentation(params)

    @mcp.tool(
        name="ppt_save_presentation_as",
        annotations={
            "title": "Save Presentation As",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        },
    )
    async def tool_save_presentation_as(params: SavePresentationAsInput) -> str:
        """Save a presentation to a new file path and/or format.

        Supported formats: 'pptx', 'pdf', 'png', 'jpg', 'default'.
        Note: SaveAs changes the presentation's name to the new path.
        For image formats (png/jpg), a folder of individual slide images is created.
        """
        return save_presentation_as(params)

    @mcp.tool(
        name="ppt_close_presentation",
        annotations={
            "title": "Close Presentation",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_close_presentation(params: ClosePresentationInput) -> str:
        """Close a presentation.

        Set save_changes=true to save before closing.
        If save_changes=false (default), unsaved changes are discarded without
        prompting the user.
        """
        return close_presentation(params)

    @mcp.tool(
        name="ppt_get_presentation_info",
        annotations={
            "title": "Get Presentation Info",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_get_presentation_info(
        params: GetPresentationInfoInput,
    ) -> str:
        """Get detailed information about a presentation.

        Returns name, file path, slide count, dimensions, read-only status,
        save status, and template name.
        """
        return get_presentation_info(params)

    @mcp.tool(
        name="ppt_list_templates",
        annotations={
            "title": "List Templates",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": True,
        },
    )
    async def tool_list_templates(params: ListTemplatesInput) -> str:
        """List available PowerPoint template files (.potx, .potm).

        Scans the personal templates folder (auto-detected) or a
        specified directory. Returns template names, file names, and
        absolute paths. Use a returned file_path with
        ppt_create_presentation's template_path to create from a template.
        """
        return list_templates(params)
