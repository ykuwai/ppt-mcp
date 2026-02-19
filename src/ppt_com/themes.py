"""Theme and presentation-level formatting operations for PowerPoint COM automation.

Handles applying themes, reading theme colors, and setting headers/footers
across all slides.
"""

import json
import logging
import os
from typing import Optional

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from utils.color import int_to_hex
from ppt_com.constants import (
    msoTrue, msoFalse,
    msoThemeColorDark1, msoThemeColorLight1,
    msoThemeColorDark2, msoThemeColorLight2,
    msoThemeColorAccent1, msoThemeColorAccent2,
    msoThemeColorAccent3, msoThemeColorAccent4,
    msoThemeColorAccent5, msoThemeColorAccent6,
    msoThemeColorHyperlink, msoThemeColorFollowedHyperlink,
)

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Theme color name map
# ---------------------------------------------------------------------------
THEME_COLOR_NAMES: dict[int, str] = {
    1: "dark1", 2: "light1", 3: "dark2", 4: "light2",
    5: "accent1", 6: "accent2", 7: "accent3", 8: "accent4",
    9: "accent5", 10: "accent6", 11: "hyperlink", 12: "followed_hyperlink",
}


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class ApplyThemeInput(BaseModel):
    """Input for applying a theme to the presentation."""
    model_config = ConfigDict(str_strip_whitespace=True)

    theme_path: str = Field(
        ...,
        description=(
            "Path to the theme file (.thmx) or themed presentation. "
            "Can be relative or absolute; will be normalized to an absolute path."
        ),
    )


class GetThemeColorsInput(BaseModel):
    """Input for getting theme colors (no parameters required)."""
    model_config = ConfigDict(str_strip_whitespace=True)


class SetHeadersFootersInput(BaseModel):
    """Input for setting headers and footers across all slides."""
    model_config = ConfigDict(str_strip_whitespace=True)

    footer_text: Optional[str] = Field(
        default=None, description="Footer text content"
    )
    footer_visible: Optional[bool] = Field(
        default=None, description="Show or hide the footer"
    )
    slide_number_visible: Optional[bool] = Field(
        default=None, description="Show or hide slide numbers"
    )
    date_visible: Optional[bool] = Field(
        default=None, description="Show or hide the date/time"
    )
    date_format: Optional[int] = Field(
        default=None,
        description=(
            "Date format as PpDateTimeFormat integer (e.g. 1=M/d/yy, "
            "2=DayOfWeek Month dd yyyy, 9=H:mm, 12=h:mm:ss AM/PM)"
        ),
    )
    date_fixed_text: Optional[str] = Field(
        default=None,
        description="Fixed date/time text (overrides auto-updating format)",
    )


# ---------------------------------------------------------------------------
# COM implementation functions (run on COM thread via ppt.execute)
# ---------------------------------------------------------------------------
def _apply_theme_impl(theme_path):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()

    abs_path = os.path.abspath(theme_path)
    if not os.path.exists(abs_path):
        raise ValueError(f"Theme file not found: {abs_path}")

    pres.ApplyTheme(abs_path)

    return {
        "success": True,
        "theme_path": abs_path,
    }


def _get_theme_colors_impl():
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()

    theme = pres.SlideMaster.Theme
    color_scheme = theme.ThemeColorScheme

    colors = []
    for i in range(1, 13):
        rgb_int = color_scheme(i).RGB
        colors.append({
            "index": i,
            "name": THEME_COLOR_NAMES[i],
            "color": int_to_hex(rgb_int),
        })

    return {
        "success": True,
        "colors": colors,
    }


def _set_headers_footers_impl(
    footer_text, footer_visible, slide_number_visible,
    date_visible, date_format, date_fixed_text,
):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()

    slide_count = pres.Slides.Count
    for i in range(1, slide_count + 1):
        hf = pres.Slides(i).HeadersFooters

        # Visibility must be set BEFORE text to avoid COM errors
        if footer_visible is not None:
            hf.Footer.Visible = msoTrue if footer_visible else msoFalse
        if footer_text is not None:
            try:
                hf.Footer.Visible = msoTrue
                hf.Footer.Text = footer_text
            except Exception:
                pass  # Slide layout may not support footer
        if slide_number_visible is not None:
            hf.SlideNumber.Visible = msoTrue if slide_number_visible else msoFalse
        if date_visible is not None:
            hf.DateAndTime.Visible = msoTrue if date_visible else msoFalse
        if date_format is not None:
            try:
                hf.DateAndTime.Format = date_format
            except Exception:
                pass
        if date_fixed_text is not None:
            try:
                hf.DateAndTime.UseFormat = msoFalse
                hf.DateAndTime.Text = date_fixed_text
            except Exception:
                pass

    return {
        "success": True,
        "slides_updated": slide_count,
    }


# ---------------------------------------------------------------------------
# MCP tool functions (sync wrappers that delegate to COM thread)
# ---------------------------------------------------------------------------
def apply_theme(params: ApplyThemeInput) -> str:
    """Apply a theme file to the active presentation.

    Args:
        params: Path to the theme file.

    Returns:
        JSON confirming the theme was applied.
    """
    try:
        result = ppt.execute(
            _apply_theme_impl,
            params.theme_path,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to apply theme: {str(e)}"})


def get_theme_colors(params: GetThemeColorsInput) -> str:
    """Get the current theme color scheme.

    Args:
        params: No parameters required.

    Returns:
        JSON with the 12 theme colors (name and hex value).
    """
    try:
        result = ppt.execute(_get_theme_colors_impl)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to get theme colors: {str(e)}"})


def set_headers_footers(params: SetHeadersFootersInput) -> str:
    """Set headers and footers across all slides.

    Args:
        params: Footer text, visibility flags, and date format options.

    Returns:
        JSON confirming how many slides were updated.
    """
    try:
        result = ppt.execute(
            _set_headers_footers_impl,
            params.footer_text, params.footer_visible,
            params.slide_number_visible, params.date_visible,
            params.date_format, params.date_fixed_text,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to set headers/footers: {str(e)}"})


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register all theme tools with the MCP server.

    Note: ppt_get_slide_master_info is already registered in placeholders.py
    and is NOT duplicated here.
    """

    @mcp.tool(
        name="ppt_apply_theme",
        annotations={
            "title": "Apply Theme",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_apply_theme(params: ApplyThemeInput) -> str:
        """Apply a theme file to the active presentation.

        Provide the path to a .thmx theme file or a themed presentation.
        The path will be normalized to an absolute Windows path for COM.
        """
        return apply_theme(params)

    @mcp.tool(
        name="ppt_get_theme_colors",
        annotations={
            "title": "Get Theme Colors",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_get_theme_colors(params: GetThemeColorsInput) -> str:
        """Get the current theme color scheme of the active presentation.

        Returns all 12 theme colors (dark1, light1, dark2, light2,
        accent1-6, hyperlink, followed_hyperlink) with their hex values.
        """
        return get_theme_colors(params)

    @mcp.tool(
        name="ppt_set_headers_footers",
        annotations={
            "title": "Set Headers & Footers",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_set_headers_footers(params: SetHeadersFootersInput) -> str:
        """Set headers and footers across all slides in the presentation.

        Configure footer text, slide numbers, and date/time visibility.
        Use date_fixed_text for a fixed date string, or date_format for
        an auto-updating date (PpDateTimeFormat integer).
        """
        return set_headers_footers(params)
