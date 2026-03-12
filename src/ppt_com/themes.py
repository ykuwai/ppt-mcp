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
from utils.color import int_to_hex, hex_to_int, THEME_COLOR_MAP
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
# Preset color palettes
# ---------------------------------------------------------------------------
PRESET_PALETTES: dict[str, dict[str, str]] = {
    # --- Classic / Professional ---
    "corporate_blue": {
        "dark1": "#1B2A4A", "light1": "#FFFFFF", "dark2": "#44546A", "light2": "#F2F2F2",
        "accent1": "#4472C4", "accent2": "#ED7D31", "accent3": "#A5A5A5",
        "accent4": "#FFC000", "accent5": "#5B9BD5", "accent6": "#70AD47",
    },
    "executive_charcoal": {
        "dark1": "#2D2D2D", "light1": "#FFFFFF", "dark2": "#404040", "light2": "#F5F5F5",
        "accent1": "#3A7CA5", "accent2": "#D4A84B", "accent3": "#81A88D",
        "accent4": "#C96B6B", "accent5": "#7B8EB8", "accent6": "#D4896A",
    },
    "consulting": {
        "dark1": "#002B49", "light1": "#FFFFFF", "dark2": "#3C3C3C", "light2": "#F0F4F8",
        "accent1": "#005587", "accent2": "#00A3AD", "accent3": "#6D2077",
        "accent4": "#E87722", "accent5": "#8DC63F", "accent6": "#C4262E",
    },
    # --- Tech / Modern ---
    "nord": {
        "dark1": "#2E3440", "light1": "#ECEFF4", "dark2": "#3B4252", "light2": "#E5E9F0",
        "accent1": "#5E81AC", "accent2": "#81A1C1", "accent3": "#88C0D0",
        "accent4": "#8FBCBB", "accent5": "#A3BE8C", "accent6": "#BF616A",
    },
    "dracula": {
        "dark1": "#282A36", "light1": "#F8F8F2", "dark2": "#44475A", "light2": "#F8F8F2",
        "accent1": "#BD93F9", "accent2": "#FF79C6", "accent3": "#8BE9FD",
        "accent4": "#50FA7B", "accent5": "#FFB86C", "accent6": "#FF5555",
    },
    "tokyo_night": {
        "dark1": "#1A1B26", "light1": "#C0CAF5", "dark2": "#16161E", "light2": "#A9B1D6",
        "accent1": "#7AA2F7", "accent2": "#BB9AF7", "accent3": "#7DCFFF",
        "accent4": "#9ECE6A", "accent5": "#FF9E64", "accent6": "#F7768E",
    },
    "catppuccin_mocha": {
        "dark1": "#1E1E2E", "light1": "#CDD6F4", "dark2": "#181825", "light2": "#BAC2DE",
        "accent1": "#89B4FA", "accent2": "#CBA6F7", "accent3": "#F5C2E7",
        "accent4": "#A6E3A1", "accent5": "#FAB387", "accent6": "#F38BA8",
    },
    "catppuccin_latte": {
        "dark1": "#4C4F69", "light1": "#EFF1F5", "dark2": "#5C5F77", "light2": "#E6E9EF",
        "accent1": "#1E66F5", "accent2": "#8839EF", "accent3": "#EA76CB",
        "accent4": "#40A02B", "accent5": "#FE640B", "accent6": "#D20F39",
    },
    "solarized_dark": {
        "dark1": "#002B36", "light1": "#FDF6E3", "dark2": "#073642", "light2": "#EEE8D5",
        "accent1": "#268BD2", "accent2": "#2AA198", "accent3": "#859900",
        "accent4": "#B58900", "accent5": "#CB4B16", "accent6": "#DC322F",
    },
    "solarized_light": {
        "dark1": "#657B83", "light1": "#FDF6E3", "dark2": "#586E75", "light2": "#EEE8D5",
        "accent1": "#268BD2", "accent2": "#2AA198", "accent3": "#859900",
        "accent4": "#B58900", "accent5": "#CB4B16", "accent6": "#D33682",
    },
    "gruvbox_dark": {
        "dark1": "#282828", "light1": "#FBF1C7", "dark2": "#3C3836", "light2": "#EBDBB2",
        "accent1": "#458588", "accent2": "#B16286", "accent3": "#689D6A",
        "accent4": "#D79921", "accent5": "#D65D0E", "accent6": "#CC241D",
    },
    "one_dark": {
        "dark1": "#282C34", "light1": "#ABB2BF", "dark2": "#21252B", "light2": "#D7DAE0",
        "accent1": "#61AFEF", "accent2": "#C678DD", "accent3": "#56B6C2",
        "accent4": "#98C379", "accent5": "#E5C07B", "accent6": "#E06C75",
    },
    # --- Vibrant / Creative ---
    "sunset": {
        "dark1": "#2C1810", "light1": "#FFF8F0", "dark2": "#4A3228", "light2": "#FAEBD7",
        "accent1": "#FF6B35", "accent2": "#F7C548", "accent3": "#D64045",
        "accent4": "#7B2D8E", "accent5": "#1B998B", "accent6": "#3185FC",
    },
    "ocean": {
        "dark1": "#0B132B", "light1": "#FFFFFF", "dark2": "#1C2541", "light2": "#E8F4F8",
        "accent1": "#3A86FF", "accent2": "#8338EC", "accent3": "#FF006E",
        "accent4": "#FB5607", "accent5": "#FFBE0B", "accent6": "#06D6A0",
    },
    # --- Minimal / Clean ---
    "stone": {
        "dark1": "#292524", "light1": "#FAFAF9", "dark2": "#44403C", "light2": "#F5F5F4",
        "accent1": "#78716C", "accent2": "#A8A29E", "accent3": "#B45309",
        "accent4": "#0F766E", "accent5": "#6D28D9", "accent6": "#BE185D",
    },
    "swiss": {
        "dark1": "#1A1A1A", "light1": "#FFFFFF", "dark2": "#4A4A4A", "light2": "#F0F0F0",
        "accent1": "#E30613", "accent2": "#2D2D2D", "accent3": "#6B6B6B",
        "accent4": "#A0A0A0", "accent5": "#D4D4D4", "accent6": "#0055A4",
    },
    "sage": {
        "dark1": "#2F3E46", "light1": "#FFFFFF", "dark2": "#354F52", "light2": "#F0F7F4",
        "accent1": "#52796F", "accent2": "#84A98C", "accent3": "#CAD2C5",
        "accent4": "#BC6C25", "accent5": "#606C38", "accent6": "#283618",
    },
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


class SetThemeColorsInput(BaseModel):
    """Input for setting individual theme colors."""
    model_config = ConfigDict(str_strip_whitespace=True)

    preset: Optional[str] = Field(
        default=None,
        description=(
            "Apply a preset color palette by name. "
            "Available: corporate_blue, executive_charcoal, consulting, "
            "nord, dracula, tokyo_night, catppuccin_mocha, catppuccin_latte, "
            "solarized_dark, solarized_light, gruvbox_dark, one_dark, "
            "sunset, ocean, stone, swiss, sage. "
            "Individual color fields override preset values."
        ),
    )
    dark1: Optional[str] = Field(
        default=None, description="Dark 1 color (main text/heading) as #RRGGBB",
    )
    light1: Optional[str] = Field(
        default=None, description="Light 1 color (main background) as #RRGGBB",
    )
    dark2: Optional[str] = Field(
        default=None, description="Dark 2 color (secondary text) as #RRGGBB",
    )
    light2: Optional[str] = Field(
        default=None, description="Light 2 color (secondary background) as #RRGGBB",
    )
    accent1: Optional[str] = Field(
        default=None, description="Accent 1 color (primary accent) as #RRGGBB",
    )
    accent2: Optional[str] = Field(
        default=None, description="Accent 2 color as #RRGGBB",
    )
    accent3: Optional[str] = Field(
        default=None, description="Accent 3 color as #RRGGBB",
    )
    accent4: Optional[str] = Field(
        default=None, description="Accent 4 color as #RRGGBB",
    )
    accent5: Optional[str] = Field(
        default=None, description="Accent 5 color as #RRGGBB",
    )
    accent6: Optional[str] = Field(
        default=None, description="Accent 6 color as #RRGGBB",
    )
    hyperlink: Optional[str] = Field(
        default=None, description="Hyperlink color as #RRGGBB",
    )
    followed_hyperlink: Optional[str] = Field(
        default=None, description="Followed hyperlink color as #RRGGBB",
    )


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


def _set_theme_colors_impl(color_map):
    """Set individual theme colors.

    Args:
        color_map: dict of {theme_color_index: bgr_int} pairs.
    """
    pres = ppt._get_pres_impl()

    theme = pres.SlideMaster.Theme
    color_scheme = theme.ThemeColorScheme

    changed = []
    for idx, bgr in color_map.items():
        color_scheme(idx).RGB = bgr
        changed.append({
            "name": THEME_COLOR_NAMES[idx],
            "color": int_to_hex(bgr),
        })

    return {
        "success": True,
        "changed": changed,
        "changed_count": len(changed),
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


def set_theme_colors(params: SetThemeColorsInput) -> str:
    """Set individual theme colors.

    Args:
        params: Color values for specific theme color slots.

    Returns:
        JSON confirming which colors were changed.
    """
    try:
        # Resolve preset palette as base, then overlay individual fields
        merged: dict[str, str] = {}
        preset_name = None
        if params.preset is not None:
            preset_name = params.preset.lower().strip()
            if preset_name not in PRESET_PALETTES:
                available = ", ".join(sorted(PRESET_PALETTES.keys()))
                return json.dumps({
                    "error": f"Unknown preset: '{params.preset}'. "
                             f"Available: {available}",
                })
            merged.update(PRESET_PALETTES[preset_name])

        # Individual fields override preset values
        for name in THEME_COLOR_MAP:
            val = getattr(params, name, None)
            if val is not None:
                merged[name] = val

        if not merged:
            return json.dumps({"error": "No colors specified"})

        # Build color_map: {theme_color_index: bgr_int}
        color_map = {}
        for name, hex_val in merged.items():
            color_map[THEME_COLOR_MAP[name]] = hex_to_int(hex_val)

        result = ppt.execute(_set_theme_colors_impl, color_map)
        if preset_name:
            result["preset"] = preset_name
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to set theme colors: {str(e)}"})


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
        name="ppt_set_theme_colors",
        annotations={
            "title": "Set Theme Colors",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_set_theme_colors(params: SetThemeColorsInput) -> str:
        """Set theme colors of the active presentation.

        Three modes:
        1. **Preset only**: `preset="nord"` applies all colors from a curated palette.
        2. **Manual**: specify individual color slots (dark1, light1, accent1, etc.).
        3. **Preset + override**: start from a preset, then override specific slots.

        17 presets available across 4 categories:
        - Classic: corporate_blue, executive_charcoal, consulting
        - Tech: nord, dracula, tokyo_night, catppuccin_mocha, catppuccin_latte,
          solarized_dark, solarized_light, gruvbox_dark, one_dark
        - Vibrant: sunset, ocean
        - Minimal: stone, swiss, sage

        Color values are #RRGGBB hex strings. Only specified colors are changed;
        omitted colors remain unchanged.
        """
        return set_theme_colors(params)

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
