"""Theme and presentation-level formatting operations for PowerPoint COM automation.

Handles applying themes, reading theme colors, and setting headers/footers
across all slides.
"""

import colorsys
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
# Preset color palettes (all accents WCAG AA Large Text 3:1+ on light1)
# ---------------------------------------------------------------------------
PRESET_PALETTES: dict[str, dict[str, str]] = {
    # --- Classic / Professional ---
    "corporate_blue": {
        "dark1": "#1B2A4A", "light1": "#FFFFFF", "dark2": "#44546A", "light2": "#F2F2F2",
        "accent1": "#2B579A", "accent2": "#BF4B28", "accent3": "#2E7D32",
        "accent4": "#7B3FA0", "accent5": "#C4652A", "accent6": "#1A7A8A",
    },
    "executive": {
        "dark1": "#1C1C1C", "light1": "#FFFFFF", "dark2": "#3D3D3D", "light2": "#F5F5F0",
        "accent1": "#2F4858", "accent2": "#6B3A3A", "accent3": "#4A6741",
        "accent4": "#5B4A6E", "accent5": "#7A5C3E", "accent6": "#3A5F6F",
    },
    "consulting": {
        "dark1": "#0A1628", "light1": "#FFFFFF", "dark2": "#2C3E50", "light2": "#F0F3F7",
        "accent1": "#003A70", "accent2": "#00796B", "accent3": "#6A1B9A",
        "accent4": "#B71C1C", "accent5": "#2E5090", "accent6": "#455A64",
    },
    # --- Design System Based ---
    "tailwind": {
        "dark1": "#1E293B", "light1": "#FFFFFF", "dark2": "#334155", "light2": "#F1F5F9",
        "accent1": "#1D4ED8", "accent2": "#B91C1C", "accent3": "#15803D",
        "accent4": "#7E22CE", "accent5": "#B45309", "accent6": "#0F766E",
    },
    "chakra": {
        "dark1": "#1A202C", "light1": "#FFFFFF", "dark2": "#2D3748", "light2": "#EDF2F7",
        "accent1": "#2C5282", "accent2": "#9B2C2C", "accent3": "#276749",
        "accent4": "#553C9A", "accent5": "#9C4221", "accent6": "#285E61",
    },
    "open_color": {
        "dark1": "#212529", "light1": "#FFFFFF", "dark2": "#343A40", "light2": "#F1F3F5",
        "accent1": "#1864AB", "accent2": "#C92A2A", "accent3": "#2F9E44",
        "accent4": "#6741D9", "accent5": "#E8590C", "accent6": "#099268",
    },
    "radix": {
        "dark1": "#11181C", "light1": "#FFFFFF", "dark2": "#1C2024", "light2": "#F0F2F4",
        "accent1": "#006ADC", "accent2": "#D31E66", "accent3": "#18794E",
        "accent4": "#793AAF", "accent5": "#BD4B00", "accent6": "#067A6F",
    },
    # --- Nature / Mood ---
    "ocean": {
        "dark1": "#0A1929", "light1": "#FFFFFF", "dark2": "#132F4C", "light2": "#E3F2FD",
        "accent1": "#0D47A1", "accent2": "#006064", "accent3": "#01579B",
        "accent4": "#0277BD", "accent5": "#00695C", "accent6": "#1A237E",
    },
    "forest": {
        "dark1": "#1B2418", "light1": "#FFFFFF", "dark2": "#33402E", "light2": "#F1F4EC",
        "accent1": "#2E7D32", "accent2": "#5D4037", "accent3": "#1B5E20",
        "accent4": "#795548", "accent5": "#827717", "accent6": "#4E342E",
    },
    "sunset": {
        "dark1": "#2C1810", "light1": "#FFFFFF", "dark2": "#4E342E", "light2": "#FFF3E0",
        "accent1": "#C62828", "accent2": "#BF360C", "accent3": "#AD1457",
        "accent4": "#880E4F", "accent5": "#6A1B9A", "accent6": "#B71C1C",
    },
    "sage": {
        "dark1": "#263238", "light1": "#FFFFFF", "dark2": "#37474F", "light2": "#F1F4EC",
        "accent1": "#558B2F", "accent2": "#6D4C41", "accent3": "#33691E",
        "accent4": "#4E6B45", "accent5": "#8D6E63", "accent6": "#5D6B3C",
    },
    # --- Modern / Trendy ---
    "nord_light": {
        "dark1": "#2E3440", "light1": "#FFFFFF", "dark2": "#3B4252", "light2": "#ECEFF4",
        "accent1": "#3868A6", "accent2": "#A3394B", "accent3": "#6B5DAD",
        "accent4": "#2C7A5D", "accent5": "#B06D2F", "accent6": "#2E6C8F",
    },
    "pastel_deep": {
        "dark1": "#2D2D2D", "light1": "#FFFFFF", "dark2": "#4A4A4A", "light2": "#FAF8F5",
        "accent1": "#5B72A8", "accent2": "#A85B6E", "accent3": "#6B8F5E",
        "accent4": "#8B6BA8", "accent5": "#B07D4F", "accent6": "#4F8F8B",
    },
    "swiss": {
        "dark1": "#000000", "light1": "#FFFFFF", "dark2": "#1A1A1A", "light2": "#F5F5F5",
        "accent1": "#CC0000", "accent2": "#000000", "accent3": "#004F9E",
        "accent4": "#6B6B6B", "accent5": "#8B0000", "accent6": "#2B5B2B",
    },
    # --- Vibrant ---
    "vivid": {
        "dark1": "#0D0D0D", "light1": "#FFFFFF", "dark2": "#262626", "light2": "#F5F5F5",
        "accent1": "#0050D0", "accent2": "#C60040", "accent3": "#007A33",
        "accent4": "#7B00B5", "accent5": "#B85000", "accent6": "#00787C",
    },
    "rainbow": {
        "dark1": "#1A1A1A", "light1": "#FFFFFF", "dark2": "#333333", "light2": "#F7F7F7",
        "accent1": "#C62828", "accent2": "#C45000", "accent3": "#2E7D32",
        "accent4": "#1565C0", "accent5": "#6A1B9A", "accent6": "#AD1457",
    },
    "neon_safe": {
        "dark1": "#0F0F1A", "light1": "#FFFFFF", "dark2": "#1A1A2E", "light2": "#F0F0FA",
        "accent1": "#0055CC", "accent2": "#CC0066", "accent3": "#008844",
        "accent4": "#7700BB", "accent5": "#CC5500", "accent6": "#007788",
    },
}


# ---------------------------------------------------------------------------
# Primary color → palette generation (HSL color harmony + WCAG contrast)
# ---------------------------------------------------------------------------
def _relative_luminance(r: int, g: int, b: int) -> float:
    """Calculate relative luminance per WCAG 2.x definition."""
    def linearize(c: int) -> float:
        c_norm = c / 255.0
        return c_norm / 12.92 if c_norm <= 0.04045 else ((c_norm + 0.055) / 1.055) ** 2.4
    return 0.2126 * linearize(r) + 0.7152 * linearize(g) + 0.0722 * linearize(b)


def _contrast_ratio(hex1: str, hex2: str) -> float:
    """Calculate WCAG contrast ratio between two #RRGGBB colors."""
    def to_rgb(h: str) -> tuple[int, int, int]:
        h = h.lstrip('#')
        return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    l1 = _relative_luminance(*to_rgb(hex1))
    l2 = _relative_luminance(*to_rgb(hex2))
    lighter, darker = max(l1, l2), min(l1, l2)
    return (lighter + 0.05) / (darker + 0.05)


def _hsl_to_hex(h: float, s: float, l: float) -> str:
    """Convert HSL (0-1 range) to #RRGGBB hex string."""
    r, g, b = colorsys.hls_to_rgb(h, l, s)  # Note: colorsys uses HLS order
    return f"#{int(r * 255 + 0.5):02X}{int(g * 255 + 0.5):02X}{int(b * 255 + 0.5):02X}"


def _hex_to_hsl(hex_color: str) -> tuple[float, float, float]:
    """Convert #RRGGBB hex string to (h, s, l) tuple in 0-1 range."""
    hex_color = hex_color.lstrip('#')
    r, g, b = int(hex_color[0:2], 16) / 255.0, int(hex_color[2:4], 16) / 255.0, int(hex_color[4:6], 16) / 255.0
    h, l, s = colorsys.rgb_to_hls(r, g, b)  # Note: colorsys returns HLS
    return h, s, l


def _ensure_contrast(hex_color: str, min_ratio: float = 3.0) -> str:
    """Darken a color until it meets the minimum contrast ratio against white."""
    if _contrast_ratio(hex_color, "#FFFFFF") >= min_ratio:
        return hex_color
    h, s, l = _hex_to_hsl(hex_color)
    # Progressively reduce lightness until contrast passes
    while l > 0.0:
        l = max(0.0, l - 0.01)
        candidate = _hsl_to_hex(h, s, l)
        if _contrast_ratio(candidate, "#FFFFFF") >= min_ratio:
            return candidate
    return _hsl_to_hex(h, s, 0.0)


def generate_palette_from_primary(primary_hex: str) -> dict[str, str]:
    """Generate a harmonious 10-color theme palette from a single primary color.

    Uses split-complementary + analogous color harmony in HSL space.
    All generated accent colors are guaranteed WCAG AA Large Text (3:1+)
    contrast against white.

    Args:
        primary_hex: Primary color as #RRGGBB hex string.

    Returns:
        Dict with keys: dark1, light1, dark2, light2, accent1-accent6.
    """
    h, s, l = _hex_to_hsl(primary_hex)

    # Clamp primary lightness for WCAG compliance
    accent_l = max(0.25, min(0.45, l))
    accent_s = max(0.5, s)  # Ensure enough saturation for vibrant accents

    # Generate 6 accents using color harmony offsets
    offsets = [0, 30, 120, 180, 210, 330]  # degrees
    accents = []
    for offset in offsets:
        accent_h = (h + offset / 360.0) % 1.0
        hex_val = _hsl_to_hex(accent_h, accent_s, accent_l)
        hex_val = _ensure_contrast(hex_val)
        accents.append(hex_val)

    # Dark/light base colors derived from primary hue
    dark1 = _hsl_to_hex(h, max(0.3, s * 0.6), 0.12)  # Very dark for text
    dark2 = _hsl_to_hex(h, max(0.2, s * 0.4), 0.22)  # Slightly lighter
    light1 = "#FFFFFF"
    light2 = _hsl_to_hex(h, max(0.1, s * 0.3), 0.95)  # Very light tint

    return {
        "dark1": dark1, "light1": light1, "dark2": dark2, "light2": light2,
        "accent1": accents[0], "accent2": accents[1], "accent3": accents[2],
        "accent4": accents[3], "accent5": accents[4], "accent6": accents[5],
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
            "Apply a preset color palette by name. All presets are "
            "WCAG AA accessible (3:1+ contrast on white). "
            "Classic: corporate_blue, executive, consulting. "
            "Design systems: tailwind, chakra, open_color, radix. "
            "Nature: ocean, forest, sunset, sage. "
            "Modern: nord_light, pastel_deep, swiss. "
            "Vibrant: vivid, rainbow, neon_safe. "
            "Individual color fields override preset values."
        ),
    )
    primary: Optional[str] = Field(
        default=None,
        description=(
            "Auto-generate a harmonious 6-color accent palette from a single "
            "primary color (#RRGGBB). Uses color harmony (split-complementary "
            "+ analogous). Individual accent fields and preset override "
            "generated values."
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

    # Read from the first design's slide master
    theme = pres.Designs(1).SlideMaster.Theme
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
        "design_count": pres.Designs.Count,
    }


def _set_theme_colors_impl(color_map):
    """Set individual theme colors across ALL slide masters.

    Args:
        color_map: dict of {theme_color_index: bgr_int} pairs.
    """
    pres = ppt._get_pres_impl()

    design_count = pres.Designs.Count
    for d in range(1, design_count + 1):
        color_scheme = pres.Designs(d).SlideMaster.Theme.ThemeColorScheme
        for idx, bgr in color_map.items():
            color_scheme(idx).RGB = bgr

    # Build response from first design
    changed = []
    for idx, bgr in color_map.items():
        changed.append({
            "name": THEME_COLOR_NAMES[idx],
            "color": int_to_hex(bgr),
        })

    return {
        "success": True,
        "changed": changed,
        "changed_count": len(changed),
        "designs_updated": design_count,
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
        # Priority order: preset → primary → individual fields
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

        # Primary generates palette as base (overridden by preset above,
        # then individual fields override both)
        if params.primary is not None:
            generated = generate_palette_from_primary(params.primary)
            # Only fill slots not already set by preset
            for k, v in generated.items():
                if k not in merged:
                    merged[k] = v

        # Individual fields override preset and primary values
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
        if params.primary is not None:
            result["primary"] = params.primary
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

        Four modes:
        1. **Preset only**: `preset="tailwind"` applies all colors from a curated palette.
        2. **Primary only**: `primary="#2B579A"` auto-generates a full harmonious palette
           from a single color using color harmony (split-complementary + analogous).
           All generated accents are WCAG AA Large Text accessible (3:1+ on white).
        3. **Manual**: specify individual color slots (dark1, light1, accent1, etc.).
        4. **Combined**: preset or primary as base, then override specific slots.

        Priority: preset > primary > individual fields (each layer overrides the previous).

        17 WCAG-accessible presets (all accents 3:1+ on white):
        - Classic: corporate_blue, executive, consulting
        - Design systems: tailwind, chakra, open_color, radix
        - Nature: ocean, forest, sunset, sage
        - Modern: nord_light, pastel_deep, swiss
        - Vibrant: vivid, rainbow, neon_safe

        Colors are applied to ALL slide masters. Values are #RRGGBB hex strings.
        Only specified colors are changed; omitted colors remain unchanged.
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
