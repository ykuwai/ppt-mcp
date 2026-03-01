"""ppt-mcp: The world's best PowerPoint MCP server.

Real-time PowerPoint control via COM automation.
"""

import logging
import os
import sys
import tempfile
from pathlib import Path

from pydantic import BaseModel, Field

# When installed via PyPI (entry point: src.server:main), ensure the src/
# directory is in sys.path so that internal imports like
# `from utils.com_wrapper import ppt` resolve correctly.
_src_dir = str(Path(__file__).parent)
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)

from contextlib import asynccontextmanager

from mcp.server.fastmcp import FastMCP, Image

# Configure logging to stderr (stdout is used for MCP protocol)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger("ppt-mcp")


@asynccontextmanager
async def app_lifespan(server: FastMCP):
    """Manage COM lifecycle for the MCP server."""
    from utils.com_wrapper import ppt

    from utils.com_wrapper import AUTO_DISMISS_DIALOG
    logger.info("AUTO_DISMISS_DIALOG=%s (set PPT_AUTO_DISMISS_DIALOG=true to enable)", AUTO_DISMISS_DIALOG)
    logger.info("Starting PowerPoint COM worker thread...")
    ppt.start()
    try:
        # Connect to PowerPoint at startup
        ppt.connect()
        logger.info("PowerPoint COM connection established")
        yield {}
    finally:
        logger.info("Shutting down PowerPoint COM worker thread...")
        ppt.stop()


mcp = FastMCP(
    "powerpoint_mcp",
    lifespan=app_lifespan,
    instructions="""
## Getting started

1. Call `ppt_activate_presentation` first — locks all tools to a specific file and prevents accidental edits to the wrong presentation.
2. Call `ppt_get_presentation_info` to understand the presentation — slide count, dimensions, template, current default fonts, and accent colors. Use this to inform all subsequent decisions.
3. After placing text, set fonts explicitly with `ppt_batch_apply_formatting` or `ppt_set_default_fonts`. On Japanese-locale Windows, the slide master default is often 游ゴシック, which renders thin and illegible when projected. Preferred fonts: BIZ UDPゴシック (Japanese) + Segoe UI (Latin).
4. For visual symbols, `ppt_search_icons` + `ppt_add_svg_icon` produce crisper, scalable results than emoji characters and are generally preferred in presentations.
5. Use `ppt_get_slide_preview` to visually inspect slides as you work.

## Tips

- Use accent colors from `ppt_get_presentation_info` (or `ppt_get_theme_colors`) instead of hardcoding RGB values — theme names adapt automatically to the presentation's palette.
- Standard 16:9 slide = 960 × 540 pt.
- `ppt_batch_apply_formatting` applies multiple operations to multiple shapes in one call. Supported operations: `set_fill`, `set_line`, `set_shadow`, `set_glow`, `set_reflection`, `set_soft_edge`, `format_text`. Specify shapes by name or 1-based index. Use this whenever you want consistent styling across several shapes — much more efficient than calling individual tools per shape.
- For consistent shape styling, call `ppt_set_default_shape_style` before inserting shapes. Shape-based mode (`slide_index` + `shape_name_or_index`) captures all properties of a template shape including gradients and effects. Property-based mode sets fill/border/font directly without needing a pre-existing shape.
""",
)


# =============================================================================
# App tools
# =============================================================================
from ppt_com.app import (
    ConnectInput,
    SetWindowStateInput,
    connect_to_powerpoint,
    get_app_info,
    get_active_window_info,
    list_presentations,
    set_window_state,
)


@mcp.tool(
    name="ppt_connect",
    annotations={
        "title": "Connect to PowerPoint",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def tool_ppt_connect(params: ConnectInput) -> str:
    """Connect to a running PowerPoint instance or launch a new one.

    Attempts to connect to an already-running PowerPoint via COM.
    If no instance is found, launches a new one.
    Set visible=false for headless mode (background operation).
    """
    return connect_to_powerpoint(params)


@mcp.tool(
    name="ppt_get_app_info",
    annotations={
        "title": "Get PowerPoint App Info",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def tool_ppt_get_app_info() -> str:
    """Get information about the connected PowerPoint application.

    Returns version, visibility, window state, presentation count,
    and active presentation name.
    """
    return get_app_info()


@mcp.tool(
    name="ppt_get_active_window",
    annotations={
        "title": "Get Active Window Info",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def tool_ppt_get_active_window() -> str:
    """Get info about the active PowerPoint window and current selection.

    Returns window caption, view type, current slide index,
    and what is selected (shapes, text, or nothing).
    """
    return get_active_window_info()


@mcp.tool(
    name="ppt_list_presentations",
    annotations={
        "title": "List Open Presentations",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def tool_ppt_list_presentations() -> str:
    """List all currently open presentations in PowerPoint.

    Returns name, path, slide count, and status for each.
    """
    return list_presentations()


@mcp.tool(
    name="ppt_set_window_state",
    annotations={
        "title": "Set PowerPoint Window State",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    },
)
async def tool_ppt_set_window_state(params: SetWindowStateInput) -> str:
    """Set the PowerPoint application window state.

    Controls whether the PowerPoint window is maximized, minimized, or
    restored to normal size.
    """
    return set_window_state(params)


# =============================================================================
# Import and register additional tool modules as they are implemented.
# Each module registers its tools below.
# =============================================================================

# Presentation tools
try:
    from ppt_com.presentation import register_tools as register_presentation_tools
    register_presentation_tools(mcp)
except ImportError:
    logger.debug("presentation module not yet available")

# Slide tools
try:
    from ppt_com.slides import register_tools as register_slide_tools
    register_slide_tools(mcp)
except ImportError:
    logger.debug("slides module not yet available")

# Shape tools
try:
    from ppt_com.shapes import register_tools as register_shape_tools
    register_shape_tools(mcp)
except ImportError:
    logger.debug("shapes module not yet available")

# Text tools
try:
    from ppt_com.text import register_tools as register_text_tools
    register_text_tools(mcp)
except ImportError:
    logger.debug("text module not yet available")

# Placeholder tools
try:
    from ppt_com.placeholders import register_tools as register_placeholder_tools
    register_placeholder_tools(mcp)
except ImportError:
    logger.debug("placeholders module not yet available")

# Formatting tools
try:
    from ppt_com.formatting import register_tools as register_formatting_tools
    register_formatting_tools(mcp)
except ImportError:
    logger.debug("formatting module not yet available")

# Table tools
try:
    from ppt_com.tables import register_tools as register_table_tools
    register_table_tools(mcp)
except ImportError:
    logger.debug("tables module not yet available")

# Export tools
try:
    from ppt_com.export import register_tools as register_export_tools
    register_export_tools(mcp)
except ImportError:
    logger.debug("export module not yet available")

# SlideShow tools
try:
    from ppt_com.slideshow import register_tools as register_slideshow_tools
    register_slideshow_tools(mcp)
except ImportError:
    logger.debug("slideshow module not yet available")

# Groups tools
try:
    from ppt_com.groups import register_tools as register_groups_tools
    register_groups_tools(mcp)
except ImportError:
    logger.debug("groups module not yet available")

# Connectors tools
try:
    from ppt_com.connectors import register_tools as register_connectors_tools
    register_connectors_tools(mcp)
except ImportError:
    logger.debug("connectors module not yet available")

# Hyperlinks tools
try:
    from ppt_com.hyperlinks import register_tools as register_hyperlinks_tools
    register_hyperlinks_tools(mcp)
except ImportError:
    logger.debug("hyperlinks module not yet available")

# Sections tools
try:
    from ppt_com.sections import register_tools as register_sections_tools
    register_sections_tools(mcp)
except ImportError:
    logger.debug("sections module not yet available")

# Properties tools
try:
    from ppt_com.properties import register_tools as register_properties_tools
    register_properties_tools(mcp)
except ImportError:
    logger.debug("properties module not yet available")

# Charts tools
try:
    from ppt_com.charts import register_tools as register_charts_tools
    register_charts_tools(mcp)
except ImportError:
    logger.debug("charts module not yet available")

# Animation tools
try:
    from ppt_com.animation import register_tools as register_animation_tools
    register_animation_tools(mcp)
except ImportError:
    logger.debug("animation module not yet available")

# Themes tools
try:
    from ppt_com.themes import register_tools as register_themes_tools
    register_themes_tools(mcp)
except ImportError:
    logger.debug("themes module not yet available")

# Media tools
try:
    from ppt_com.media import register_tools as register_media_tools
    register_media_tools(mcp)
except ImportError:
    logger.debug("media module not yet available")

# SmartArt tools
try:
    from ppt_com.smartart import register_tools as register_smartart_tools
    register_smartart_tools(mcp)
except ImportError:
    logger.debug("smartart module not yet available")

# Edit operations tools (undo, redo, clipboard, format copy)
try:
    from ppt_com.edit_ops import register_tools as register_edit_ops_tools
    register_edit_ops_tools(mcp)
except ImportError:
    logger.debug("edit_ops module not yet available")

# Layout tools (align, distribute, slide size, background, flip)
try:
    from ppt_com.layout import register_tools as register_layout_tools
    register_layout_tools(mcp)
except ImportError:
    logger.debug("layout module not yet available")

# Visual effects tools (glow, reflection, soft edge)
try:
    from ppt_com.effects import register_tools as register_effects_tools
    register_effects_tools(mcp)
except ImportError:
    logger.debug("effects module not yet available")

# Comments tools
try:
    from ppt_com.comments import register_tools as register_comments_tools
    register_comments_tools(mcp)
except ImportError:
    logger.debug("comments module not yet available")

# Advanced operations tools (tags, fonts, crop, merge, export, hidden, select, view)
try:
    from ppt_com.advanced_ops import register_tools as register_advanced_ops_tools
    register_advanced_ops_tools(mcp)
except ImportError:
    logger.debug("advanced_ops module not yet available")

# Batch apply formatting tools
try:
    from ppt_com.batch_apply import register_tools as register_batch_apply_tools
    register_batch_apply_tools(mcp)
except ImportError:
    logger.debug("batch_apply module not yet available")

# Freeform shape tools
try:
    from ppt_com.freeform import register_tools as register_freeform_tools
    register_freeform_tools(mcp)
except ImportError:
    logger.debug("freeform module not yet available")


# =============================================================================
# Tools: Slide Preview (Visual Inspection)
# =============================================================================


class GetSlidePreviewInput(BaseModel):
    slide_index: int = Field(1, ge=1, description="1-based slide index")


@mcp.tool(
    name="ppt_get_slide_preview",
    annotations={
        "title": "Get Slide Preview Image",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def tool_ppt_get_slide_preview(params: GetSlidePreviewInput) -> Image:
    """Get a visual preview of a PowerPoint slide as an image.

    This is the RECOMMENDED way to visually inspect slides for appearance, design,
    layout, colors, text readability, and overall quality. Much more efficient
    than exporting all slides to files.

    Also navigates the PowerPoint editor window to the target slide so the user
    can see which slide is being inspected.

    Returns:
        Image: PNG image of the slide for visual inspection
    """
    from utils.com_wrapper import ppt
    from utils.navigation import goto_slide

    def _export_slide_impl(slide_idx: int):
        app = ppt._get_app_impl()
        pres = ppt._get_pres_impl()
        goto_slide(app, slide_idx)

        # Validate slide
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            raise ValueError(
                f"Slide index {slide_idx} out of range (1-{pres.Slides.Count})"
            )

        slide = pres.Slides(slide_idx)

        # Generate temp file path
        temp_dir = tempfile.gettempdir()
        temp_file = os.path.join(temp_dir, f"ppt_slide_preview_{slide_idx}.png")

        try:
            # Export slide as PNG
            slide.Export(temp_file, "PNG")

            # Read binary data
            with open(temp_file, "rb") as f:
                image_data = f.read()

            return image_data
        finally:
            # Cleanup
            if os.path.exists(temp_file):
                os.remove(temp_file)

    image_data = ppt.execute(_export_slide_impl, params.slide_index)
    return Image(data=image_data, format="png")


def main():
    """Entry point for the MCP server."""
    mcp.run()


if __name__ == "__main__":
    main()
