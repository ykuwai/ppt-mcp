"""ppt-com-mcp: The world's best PowerPoint MCP server.

Real-time PowerPoint control via COM automation.
"""

import logging
import sys
from contextlib import asynccontextmanager

from mcp.server.fastmcp import FastMCP

# Configure logging to stderr (stdout is used for MCP protocol)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger("ppt-com-mcp")


@asynccontextmanager
async def app_lifespan(server: FastMCP):
    """Manage COM lifecycle for the MCP server."""
    from utils.com_wrapper import ppt

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
)


# =============================================================================
# App tools
# =============================================================================
from ppt_com.app import (
    ConnectInput,
    connect_to_powerpoint,
    get_app_info,
    get_active_window_info,
    list_presentations,
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


def main():
    """Entry point for the MCP server."""
    mcp.run()


if __name__ == "__main__":
    main()
