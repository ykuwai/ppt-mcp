"""Connector tools for PowerPoint COM automation.

Handles adding connectors between shapes and formatting connector lines
including color, weight, dash style, and arrowheads.
"""

import json
import logging
from typing import Optional, Union

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from utils.color import hex_to_int
from ppt_com.constants import (
    msoLineSolid, msoLineRoundDot, msoLineDash,
    msoLineDashDot, msoLineLongDash,
    msoArrowheadNone, msoArrowheadTriangle, msoArrowheadOpen,
    msoArrowheadStealth, msoArrowheadDiamond, msoArrowheadOval,
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Constant maps
# ---------------------------------------------------------------------------
CONNECTOR_TYPE_MAP: dict[str, int] = {
    "straight": 1,
    "elbow": 2,
    "curve": 3,
}

ARROW_STYLE_MAP: dict[str, int] = {
    "none": msoArrowheadNone,
    "triangle": msoArrowheadTriangle,
    "open": msoArrowheadOpen,
    "stealth": msoArrowheadStealth,
    "diamond": msoArrowheadDiamond,
    "oval": msoArrowheadOval,
}

DASH_STYLE_MAP: dict[str, int] = {
    "solid": msoLineSolid,
    "round_dot": msoLineRoundDot,
    "dash": msoLineDash,
    "dash_dot": msoLineDashDot,
    "long_dash": msoLineLongDash,
}


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class AddConnectorInput(BaseModel):
    """Input for adding a connector between two shapes."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    connector_type: str = Field(
        default="elbow",
        description="Connector type: 'straight', 'elbow', or 'curve'",
    )
    begin_shape: str = Field(
        ..., description="Name of the shape where the connector begins"
    )
    begin_site: int = Field(
        default=1, ge=1,
        description="1-based connection site index on the begin shape",
    )
    end_shape: str = Field(
        ..., description="Name of the shape where the connector ends"
    )
    end_site: int = Field(
        default=1, ge=1,
        description="1-based connection site index on the end shape",
    )


class FormatConnectorInput(BaseModel):
    """Input for formatting a connector's line properties."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Connector shape name (string) or 1-based index (int)"
    )
    color: Optional[str] = Field(
        default=None, description="Line color as '#RRGGBB'"
    )
    weight: Optional[float] = Field(
        default=None, description="Line weight in points"
    )
    dash_style: Optional[str] = Field(
        default=None,
        description="'solid', 'round_dot', 'dash', 'dash_dot', or 'long_dash'",
    )
    begin_arrow: Optional[str] = Field(
        default=None,
        description="Begin arrowhead: 'none', 'triangle', 'open', 'stealth', 'diamond', or 'oval'",
    )
    end_arrow: Optional[str] = Field(
        default=None,
        description="End arrowhead: 'none', 'triangle', 'open', 'stealth', 'diamond', or 'oval'",
    )


# ---------------------------------------------------------------------------
# Helper: find a shape by name or index
# ---------------------------------------------------------------------------
def _get_shape(slide, name_or_index: Union[str, int]):
    """Find a shape on a slide by name (str) or 1-based index (int).

    Args:
        slide: Slide COM object
        name_or_index: Shape name (str) or 1-based index (int)

    Returns:
        Shape COM object

    Raises:
        ValueError: If shape not found or index out of range
    """
    if isinstance(name_or_index, int):
        if name_or_index < 1 or name_or_index > slide.Shapes.Count:
            raise ValueError(
                f"Shape index {name_or_index} out of range "
                f"(1-{slide.Shapes.Count})"
            )
        return slide.Shapes(name_or_index)
    else:
        for i in range(1, slide.Shapes.Count + 1):
            if slide.Shapes(i).Name == name_or_index:
                return slide.Shapes(i)
        raise ValueError(f"Shape '{name_or_index}' not found on slide")


# ---------------------------------------------------------------------------
# COM implementation functions (run on COM thread via ppt.execute)
# ---------------------------------------------------------------------------
def _add_connector_impl(slide_index, connector_type, begin_shape, begin_site,
                         end_shape, end_site):
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    # Resolve connector type
    type_key = connector_type.strip().lower()
    type_int = CONNECTOR_TYPE_MAP.get(type_key)
    if type_int is None:
        raise ValueError(
            f"Unknown connector_type '{connector_type}'. "
            f"Valid values: {list(CONNECTOR_TYPE_MAP.keys())}"
        )

    # Create connector with dummy coordinates (will be repositioned by connections)
    connector = slide.Shapes.AddConnector(type_int, 0, 0, 100, 100)

    # Find begin and end shapes
    begin = _get_shape(slide, begin_shape)
    end = _get_shape(slide, end_shape)

    # Connect using positional args (keyword args unreliable with late binding)
    connector.ConnectorFormat.BeginConnect(begin, begin_site)
    connector.ConnectorFormat.EndConnect(end, end_site)
    connector.RerouteConnections()

    return {
        "success": True,
        "shape_name": connector.Name,
        "connector_type": type_key,
    }


def _format_connector_impl(slide_index, shape_name_or_index,
                             color, weight, dash_style,
                             begin_arrow, end_arrow):
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    line = shape.Line

    if color is not None:
        line.ForeColor.RGB = hex_to_int(color)

    if weight is not None:
        line.Weight = weight

    if dash_style is not None:
        dash_val = DASH_STYLE_MAP.get(dash_style.strip().lower())
        if dash_val is None:
            raise ValueError(
                f"Unknown dash_style '{dash_style}'. "
                f"Valid values: {list(DASH_STYLE_MAP.keys())}"
            )
        line.DashStyle = dash_val

    if begin_arrow is not None:
        arrow_val = ARROW_STYLE_MAP.get(begin_arrow.strip().lower())
        if arrow_val is None:
            raise ValueError(
                f"Unknown begin_arrow '{begin_arrow}'. "
                f"Valid values: {list(ARROW_STYLE_MAP.keys())}"
            )
        line.BeginArrowheadStyle = arrow_val

    if end_arrow is not None:
        arrow_val = ARROW_STYLE_MAP.get(end_arrow.strip().lower())
        if arrow_val is None:
            raise ValueError(
                f"Unknown end_arrow '{end_arrow}'. "
                f"Valid values: {list(ARROW_STYLE_MAP.keys())}"
            )
        line.EndArrowheadStyle = arrow_val

    return {
        "success": True,
        "shape_name": shape.Name,
    }


# ---------------------------------------------------------------------------
# MCP tool functions (async wrappers that delegate to COM thread)
# ---------------------------------------------------------------------------
def add_connector(params: AddConnectorInput) -> str:
    """Add a connector between two shapes.

    Args:
        params: Connector parameters including type, begin/end shapes and sites.

    Returns:
        JSON with connector shape name and type.
    """
    try:
        result = ppt.execute(
            _add_connector_impl,
            params.slide_index, params.connector_type,
            params.begin_shape, params.begin_site,
            params.end_shape, params.end_site,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to add connector: {str(e)}"})


def format_connector(params: FormatConnectorInput) -> str:
    """Format a connector's line properties.

    Args:
        params: Connector identifier and optional line formatting properties.

    Returns:
        JSON confirming the format update.
    """
    try:
        result = ppt.execute(
            _format_connector_impl,
            params.slide_index, params.shape_name_or_index,
            params.color, params.weight, params.dash_style,
            params.begin_arrow, params.end_arrow,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to format connector: {str(e)}"})


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register all connector tools with the MCP server."""

    @mcp.tool(
        name="ppt_add_connector",
        annotations={
            "title": "Add Connector",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_add_connector(params: AddConnectorInput) -> str:
        """Add a connector between two shapes.

        Creates a connector of the specified type (straight, elbow, or curve)
        and attaches it to connection sites on the begin and end shapes.
        Connection sites are 1-based indices on the shape perimeter.
        """
        return add_connector(params)

    @mcp.tool(
        name="ppt_format_connector",
        annotations={
            "title": "Format Connector",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_format_connector(params: FormatConnectorInput) -> str:
        """Format a connector's line properties.

        Configure color, weight, dash style, and arrowheads (begin/end).
        Identify the connector by shape name or 1-based shape index.
        """
        return format_connector(params)
