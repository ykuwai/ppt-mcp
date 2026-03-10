"""Connector tools for PowerPoint COM automation.

Handles adding connectors between shapes and formatting connector lines
including color, weight, dash style, and arrowheads.
"""

import json
import logging
from typing import Optional, Union

from pydantic import BaseModel, Field, ConfigDict, model_validator

from utils.com_wrapper import ppt
from utils.color import hex_to_int
from utils.navigation import goto_slide
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

ARROW_LENGTH_MAP: dict[str, int] = {
    "short": 1,    # msoArrowheadShort
    "medium": 2,   # msoArrowheadLengthMedium
    "long": 3,     # msoArrowheadLong
}

ARROW_WIDTH_MAP: dict[str, int] = {
    "narrow": 1,   # msoArrowheadNarrow
    "medium": 2,   # msoArrowheadWidthMedium
    "wide": 3,     # msoArrowheadWide
}

DASH_STYLE_MAP: dict[str, int] = {
    "solid": msoLineSolid,
    "round_dot": msoLineRoundDot,
    "dash": msoLineDash,
    "dash_dot": msoLineDashDot,
    "long_dash": msoLineLongDash,
}

# Friendly names for connection sites.
# Maps a direction name to a unit vector (dx, dy) used to find the closest
# connection site on a shape.  Coordinates grow right (+x) and down (+y).
SITE_DIRECTION_VECTORS: dict[str, tuple[float, float]] = {
    "top": (0, -1),
    "bottom": (0, 1),
    "left": (-1, 0),
    "right": (1, 0),
}

VALID_SITE_NAMES = list(SITE_DIRECTION_VECTORS.keys())


def _resolve_site(shape, site: Union[int, str]) -> int:
    """Resolve a connection site value to a 1-based integer index.

    If *site* is already an int it is returned as-is.  If it is a string
    (e.g. "top", "right"), the function inspects the shape's ConnectionSites
    and returns the index of the site closest to the named direction relative
    to the shape's centre.

    Args:
        shape: COM Shape object that has ConnectionSites.
        site: 1-based index (int) or direction name (str).

    Returns:
        1-based connection site index.

    Raises:
        ValueError: If the name is unrecognised or the shape has no sites.
    """
    if isinstance(site, int):
        return site

    name = site.strip().lower()
    vec = SITE_DIRECTION_VECTORS.get(name)
    if vec is None:
        raise ValueError(
            f"Unknown connection site name '{site}'. "
            f"Valid names: {VALID_SITE_NAMES}"
        )

    # Shape centre in slide coordinates
    cx = shape.Left + shape.Width / 2
    cy = shape.Top + shape.Height / 2

    count = shape.ConnectionSites.Count
    if count == 0:
        raise ValueError(
            f"Shape '{shape.Name}' has no connection sites"
        )

    best_index = 1
    best_dot = float("-inf")
    for i in range(1, count + 1):
        # ConnectionSites(i) returns (x, y) tuple in slide coordinates
        pt = shape.ConnectionSites(i)
        dx = pt[0] - cx
        dy = pt[1] - cy
        # Normalise to avoid bias towards distant sites
        length = (dx * dx + dy * dy) ** 0.5
        if length == 0:
            continue
        dot = (dx / length) * vec[0] + (dy / length) * vec[1]
        if dot > best_dot:
            best_dot = dot
            best_index = i

    return best_index


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
    begin_site: Union[int, str] = Field(
        default=1,
        description=(
            "Connection site on the begin shape. "
            "Either a 1-based index (int) or a direction name: "
            "'top', 'bottom', 'left', 'right'"
        ),
    )
    end_shape: str = Field(
        ..., description="Name of the shape where the connector ends"
    )
    end_site: Union[int, str] = Field(
        default=1,
        description=(
            "Connection site on the end shape. "
            "Either a 1-based index (int) or a direction name: "
            "'top', 'bottom', 'left', 'right'"
        ),
    )

    @model_validator(mode="after")
    def validate_sites(self) -> "AddConnectorInput":
        for field_name in ("begin_site", "end_site"):
            val = getattr(self, field_name)
            if isinstance(val, int) and val < 1:
                raise ValueError(f"{field_name} must be >= 1 when specified as int")
            if isinstance(val, str) and val.strip().lower() not in VALID_SITE_NAMES:
                raise ValueError(
                    f"Unknown {field_name} name '{val}'. "
                    f"Valid names: {VALID_SITE_NAMES}"
                )
        return self


class FormatConnectorInput(BaseModel):
    """Input for formatting a connector's line properties."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Connector shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
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
    begin_arrow_length: Optional[str] = Field(
        default=None,
        description="Begin arrowhead length: 'short', 'medium', or 'long'",
    )
    begin_arrow_width: Optional[str] = Field(
        default=None,
        description="Begin arrowhead width: 'narrow', 'medium', or 'wide'",
    )
    end_arrow: Optional[str] = Field(
        default=None,
        description="End arrowhead: 'none', 'triangle', 'open', 'stealth', 'diamond', or 'oval'",
    )
    end_arrow_length: Optional[str] = Field(
        default=None,
        description="End arrowhead length: 'short', 'medium', or 'long'",
    )
    end_arrow_width: Optional[str] = Field(
        default=None,
        description="End arrowhead width: 'narrow', 'medium', or 'wide'",
    )
    begin_shape: Optional[str] = Field(
        default=None,
        description="Reconnect the start to this shape (by name)",
    )
    begin_site: Optional[Union[int, str]] = Field(
        default=None,
        description=(
            "Connection site on the new begin shape. "
            "Either a 1-based index (int) or a direction name: "
            "'top', 'bottom', 'left', 'right'. Defaults to 1 if omitted"
        ),
    )
    end_shape: Optional[str] = Field(
        default=None,
        description="Reconnect the end to this shape (by name)",
    )
    end_site: Optional[Union[int, str]] = Field(
        default=None,
        description=(
            "Connection site on the new end shape. "
            "Either a 1-based index (int) or a direction name: "
            "'top', 'bottom', 'left', 'right'. Defaults to 1 if omitted"
        ),
    )

    @model_validator(mode="after")
    def check_site_requires_shape(self) -> "FormatConnectorInput":
        if self.begin_site is not None and self.begin_shape is None:
            raise ValueError("begin_site requires begin_shape to be set")
        if self.end_site is not None and self.end_shape is None:
            raise ValueError("end_site requires end_shape to be set")
        for field_name in ("begin_site", "end_site"):
            val = getattr(self, field_name)
            if val is None:
                continue
            if isinstance(val, int) and val < 1:
                raise ValueError(f"{field_name} must be >= 1 when specified as int")
            if isinstance(val, str) and val.strip().lower() not in VALID_SITE_NAMES:
                raise ValueError(
                    f"Unknown {field_name} name '{val}'. "
                    f"Valid names: {VALID_SITE_NAMES}"
                )
        return self


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
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
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

    # Resolve friendly site names to 1-based indices
    resolved_begin = _resolve_site(begin, begin_site)
    resolved_end = _resolve_site(end, end_site)

    # Connect using positional args (keyword args unreliable with late binding)
    connector.ConnectorFormat.BeginConnect(begin, resolved_begin)
    connector.ConnectorFormat.EndConnect(end, resolved_end)
    connector.RerouteConnections()

    return {
        "success": True,
        "shape_name": connector.Name,
        "connector_type": type_key,
    }


def _format_connector_impl(slide_index, shape_name_or_index,
                             color, weight, dash_style,
                             begin_arrow, begin_arrow_length, begin_arrow_width,
                             end_arrow, end_arrow_length, end_arrow_width,
                             begin_shape, begin_site,
                             end_shape, end_site):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
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

    if begin_arrow_length is not None:
        length_val = ARROW_LENGTH_MAP.get(begin_arrow_length.strip().lower())
        if length_val is None:
            raise ValueError(
                f"Unknown begin_arrow_length '{begin_arrow_length}'. "
                f"Valid values: {list(ARROW_LENGTH_MAP.keys())}"
            )
        line.BeginArrowheadLength = length_val

    if begin_arrow_width is not None:
        width_val = ARROW_WIDTH_MAP.get(begin_arrow_width.strip().lower())
        if width_val is None:
            raise ValueError(
                f"Unknown begin_arrow_width '{begin_arrow_width}'. "
                f"Valid values: {list(ARROW_WIDTH_MAP.keys())}"
            )
        line.BeginArrowheadWidth = width_val

    if end_arrow is not None:
        arrow_val = ARROW_STYLE_MAP.get(end_arrow.strip().lower())
        if arrow_val is None:
            raise ValueError(
                f"Unknown end_arrow '{end_arrow}'. "
                f"Valid values: {list(ARROW_STYLE_MAP.keys())}"
            )
        line.EndArrowheadStyle = arrow_val

    if end_arrow_length is not None:
        length_val = ARROW_LENGTH_MAP.get(end_arrow_length.strip().lower())
        if length_val is None:
            raise ValueError(
                f"Unknown end_arrow_length '{end_arrow_length}'. "
                f"Valid values: {list(ARROW_LENGTH_MAP.keys())}"
            )
        line.EndArrowheadLength = length_val

    if end_arrow_width is not None:
        width_val = ARROW_WIDTH_MAP.get(end_arrow_width.strip().lower())
        if width_val is None:
            raise ValueError(
                f"Unknown end_arrow_width '{end_arrow_width}'. "
                f"Valid values: {list(ARROW_WIDTH_MAP.keys())}"
            )
        line.EndArrowheadWidth = width_val

    # Reconnect begin/end to different shapes
    reroute = False
    if begin_shape is not None:
        target = _get_shape(slide, begin_shape)
        site = _resolve_site(target, begin_site) if begin_site is not None else 1
        shape.ConnectorFormat.BeginConnect(target, site)
        reroute = True

    if end_shape is not None:
        target = _get_shape(slide, end_shape)
        site = _resolve_site(target, end_site) if end_site is not None else 1
        shape.ConnectorFormat.EndConnect(target, site)
        reroute = True

    if reroute:
        shape.RerouteConnections()

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
            params.begin_arrow, params.begin_arrow_length, params.begin_arrow_width,
            params.end_arrow, params.end_arrow_length, params.end_arrow_width,
            params.begin_shape, params.begin_site,
            params.end_shape, params.end_site,
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
        Connection sites can be specified as 1-based indices or direction
        names: 'top', 'bottom', 'left', 'right'.
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
        """Format a connector's line properties and reconnect endpoints.

        Configure color, weight, dash style, arrowheads, and arrowhead size.
        Reconnect begin/end to different shapes via begin_shape/end_shape.
        Connection sites accept 1-based indices or direction names
        ('top', 'bottom', 'left', 'right').
        Identify the connector by shape name or 1-based shape index.
        """
        return format_connector(params)
