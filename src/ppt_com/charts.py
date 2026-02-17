"""Chart operations for PowerPoint COM automation.

Handles creating charts, setting chart data via Excel workbook interaction,
formatting charts (title, legend, style), formatting individual series,
reading chart data back, and changing chart types.

CRITICAL: Every chart.ChartData.Activate() opens an Excel workbook process.
You MUST close it with wb.Close(False) in a finally block to prevent leaked
Excel processes.
"""

import json
import logging
from typing import Optional, Union

import pythoncom
from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from utils.color import hex_to_int, int_to_hex
from utils.navigation import goto_slide
from ppt_com.constants import msoChart

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Friendly name maps
# ---------------------------------------------------------------------------
CHART_TYPE_MAP: dict[str, int] = {
    "column": 51, "column_stacked": 52, "column_stacked_100": 53,
    "bar": 57, "bar_stacked": 58,
    "line": 4, "line_markers": 65, "line_stacked": 63,
    "pie": 5, "pie_exploded": 69, "doughnut": -4120,
    "area": 1, "area_stacked": 76,
    "scatter": -4169, "scatter_lines": 74,
    "radar": -4151, "bubble": 15,
    "3d_column": 54, "3d_pie": -4102, "3d_line": -4101,
}

LEGEND_POSITION_MAP: dict[str, int] = {
    "bottom": -4107, "left": -4131, "right": -4152, "top": -4160, "corner": 2,
}

# Reverse map for chart type display
CHART_TYPE_NAMES: dict[int, str] = {v: k for k, v in CHART_TYPE_MAP.items()}


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class AddChartInput(BaseModel):
    """Input for adding a chart to a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    chart_type: Union[int, str] = Field(
        default="column",
        description=(
            "Chart type as friendly name (e.g. 'column', 'bar', 'line', 'pie', "
            "'scatter', 'area', 'doughnut', '3d_column', etc.) or XlChartType integer"
        ),
    )
    left: float = Field(default=50.0, description="Left position in points")
    top: float = Field(default=50.0, description="Top position in points")
    width: float = Field(default=500.0, description="Width in points")
    height: float = Field(default=350.0, description="Height in points")


class SetChartDataInput(BaseModel):
    """Input for setting chart data via Excel workbook."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Chart shape name (string) or 1-based index (int)"
    )
    categories: list[str] = Field(
        ..., description="List of category labels (e.g. ['Q1', 'Q2', 'Q3', 'Q4'])"
    )
    series: list[dict] = Field(
        ...,
        description=(
            "List of series dicts, each with 'name' (str) and 'values' (list[float]). "
            "Example: [{'name': 'Sales', 'values': [100, 200, 150, 300]}]"
        ),
    )


class GetChartDataInput(BaseModel):
    """Input for reading chart data from the Excel workbook."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Chart shape name (string) or 1-based index (int)"
    )


class FormatChartInput(BaseModel):
    """Input for formatting chart properties (title, legend, style)."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Chart shape name (string) or 1-based index (int)"
    )
    title: Optional[str] = Field(
        default=None, description="Chart title text (sets HasTitle=True automatically)"
    )
    has_legend: Optional[bool] = Field(
        default=None, description="Show or hide the chart legend"
    )
    legend_position: Optional[str] = Field(
        default=None,
        description="Legend position: 'bottom', 'left', 'right', 'top', or 'corner'",
    )
    chart_style: Optional[int] = Field(
        default=None, description="Built-in chart style index (1-48 typically)"
    )


class SetChartSeriesInput(BaseModel):
    """Input for formatting an individual chart series."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Chart shape name (string) or 1-based index (int)"
    )
    series_index: int = Field(..., ge=1, description="1-based series index")
    color: Optional[str] = Field(
        default=None, description="Series fill color as '#RRGGBB'"
    )
    show_data_labels: Optional[bool] = Field(
        default=None, description="Show or hide data labels for this series"
    )
    line_weight: Optional[float] = Field(
        default=None, description="Line weight in points (for line/scatter charts)"
    )


class ChangeChartTypeInput(BaseModel):
    """Input for changing a chart's type."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Chart shape name (string) or 1-based index (int)"
    )
    chart_type: Union[int, str] = Field(
        ...,
        description=(
            "New chart type as friendly name (e.g. 'column', 'bar', 'line', 'pie') "
            "or XlChartType integer"
        ),
    )


# ---------------------------------------------------------------------------
# Helper: resolve chart type from string or int
# ---------------------------------------------------------------------------
def _resolve_chart_type(chart_type: Union[str, int]) -> int:
    """Convert a friendly chart type name to its XlChartType integer value.

    If chart_type is already an int, returns it unchanged.
    If it is a string, looks it up in CHART_TYPE_MAP.

    Raises:
        ValueError: If the string name is not recognized.
    """
    if isinstance(chart_type, int):
        return chart_type
    key = chart_type.strip().lower()
    if key not in CHART_TYPE_MAP:
        raise ValueError(
            f"Unknown chart type '{chart_type}'. "
            f"Valid names: {', '.join(sorted(CHART_TYPE_MAP.keys()))}"
        )
    return CHART_TYPE_MAP[key]


# ---------------------------------------------------------------------------
# Helper: find a chart shape
# ---------------------------------------------------------------------------
def _get_chart_shape(slide, name_or_index: Union[str, int]):
    """Find a shape on a slide and verify it has a chart.

    Args:
        slide: Slide COM object
        name_or_index: Shape name (str) or 1-based index (int)

    Returns:
        Shape COM object that contains a chart

    Raises:
        ValueError: If shape not found or is not a chart
    """
    if isinstance(name_or_index, int):
        if name_or_index < 1 or name_or_index > slide.Shapes.Count:
            raise ValueError(
                f"Shape index {name_or_index} out of range "
                f"(1-{slide.Shapes.Count})"
            )
        shape = slide.Shapes(name_or_index)
    else:
        shape = None
        for i in range(1, slide.Shapes.Count + 1):
            if slide.Shapes(i).Name == name_or_index:
                shape = slide.Shapes(i)
                break
        if shape is None:
            raise ValueError(f"Shape '{name_or_index}' not found on slide")

    if not shape.HasChart:
        raise ValueError(f"Shape '{shape.Name}' is not a chart")
    return shape


# ---------------------------------------------------------------------------
# COM implementation functions (run on COM thread via ppt.execute)
# ---------------------------------------------------------------------------
def _add_chart_impl(slide_index, chart_type, left, top, width, height):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    type_int = _resolve_chart_type(chart_type)
    chart_shape = slide.Shapes.AddChart2(-1, type_int, left, top, width, height)
    chart = chart_shape.Chart

    # CRITICAL: AddChart2 auto-opens an Excel workbook. MUST close it to
    # prevent leaked Excel processes.
    try:
        chart.ChartData.Activate()
        wb = chart.ChartData.Workbook
        wb.Close(False)
    except Exception:
        pass  # Workbook may not always be open

    type_name = CHART_TYPE_NAMES.get(type_int, str(type_int))
    return {
        "success": True,
        "shape_name": chart_shape.Name,
        "shape_index": chart_shape.ZOrderPosition,
        "chart_type": type_name,
        "chart_type_int": type_int,
    }


def _set_chart_data_impl(slide_index, shape_name_or_index, categories, series):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_chart_shape(slide, shape_name_or_index)
    chart = shape.Chart

    chart.ChartData.Activate()
    wb = chart.ChartData.Workbook
    try:
        ws = wb.Worksheets(1)
        ws.Cells.Clear()

        # Write categories in column A starting at row 2
        for i, cat in enumerate(categories):
            ws.Cells(i + 2, 1).Value = cat

        # Write series headers (row 1) and values (rows 2+)
        for s_idx, s in enumerate(series):
            col = s_idx + 2  # columns B, C, D, ...
            ws.Cells(1, col).Value = s["name"]
            for i, val in enumerate(s["values"]):
                ws.Cells(i + 2, col).Value = val

        # Set the chart's source data range.
        # chart.SetSourceData() has a pywin32 VT_BYREF bug on the PlotBy
        # parameter, so we call _oleobj_.InvokeTypes directly with fixed
        # type flags (12, 49) instead of (12, 17) for PlotBy.
        last_row = len(categories) + 1
        last_col = len(series) + 1
        data_range = ws.Range(ws.Cells(1, 1), ws.Cells(last_row, last_col))
        addr = "Sheet1!" + data_range.Address
        chart._oleobj_.InvokeTypes(
            1413, 0, 1, (24, 0),
            ((8, 1), (12, 49)),
            addr, pythoncom.Empty,
        )
    finally:
        wb.Close(False)

    return {
        "success": True,
        "shape_name": shape.Name,
        "categories_count": len(categories),
        "series_count": len(series),
    }


def _get_chart_data_impl(slide_index, shape_name_or_index):
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_chart_shape(slide, shape_name_or_index)
    chart = shape.Chart

    # Read data directly from SeriesCollection (avoids opening Excel workbook)
    categories = []
    series_list = []
    sc_count = chart.SeriesCollection().Count
    for s in range(1, sc_count + 1):
        ser = chart.SeriesCollection(s)
        name = ser.Name
        try:
            values = list(ser.Values)
        except Exception:
            values = []
        series_list.append({"name": name, "values": values})
        if not categories:
            try:
                categories = [str(x) for x in ser.XValues]
            except Exception:
                pass

    return {
        "success": True,
        "shape_name": shape.Name,
        "categories": categories,
        "series": series_list,
    }


def _format_chart_impl(
    slide_index, shape_name_or_index,
    title, has_legend, legend_position, chart_style,
):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_chart_shape(slide, shape_name_or_index)
    chart = shape.Chart

    if title is not None:
        chart.HasTitle = True
        chart.ChartTitle.Text = title

    if has_legend is not None:
        chart.HasLegend = has_legend

    if legend_position is not None:
        if not chart.HasLegend:
            raise ValueError(
                "Cannot set legend position when chart has no legend. "
                "Set has_legend=true first."
            )
        key = legend_position.strip().lower()
        if key not in LEGEND_POSITION_MAP:
            raise ValueError(
                f"Unknown legend position '{legend_position}'. "
                f"Valid: {', '.join(LEGEND_POSITION_MAP.keys())}"
            )
        chart.Legend.Position = LEGEND_POSITION_MAP[key]

    if chart_style is not None:
        chart.ChartStyle = chart_style

    return {
        "success": True,
        "shape_name": shape.Name,
        "has_title": bool(chart.HasTitle),
        "has_legend": bool(chart.HasLegend),
    }


def _set_chart_series_impl(
    slide_index, shape_name_or_index,
    series_index, color, show_data_labels, line_weight,
):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_chart_shape(slide, shape_name_or_index)
    chart = shape.Chart

    series = chart.SeriesCollection(series_index)

    if color is not None:
        series.Format.Fill.ForeColor.RGB = hex_to_int(color)

    if show_data_labels is not None:
        series.HasDataLabels = show_data_labels

    if line_weight is not None:
        series.Format.Line.Weight = line_weight

    return {
        "success": True,
        "shape_name": shape.Name,
        "series_index": series_index,
    }


def _change_chart_type_impl(slide_index, shape_name_or_index, chart_type):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_chart_shape(slide, shape_name_or_index)
    chart = shape.Chart

    type_int = _resolve_chart_type(chart_type)
    chart.ChartType = type_int

    type_name = CHART_TYPE_NAMES.get(type_int, str(type_int))
    return {
        "success": True,
        "shape_name": shape.Name,
        "new_chart_type": type_name,
        "new_chart_type_int": type_int,
    }


# ---------------------------------------------------------------------------
# MCP tool functions (sync wrappers that delegate to COM thread)
# ---------------------------------------------------------------------------
def add_chart(params: AddChartInput) -> str:
    """Add a chart to a slide.

    Args:
        params: Chart parameters including type, position, and size.

    Returns:
        JSON with shape name, index, and chart type.
    """
    try:
        result = ppt.execute(
            _add_chart_impl,
            params.slide_index, params.chart_type,
            params.left, params.top, params.width, params.height,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to add chart: {str(e)}"})


def set_chart_data(params: SetChartDataInput) -> str:
    """Set chart data by writing to the underlying Excel workbook.

    Args:
        params: Slide index, shape identifier, categories, and series data.

    Returns:
        JSON with categories count and series count.
    """
    try:
        result = ppt.execute(
            _set_chart_data_impl,
            params.slide_index, params.shape_name_or_index,
            params.categories, params.series,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to set chart data: {str(e)}"})


def get_chart_data(params: GetChartDataInput) -> str:
    """Read chart data from the underlying Excel workbook.

    Args:
        params: Slide index and chart shape identifier.

    Returns:
        JSON with categories list and series list (name + values).
    """
    try:
        result = ppt.execute(
            _get_chart_data_impl,
            params.slide_index, params.shape_name_or_index,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to get chart data: {str(e)}"})


def format_chart(params: FormatChartInput) -> str:
    """Format chart properties such as title, legend, and style.

    Args:
        params: Chart identifier and optional formatting properties.

    Returns:
        JSON confirming the formatting changes.
    """
    try:
        result = ppt.execute(
            _format_chart_impl,
            params.slide_index, params.shape_name_or_index,
            params.title, params.has_legend,
            params.legend_position, params.chart_style,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to format chart: {str(e)}"})


def set_chart_series(params: SetChartSeriesInput) -> str:
    """Format an individual chart series (color, data labels, line weight).

    Args:
        params: Series index and optional formatting properties.

    Returns:
        JSON confirming the series formatting.
    """
    try:
        result = ppt.execute(
            _set_chart_series_impl,
            params.slide_index, params.shape_name_or_index,
            params.series_index, params.color,
            params.show_data_labels, params.line_weight,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to set chart series: {str(e)}"})


def change_chart_type(params: ChangeChartTypeInput) -> str:
    """Change the chart type of an existing chart.

    Args:
        params: Chart identifier and new chart type.

    Returns:
        JSON confirming the type change.
    """
    try:
        result = ppt.execute(
            _change_chart_type_impl,
            params.slide_index, params.shape_name_or_index,
            params.chart_type,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to change chart type: {str(e)}"})


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register all chart tools with the MCP server."""

    @mcp.tool(
        name="ppt_add_chart",
        annotations={
            "title": "Add Chart",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_add_chart(params: AddChartInput) -> str:
        """Add a chart to a slide.

        Creates a chart with the specified type and position.
        Supported types: column, bar, line, pie, scatter, area, doughnut, radar,
        bubble, 3d_column, 3d_pie, 3d_line, and variants (e.g. column_stacked,
        line_markers, pie_exploded). You can also pass an XlChartType integer.
        All positions and sizes are in points (72 points = 1 inch).
        """
        return add_chart(params)

    @mcp.tool(
        name="ppt_set_chart_data",
        annotations={
            "title": "Set Chart Data",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_set_chart_data(params: SetChartDataInput) -> str:
        """Set chart data by writing categories and series to the Excel workbook.

        Provide category labels and one or more series, each with a name and
        numeric values. The number of values in each series should match the
        number of categories.
        Example series: [{"name": "Revenue", "values": [100, 200, 150, 300]}]
        """
        return set_chart_data(params)

    @mcp.tool(
        name="ppt_get_chart_data",
        annotations={
            "title": "Get Chart Data",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_get_chart_data(params: GetChartDataInput) -> str:
        """Read chart data from the underlying Excel workbook.

        Returns the category labels and all series (name + values).
        Identify the chart by shape name or 1-based shape index.
        """
        return get_chart_data(params)

    @mcp.tool(
        name="ppt_format_chart",
        annotations={
            "title": "Format Chart",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_format_chart(params: FormatChartInput) -> str:
        """Format chart properties: title, legend, and chart style.

        Set a chart title (enables HasTitle automatically), show/hide the legend,
        set legend position ('bottom', 'left', 'right', 'top', 'corner'),
        and apply a built-in chart style by index number.
        """
        return format_chart(params)

    @mcp.tool(
        name="ppt_set_chart_series",
        annotations={
            "title": "Set Chart Series Format",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_set_chart_series(params: SetChartSeriesInput) -> str:
        """Format an individual chart series.

        Set the fill color (as '#RRGGBB'), toggle data labels, and set
        line weight (in points, for line/scatter charts).
        Series are 1-based indexed.
        """
        return set_chart_series(params)

    @mcp.tool(
        name="ppt_change_chart_type",
        annotations={
            "title": "Change Chart Type",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_change_chart_type(params: ChangeChartTypeInput) -> str:
        """Change the type of an existing chart.

        Accepts a friendly name (e.g. 'bar', 'line', 'pie') or XlChartType integer.
        The chart data is preserved when changing types.
        """
        return change_chart_type(params)
