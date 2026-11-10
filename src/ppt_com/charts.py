"""Chart operations for PowerPoint COM automation.

Handles creating charts, setting chart data via Excel workbook interaction,
formatting charts (title, legend, style), formatting individual series,
reading chart data back, and changing chart types.

A new chart of a type gvml can write is written without Excel and pasted
(ppt_com/chart_defaults.py, #299), and its data is swapped the same way where
it can be. Every other way in, AddChart2 and chart.ChartData.Activate(),
opens an Excel workbook process.
You MUST close it with wb.Close(False) in a finally block to prevent leaked
Excel processes.
"""

import contextvars
import json
import logging
import sys
import time
from typing import Literal, Optional, Union

# pythoncom does not exist on macOS, and this module has to stay importable
# there so the Apple Event implementations at the bottom can take the place of
# the COM ones (#185, the same guard ppt_com/export.py already carries). It is
# used only by the Windows impls, which are replaced wholesale on macOS and
# never run there.
if sys.platform == "win32":
    import pythoncom
else:
    pythoncom = None

from pydantic import BaseModel, Field, ConfigDict, model_validator

from utils import clipboard
from utils.clipboard import kept_clipboard
from utils.held_refs import held, window_process
from utils.offload import run_offloaded
from backend import ppt, vba
from ppt_com.layout_check import report_layout, touched_names
from utils.color import hex_to_int, int_to_hex
from utils.navigation import goto_slide
from ppt_com.shape_lookup import resolve_shape
from ppt_com.constants import msoChart, msoGroup, msoPlaceholder
from ppt_com import animation_carry, chart_defaults
from gvml import charts as gvml_charts
from gvml import package as gvml_package
from gvml.shapes import strip_package_ids

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

# XlAxisType constants
AXIS_TYPE_MAP: dict[str, int] = {
    "category": 1,         # xlCategory
    "value": 2,            # xlValue
    "secondary_value": 2,  # xlValue with xlSecondary group
    "series": 3,           # xlSeriesAxis (3D charts only)
}

# XlAxisGroup constants
AXIS_GROUP_PRIMARY = 1    # xlPrimary
AXIS_GROUP_SECONDARY = 2  # xlSecondary

# XlTickMark constants
TICK_MARK_MAP: dict[str, int] = {
    "none": -4142,    # xlTickMarkNone
    "inside": 2,      # xlTickMarkInside
    "outside": 3,     # xlTickMarkOutside
    "cross": 4,       # xlTickMarkCross
}

# XlScaleType constants
SCALE_LINEAR = -4132       # xlLinear
SCALE_LOGARITHMIC = -4133  # xlLogarithmic

# Reverse map for chart type display
CHART_TYPE_NAMES: dict[int, str] = {v: k for k, v in CHART_TYPE_MAP.items()}


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class AddChartInput(BaseModel):
    """Input for adding a chart to a slide."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

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
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Chart shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    categories: list[str] = Field(
        ..., description="List of category labels (e.g. ['Q1', 'Q2', 'Q3', 'Q4'])"
    )
    series: list[dict] = Field(
        ...,
        description=(
            "List of series dicts, each with 'name' (str) and 'values' (list[float]). "
            "Example: [{'name': 'Sales', 'values': [100, 200, 150, 300]}]. "
            "For a scatter or bubble chart the categories are the x values, and "
            "a bubble chart's series also need 'sizes' (list[float])."
        ),
    )

    @model_validator(mode="after")
    def _one_size_per_point(self):
        for entry in self.series:
            if "sizes" in entry and len(entry["sizes"]) != len(self.categories):
                raise ValueError(
                    f"series {entry.get('name')!r} has {len(entry['sizes'])} sizes "
                    f"for {len(self.categories)} categories."
                )
        return self


class GetChartDataInput(BaseModel):
    """Input for reading chart data from the Excel workbook."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Chart shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )


class FormatChartInput(BaseModel):
    """Input for formatting chart properties (title, legend, style)."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Chart shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    title: Optional[str] = Field(
        default=None, description="Chart title text (sets HasTitle=True automatically)"
    )
    has_legend: Optional[bool] = Field(
        default=None, description="Show or hide the chart legend"
    )
    legend_position: Optional[str] = Field(
        default=None,
        description=(
            "Legend position. PowerPoint presets: 'bottom', 'left', 'right', 'top', 'corner'. "
            "8-direction presets (coordinate-based): 'top-left', 'top-center', 'top-right', "
            "'middle-left', 'middle-right', 'bottom-left', 'bottom-center', 'bottom-right'. "
            "Applied before legend_top/legend_left."
        ),
    )
    chart_style: Optional[int] = Field(
        default=None, description="Built-in chart style index (1-48 typically)"
    )
    legend_font_size: Optional[float] = Field(
        default=None, description="Legend font size in points", gt=0
    )
    legend_top: Optional[float] = Field(
        default=None, ge=0,
        description="Legend top position in points (relative to chart area). Overrides legend_position."
    )
    legend_left: Optional[float] = Field(
        default=None, ge=0,
        description="Legend left position in points (relative to chart area). Overrides legend_position."
    )
    title_position: Optional[Literal["top", "bottom", "center"]] = Field(
        default=None,
        description=(
            "Chart title position preset: 'top' (restores PowerPoint auto-placement), "
            "'bottom' (below plot area, above legend), "
            "'center' (vertically centered, overlaps plot area). "
            "Applied before title_top/title_left."
        ),
    )
    title_top: Optional[float] = Field(
        default=None, ge=0,
        description="Chart title top position in points (relative to chart area). Overrides title_position."
    )
    title_left: Optional[float] = Field(
        default=None, ge=0,
        description="Chart title left position in points (relative to chart area). Overrides title_position."
    )


class FormatChartAxisInput(BaseModel):
    """Input for formatting a chart axis (scale, ticks, labels, title)."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Chart shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    axis: Literal["category", "value", "secondary_value", "series"] = Field(
        default="category",
        description=(
            "Which axis to format. 'category' = horizontal x-axis on column/line/bar; "
            "'value' = primary y-axis; 'secondary_value' = secondary y-axis; "
            "'series' = 3D charts only."
        ),
    )
    title: Optional[str] = Field(
        default=None, description="Axis title text (sets HasTitle=True automatically)"
    )
    min_scale: Optional[float] = Field(
        default=None,
        description="MinimumScale. Value-axis only; raises an error if set on a category axis.",
    )
    max_scale: Optional[float] = Field(
        default=None,
        description="MaximumScale. Value-axis only; raises an error if set on a category axis.",
    )
    major_unit: Optional[float] = Field(
        default=None, gt=0,
        description=(
            "Distance between major tick marks / gridlines (in axis units). "
            "Value-axis only; raises an error on category or series axes."
        ),
    )
    minor_unit: Optional[float] = Field(
        default=None, gt=0,
        description=(
            "Distance between minor tick marks (in axis units). "
            "Value-axis only; raises an error on category or series axes."
        ),
    )
    tick_label_spacing: Optional[int] = Field(
        default=None, ge=1,
        description=(
            "TickLabelSpacing — show every Nth category label. "
            "Category-axis only; raises an error otherwise. Example: 3 on a 0–60 day axis "
            "renders labels at 0, 3, 6, … 60."
        ),
    )
    tick_mark_spacing: Optional[int] = Field(
        default=None, ge=1,
        description="TickMarkSpacing — draw a tick every Nth category. Category-axis only.",
    )
    major_tick_mark: Optional[Literal["none", "inside", "outside", "cross"]] = Field(
        default=None, description="MajorTickMark style."
    )
    minor_tick_mark: Optional[Literal["none", "inside", "outside", "cross"]] = Field(
        default=None, description="MinorTickMark style."
    )
    reverse_order: Optional[bool] = Field(
        default=None,
        description="ReversePlotOrder — reverse the direction of the axis.",
    )
    log_scale: Optional[bool] = Field(
        default=None,
        description="Switch the axis between linear and logarithmic scale. Value-axis only.",
    )
    log_base: Optional[float] = Field(
        default=None, gt=1,
        description="LogBase (default 10). Only meaningful when log_scale=true.",
    )
    tick_label_font_size: Optional[float] = Field(
        default=None, gt=0,
        description="Font size of the axis tick labels, in points.",
    )
    number_format: Optional[str] = Field(
        default=None,
        description=(
            "Excel-style number format for tick labels (e.g. '0', '0.0%', '#,##0', "
            "'yyyy/m/d'). Applied to TickLabels.NumberFormat."
        ),
    )


class SetChartSeriesInput(BaseModel):
    """Input for formatting an individual chart series."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Chart shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
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
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Chart shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
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
    shape = resolve_shape(slide, name_or_index)

    if not shape.HasChart:
        raise ValueError(f"Shape '{shape.Name}' is not a chart")
    return shape


# ---------------------------------------------------------------------------
# COM implementation functions (run on COM thread via ppt.execute)
# ---------------------------------------------------------------------------
# The Mac's VBA bridge puts away the Excel a chart's workbook brought up
# (backend/mac_excel.py), but only after a call that opened one. It sets this
# to its record of the call, and _close_chart_workbook, which every workbook
# opened here goes through, notes on it what happened.
excel_call = contextvars.ContextVar("excel_call", default=None)

WORKBOOK_LEFT_OPEN = (
    "The chart's data workbook could not be closed, so Excel's 'Chart in "
    "Microsoft PowerPoint' window may still be open. Close it before adding or "
    "editing another chart, or PowerPoint will refuse or wait for it."
)


def _wait_pumping(seconds) -> None:
    """Wait on the STA thread while still answering COM calls into it."""
    end = time.monotonic() + seconds
    while True:
        if pythoncom is not None:
            pythoncom.PumpWaitingMessages()
        left = end - time.monotonic()
        if left <= 0:
            return
        time.sleep(min(0.05, left))


class _ExcelRefs:
    """The Excel objects one tool call got, and the Excel process behind them.

    Handed to ``held`` once the workbook is closed, so that letting go of them
    waits for nothing (#299). The process is looked up while the workbook is
    still open, from its Excel's main window.
    """

    def __init__(self):
        self.objects, self.pid = [], None

    def add(self, obj):
        self.objects.append(obj)
        if self.pid is None:
            try:
                self.pid = window_process(obj.Application.Hwnd)
            except Exception:  # noqa: BLE001 - then they are let go of at once
                logger.debug("Could not find the Excel behind a chart's workbook", exc_info=True)
        return obj

    def hand_over(self):
        held.hold(self.objects, self.pid)
        self.objects = []


def _close_chart_workbook(chart, workbook=None, attempts=5, wait=0.5, keep=None) -> bool:
    """Close a chart's data workbook, and say whether it closed.

    A workbook left open keeps the chart in edit mode: the next chart PowerPoint
    is asked to insert or edit answers with a dialog asking for Excel's dialogs
    to be closed, or waits for it, and every call after that times out. Excel
    still starting up can answer with an error, or hand back a workbook its type
    information does not describe yet, so that ``Close`` is not even found. So
    the workbook is asked for again and closed again, a few times, before giving
    up; giving up is logged and reported, never silent. Closing a workbook that
    is already closed does no harm, and the chart hands out its workbook object
    even then, so a failure to get it is retried rather than taken as closed.

    ``workbook`` is the one the caller holds, or None to open it first, as
    ``ChartData.Activate`` does for a chart just added. Each workbook got here
    goes into ``keep``, an ``_ExcelRefs``, when one is given.
    """
    call = excel_call.get()
    if call is not None:
        call.opened = True
    last = None
    for attempt in range(attempts):
        if attempt:
            _wait_pumping(wait)
        try:
            if workbook is None or attempt:
                if workbook is None and not attempt:
                    chart.ChartData.Activate()
                workbook = chart.ChartData.Workbook
                if keep is not None:
                    keep.add(workbook)
            workbook.Close(False)
            return True
        except Exception as e:  # noqa: BLE001 - retried, then reported
            # Only its text: the error's frames would hold the workbook.
            last = f"{type(e).__name__}: {e}"
    logger.warning("Could not close a chart's data workbook after %d tries: %s", attempts, last)
    if call is not None:
        call.left_open = True
    return False


def _office_language(app) -> str:
    """The language the sample data of a new chart is in, from Office's UI."""
    try:
        return chart_defaults.language(app.LanguageSettings.LanguageID(2))  # msoLanguageIDUI
    except Exception:  # noqa: BLE001 - English then
        return "en"


GVML_FORMAT = "Art::GVML ClipFormat"
EMU_PER_POINT = 12700


def _paste_from(copy, slide, attempts=10, wait=0.1):
    """Put a shape on the clipboard with ``copy()`` and paste it onto
    ``slide``, answering the pasted shape. Right after a copy the clipboard
    can still be empty, so the paste is tried again; but PowerPoint can also
    raise after the shape has landed, so a slide that gained a shape is not
    pasted on twice. The clipboard is emptied first: a paste that comes
    before the copy has landed then finds nothing, rather than what was on
    the clipboard before, which it would paste as a shape of its own. The
    caller puts the user's clipboard back."""
    clipboard.empty()
    copy()
    count = slide.Shapes.Count
    last = None
    for attempt in range(attempts):
        if attempt:
            _wait_pumping(wait)
        try:
            return slide.Shapes.Paste()(1)
        except Exception as e:  # noqa: BLE001 - tried again, then raised
            last = e
            try:
                now = slide.Shapes.Count
            except Exception:  # noqa: BLE001 - then the paste is tried again
                logger.debug("Could not count the shapes after a failed paste", exc_info=True)
                continue
            if now > count:
                return slide.Shapes(now)
    raise RuntimeError(f"Could not paste the chart: {last}")


def _copied_package(shape, attempts=10, wait=0.1):
    """Copy ``shape`` and answer the GVML package PowerPoint puts on the
    clipboard for it. The caller puts the user's clipboard back."""
    clipboard.empty()
    shape.Copy()
    for attempt in range(attempts):
        if attempt:
            _wait_pumping(wait)
        raw = clipboard.get(GVML_FORMAT)
        if raw:
            return gvml_package.read(raw)
    raise RuntimeError(f"Copying the chart put no {GVML_FORMAT} on the clipboard")


def _unique_name(slide, shape) -> None:
    """Rename a pasted chart that took the name of a shape already there."""
    names = [slide.Shapes(i).Name for i in range(1, slide.Shapes.Count + 1)]
    if names.count(shape.Name) < 2:
        return
    number = 1
    while f"Chart {number}" in names:
        number += 1
    shape.Name = f"Chart {number}"


def _next_chart_name(slide) -> str:
    names = {slide.Shapes(i).Name for i in range(1, slide.Shapes.Count + 1)}
    number = 1
    while f"Chart {number}" in names:
        number += 1
    return f"Chart {number}"


def _add_chart_by_paste(app, slide, type_int, left, top, width, height):
    """Add a chart written by gvml, so that no Excel starts (#299).

    The chart's XML is written with the sample data AddChart2 gives the type
    (``chart_defaults``) and pasted as ``Art::GVML ClipFormat``, with the
    user's clipboard put back afterwards. The pasted shape is read back,
    and one that is no chart or not of ``type_int`` is removed as a failure.
    Answers the shape, its chart type as read back, and any warnings:
    clipboard formats that could not be put back, or a style that did not
    take.
    """
    raw = gvml_charts.chart_package(
        _next_chart_name(slide),
        round(left * EMU_PER_POINT), round(top * EMU_PER_POINT),
        round(width * EMU_PER_POINT), round(height * EMU_PER_POINT),
        chart_defaults.chart_xml(type_int, _office_language(app)),
    )
    with kept_clipboard() as lost:
        shape = _paste_from(lambda: clipboard.put(GVML_FORMAT, raw), slide)
    style = chart_defaults.STYLES.get(type_int)
    try:
        # What landed is read back before anything else is done to it. A
        # paste lands in the middle of the slide, so it is put in its place.
        if not shape.HasChart:
            raise RuntimeError("what was pasted is not a chart")
        chart = shape.Chart
        type_now = chart.ChartType
        if type_now != type_int:
            raise RuntimeError(f"the chart came in as type {type_now}, not {type_int}")
        shape.Left, shape.Top, shape.Width, shape.Height = left, top, width, height
        if style:
            chart.ChartStyle = style
            chart.ClearToMatchStyle()
        _unique_name(slide, shape)
        style_now = chart.ChartStyle
    except Exception as e:
        # A half made chart is not left behind, where a retry would add a
        # second one.
        name = None
        try:
            name = shape.Name
            shape.Delete()
        except Exception:  # noqa: BLE001 - then the error names it
            raise RuntimeError(
                f"The chart was pasted but could not be finished ({e}), and "
                f"could not be removed either: delete the shape '{name}'.") from e
        raise RuntimeError(f"The chart was pasted but could not be finished: {e}") from e
    warnings = []
    if lost:
        warnings.append(
            "The chart went through the clipboard, and what was on it before "
            f"was put back without these formats: {', '.join(lost)}.")
    if style and style_now != style:
        warnings.append(f"The chart has style {style_now}, not {style} as AddChart2 gives it.")
    return shape, type_now, warnings


def _add_chart_impl(slide_index, chart_type, left, top, width, height):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)

    type_int = _resolve_chart_type(chart_type)
    type_name = CHART_TYPE_NAMES.get(type_int, str(type_int))
    if type_int in chart_defaults.CHART_TYPES:
        chart_shape, type_now, warnings = _add_chart_by_paste(
            app, slide, type_int, left, top, width, height)
        result = {
            "success": True,
            "shape_name": chart_shape.Name,
            "shape_index": chart_shape.ZOrderPosition,
            "chart_type": CHART_TYPE_NAMES.get(type_now, str(type_now)),
            "chart_type_int": type_now,
        }
        if warnings:
            result["warning"] = " ".join(warnings)
        return result

    # The types gvml cannot write still go through AddChart2, which
    # starts an Excel; not while the last chart's is ending (#299).
    held.wait_until_ended()
    chart_shape = slide.Shapes.AddChart2(-1, type_int, left, top, width, height)
    chart = chart_shape.Chart

    # CRITICAL: AddChart2 auto-opens an Excel workbook. MUST close it to
    # prevent leaked Excel processes, and a chart left in edit mode stops the
    # next one being added. Its Excel is let go of once it has ended.
    refs = _ExcelRefs()
    try:
        closed = _close_chart_workbook(chart, keep=refs)
    finally:
        refs.hand_over()

    result = {
        "success": True,
        "shape_name": chart_shape.Name,
        "shape_index": chart_shape.ZOrderPosition,
        "chart_type": type_name,
        "chart_type_int": type_int,
    }
    if not closed:
        result["warning"] = WORKBOOK_LEFT_OPEN
    return result


class _NotCarried(Exception):
    """The new chart could not be made to stand in for the old one."""


def _excel_needed(slide, shape) -> Optional[str]:
    """Why a chart's data has to go through Excel, or None when it need not.

    Its data is otherwise swapped by replacing the chart with a copy, and a
    copy is a new shape: what points at the old one by its id, or what a
    copy does not take along, would be lost. A check that cannot be made
    counts as a reason, since then nothing says the copy would be safe.
    """
    if vba.active():
        # PowerPoint for Mac, through its VBA (backend.vba): Excel writes into
        # the chart itself, where a copy would come through the Windows
        # clipboard code and be a new shape.
        return "PowerPoint for Mac writes chart data through Excel"
    try:
        return _why_excel(slide, shape)
    except Exception as e:  # noqa: BLE001 - Excel then, as before
        return f"it could not be checked whether it can be replaced ({e})"


def _connectors(shapes):
    """Every connector among ``shapes``, those inside groups too."""
    for i in range(1, shapes.Count + 1):
        shape = shapes(i)
        if shape.Type == msoGroup:
            yield from _connectors(shape.GroupItems)
        elif shape.Connector == -1:
            yield shape


def _attached(connector, shape_id) -> bool:
    """Whether either end of ``connector`` is glued to the shape ``shape_id``."""
    ends = connector.ConnectorFormat
    return ((ends.BeginConnected == -1 and ends.BeginConnectedShape.Id == shape_id)
            or (ends.EndConnected == -1 and ends.EndConnectedShape.Id == shape_id))


def _why_excel(slide, shape) -> Optional[str]:
    chart = shape.Chart
    if chart.ChartType not in chart_defaults.CHART_TYPES:
        return f"its type ({CHART_TYPE_NAMES.get(chart.ChartType, chart.ChartType)}) is not swapped without Excel"
    try:
        if chart.ChartData.IsLinked:
            return "its data is linked to a workbook"
    except Exception:  # noqa: BLE001 - not linked as far as can be told
        pass
    if shape.Type == msoPlaceholder:
        return "it fills a placeholder"
    if shape.Child == -1:  # msoTrue
        return "it is in a group"
    # Its animations, the animations it starts and the connectors on it are
    # handed to the copy (animation_carry, #343).
    if slide.Comments.Count:
        # Which shape a comment is anchored to cannot be read through COM.
        return "the slide has comments, which may be anchored to it"
    return None


# What a chart's replacement takes over from it, read and set by name.
_CARRIED = ("Left", "Top", "Width", "Height", "Rotation", "Visible",
            "LockAspectRatio", "AlternativeText", "Title")


def _shape_state(shape) -> dict:
    """What ``_carry_over`` gives the replacement, read off ``shape``."""
    state = {name: getattr(shape, name) for name in _CARRIED}
    state["Name"] = shape.Name
    state["Tags"] = [(shape.Tags.Name(i), shape.Tags.Value(i))
                     for i in range(1, shape.Tags.Count + 1)]
    actions = []
    for when in (1, 2):  # ppMouseClick, ppMouseOver
        action = shape.ActionSettings(when)
        link = action.Hyperlink
        actions.append((action.Action, link.Address, link.SubAddress))
    state["Actions"] = actions
    return state


def _carry_over(state, shape) -> list:
    """Give ``shape`` what ``state`` holds, and answer what did not take."""
    failed = []
    for name in _CARRIED:
        try:
            if getattr(shape, name) != state[name]:
                setattr(shape, name, state[name])
            if getattr(shape, name) != state[name]:
                failed.append(name)
        except Exception:  # noqa: BLE001 - reported
            failed.append(name)
    try:
        shape.Name = state["Name"]
        if shape.Name != state["Name"]:
            failed.append("Name")
    except Exception:  # noqa: BLE001 - reported
        failed.append("Name")
    try:
        for name, value in state["Tags"]:
            shape.Tags.Add(name, value)
        if sorted(_shape_state_tags(shape)) != sorted(state["Tags"]):
            failed.append("Tags")
    except Exception:  # noqa: BLE001 - reported
        failed.append("Tags")
    try:
        for when, (action, address, sub_address) in zip((1, 2), state["Actions"]):
            setting = shape.ActionSettings(when)
            if (setting.Hyperlink.Address, setting.Hyperlink.SubAddress) != (address, sub_address):
                setting.Hyperlink.Address = address
                setting.Hyperlink.SubAddress = sub_address
            if setting.Action != action:
                setting.Action = action
        if _shape_state(shape)["Actions"] != state["Actions"]:
            failed.append("ActionSettings")
    except Exception:  # noqa: BLE001 - reported
        failed.append("ActionSettings")
    return failed


def _series_colours(chart) -> list:
    """The fill and line colour of each series, as shown."""
    out = []
    for i in range(1, chart.SeriesCollection().Count + 1):
        series = chart.SeriesCollection(i).Format
        out.append((series.Fill.ForeColor.RGB, series.Line.ForeColor.RGB))
    return out


def _shape_state_tags(shape) -> list:
    return [(shape.Tags.Name(i), shape.Tags.Value(i)) for i in range(1, shape.Tags.Count + 1)]


def _number(label, what: str) -> float:
    try:
        return float(label)
    except (TypeError, ValueError):
        raise ValueError(f"{what} {label!r} is not a number") from None


def _new_chart_bytes(chart_bytes, chart_type, categories, series) -> bytes:
    """``chart1.xml`` with the new data, the chart's formatting kept.

    Raises ``_NotCarried`` for a chart with more than one plot, whose other
    plots would keep their old data, and ValueError for data a scatter or
    bubble chart cannot take.
    """
    groups = gvml_charts.plot_count(chart_bytes)
    if groups != 1:
        raise _NotCarried(f"it has {groups} plot groups, and only the first would get the data")
    if chart_type in chart_defaults.XY_TYPES or chart_type in chart_defaults.BUBBLE_TYPES:
        categories = [_number(label, "the x value") for label in categories]
    if chart_type in chart_defaults.BUBBLE_TYPES:
        for entry in series:
            if "sizes" not in entry:
                raise ValueError(f"series {entry['name']!r} has no sizes for its bubbles")
            entry["sizes"] = [_number(size, "the bubble size") for size in entry["sizes"]]
    return gvml_charts.set_data(chart_bytes, categories, series)


def _replace_chart_data(app, pres, slide, shape, categories, series):
    """Swap a chart's data without Excel, by replacing the chart (#299).

    The chart is copied, its ``chart1.xml`` gets the new data with
    ``gvml.charts.set_data``, which keeps its formatting, and it is pasted
    back as ``Art::GVML ClipFormat``. Its embedded workbook is left out,
    since it would no longer match; PowerPoint makes a new one from the
    chart the first time its data is opened. The copy takes the old chart's
    name, place, size, stacking order, alternative text, tags and click
    actions before the old chart is deleted, and its animations, the
    animations it starts and the connectors on it (``animation_carry``,
    #343); and it has to show the new data
    in the same series colours. Anything short of that removes the copy,
    leaves the old chart as it was, and raises ``_NotCarried``. Answers the
    new shape.
    """
    try:
        state = _shape_state(shape)
        old_z = shape.ZOrderPosition
        colours = _series_colours(shape.Chart)
        chart_type = shape.Chart.ChartType
        old_id = shape.Id
        animated = animation_carry.animated(slide, old_id)
        connectors = [c for c in _connectors(slide.Shapes) if _attached(c, old_id)]
        timeline_before = animation_carry.timeline(slide, old_id) if animated else None
    except Exception as e:  # noqa: BLE001 - then it is not carried over
        raise _NotCarried(f"could not read what it has to take over ({e})") from e
    series = [dict(entry) for entry in series]
    new = None
    links = []
    try:
        with kept_clipboard() as lost:
            package = _copied_package(shape)
            chart_part = package.chart_part()
            chart_bytes = package.chart()
            if chart_bytes is None:
                raise _NotCarried("its copy holds no chart part")
            new_bytes = _new_chart_bytes(chart_bytes, chart_type, categories, series)
            package.parts[chart_part] = new_bytes
            workbook = gvml_charts.external_data_id(chart_bytes)
            if workbook is not None:
                package.remove_relationship(chart_part, workbook)
            strip_package_ids(package)
            raw = package.to_bytes()
            new = _paste_from(lambda: clipboard.put(GVML_FORMAT, raw), slide)
        if not new.HasChart:
            raise _NotCarried("what was pasted back is not a chart")
        problem, renamed = _check_chart_data(new.Chart, categories, series)
        if problem:
            raise _NotCarried(f"the copy does not show the data ({problem})")
        now = _series_colours(new.Chart)
        if now[:len(colours)] != colours[:len(now)]:
            raise _NotCarried("its series came back in other colours")
        failed = _carry_over(state, new)
        if failed:
            raise _NotCarried(f"could not carry over {', '.join(failed)}")
        # Just above the old chart, which then leaves its place to it. A step
        # that does not move it ends the loop, rather than looping for good.
        while new.ZOrderPosition > old_z + 1:
            before = new.ZOrderPosition
            new.ZOrder(3)  # msoSendBackward
            if new.ZOrderPosition >= before:
                break
        if new.ZOrderPosition != old_z + 1:
            raise _NotCarried("could not put it back in its stacking order")
        if animated:
            animation_carry.copy_animations(slide, shape, new)
        animation_carry.move_links(slide, shape, new, connectors, links)
        if animated and animation_carry.timeline(slide, new.Id, leave_out=old_id) != timeline_before:
            raise _NotCarried("its animations did not come over as they were")
    except Exception as e:
        unmoved = animation_carry.put_back(links)
        if unmoved:
            raise RuntimeError(
                f"The chart's copy could not be finished ({e}), and what was pointed at "
                f"the copy could not all be pointed back ({'; '.join(unmoved)}). Both "
                f"charts are on the slide: '{shape.Name}' (id {old_id}) is the old one, "
                f"id {new.Id} the copy.") from e
        if isinstance(e, animation_carry.NotCarried):
            e = _NotCarried(str(e))
        if new is not None:
            try:
                new.Delete()
            except Exception:  # noqa: BLE001 - named in the error instead
                raise RuntimeError(
                    f"The chart's copy could not be finished ({e}) and could not be "
                    f"removed either: delete the shape '{new.Name}' (id {new.Id}).") from e
        if isinstance(e, _NotCarried):
            raise e
        raise _NotCarried(str(e)) from e
    try:
        shape.Delete()
    except Exception as e:
        unmoved = animation_carry.put_back(links)
        if unmoved:
            raise RuntimeError(
                f"The old chart could not be removed ({e}), and what was pointed at its "
                f"copy could not all be pointed back ({'; '.join(unmoved)}). Both charts "
                f"are on the slide: '{shape.Name}' (id {old_id}) is the old one, id "
                f"{new.Id} the copy.") from e
        try:
            new.Delete()
        except Exception:  # noqa: BLE001 - named in the error instead
            raise RuntimeError(
                f"The old chart could not be removed ({e}), nor its copy: delete the "
                f"shape '{new.Name}' (id {new.Id}), which is the copy.") from e
        raise _NotCarried(f"the old chart could not be removed ({e})") from e
    warnings = [w for w in (renamed,) if w]
    if new.ZOrderPosition != old_z:
        warnings.append(f"The chart is at stacking position {new.ZOrderPosition}, not {old_z}.")
    if lost:
        warnings.append(
            "The chart went through the clipboard, and what was on it before "
            f"was put back without these formats: {', '.join(lost)}.")
    return new, warnings


def _set_chart_data_impl(slide_index, shape_name_or_index, categories, series):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_chart_shape(slide, shape_name_or_index)

    if shape.Chart.ChartType in chart_defaults.BUBBLE_TYPES:
        missing = [s.get("name") for s in series if "sizes" not in s]
        if missing:
            raise ValueError(
                f"A bubble chart's series need 'sizes', and {missing[0]!r} has none. "
                "Nothing was changed.")

    reason = _excel_needed(slide, shape)
    if reason is None:
        old_id = shape.Id
        try:
            new, warnings = _replace_chart_data(app, pres, slide, shape, categories, series)
        except _NotCarried as e:
            reason = f"swapping it without Excel failed: {e}"
            logger.warning("Setting chart data through Excel, since %s", reason)
        else:
            result = {
                "success": True,
                "shape_name": new.Name,
                "categories_count": len(categories),
                "series_count": len(series),
                "method": "replaced without Excel",
                "replaced": True,
                "shape_id": new.Id,
                "previous_shape_id": old_id,
            }
            if warnings:
                result["warning"] = " ".join(warnings)
            return result

    if shape.Chart.ChartType in chart_defaults.BUBBLE_TYPES:
        # The Excel way writes names and values only, so the sizes would be
        # dropped without a word.
        raise RuntimeError(
            f"Bubble sizes can only be set without Excel, and this chart has to go "
            f"through Excel, since {reason}. Nothing was changed.")
    try:
        result = _set_chart_data_through_excel(slide, shape, categories, series)
    except Exception as e:
        raise RuntimeError(f"{e} (Written through Excel, since {reason}.)") from e
    result["method"] = "Excel"
    result["reason"] = reason
    return result


XL_COLUMNS = 2  # XlRowCol.xlColumns


def _set_source_data(chart, address) -> None:
    """Point the chart at ``address``, a series per column, as written.

    Left to Excel, it takes rows when there are more series than categories.
    chart.SetSourceData() has a pywin32 VT_BYREF bug on the PlotBy parameter,
    so through pywin32 it is called with InvokeTypes and fixed type flags
    (12, 49) instead of (12, 17). PowerPoint for Mac's VBA (backend.vba) has
    no such bug and no ``_oleobj_``, and shows the new data only once the
    chart is refreshed.
    """
    oleobj = getattr(chart, "_oleobj_", None)
    if oleobj is not None:
        oleobj.InvokeTypes(
            1413, 0, 1, (24, 0),
            ((8, 1), (12, 49)),
            address, XL_COLUMNS,
        )
    else:
        chart.SetSourceData(address, XL_COLUMNS)
        chart.Refresh()


def _old_extent(chart):
    """The last row and column the chart's current data takes on its sheet."""
    try:
        categories, series = _read_chart_data(chart)
        return len(categories) + 1, len(series) + 1
    except Exception as e:
        logger.warning("Could not read the old chart data: %s", e)
        return 0, 0


def _clear_beyond(ws, old, last_row, last_col) -> None:
    """Clear what the old data left below and to the right of the new data.

    Done once the chart points at the new range, so it never shows them.
    """
    rows, cols = old
    try:
        if rows > last_row:
            ws.Range(ws.Cells(last_row + 1, 1), ws.Cells(rows, max(cols, last_col))).Clear()
        if cols > last_col:
            ws.Range(ws.Cells(1, last_col + 1), ws.Cells(last_row, cols)).Clear()
    except Exception as e:
        # The chart already shows the new data; leftovers sit outside its range.
        logger.warning("Could not clear the old chart data: %s", e)


def _write_chart_cells(ws, categories, series) -> None:
    """Write categories down column A and each series in a column after it.

    The cells are written over rather than cleared first, so the chart never
    points at an empty sheet while Excel is busy (#343, #323). A cell to be
    left empty is cleared, not given None: through PowerPoint for Mac's VBA a
    None reaches Excel as a value it refuses (error 1004, "Application-defined
    or object-defined error"), and every chart data write failed there. Clear
    is what the VBA bridge already writes out (as for _clear_beyond).
    """
    ws.Cells(1, 1).Clear()

    # Write categories in column A starting at row 2
    for i, cat in enumerate(categories):
        ws.Cells(i + 2, 1).Value = cat

    # Write series headers (row 1) and values (rows 2+)
    for s_idx, s in enumerate(series):
        col = s_idx + 2  # columns B, C, D, ...
        ws.Cells(1, col).Value = s["name"]
        # Every row is written, so a short series leaves no old value behind.
        values = list(s["values"]) + [None] * (len(categories) - len(s["values"]))
        for i, val in enumerate(values):
            if val is None:
                ws.Cells(i + 2, col).Clear()
            else:
                ws.Cells(i + 2, col).Value = val


def _set_chart_data_through_excel(slide, shape, categories, series):
    chart = shape.Chart
    old = _old_extent(chart)

    # Activate starts an Excel; not while the last chart's is ending (#299).
    held.wait_until_ended()
    # Every Excel object kept here is handed to ``held`` once the workbook is
    # closed, and let go of after its Excel has ended (#299).
    refs = _ExcelRefs()
    chart.ChartData.Activate()
    try:
        wb = refs.add(chart.ChartData.Workbook)
    except Exception:
        # Activate opened it; do not leave it open on the way out.
        try:
            _close_chart_workbook(chart, keep=refs)
        finally:
            refs.hand_over()
        raise
    try:
        ws = refs.add(wb.Worksheets(1))
        _write_chart_cells(ws, categories, series)

        # Set the chart's source data range.
        last_row = len(categories) + 1
        last_col = len(series) + 1
        data_range = refs.add(ws.Range(ws.Cells(1, 1), ws.Cells(last_row, last_col)))
        _set_source_data(chart, "Sheet1!" + data_range.Address)
        _clear_beyond(ws, old, last_row, last_col)
    finally:
        try:
            closed = _close_chart_workbook(chart, wb, keep=refs)
        finally:
            refs.hand_over()

    # Excel can stop while the data is being written and every call still
    # succeed, leaving the chart empty (#315). So the chart is read back, and
    # what it shows has to be what was written.
    problem, renamed = _check_chart_data(chart, categories, series)
    if problem:
        left_open = f" {WORKBOOK_LEFT_OPEN}" if not closed else ""
        raise RuntimeError(
            f"The chart did not take the data: {problem}. Excel may have stopped "
            f"while it was written; call ppt_set_chart_data again.{left_open}")

    result = {
        "success": True,
        "shape_name": shape.Name,
        "categories_count": len(categories),
        "series_count": len(series),
    }
    warnings = [w for w in (renamed, None if closed else WORKBOOK_LEFT_OPEN) if w]
    if warnings:
        result["warning"] = " ".join(warnings)
    return result


def _read_chart_data(chart):
    """The categories and series a chart shows, read from the chart itself.

    Read from SeriesCollection rather than the workbook, so Excel is not opened.
    """
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
    return categories, series_list


def _same_number(written, read) -> bool:
    import math

    try:
        # Relative for ordinary values, and a floor far below any value a
        # chart shows, so that 1e-10 written and 0.0 read is still a difference.
        return math.isclose(float(written), float(read), rel_tol=1e-9, abs_tol=1e-15)
    except (TypeError, ValueError):
        return False


def _same_label(written, read) -> bool:
    """Whether a label reads back as written, a number in either spelling included."""
    return str(written) == str(read) or _same_number(written, read)


def _compare_chart_data(categories, series, read_categories, read_series):
    """Compare what was written with what the chart shows.

    Answers ``(problem, renamed)``. ``problem`` says how the data did not land:
    a different number of categories, series or values, or a value that differs.
    ``renamed`` says which labels Excel read differently, since it turns some
    text into numbers or dates; that is worth a warning, not a failure.
    """
    if len(read_series) != len(series):
        return f"it has {len(read_series)} series, not {len(series)}", None
    if categories and len(read_categories) != len(categories):
        return f"it has {len(read_categories)} categories, not {len(categories)}", None
    for number, (wrote, read) in enumerate(zip(series, read_series), 1):
        # The data range has a row per category, so each series shows that
        # many points; values past them are not on the chart.
        values, shown = list(wrote.get("values", [])), list(read["values"])
        if categories and len(shown) != len(categories):
            return (f"series {number} has {len(shown)} values, not {len(categories)}"), None
        for place, (value, seen) in enumerate(zip(values, shown), 1):
            if (isinstance(value, (int, float)) and not isinstance(value, bool)
                    and not _same_number(value, seen)):
                return f"series {number} value {place} reads {seen!r}, not {value!r}", None
    renamed = [f"category {place} reads {seen!r}"
               for place, (label, seen) in enumerate(zip(categories, read_categories), 1)
               if not _same_label(label, seen)]
    renamed += [f"series {number} is named {read['name']!r}"
                for number, (wrote, read) in enumerate(zip(series, read_series), 1)
                if wrote.get("name") is not None and not _same_label(wrote["name"], read["name"])]
    if renamed:
        return None, ("Excel read some labels differently: " + "; ".join(renamed[:5]) + ".")
    return None, None


def _check_chart_data(chart, categories, series, attempts=3, wait=0.5):
    """Read the chart back until it shows what was written, a few times at most."""
    for attempt in range(attempts):
        if attempt:
            _wait_pumping(wait)
        try:
            read_categories, read_series = _read_chart_data(chart)
        except Exception as e:  # noqa: BLE001 - reported as the problem
            problem, renamed = f"reading it back failed ({e})", None
            continue
        problem, renamed = _compare_chart_data(categories, series, read_categories, read_series)
        if not problem:
            break
    return problem, renamed


def _get_chart_data_impl(slide_index, shape_name_or_index):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_chart_shape(slide, shape_name_or_index)
    chart = shape.Chart

    categories, series_list = _read_chart_data(chart)
    return {
        "success": True,
        "shape_name": shape.Name,
        "categories": categories,
        "series": series_list,
    }


def _format_chart_impl(
    slide_index, shape_name_or_index,
    title, has_legend, legend_position, chart_style, legend_font_size,
    legend_top, legend_left,
    title_position, title_top, title_left,
):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
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
        if key in LEGEND_POSITION_MAP:
            chart.Legend.Position = LEGEND_POSITION_MAP[key]
        else:
            # 8-direction coordinate-based presets
            parts = key.split("-")
            if len(parts) != 2 or parts[0] not in ("top", "middle", "bottom") or parts[1] not in ("left", "center", "right"):
                raise ValueError(
                    f"Unknown legend position '{legend_position}'. "
                    f"PowerPoint presets: {', '.join(LEGEND_POSITION_MAP.keys())}. "
                    "8-direction presets: top-left, top-center, top-right, "
                    "middle-left, middle-right, bottom-left, bottom-center, bottom-right."
                )
            row, col = parts
            leg = chart.Legend
            ca_w = chart.ChartArea.Width
            ca_h = chart.ChartArea.Height
            # Note: leg.Width/leg.Height may be 0 if the chart hasn't re-rendered.
            # If placement is off, ensure the chart is fully rendered before this call.
            leg_w = leg.Width
            leg_h = leg.Height
            gap = 5  # pt — small gap from chart edge
            if row == "top":
                leg_top = gap
            elif row == "middle":
                leg_top = (ca_h - leg_h) / 2
            else:  # bottom
                leg_top = ca_h - leg_h - gap
            if col == "left":
                leg_left = gap
            elif col == "center":
                leg_left = (ca_w - leg_w) / 2
            else:  # right
                leg_left = ca_w - leg_w - gap
            leg.Top = leg_top
            leg.Left = leg_left

    if legend_top is not None or legend_left is not None:
        if not chart.HasLegend:
            raise ValueError(
                "Cannot set legend coordinates when chart has no legend. "
                "Set has_legend=true first."
            )
        if legend_top is not None:
            chart.Legend.Top = legend_top
        if legend_left is not None:
            chart.Legend.Left = legend_left

    if chart_style is not None:
        chart.ChartStyle = chart_style

    if legend_font_size is not None:
        if not chart.HasLegend:
            raise ValueError(
                "Cannot set legend font size when chart has no legend. "
                "Set has_legend=true first."
            )
        chart.Legend.Font.Size = legend_font_size

    if title_position is not None or title_top is not None or title_left is not None:
        if not chart.HasTitle:
            raise ValueError(
                "Cannot set title position when chart has no title. "
                "Set title first."
            )
        ct = chart.ChartTitle
        gap = 5  # pt — small gap from chart edge / between elements
        # Note: ct.Width and ct.Height may be 0 if the chart has not yet
        # re-rendered after a recent title or data change. If positioning
        # appears off, ensure the chart is fully rendered before this call.
        if title_position is not None:
            if title_position == "top":
                # Restore PowerPoint automatic title placement by toggling HasTitle.
                # This clears any manual Top/Left and lets PowerPoint re-layout
                # the title at the top and restore the PlotArea accordingly.
                # NOTE: toggling HasTitle resets title formatting (font, color, size)
                # to defaults, so we snapshot and restore basic font properties.
                title_text = ct.Text
                try:
                    saved_name = ct.Font.Name
                    saved_size = ct.Font.Size
                    saved_bold = ct.Font.Bold
                    try:
                        saved_color = ct.Font.Color.RGB
                    except Exception:
                        saved_color = None
                except Exception:
                    saved_name = saved_size = saved_bold = saved_color = None
                chart.HasTitle = False
                chart.HasTitle = True
                chart.ChartTitle.Text = title_text
                ct = chart.ChartTitle  # re-bind after toggle
                if saved_name is not None:
                    ct.Font.Name = saved_name
                if saved_size is not None:
                    ct.Font.Size = saved_size
                if saved_bold is not None:
                    ct.Font.Bold = saved_bold
                if saved_color is not None:
                    ct.Font.Color.RGB = saved_color
            elif title_position == "bottom":
                # Shrink PlotArea from the bottom to make room for the title,
                # then position the title in the freed space.
                pa = chart.PlotArea
                # Determine target top for the title.
                # TODO: also detect legends placed at the bottom via 8-direction
                # coordinate presets (bottom-left/center/right), which set Legend.Top/Left
                # directly and don't change Legend.Position.
                if chart.HasLegend and chart.Legend.Position == LEGEND_POSITION_MAP["bottom"]:
                    title_top_target = chart.Legend.Top - ct.Height - gap
                else:
                    title_top_target = chart.ChartArea.Height - ct.Height - gap
                # Shrink PlotArea so its bottom does not overlap the title
                available_bottom = title_top_target - gap
                if pa.Top + pa.Height > available_bottom:
                    pa.Height = available_bottom - pa.Top
                ct.Top = title_top_target
                ct.Left = (chart.ChartArea.Width - ct.Width) / 2
            elif title_position == "center":
                ct.Top = (chart.ChartArea.Height - ct.Height) / 2
                ct.Left = (chart.ChartArea.Width - ct.Width) / 2
        if title_top is not None:
            ct.Top = title_top
        if title_left is not None:
            ct.Left = title_left

    return {
        "success": True,
        "shape_name": shape.Name,
        "has_title": bool(chart.HasTitle),
        "has_legend": bool(chart.HasLegend),
    }


def _format_chart_axis_impl(
    slide_index, shape_name_or_index, axis,
    title,
    min_scale, max_scale, major_unit, minor_unit,
    tick_label_spacing, tick_mark_spacing,
    major_tick_mark, minor_tick_mark,
    reverse_order, log_scale, log_base,
    tick_label_font_size, number_format,
):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_chart_shape(slide, shape_name_or_index)
    chart = shape.Chart

    axis_key = axis.strip().lower()
    if axis_key not in AXIS_TYPE_MAP:
        raise ValueError(
            f"Unknown axis '{axis}'. Valid: {', '.join(AXIS_TYPE_MAP.keys())}"
        )
    is_category = axis_key == "category"
    is_secondary = axis_key == "secondary_value"
    is_value_axis = axis_key in {"value", "secondary_value"}
    axis_type = AXIS_TYPE_MAP[axis_key]
    axis_group = AXIS_GROUP_SECONDARY if is_secondary else AXIS_GROUP_PRIMARY

    try:
        ax = chart.Axes(axis_type, axis_group)
    except Exception as e:
        raise ValueError(
            f"Axis '{axis}' is not available on this chart type "
            f"(e.g. pie/doughnut charts have no axes). Underlying error: {e}"
        )

    applied: list[str] = []

    # Category-axis-only guards
    if not is_category:
        if tick_label_spacing is not None:
            raise ValueError("tick_label_spacing is only valid for axis='category'.")
        if tick_mark_spacing is not None:
            raise ValueError("tick_mark_spacing is only valid for axis='category'.")
    # Value-axis-only guards (category and 3D series axes don't support these)
    if not is_value_axis:
        if min_scale is not None or max_scale is not None:
            raise ValueError("min_scale/max_scale are only valid for value axes.")
        if major_unit is not None or minor_unit is not None:
            raise ValueError("major_unit/minor_unit are only valid for value axes.")
        if log_scale is not None:
            raise ValueError("log_scale is only valid for value axes.")
    if log_base is not None and log_scale is not True:
        raise ValueError("log_base requires log_scale=true.")

    if title is not None:
        ax.HasTitle = True
        ax.AxisTitle.Text = title
        applied.append("title")

    if min_scale is not None:
        ax.MinimumScale = min_scale
        applied.append("min_scale")
    if max_scale is not None:
        ax.MaximumScale = max_scale
        applied.append("max_scale")
    if major_unit is not None:
        ax.MajorUnit = major_unit
        applied.append("major_unit")
    if minor_unit is not None:
        ax.MinorUnit = minor_unit
        applied.append("minor_unit")

    if tick_label_spacing is not None:
        ax.TickLabelSpacing = tick_label_spacing
        applied.append("tick_label_spacing")
    if tick_mark_spacing is not None:
        ax.TickMarkSpacing = tick_mark_spacing
        applied.append("tick_mark_spacing")

    if major_tick_mark is not None:
        ax.MajorTickMark = TICK_MARK_MAP[major_tick_mark]
        applied.append("major_tick_mark")
    if minor_tick_mark is not None:
        ax.MinorTickMark = TICK_MARK_MAP[minor_tick_mark]
        applied.append("minor_tick_mark")

    if reverse_order is not None:
        ax.ReversePlotOrder = reverse_order
        applied.append("reverse_order")

    if log_scale is not None:
        ax.ScaleType = SCALE_LOGARITHMIC if log_scale else SCALE_LINEAR
        applied.append("log_scale")
    if log_base is not None:
        ax.LogBase = log_base
        applied.append("log_base")

    if tick_label_font_size is not None:
        ax.TickLabels.Font.Size = tick_label_font_size
        applied.append("tick_label_font_size")
    if number_format is not None:
        ax.TickLabels.NumberFormat = number_format
        applied.append("number_format")

    return {
        "success": True,
        "shape_name": shape.Name,
        "axis": axis_key,
        "applied": applied,
    }


def _set_chart_series_impl(
    slide_index, shape_name_or_index,
    series_index, color, show_data_labels, line_weight,
):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
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
    pres = ppt._get_pres_impl()
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
        report_layout(result, params.slide_index, touched_names(result))
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
        report_layout(result, params.slide_index,
                      touched_names(result) or [params.shape_name_or_index])
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to set chart data: {str(e)}"})


def get_chart_data(params: GetChartDataInput) -> str:
    """Read the data a chart shows, from the chart itself, without Excel.

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
            params.legend_font_size,
            params.legend_top, params.legend_left,
            params.title_position, params.title_top, params.title_left,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to format chart: {str(e)}"})


def format_chart_axis(params: FormatChartAxisInput) -> str:
    """Format axis properties: scale, ticks, labels, axis title.

    Args:
        params: Axis identifier and optional formatting properties.

    Returns:
        JSON listing which properties were applied.
    """
    try:
        result = ppt.execute(
            _format_chart_axis_impl,
            params.slide_index, params.shape_name_or_index, params.axis,
            params.title,
            params.min_scale, params.max_scale,
            params.major_unit, params.minor_unit,
            params.tick_label_spacing, params.tick_mark_spacing,
            params.major_tick_mark, params.minor_tick_mark,
            params.reverse_order, params.log_scale, params.log_base,
            params.tick_label_font_size, params.number_format,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to format chart axis: {str(e)}"})


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
        On Windows the chart is pasted without starting Excel. The types that
        still open Excel (3-D area and bar, cone, cylinder, pyramid, stock,
        surface and the like) can first wait up to 15 s for the previous
        chart's Excel to close.
        """
        return await run_offloaded(add_chart, params)

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
        """Set a chart's categories and series.

        Provide category labels and one or more series, each with a name and
        numeric values. The number of values in each series should match the
        number of categories.
        Example series: [{"name": "Revenue", "values": [100, 200, 150, 300]}]
        On Windows the chart is read back afterwards, and the call fails when
        it does not show the data written. Labels Excel reads differently, such
        as text it takes for a date, come back in `warning`.
        On Windows the data is swapped without starting Excel where it can be:
        the chart is replaced by a copy holding the new data, with the old
        one's formatting, name, place, size and stacking order, its animations
        in their places, the animations clicking it starts, and the connectors
        on it. The copy has a new shape id, given as `shape_id` (the old one as
        `previous_shape_id`). A chart that is grouped, in a placeholder, sits
        on a slide with comments, has animations that cannot be copied as
        they were, combines more than one chart type, or is of a type that still opens
        Excel (3-D area and bar, cone, stock, surface and the like), is written
        through Excel as before, and so is a scatter chart whose categories
        are not numbers. A bubble chart cannot go through Excel, which would
        drop its sizes, so such a call, or one with a series that has no
        'sizes', is refused. `method` says which was used, and `reason` why
        Excel was.
        Either way the values' number format goes back to General.
        """
        return await run_offloaded(set_chart_data, params)

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
        """Read the data a chart shows, from the chart itself, without Excel.

        Returns the category labels and all series (name + values).
        Identify the chart by shape name or 1-based shape index.
        """
        return await run_offloaded(get_chart_data, params)

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
        set legend position using PowerPoint presets ('bottom', 'left', 'right',
        'top', 'corner') or 8-direction coordinate presets ('top-left',
        'top-center', 'top-right', 'middle-left', 'middle-right', 'bottom-left',
        'bottom-center', 'bottom-right'), or use legend_top/legend_left for
        explicit coordinate positioning.
        Apply a built-in chart style by index number, and set the legend font
        size in points via legend_font_size.
        Control chart title position via title_position preset ('top', 'bottom',
        'center') or explicit title_top/title_left coordinates in points.
        """
        return await run_offloaded(format_chart, params)

    @mcp.tool(
        name="ppt_format_chart_axis",
        annotations={
            "title": "Format Chart Axis",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_format_chart_axis(params: FormatChartAxisInput) -> str:
        """Format an axis: scale (min/max/major_unit/minor_unit), tick spacing,
        tick marks, log scale, reverse order, tick-label font size, number
        format, and axis title.

        Select the axis via the `axis` parameter:
        - 'category' — horizontal x-axis on column/line/bar charts (default)
        - 'value' — primary y-axis
        - 'secondary_value' — secondary y-axis
        - 'series' — 3D charts only

        `tick_label_spacing` / `tick_mark_spacing` are category-axis only:
        for example, on a 0–60 day axis, `tick_label_spacing=3` renders
        labels at 0, 3, 6, …, 60.

        `min_scale` / `max_scale` / `major_unit` / `minor_unit` / `log_scale`
        are value-axis only. `log_base` requires `log_scale=true`.
        """
        return await run_offloaded(format_chart_axis, params)

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
        return await run_offloaded(set_chart_series, params)

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
        return await run_offloaded(change_chart_type, params)


# ---------------------------------------------------------------------------
# macOS
# ---------------------------------------------------------------------------
# The implementations above walk COM. Their Apple Event counterparts have the
# same names and signatures, so on macOS they simply take their place; nothing
# else in this module changes.
from backend import IS_MACOS, use_mac_impls  # noqa: E402

if IS_MACOS:  # pragma: no cover - platform specific
    from ppt_mac import charts as _mac_charts

    use_mac_impls(globals(), _mac_charts)
