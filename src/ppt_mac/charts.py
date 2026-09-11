"""Chart tools, on Apple Events and the clipboard.

Mirrors ``ppt_com/charts.py``. ``ppt_add_chart`` and ``ppt_get_chart_data``
work; the five that edit an existing chart still refuse, and the reason is
stated once here.

**There is no chart in this dictionary.** Not a thin one, not a read only one,
none. The only word in it containing "chart" is ``chart unit effect``, an
animation setting. There is no ``chart`` class, no ``has chart`` beside ``has
table`` on ``shape``, no ``series``, no ``axis``, no ``chart data``. Windows
reaches all seven tools through ``Shape.Chart``, and that step has no
counterpart.

**The clipboard has one.** A chart copied with ``copy shape`` lands on the
pasteboard as a DrawingML package holding ``chart1.xml``, caches and all, and
``paste object`` takes such a package back and makes a real chart of it,
without a workbook. So a chart is added by writing ``chart1.xml`` from the
``XlChartType`` integer and pasting it, and its data is read by copying it
and reading the caches. Both were measured before they were written
(docs/gvml-design.md section 0). The XML writer and reader are in
``gvml/charts.py``, the paste procedure in ``ppt_mac/gvml_paste.py``.

**What is not yet written is the editing.** ``ppt_set_chart_data``,
``ppt_change_chart_type``, ``ppt_format_chart``, ``ppt_format_chart_axis``
and ``ppt_set_chart_series`` are the design's second tier: copy, rewrite the
XML, paste, delete the original, restore position and z order. They refuse
until then, naming the shape they were asked about first, so a caller who
picked the wrong shape hears that rather than a platform note.

Nothing that refuses moves the view, so a refused call leaves the user
looking where they were.
"""

import logging

from backend.mac_ae import ppt, slide_at as _slide
from backend.mac_enums import MsoShapeType
from backend.unsupported import refusal as _refusal
from gvml import PackageError
from gvml import canvas as _canvas
from gvml import charts as _gvml_charts
from ppt_com.constants import msoChart
from ppt_mac.gvml_paste import (
    Clipboard,
    Refused,
    copy_shape_package,
    paste_package,
    unused_name,
    with_warnings,
)
from ppt_mac.shapes import _WIN_SHAPE_TYPE, _get_shape, _shape_names, _win_constant

logger = logging.getLogger(__name__)


def _com_charts():
    """Reach the COM module's validation tables without importing them up here.

    ``ppt_com/charts.py`` ends by importing this module, so importing it back
    at load time is a cycle, and which of the two wins depends on which one is
    imported first. Whichever loses gets a half built module and the swap
    silently finds nothing to swap. The tables are only wanted inside a call,
    by which point both modules are finished, so the import waits until then.

    ``_resolve_chart_type``, ``CHART_TYPE_NAMES`` and ``AXIS_TYPE_MAP`` are
    taken from there rather than copied so a chart type added on the Windows
    side is still spelled the same way in a Mac caller's error.
    """
    from ppt_com import charts

    return charts


# The one finding every remaining refusal in this module rests on.
_NO_CHART_OBJECT = (
    "PowerPoint for Mac has no `chart` class in its Apple Event dictionary. "
    "The only property, element or command in it containing \"chart\" is "
    "`chart unit effect`, which is an animation setting, and `shape` carries "
    "`has table` with no `has chart` beside it. Windows reaches this through "
    "Shape.Chart, and that step does not exist here. The chart's XML can be "
    "reached through the clipboard, which is how ppt_add_chart and "
    "ppt_get_chart_data work, and rewriting it that way is not yet written."
)

# What survives, said the same way each time.
_SHAPE_TOOLS = [
    "ppt_get_chart_data, which reads the categories and series",
    "ppt_get_shape_info, which reports a chart shape's type, name and box",
    "ppt_update_shape, which moves and resizes it",
    "ppt_delete_shape",
]

_ADD_ALTERNATIVES = [
    "ppt_add_table, which carries the same numbers as a grid",
    "ppt_add_picture, for a chart rendered elsewhere",
]


def _chart_shape(slide, name_or_index):
    """Find a shape on a slide and verify it is a chart.

    The counterpart of ``_get_chart_shape`` on the COM side, which asks
    ``HasChart``. There is no such property here, so the shape type carries the
    answer instead, read through the generated table so the comparison is
    against the Windows constant a caller already knows. The message is the one
    Windows gives, word for word.
    """
    shape = _get_shape(slide, name_or_index)
    type_val = _win_constant(_WIN_SHAPE_TYPE, shape.shape_type())
    if type_val != msoChart:
        raise ValueError(f"Shape '{shape.name()}' is not a chart")
    return shape


def _refuse_for_shape(tool_name, slide_index, shape_name_or_index, detail, extra=None):
    """Find the chart, then refuse, naming it."""
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _chart_shape(slide, shape_name_or_index)

    return _refusal(
        tool_name,
        f"'{shape.name()}' on slide {slide_index} is a chart and {detail} "
        f"{_NO_CHART_OBJECT}",
        (extra or []) + _SHAPE_TOOLS,
    )


# ---------------------------------------------------------------------------
# Apple Event implementation functions
# ---------------------------------------------------------------------------
def _add_chart_impl(slide_index, chart_type, left, top, width, height):
    """Write ``chart1.xml`` for the type and paste it as a chart.

    The type is resolved through the Windows table first, so a caller who
    wrote 'colunm' hears about the spelling. A type the XML writer has no
    template for is refused by argument, not by tool.
    """
    com = _com_charts()
    type_int = com._resolve_chart_type(chart_type)
    type_name = com.CHART_TYPE_NAMES.get(type_int, str(type_int))
    try:
        chart_xml = _gvml_charts.chart_xml(type_int)
    except _gvml_charts.ChartTypeError as exc:
        supported = sorted(
            com.CHART_TYPE_NAMES[t] for t in _gvml_charts.supported_types()
            if t in com.CHART_TYPE_NAMES
        )
        return _refusal(
            "ppt_add_chart",
            f"{exc} On macOS a chart is made by writing its XML and pasting "
            "it, and no template exists yet for this type. Types that can be "
            f"drawn: {', '.join(supported)}.",
            _ADD_ALTERNATIVES,
            error=f"ppt_add_chart cannot draw chart_type {chart_type!r} on macOS",
        )

    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    name = unused_name(_shape_names(slide), "Chart")
    raw = _gvml_charts.chart_package(
        name, _canvas.emu(left), _canvas.emu(top), _canvas.emu(width), _canvas.emu(height), chart_xml,
    )

    clip = Clipboard.take()
    try:
        try:
            pasted = paste_package(
                pres, slide, slide_index, raw, clip, "ppt_add_chart",
                MsoShapeType[msoChart], left, top, _ADD_ALTERNATIVES,
            )
        except Refused as refused:
            return refused.payload
    finally:
        clip.restore()

    return with_warnings({
        "success": True,
        "shape_name": pasted.name,
        "shape_index": pasted.shape.z_order_position(),
        "chart_type": type_name,
        "chart_type_int": type_int,
    }, clip, pasted.warnings)


def _set_chart_data_impl(slide_index, shape_name_or_index, categories, series):
    """Refuse, because rewriting a chart's XML is not yet written here."""
    return _refuse_for_shape(
        "ppt_set_chart_data",
        slide_index,
        shape_name_or_index,
        "its data cannot be written yet. Windows writes the numbers into the "
        "chart's Excel workbook through Chart.ChartData; there is no `chart "
        "data` and no workbook here, and rewriting the caches in the chart's "
        "XML through the clipboard is the design's next step.",
        ["ppt_add_chart, then delete the old chart, until then"],
    )


def _get_chart_data_impl(slide_index, shape_name_or_index):
    """Copy the chart and read its categories and series out of ``chart1.xml``.

    The numbers come from the caches PowerPoint keeps in the XML, which are
    what the chart draws from; a chart without a workbook still has them.
    Nothing on the slide changes and the clipboard is put back afterwards.
    """
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _chart_shape(slide, shape_name_or_index)
    name = shape.name()

    clip = Clipboard.take()
    try:
        try:
            package = copy_shape_package(shape, clip, "ppt_get_chart_data")
        except Refused as refused:
            return refused.payload
        chart_bytes = package.chart()
        if chart_bytes is None:
            return _refusal(
                "ppt_get_chart_data",
                f"'{name}' reports `shape type chart`, but the package "
                "PowerPoint wrote for it holds no chart part, so there are no "
                "categories or series to read.",
                _SHAPE_TOOLS[1:],
            )
        try:
            data = _gvml_charts.read_chart(chart_bytes)
        except PackageError as exc:
            return _refusal(
                "ppt_get_chart_data",
                f"'{name}' is a chart, but its chart1.xml could not be read: {exc}",
                _SHAPE_TOOLS[1:],
            )
    finally:
        clip.restore()

    return with_warnings({
        "success": True,
        "shape_name": name,
        "categories": data["categories"],
        "series": data["series"],
    }, clip)


def _format_chart_impl(
    slide_index, shape_name_or_index,
    title, has_legend, legend_position, chart_style, legend_font_size,
    legend_top, legend_left,
    title_position, title_top, title_left,
):
    """Refuse, because a chart's title and legend are parts of the chart."""
    return _refuse_for_shape(
        "ppt_format_chart",
        slide_index,
        shape_name_or_index,
        "its title, legend and style are parts of the chart rather than of the "
        "shape. Every one of them hangs off Chart on the Windows side, so none "
        "of them survives the crossing.",
    )


def _format_chart_axis_impl(
    slide_index, shape_name_or_index, axis,
    title,
    min_scale, max_scale, major_unit, minor_unit,
    tick_label_spacing, tick_mark_spacing,
    major_tick_mark, minor_tick_mark,
    reverse_order, log_scale, log_base,
    tick_label_font_size, number_format,
):
    """Refuse, because there is no axis to format.

    The axis name is checked first for the same reason the chart type is in
    ``_add_chart_impl``. A caller who wrote 'catagory' should hear that.
    """
    axis_map = _com_charts().AXIS_TYPE_MAP
    axis_key = axis.strip().lower()
    if axis_key not in axis_map:
        raise ValueError(
            f"Unknown axis '{axis}'. Valid: {', '.join(axis_map.keys())}"
        )

    return _refuse_for_shape(
        "ppt_format_chart_axis",
        slide_index,
        shape_name_or_index,
        f"it has no addressable '{axis_key}' axis. There is no `axis` class in "
        "the dictionary and no `axes` element on anything, so scale, ticks, "
        "labels and axis titles are all out of reach.",
    )


def _set_chart_series_impl(
    slide_index, shape_name_or_index,
    series_index, color, show_data_labels, line_weight,
):
    """Refuse, because a series is not a shape.

    Tempting to reach for ``fill format`` on the chart shape, which does exist.
    It fills the chart's own background, not one series inside it, so honouring
    ``color`` that way would repaint the whole chart and report success.
    """
    return _refuse_for_shape(
        "ppt_set_chart_series",
        slide_index,
        shape_name_or_index,
        f"series {series_index} cannot be addressed. There is no `series` "
        "class in the dictionary. The chart shape does carry a `fill format`, "
        "but that paints the whole chart's background rather than one series, "
        "so it is deliberately left alone.",
    )


def _change_chart_type_impl(slide_index, shape_name_or_index, chart_type):
    """Refuse, because the type belongs to the chart and not to the shape.

    The requested type is resolved first, so a typo is still heard.
    """
    type_int = _com_charts()._resolve_chart_type(chart_type)
    logger.debug("chart type %r resolved to %d before refusing", chart_type, type_int)

    return _refuse_for_shape(
        "ppt_change_chart_type",
        slide_index,
        shape_name_or_index,
        "its chart type cannot be changed. `shape type chart` is read only and "
        "says only that the shape is a chart, not which kind. The Windows "
        "Chart.ChartType that carries the kind has no counterpart.",
        ["Change the chart type by hand in PowerPoint"],
    )
