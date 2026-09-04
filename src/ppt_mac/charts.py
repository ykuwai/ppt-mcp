"""Chart tools, on Apple Events.

Mirrors ``ppt_com/charts.py``. All seven tools refuse, and the reason is the
same one seven times, so it is stated once here and named once in each refusal.

**There is no chart in this dictionary.** Not a thin one, not a read only one,
none. Parsing ``PowerPoint.sdef`` into the tables appscript itself builds gives
a reference table, which holds every property, element and command PowerPoint
answers to, and the only word in it containing "chart" is ``chart unit effect``,
which sets whether an animation reveals a chart by series or by category. There
is no ``chart`` class, no ``has chart`` beside ``has table`` on ``shape``, no
``series``, no ``axis``, no ``chart data``. Windows reaches all seven of these
tools through ``Shape.Chart``, and that one step has no counterpart.

**What is not missing is the shape.** ``shape type chart`` is a real enumerator,
code ``0x008c0003``, which is the Windows ``msoChart`` constant 3 in the low
byte, so a chart already in the deck reports itself as a chart. That is enough
to find it, name it, move it, resize it, read its box and delete it through the
ordinary shape tools, and it is why every refusal here points at those rather
than saying charts are absent altogether. The chart is on the slide. Only its
contents are out of reach.

**A chart shape is still checked before it is refused.** A caller who names the
wrong shape hears that, with the same message Windows gives, because being told
a platform cannot do something is no use when the real mistake was the shape
name. Only once the shape is found and is a chart does the refusal follow.

MACOS_PORT section 6.1 leaves one question open, whether ``AddChart`` exists in
the Mac VBA type library, and the honest answer there is that it was not found
by a method whose negatives are unreliable. Nothing here contradicts that. The
Apple Event dictionary is settled and empty; VBA is a separate route that would
still need ``run VB macro`` to be proven to execute at all.

Nothing here edits a slide, so nothing calls ``goto_slide``, and a refused call
leaves the user's view where it was.
"""

import logging

from backend.mac_ae import ppt, slide_at as _slide
from backend.unsupported import refusal as _refusal
from ppt_com.constants import msoChart
from ppt_mac.shapes import _WIN_SHAPE_TYPE, _get_shape, _win_constant

logger = logging.getLogger(__name__)


def _com_charts():
    """Reach the COM module's validation tables without importing them up here.

    ``ppt_com/charts.py`` ends by importing this module, so importing it back
    at load time is a cycle, and which of the two wins depends on which one is
    imported first. Whichever loses gets a half built module and the swap
    silently finds nothing to swap. The tables are only wanted inside a call,
    by which point both modules are finished, so the import waits until then.

    ``_resolve_chart_type`` and ``AXIS_TYPE_MAP`` are taken from there rather
    than copied so a chart type added on the Windows side is still spelled the
    same way in the error a Mac caller reads.
    """
    from ppt_com import charts

    return charts


# The one finding every refusal in this module rests on, kept in one place so
# that a reader who meets it twice recognises it as the same finding.
_NO_CHART_OBJECT = (
    "PowerPoint for Mac has no `chart` class in its Apple Event dictionary. "
    "The only property, element or command in it containing \"chart\" is "
    "`chart unit effect`, which is an animation setting, and `shape` carries "
    "`has table` with no `has chart` "
    "beside it. Windows reaches this through Shape.Chart, and that step does "
    "not exist here."
)

# What survives, said the same way each time. An existing chart is a shape like
# any other, and these are the tools that treat it as one.
_SHAPE_TOOLS = [
    "ppt_get_shape_info, which reports a chart shape's type, name and box",
    "ppt_update_shape, which moves and resizes it",
    "ppt_list_shapes",
    "ppt_delete_shape",
]


def _chart_shape(slide, name_or_index):
    """Find a shape on a slide and verify it is a chart.

    The counterpart of ``_get_chart_shape`` on the COM side, which asks
    ``HasChart``. There is no such property here, so the shape type carries the
    answer instead, read through the generated table so the comparison is
    against the Windows constant a caller already knows. The message is the one
    Windows gives, word for word, because a caller who moves between the two
    should not have to learn it twice.
    """
    shape = _get_shape(slide, name_or_index)
    type_val = _win_constant(_WIN_SHAPE_TYPE, shape.shape_type())
    if type_val != msoChart:
        raise ValueError(f"Shape '{shape.name()}' is not a chart")
    return shape


def _refuse_for_shape(tool_name, slide_index, shape_name_or_index, detail, extra=None):
    """Find the chart, then refuse, naming it.

    Every tool but ``ppt_add_chart`` takes a shape identifier, and resolving it
    first is what separates a caller who picked the wrong shape from a caller
    who picked a platform that cannot help. The first hears about the shape.
    """
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
    """Refuse, because nothing in the dictionary makes a chart.

    The chart type is resolved first even though it is about to be thrown
    away, so a caller who wrote 'colunm' hears about the spelling rather than
    reading a platform note and going away to fix the wrong thing.
    """
    type_int = _com_charts()._resolve_chart_type(chart_type)
    logger.debug("chart type %r resolved to %d before refusing", chart_type, type_int)

    return _refusal(
        "ppt_add_chart",
        f"{_NO_CHART_OBJECT} There is no `make new chart`, and no other "
        "command takes a chart type, so a chart cannot be put on a slide from "
        "a script at all. A chart inserted by hand is a different matter. It "
        "reports `shape type chart` and behaves as an ordinary shape from "
        "then on.",
        [
            "Insert the chart once by hand in PowerPoint, then position it "
            "with ppt_update_shape",
            "ppt_add_table, which carries the same numbers as a grid",
            "ppt_add_picture, for a chart rendered elsewhere",
        ],
    )


def _set_chart_data_impl(slide_index, shape_name_or_index, categories, series):
    """Refuse, because the data sheet behind a chart is not addressable.

    Windows opens the chart's Excel workbook and writes cells. There is no
    ``chart data`` and no workbook here, and no route to Excel either, since
    the two applications are separate Apple Event targets with nothing linking
    a shape on a slide to a sheet in a book.
    """
    return _refuse_for_shape(
        "ppt_set_chart_data",
        slide_index,
        shape_name_or_index,
        "its data sheet cannot be reached. Windows writes the numbers into the "
        "chart's Excel workbook through Chart.ChartData, and there is no "
        "`chart data` and no workbook in this dictionary.",
    )


def _get_chart_data_impl(slide_index, shape_name_or_index):
    """Refuse, because the numbers cannot be read back either.

    Symmetrical with the write. Worth its own refusal rather than a shared one,
    because a caller reading a chart is usually trying to find out what is in
    the deck, and the useful answer names the tool that does report something.
    """
    return _refuse_for_shape(
        "ppt_get_chart_data",
        slide_index,
        shape_name_or_index,
        "its categories and series cannot be read. There is no `series`, no "
        "`axis` and no `chart data` anywhere in the dictionary, so the only "
        "thing PowerPoint will say about it is that it is a chart, how big it "
        "is and where it sits.",
    )


def _format_chart_impl(
    slide_index, shape_name_or_index,
    title, has_legend, legend_position, chart_style, legend_font_size,
    legend_top, legend_left,
    title_position, title_top, title_left,
):
    """Refuse, because a chart's title and legend are parts of the chart.

    Every argument this tool takes hangs off ``Chart``. None of them is a
    property of the shape, so there is nothing here to honour partially and no
    argument worth singling out with ``error=``.
    """
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
