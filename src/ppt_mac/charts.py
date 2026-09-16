"""Chart tools, on Apple Events and the clipboard.

Mirrors ``ppt_com/charts.py``. All seven tools work; the five that edit an
existing chart do it by rewriting the chart and pasting it back, and say so.

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

**Editing is copy, rewrite, paste back, delete the original.** The five
editing tools share ``_edit_chart``: the chart is copied, one thing in its
``chart1.xml`` is changed by ``gvml.charts``, the package is pasted beside the
original, read back to confirm the change is in it, and only then is the
original deleted and the new chart walked back to the original's z order.
The chart that results is a new object with the old name. The caller is told
so in ``warnings``, along with how many animation effects went with the
original, because the package does not carry them (design section 4). A call
that left every optional argument out is not worth that price, so it stops
after the copy and says in ``note`` that the chart still stands.

``ppt_format_chart`` and ``ppt_format_chart_axis`` take the arguments that
are one element or one attribute in the XML and refuse the rest by name,
before PowerPoint is touched, so dropping the argument and calling again
works. What is refused and why is in ``_FORMAT_CHART_UNMAPPED`` and
``_FORMAT_AXIS_UNMAPPED``.

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
    replace_shape,
    strip_ids,
    unused_name,
    with_warnings,
)
from ppt_mac.shapes import (
    _WIN_SHAPE_TYPE,
    _get_shape,
    _shape_index,
    _shape_names,
    _win_constant,
)

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


# What survives a refusal, said the same way each time.
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

_EDIT_ALTERNATIVES = [
    "ppt_add_chart, then ppt_delete_shape on the old chart",
    "Edit the chart by hand in PowerPoint",
]

# Arguments of ppt_format_chart with no single element in chart1.xml to
# write. Each is refused by name, before PowerPoint is touched.
_FORMAT_CHART_UNMAPPED = {
    "chart_style": (
        "a built-in style number is applied by PowerPoint when it builds the "
        "chart's colour and effect parts, and `c:style` in the XML does not "
        "restyle a chart that already has them"
    ),
    "legend_font_size": "the legend's text properties are a whole `c:txPr` block, not one attribute",
    "legend_top": "a legend's coordinates are a manual layout PowerPoint computes from the rendered size",
    "legend_left": "a legend's coordinates are a manual layout PowerPoint computes from the rendered size",
    "title_position": "the title's placement is computed from its rendered size, which the XML does not carry",
    "title_top": "the title's coordinates are a manual layout PowerPoint computes from the rendered size",
    "title_left": "the title's coordinates are a manual layout PowerPoint computes from the rendered size",
}

# The 8-direction legend presets Windows computes from the legend's rendered
# width and height, which the XML does not carry.
_LEGEND_DIRECTIONS = {
    "top-left", "top-center", "top-right", "middle-left", "middle-right",
    "bottom-left", "bottom-center", "bottom-right",
}

# What a call that left every optional argument out is told. It is a success,
# because it is one on Windows too, where no argument means no property set;
# the sentence is here so nobody reads the success as work done.
_NOTHING_ASKED = (
    "No optional argument was given, so nothing was changed. On macOS this "
    "tool rewrites a chart by replacing it, which would cost the chart its "
    "animations, and a call that asks for nothing is not worth that. The "
    "chart is the object it was."
)

_FORMAT_AXIS_UNMAPPED = {
    "tick_label_font_size": "the tick labels' text properties are a whole `c:txPr` block, not one attribute",
}


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


def _refuse_arguments(tool_name, given: dict, table: dict):
    """Refuse the arguments a tool cannot honour, naming each, or None."""
    names = [name for name in table if given.get(name) is not None]
    if not names:
        return None
    reasons = "; ".join(f"{name}: {table[name]}" for name in names)
    return _refusal(
        tool_name,
        f"On macOS this tool rewrites the chart's XML through the clipboard, "
        f"and {', '.join(names)} cannot be written that way ({reasons}). "
        "Nothing was changed; drop the argument and call again.",
        [f"{tool_name} without {', '.join(names)}"],
        error=f"{tool_name} cannot set {', '.join(names)} on macOS",
    )


def _edit_chart(tool_name, slide_index, shape_name_or_index, edit, verify, describe,
                requested=True):
    """Copy the chart, rewrite its XML with ``edit``, paste it back, check it.

    ``edit(chart_bytes) -> (chart_bytes, state)`` does the rewrite and may
    raise ``ValueError`` with the Windows wording for an argument the chart
    cannot take (a legend position with no legend). ``verify(readback_bytes,
    state) -> problem or None`` says whether the chart PowerPoint pasted
    carries the change. ``describe(readback_bytes, state) -> dict`` builds
    the success body, minus ``shape_name`` and ``warnings``.

    ``requested`` is False when every optional argument was left out. Such a
    call still copies the chart, because ``edit`` is what finds out whether
    the axis or the series it names is there at all, and Windows raises for
    one that is not. What it does not do is paste the chart back; replacing
    a chart costs it its animations and makes a new object of it, which is
    a high price for a call that asked for nothing.
    """
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _chart_shape(slide, shape_name_or_index)
    index = _shape_index(slide, shape_name_or_index)
    name = shape.name()
    left, top = shape.left_position(), shape.top()

    clip = Clipboard.take()
    kept = None
    try:
        try:
            package = copy_shape_package(shape, clip, tool_name)
        except Refused as refused:
            return refused.payload
        chart_bytes = package.chart()
        if chart_bytes is None:
            return _refusal(
                tool_name,
                f"'{name}' reports `shape type chart`, but the package "
                "PowerPoint wrote for it holds no chart part, so there is no "
                "chart1.xml to rewrite. Nothing was changed.",
                _SHAPE_TOOLS[1:],
            )
        try:
            new_bytes, state = edit(chart_bytes)
        except PackageError as exc:
            return _refusal(
                tool_name,
                f"'{name}' is a chart, but its chart1.xml could not be "
                f"rewritten: {exc}. Nothing was changed.",
                _SHAPE_TOOLS,
            )
        if not requested:
            # Whatever the call names is there, and nothing was asked of it.
            # The chart stays the object it was, animations and all.
            kept = new_bytes
        else:
            chart_part = package.chart_part()
            package.parts[chart_part] = new_bytes
            # A workbook the chart no longer matches is let go of; PowerPoint
            # pastes a chart without one (design section 0).
            workbook_id = _gvml_charts.external_data_id(chart_bytes)
            if workbook_id is not None and _gvml_charts.external_data_id(new_bytes) is None:
                package.remove_relationship(chart_part, workbook_id)
            strip_ids(package)

            def check(readback):
                chart = readback.chart()
                if chart is None:
                    return "the pasted shape's package holds no chart part"
                return verify(chart, state)

            try:
                replaced = replace_shape(
                    pres, slide, slide_index, index, package.to_bytes(), clip, tool_name,
                    MsoShapeType[msoChart], left, top, "chart", check, _EDIT_ALTERNATIVES,
                )
            except Refused as refused:
                return refused.payload
    finally:
        clip.restore()

    if kept is not None:
        body = {"success": True, "shape_name": name, "note": _NOTHING_ASKED}
        body.update(describe(kept, state))
        return with_warnings(body, clip)

    body = {"success": True, "shape_name": replaced.name}
    body.update(describe(replaced.package.chart(), state))
    return with_warnings(body, clip, replaced.warnings)


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
    """Rewrite the series caches in the chart's XML and paste it back.

    Windows writes the numbers into the chart's Excel workbook. There is no
    workbook here; the caches are what the chart draws from, so they are
    what is written, and a workbook the copied package carried is dropped
    rather than left saying something else.
    """
    for s in series:
        if "name" not in s or "values" not in s:
            raise ValueError("each series needs a 'name' and a 'values' list")

    def edit(chart_bytes):
        return _gvml_charts.set_data(chart_bytes, categories, series), None

    def verify(readback, _state):
        got = _gvml_charts.read_chart(readback)
        want_cats = [str(c) for c in categories]
        if got["categories"] != want_cats:
            return f"categories read back as {got['categories']}, not {want_cats}"
        want = [(s["name"], [None if v is None else float(v) for v in s["values"]]) for s in series]
        have = [(s["name"], s["values"]) for s in got["series"]]
        if have != want:
            return f"series read back as {have}, not {want}"
        return None

    def describe(_readback, _state):
        return {"categories_count": len(categories), "series_count": len(series)}

    return _edit_chart("ppt_set_chart_data", slide_index, shape_name_or_index, edit, verify, describe)


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
    """Title, legend and legend position, written into the chart's XML.

    The arguments that need the chart's rendered geometry or a whole text
    property block are refused by name first, so nothing is copied for a
    call that cannot be honoured.
    """
    refused = _refuse_arguments("ppt_format_chart", {
        "chart_style": chart_style, "legend_font_size": legend_font_size,
        "legend_top": legend_top, "legend_left": legend_left,
        "title_position": title_position, "title_top": title_top,
        "title_left": title_left,
    }, _FORMAT_CHART_UNMAPPED)
    if refused is not None:
        return refused
    if legend_position is not None:
        key = legend_position.strip().lower()
        if key in _LEGEND_DIRECTIONS:
            return _refusal(
                "ppt_format_chart",
                f"legend_position '{legend_position}' is placed from the "
                "legend's rendered width and height, which the chart's XML "
                "does not carry. The PowerPoint presets "
                f"({', '.join(_gvml_charts.LEGEND_POS)}) are written as "
                "`c:legendPos` and work. Nothing was changed.",
                ["ppt_format_chart with legend_position 'bottom', 'left', 'right', 'top' or 'corner'"],
                error="ppt_format_chart cannot set an 8-direction legend_position on macOS",
            )
        if key not in _gvml_charts.LEGEND_POS:
            raise ValueError(
                f"Unknown legend position '{legend_position}'. "
                f"PowerPoint presets: {', '.join(_gvml_charts.LEGEND_POS)}. "
                "8-direction presets: top-left, top-center, top-right, "
                "middle-left, middle-right, bottom-left, bottom-center, bottom-right."
            )

    def edit(chart_bytes):
        new = _gvml_charts.format_chart(chart_bytes, title, has_legend, legend_position)
        return new, _gvml_charts.chart_summary(new)

    def verify(readback, wanted):
        got = _gvml_charts.chart_summary(readback)
        if title is not None and got["title"] != title:
            return f"the title read back as {got['title']!r}, not {title!r}"
        if has_legend is not None and got["has_legend"] != has_legend:
            return f"has_legend read back as {got['has_legend']}"
        if legend_position is not None and got["legend_position"] != wanted["legend_position"]:
            return f"the legend position read back as {got['legend_position']!r}"
        return None

    def describe(readback, _wanted):
        got = _gvml_charts.chart_summary(readback)
        return {"has_title": got["has_title"], "has_legend": got["has_legend"]}

    # The three arguments below are every one this tool writes; the rest are
    # refused by name above. An argument added to `edit` belongs here too, or
    # a call carrying it would be taken for a call that asked for nothing.
    return _edit_chart(
        "ppt_format_chart", slide_index, shape_name_or_index, edit, verify, describe,
        requested=any(v is not None for v in (title, has_legend, legend_position)),
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
    """One axis's scale, ticks, title and number format, in the XML.

    The axis name and the Windows argument guards are checked first, the
    same wording as Windows, then the one argument with no single place in
    the XML is refused by name. Only then is the chart copied.
    """
    axis_map = _com_charts().AXIS_TYPE_MAP
    axis_key = axis.strip().lower()
    if axis_key not in axis_map:
        raise ValueError(
            f"Unknown axis '{axis}'. Valid: {', '.join(axis_map.keys())}"
        )
    is_category = axis_key == "category"
    is_value_axis = axis_key in {"value", "secondary_value"}
    if not is_category:
        if tick_label_spacing is not None:
            raise ValueError("tick_label_spacing is only valid for axis='category'.")
        if tick_mark_spacing is not None:
            raise ValueError("tick_mark_spacing is only valid for axis='category'.")
    if not is_value_axis:
        if min_scale is not None or max_scale is not None:
            raise ValueError("min_scale/max_scale are only valid for value axes.")
        if major_unit is not None or minor_unit is not None:
            raise ValueError("major_unit/minor_unit are only valid for value axes.")
        if log_scale is not None:
            raise ValueError("log_scale is only valid for value axes.")
    if log_base is not None and log_scale is not True:
        raise ValueError("log_base requires log_scale=true.")
    refused = _refuse_arguments(
        "ppt_format_chart_axis", {"tick_label_font_size": tick_label_font_size}, _FORMAT_AXIS_UNMAPPED,
    )
    if refused is not None:
        return refused

    fields = {
        "title": title, "min_scale": min_scale, "max_scale": max_scale,
        "major_unit": major_unit, "minor_unit": minor_unit,
        "tick_label_spacing": tick_label_spacing, "tick_mark_spacing": tick_mark_spacing,
        "major_tick_mark": major_tick_mark, "minor_tick_mark": minor_tick_mark,
        "reverse_order": reverse_order, "log_scale": log_scale, "log_base": log_base,
        "number_format": number_format,
    }

    def edit(chart_bytes):
        new, applied = _gvml_charts.format_axis(chart_bytes, axis_key, **fields)
        return new, (applied, _gvml_charts.read_axis(new, axis_key))

    def verify(readback, state):
        applied, wanted = state
        got = _gvml_charts.read_axis(readback, axis_key)
        if got is None:
            return f"the pasted chart has no '{axis_key}' axis"
        for field in applied:
            if got.get(field) != wanted.get(field):
                return f"{field} read back as {got.get(field)!r}, not {wanted.get(field)!r}"
        return None

    def describe(_readback, state):
        return {"axis": axis_key, "applied": state[0]}

    return _edit_chart(
        "ppt_format_chart_axis", slide_index, shape_name_or_index, edit, verify, describe,
        requested=any(v is not None for v in fields.values()),
    )


def _set_chart_series_impl(
    slide_index, shape_name_or_index,
    series_index, color, show_data_labels, line_weight,
):
    """One series's colour, data labels and line weight, in the XML.

    ``fill format`` on the chart shape does exist and paints the chart's
    background, not a series, so it is left alone; the series' own ``c:spPr``
    is written instead.
    """
    def edit(chart_bytes):
        new = _gvml_charts.set_series(chart_bytes, series_index, color, show_data_labels, line_weight)
        return new, _gvml_charts.read_series(new, series_index)

    def verify(readback, wanted):
        got = _gvml_charts.read_series(readback, series_index)
        if got is None:
            return f"the pasted chart has no series {series_index}"
        for field, value in (("color", color), ("show_data_labels", show_data_labels), ("line_weight", line_weight)):
            if value is not None and got.get(field) != wanted.get(field):
                return f"{field} read back as {got.get(field)!r}, not {wanted.get(field)!r}"
        return None

    def describe(_readback, _wanted):
        return {"series_index": series_index}

    return _edit_chart(
        "ppt_set_chart_series", slide_index, shape_name_or_index, edit, verify, describe,
        requested=any(v is not None for v in (color, show_data_labels, line_weight)),
    )


def _change_chart_type_impl(slide_index, shape_name_or_index, chart_type):
    """Swap the plot for one of another kind, keeping the data, and paste.

    The requested type is resolved first, so a typo is heard before the
    chart is touched, and a type the writer has no template for is refused
    by argument the way ``ppt_add_chart`` refuses it.
    """
    com = _com_charts()
    type_int = com._resolve_chart_type(chart_type)
    type_name = com.CHART_TYPE_NAMES.get(type_int, str(type_int))
    if type_int not in _gvml_charts.SPECS:
        supported = sorted(
            com.CHART_TYPE_NAMES[t] for t in _gvml_charts.supported_types()
            if t in com.CHART_TYPE_NAMES
        )
        return _refusal(
            "ppt_change_chart_type",
            f"XlChartType {type_int} has no GVML template yet. On macOS a "
            "chart's kind is changed by rewriting its XML and pasting it, and "
            f"no template exists for this kind. Kinds that can be drawn: "
            f"{', '.join(supported)}. Nothing was changed.",
            _EDIT_ALTERNATIVES[1:],
            error=f"ppt_change_chart_type cannot draw chart_type {chart_type!r} on macOS",
        )
    expected_kind = _gvml_charts.SPECS[type_int].kind + "Chart"

    def edit(chart_bytes):
        return _gvml_charts.set_type(chart_bytes, type_int), _gvml_charts.read_chart(chart_bytes)

    def verify(readback, before):
        kind = _gvml_charts.kind_of(readback)
        if kind != expected_kind:
            return f"the plot read back as {kind}, not {expected_kind}"
        after = _gvml_charts.read_chart(readback)
        if [s["values"] for s in after["series"]] != [s["values"] for s in before["series"]]:
            return "the series values did not survive the change of kind"
        return None

    def describe(_readback, _before):
        return {"new_chart_type": type_name, "new_chart_type_int": type_int}

    return _edit_chart("ppt_change_chart_type", slide_index, shape_name_or_index, edit, verify, describe)
