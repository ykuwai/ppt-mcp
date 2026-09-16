"""Charts: an ``XlChartType`` integer to ``chart1.xml``, and back to its data.

The writer takes the Windows integer, not the friendly name. Resolving
``'column'`` to 51 stays where it is in ``ppt_com/charts.py``, so a name added
there is spelled the same in a Mac caller's error, and this table only has to
know what each integer looks like in XML. Two tables, one step each; see
docs/gvml-design.md section 1.

Every template's element order was checked against ``chart1.xml`` as
PowerPoint itself writes it after a paste, which is the fixture
``tests/fixtures/gvml/chart.gvml.zip`` holds, and a chart written here has to
stay an ordered subset of that. The hand written 1.6 KB chart the design
measured is what these grew from.

No workbook. PowerPoint pastes a chart whose data lives only in the caches
and does not add one afterwards (design section 0, probe 5 and 6). What the
"Edit Data" button does to such a chart is not known and is listed among the
things only a hand can check.
"""

import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import Dict, List, Optional, Sequence

from gvml.package import NS_A, NS_C, NS_R, XML_DECL, PackageError

_C = f"{{{NS_C}}}"


class ChartTypeError(ValueError):
    """An ``XlChartType`` this writer has no template for."""


@dataclass(frozen=True)
class _Spec:
    kind: str                 # element name without "Chart": bar, line, pie...
    bar_dir: str = "col"
    grouping: str = "clustered"
    markers: bool = True      # line, scatter, radar
    lines: bool = True        # scatter
    explosion: int = 0        # pie
    three_d: bool = False


# XlChartType -> what to write. The integers are Excel's and match
# ppt_com/charts.py CHART_TYPE_MAP.
SPECS: Dict[int, _Spec] = {
    51: _Spec("bar", "col", "clustered"),
    52: _Spec("bar", "col", "stacked"),
    53: _Spec("bar", "col", "percentStacked"),
    57: _Spec("bar", "bar", "clustered"),
    58: _Spec("bar", "bar", "stacked"),
    4: _Spec("line", grouping="standard", markers=False),
    65: _Spec("line", grouping="standard", markers=True),
    63: _Spec("line", grouping="stacked", markers=False),
    5: _Spec("pie"),
    69: _Spec("pie", explosion=25),
    -4120: _Spec("doughnut"),
    1: _Spec("area", grouping="standard"),
    76: _Spec("area", grouping="stacked"),
    -4169: _Spec("scatter", lines=False),
    74: _Spec("scatter", lines=True),
    -4151: _Spec("radar", markers=False),
    15: _Spec("bubble"),
    54: _Spec("bar3D", "col", "clustered", three_d=True),
    -4102: _Spec("pie3D", three_d=True),
    -4101: _Spec("line3D", grouping="standard", three_d=True),
}

# What PowerPoint puts in a new chart, so a chart made here and one made on
# Windows start from the same numbers.
DEFAULT_CATEGORIES = ["Category 1", "Category 2", "Category 3", "Category 4"]
DEFAULT_SERIES = [
    {"name": "Series 1", "values": [4.3, 2.5, 3.5, 4.5]},
    {"name": "Series 2", "values": [2.4, 4.4, 1.8, 2.8]},
    {"name": "Series 3", "values": [2.0, 2.0, 3.0, 5.0]},
]
DEFAULT_PIE_CATEGORIES = ["1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr"]
DEFAULT_PIE_SERIES = [{"name": "Sales", "values": [8.2, 3.2, 1.4, 1.2]}]
DEFAULT_XY_X = [0.7, 1.8, 2.6, 3.2, 4.1, 5.0]
DEFAULT_XY_SERIES = [{"name": "Y Values", "values": [2.7, 3.2, 0.8, 1.2, 1.1, 2.5]}]
DEFAULT_BUBBLE_X = [0.7, 1.8, 2.6]
DEFAULT_BUBBLE_SERIES = [{"name": "Y Values", "values": [2.7, 3.2, 0.8], "sizes": [10.0, 4.0, 8.0]}]


def _escape(text) -> str:
    return (
        str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    )


def _column(i: int) -> str:
    """0 -> B, 1 -> C, ... the sheet column a series would live in."""
    n = i + 1
    letters = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return letters


def _num(value) -> str:
    if value is None:
        return ""
    text = repr(float(value))
    return text[:-2] if text.endswith(".0") else text


def _str_cache(values: Sequence) -> str:
    pts = "".join(
        f'<c:pt idx="{i}"><c:v>{_escape(v)}</c:v></c:pt>'
        for i, v in enumerate(values) if v is not None
    )
    return f'<c:strCache><c:ptCount val="{len(values)}"/>{pts}</c:strCache>'


def _num_cache(values: Sequence) -> str:
    pts = "".join(
        f'<c:pt idx="{i}"><c:v>{_num(v)}</c:v></c:pt>'
        for i, v in enumerate(values) if v is not None
    )
    return (
        '<c:numCache><c:formatCode>General</c:formatCode>'
        f'<c:ptCount val="{len(values)}"/>{pts}</c:numCache>'
    )


def _str_ref(formula: str, values: Sequence) -> str:
    return f"<c:strRef><c:f>{formula}</c:f>{_str_cache(values)}</c:strRef>"


def _num_ref(formula: str, values: Sequence) -> str:
    return f"<c:numRef><c:f>{formula}</c:f>{_num_cache(values)}</c:numRef>"


def _series_xml(spec: _Spec, i: int, series: dict, categories: Sequence, n: int) -> str:
    """One ``c:ser``. Element order follows the schema for each chart kind."""
    col = _column(i + 1)
    head = (
        f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/>'
        f'<c:tx>{_str_ref(f"Sheet1!${col}$1", [series["name"]])}</c:tx>'
    )
    cat = f'<c:cat>{_str_ref(f"Sheet1!$A$2:$A${n + 1}", categories)}</c:cat>'
    val = f'<c:val>{_num_ref(f"Sheet1!${col}$2:${col}${n + 1}", series["values"])}</c:val>'
    kind = spec.kind
    if kind in ("bar", "bar3D"):
        return head + '<c:invertIfNegative val="0"/>' + cat + val + "</c:ser>"
    if kind in ("line", "radar"):
        marker = "" if spec.markers else '<c:marker><c:symbol val="none"/></c:marker>'
        return head + marker + cat + val + ("<c:smooth val=\"0\"/>" if kind == "line" else "") + "</c:ser>"
    if kind == "line3D":
        return head + cat + val + "</c:ser>"
    if kind in ("pie", "pie3D", "doughnut"):
        explosion = f'<c:explosion val="{spec.explosion}"/>' if spec.explosion else ""
        return head + explosion + cat + val + "</c:ser>"
    if kind == "area":
        return head + cat + val + "</c:ser>"
    if kind == "scatter":
        line = "" if spec.lines else '<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>'
        x_val = f'<c:xVal>{_num_ref(f"Sheet1!$A$2:$A${n + 1}", categories)}</c:xVal>'
        y_val = f'<c:yVal>{_num_ref(f"Sheet1!${col}$2:${col}${n + 1}", series["values"])}</c:yVal>'
        return head + line + x_val + y_val + '<c:smooth val="0"/></c:ser>'
    if kind == "bubble":
        size_col = _column(i + 2)
        x_val = f'<c:xVal>{_num_ref(f"Sheet1!$A$2:$A${n + 1}", categories)}</c:xVal>'
        y_val = f'<c:yVal>{_num_ref(f"Sheet1!${col}$2:${col}${n + 1}", series["values"])}</c:yVal>'
        sizes = series.get("sizes") or [1.0] * n
        size = f'<c:bubbleSize>{_num_ref(f"Sheet1!${size_col}$2:${size_col}${n + 1}", sizes)}</c:bubbleSize>'
        return head + '<c:invertIfNegative val="0"/>' + x_val + y_val + size + "</c:ser>"
    raise ChartTypeError(f"no series template for {kind}")


def _axes(spec: _Spec) -> str:
    kind = spec.kind
    if kind in ("pie", "pie3D", "doughnut"):
        return ""
    cat_ax = (
        '<c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="0"/>'
        '<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
        '<c:crossAx val="20"/><c:crosses val="autoZero"/><c:auto val="1"/>'
        '<c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'
    )
    val_ax = (
        '<c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/>'
        '<c:numFmt formatCode="General" sourceLinked="1"/>'
        '<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
        '<c:crossAx val="10"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>'
    )
    if kind in ("scatter", "bubble"):
        x_ax = (
            '<c:valAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
            '<c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/>'
            '<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
            '<c:crossAx val="20"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>'
        )
        return x_ax + val_ax
    if kind == "line3D":
        ser_ax = (
            '<c:serAx><c:axId val="30"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
            '<c:delete val="0"/><c:axPos val="b"/><c:majorTickMark val="out"/>'
            '<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="20"/>'
            '<c:crosses val="autoZero"/></c:serAx>'
        )
        return cat_ax + val_ax + ser_ax
    return cat_ax + val_ax


def _plot(spec: _Spec, series_xml: str) -> str:
    kind = spec.kind
    ax = '<c:axId val="10"/><c:axId val="20"/>'
    if kind == "bar":
        overlap = '<c:overlap val="100"/>' if spec.grouping != "clustered" else ""
        return (
            f'<c:barChart><c:barDir val="{spec.bar_dir}"/><c:grouping val="{spec.grouping}"/>'
            f'<c:varyColors val="0"/>{series_xml}<c:gapWidth val="150"/>{overlap}{ax}</c:barChart>'
        )
    if kind == "bar3D":
        return (
            f'<c:bar3DChart><c:barDir val="{spec.bar_dir}"/><c:grouping val="{spec.grouping}"/>'
            f'<c:varyColors val="0"/>{series_xml}<c:gapWidth val="150"/><c:shape val="box"/>{ax}</c:bar3DChart>'
        )
    if kind == "line":
        return (
            f'<c:lineChart><c:grouping val="{spec.grouping}"/><c:varyColors val="0"/>'
            f'{series_xml}<c:marker val="1"/>{ax}</c:lineChart>'
        )
    if kind == "line3D":
        return (
            f'<c:line3DChart><c:grouping val="{spec.grouping}"/><c:varyColors val="0"/>'
            f'{series_xml}{ax}<c:axId val="30"/></c:line3DChart>'
        )
    if kind == "pie":
        return f'<c:pieChart><c:varyColors val="1"/>{series_xml}<c:firstSliceAng val="0"/></c:pieChart>'
    if kind == "pie3D":
        return f'<c:pie3DChart><c:varyColors val="1"/>{series_xml}</c:pie3DChart>'
    if kind == "doughnut":
        return (
            f'<c:doughnutChart><c:varyColors val="1"/>{series_xml}'
            '<c:firstSliceAng val="0"/><c:holeSize val="75"/></c:doughnutChart>'
        )
    if kind == "area":
        return (
            f'<c:areaChart><c:grouping val="{spec.grouping}"/><c:varyColors val="0"/>'
            f'{series_xml}{ax}</c:areaChart>'
        )
    if kind == "scatter":
        return (
            f'<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>'
            f'{series_xml}{ax}</c:scatterChart>'
        )
    if kind == "radar":
        style = "marker" if spec.markers else "standard"
        return (
            f'<c:radarChart><c:radarStyle val="{style}"/><c:varyColors val="0"/>'
            f'{series_xml}{ax}</c:radarChart>'
        )
    if kind == "bubble":
        return (
            f'<c:bubbleChart><c:varyColors val="0"/>{series_xml}'
            f'<c:bubbleScale val="100"/><c:showNegBubbles val="0"/>{ax}</c:bubbleChart>'
        )
    raise ChartTypeError(f"no plot template for {kind}")


def default_data(type_int: int):
    """The categories and series a new chart of this type starts with."""
    spec = SPECS[type_int]
    if spec.kind in ("pie", "pie3D", "doughnut"):
        return list(DEFAULT_PIE_CATEGORIES), [dict(s) for s in DEFAULT_PIE_SERIES]
    if spec.kind == "scatter":
        return list(DEFAULT_XY_X), [dict(s) for s in DEFAULT_XY_SERIES]
    if spec.kind == "bubble":
        return list(DEFAULT_BUBBLE_X), [dict(s) for s in DEFAULT_BUBBLE_SERIES]
    return list(DEFAULT_CATEGORIES), [dict(s) for s in DEFAULT_SERIES]


def supported_types() -> List[int]:
    return list(SPECS)


def chart_xml(
    type_int: int,
    categories: Optional[Sequence] = None,
    series: Optional[Sequence[dict]] = None,
) -> str:
    """``chart1.xml`` for an ``XlChartType``, with the given or default data.

    Raises ``ChartTypeError`` for a type with no template, naming it, so the
    caller can refuse by argument rather than by tool.
    """
    spec = SPECS.get(type_int)
    if spec is None:
        raise ChartTypeError(
            f"XlChartType {type_int} has no GVML template yet. Supported: "
            + ", ".join(str(t) for t in SPECS)
        )
    if categories is None or series is None:
        default_cats, default_series = default_data(type_int)
        categories = default_cats if categories is None else categories
        series = default_series if series is None else series
    if not series:
        raise ValueError("a chart needs at least one series")
    n = len(categories)
    for s in series:
        if len(s["values"]) != n:
            raise ValueError(
                f"series '{s['name']}' has {len(s['values'])} values for {n} categories"
            )
    series_xml = "".join(_series_xml(spec, i, s, categories, n) for i, s in enumerate(series))
    view3d = ""
    if spec.three_d:
        if spec.kind == "pie3D":
            view3d = '<c:view3D><c:rotX val="30"/><c:rotY val="0"/><c:rAngAx val="0"/></c:view3D>'
        else:
            view3d = '<c:view3D><c:rotX val="15"/><c:rotY val="20"/><c:rAngAx val="1"/></c:view3D>'
    legend = '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>'
    return (
        XML_DECL
        + f'<c:chartSpace xmlns:c="{NS_C}" xmlns:a="{NS_A}" xmlns:r="{NS_R}">'
        + '<c:roundedCorners val="0"/>'
        + f'<c:chart><c:autoTitleDeleted val="0"/>{view3d}<c:plotArea><c:layout/>'
        + _plot(spec, series_xml)
        + _axes(spec)
        + "</c:plotArea>"
        + legend
        + '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>'
    )


# ---------------------------------------------------------------------------
# Reading
# ---------------------------------------------------------------------------

def _cache_values(container: Optional[ET.Element], numeric: bool) -> List:
    """The points of a ``c:cat``, ``c:val``, ``c:xVal`` or ``c:yVal``.

    Handles the reference forms with their caches and the literal forms.
    Gaps, a point index with no ``c:pt``, come back as None so the list is
    as long as ``ptCount`` says and the indices still line up.
    """
    if container is None:
        return []
    cache = None
    for path in ("numRef/numCache", "strRef/strCache", "numLit", "strLit",
                 "multiLvlStrRef/multiLvlStrCache/lvl"):
        cache = container.find("/".join(_C + step for step in path.split("/")))
        if cache is not None:
            break
    if cache is None:
        return []
    count_el = cache.find(f"{_C}ptCount")
    pts = cache.findall(f"{_C}pt")
    if count_el is not None:
        count = int(count_el.get("val", 0))
    else:
        count = max((int(p.get("idx", 0)) for p in pts), default=-1) + 1
    if count_el is None:
        # A level of a multi level cache carries no ptCount of its own.
        parent_count = container.find(f"{_C}multiLvlStrRef/{_C}multiLvlStrCache/{_C}ptCount")
        if parent_count is not None:
            count = int(parent_count.get("val", count))
    values: List = [None] * count
    for p in pts:
        idx = int(p.get("idx", 0))
        v = p.find(f"{_C}v")
        text = v.text if v is not None else None
        if idx >= count:
            values.extend([None] * (idx - count + 1))
            count = idx + 1
        if numeric:
            try:
                values[idx] = float(text) if text not in (None, "") else None
            except ValueError:
                values[idx] = text
        else:
            values[idx] = "" if text is None else text
    return values


def _series_name(ser: ET.Element) -> str:
    tx = ser.find(f"{_C}tx")
    if tx is None:
        return ""
    v = tx.find(f"{_C}strRef/{_C}strCache/{_C}pt/{_C}v")
    if v is None:
        v = tx.find(f"{_C}v")
    return (v.text or "") if v is not None else ""


def kind_of(chart_xml_bytes: bytes) -> Optional[str]:
    """The element name of the first plot, e.g. ``barChart``, or None."""
    root = _parse(chart_xml_bytes)
    plot = root.find(f"{_C}chart/{_C}plotArea")
    if plot is None:
        return None
    for child in plot:
        if child.tag.startswith(_C) and child.tag.endswith("Chart"):
            return child.tag[len(_C):]
    return None


def _parse(chart_xml_bytes: bytes) -> ET.Element:
    try:
        root = ET.fromstring(chart_xml_bytes)
    except ET.ParseError as exc:
        raise PackageError(f"chart1.xml is not well formed XML: {exc}") from exc
    if root.tag != f"{_C}chartSpace":
        raise PackageError(f"chart1.xml's root is {root.tag}, not c:chartSpace")
    return root


def read_chart(chart_xml_bytes: bytes) -> dict:
    """Categories and series, in the form ``ppt_get_chart_data`` returns.

    ``categories`` come from the first series that has any, as strings, and
    each series is ``{"name", "values"}`` with floats. Scatter and bubble
    charts keep their x values as the categories, which is what Windows
    reports through ``XValues`` too.
    """
    root = _parse(chart_xml_bytes)
    plot = root.find(f"{_C}chart/{_C}plotArea")
    if plot is None:
        raise PackageError("chart1.xml has no plot area")
    categories: List[str] = []
    series_out: List[dict] = []
    for chart in plot:
        if not (chart.tag.startswith(_C) and chart.tag.endswith("Chart")):
            continue
        for ser in chart.findall(f"{_C}ser"):
            values = _cache_values(ser.find(f"{_C}val"), numeric=True)
            if not values:
                values = _cache_values(ser.find(f"{_C}yVal"), numeric=True)
            series_out.append({"name": _series_name(ser), "values": values})
            if not categories:
                cats = _cache_values(ser.find(f"{_C}cat"), numeric=False)
                if not cats:
                    # Numeric x values come back as str(float), which is what
                    # Windows reports through XValues: "5.0", not "5".
                    cats = _cache_values(ser.find(f"{_C}xVal"), numeric=True)
                categories = ["" if c is None else str(c) for c in cats]
    return {"categories": categories, "series": series_out}


def chart_package(name: str, x: int, y: int, cx: int, cy: int, chart_xml_text: str) -> bytes:
    """A whole package holding one chart frame at ``x, y, cx, cy`` EMU."""
    from gvml.canvas import wrap
    from gvml.package import CHART_PART, CT_CHART, REL_CHART, Relationship, build
    from gvml.shapes import chart_frame_xml

    drawing = wrap(chart_frame_xml(2, name, x, y, cx, cy), x, y, cx, cy)
    return build(
        drawing,
        parts={CHART_PART: chart_xml_text.encode("utf-8")},
        drawing_rels=[Relationship("rId1", REL_CHART, CHART_PART)],
        overrides={CHART_PART: CT_CHART},
    )


# ---------------------------------------------------------------------------
# Editing a chart1.xml PowerPoint wrote
# ---------------------------------------------------------------------------
# The five editing tools copy a chart off the slide, change one thing in its
# chart1.xml and paste it back (design section 6, second tier). Everything the
# caller did not ask about is left exactly as PowerPoint wrote it, which is why
# these edit the tree rather than regenerating it from the writer above.
#
# Element order is what PowerPoint checks and does not report, so every
# insertion goes through ``_insert_ordered`` with the schema's order for that
# parent. The lists below are the CT_* sequences from the chart schema, with
# the children of every axis kind merged into one list so one helper serves
# catAx, valAx, dateAx and serAx.

import copy as _copy

_MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"

_CHART_ORDER = [
    "title", "autoTitleDeleted", "pivotFmts", "view3D", "floor", "sideWall",
    "backWall", "plotArea", "legend", "plotVisOnly", "dispBlanksAs",
    "showDLblsOverMax", "extLst",
]
_CHART_SPACE_ORDER = [
    "date1904", "lang", "roundedCorners", "AlternateContent", "style", "clrMapOvr",
    "pivotSource", "protection", "chart", "spPr", "txPr", "externalData",
    "printSettings", "userShapes", "extLst",
]
_SER_ORDER = [
    "idx", "order", "tx", "spPr", "invertIfNegative", "pictureOptions", "marker",
    "explosion", "dPt", "dLbls", "trendline", "errBars", "cat", "val", "xVal",
    "yVal", "bubbleSize", "bubble3D", "shape", "smooth", "extLst",
]
_AXIS_ORDER = [
    "axId", "scaling", "delete", "axPos", "majorGridlines", "minorGridlines",
    "title", "numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr",
    "txPr", "crossAx", "crosses", "crossesAt", "auto", "lblAlgn", "lblOffset",
    "tickLblSkip", "tickMarkSkip", "noMultiLvlLbl", "crossBetween",
    "baseTimeUnit", "majorUnit", "majorTimeUnit", "minorUnit", "minorTimeUnit",
    "dispUnits", "extLst",
]
_SCALING_ORDER = ["logBase", "orientation", "max", "min", "extLst"]
_LEGEND_ORDER = ["legendPos", "legendEntry", "layout", "overlay", "spPr", "txPr", "extLst"]
_DLBLS_ORDER = [
    "dLbl", "delete", "numFmt", "spPr", "txPr", "dLblPos", "showLegendKey",
    "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize",
    "separator", "showLeaderLines", "leaderLines", "extLst",
]
# What a plot element holds before its series, by kind, so new series go
# after the last of these and before dLbls, gapWidth and the axis ids.
_PLOT_HEAD = {"barDir", "grouping", "varyColors", "scatterStyle", "radarStyle",
              "wireframe", "ofPieType"}

AXIS_TAGS = ("catAx", "valAx", "dateAx", "serAx")
PLOT_KINDS = {tag for tag in (
    "barChart", "bar3DChart", "lineChart", "line3DChart", "pieChart", "pie3DChart",
    "doughnutChart", "areaChart", "area3DChart", "scatterChart", "radarChart",
    "bubbleChart", "stockChart", "surfaceChart", "surface3DChart", "ofPieChart",
)}

# Windows legend positions to c:legendPos. The 8-direction presets Windows
# computes from the legend's rendered size are not here; the XML does not
# carry that size.
LEGEND_POS = {"bottom": "b", "left": "l", "right": "r", "top": "t", "corner": "tr"}
TICK_MARK = {"none": "none", "inside": "in", "outside": "out", "cross": "cross"}


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _insert_ordered(parent: ET.Element, child: ET.Element, order: List[str]) -> ET.Element:
    """Put ``child`` where the schema says it goes, replacing one already there.

    An element the schema lists once (a title, a legend) is replaced in place
    when present. Otherwise the child lands after the last existing sibling
    whose name precedes it in ``order``.
    """
    name = _local(child.tag)
    if name not in order:
        raise ValueError(f"{name} has no place in this parent's order")
    rank = order.index(name)
    existing = [c for c in parent if _local(c.tag) == name]
    if existing and name not in ("dPt", "legendEntry", "trendline", "dLbl", "pivotFmt"):
        parent.insert(list(parent).index(existing[0]), child)
        parent.remove(existing[0])
        return child
    position = 0
    for i, sibling in enumerate(parent):
        local = _local(sibling.tag)
        if local in order and order.index(local) <= rank:
            position = i + 1
    parent.insert(position, child)
    return child


def _remove(parent: ET.Element, name: str) -> bool:
    found = [c for c in parent if _local(c.tag) == name]
    for c in found:
        parent.remove(c)
    return bool(found)


def _el(name: str, **attrs) -> ET.Element:
    e = ET.Element(_C + name)
    for key, value in attrs.items():
        e.set(key, str(value))
    return e


def _val(name: str, value) -> ET.Element:
    return _el(name, val=value)


def load(chart_xml_bytes: bytes) -> ET.Element:
    """Parse a chart part with its prefixes kept, or say what is wrong."""
    from gvml.package import register_prefixes

    register_prefixes(chart_xml_bytes)
    return _parse(chart_xml_bytes)


def dump(root: ET.Element) -> bytes:
    return (XML_DECL + ET.tostring(root, encoding="unicode")).encode("utf-8")


def _plot_area(root: ET.Element) -> ET.Element:
    plot = root.find(f"{_C}chart/{_C}plotArea")
    if plot is None:
        raise PackageError("chart1.xml has no plot area")
    return plot


def _plots(plot_area: ET.Element) -> List[ET.Element]:
    return [c for c in plot_area if _local(c.tag) in PLOT_KINDS]


def _all_series(plot_area: ET.Element) -> List[ET.Element]:
    return [ser for plot in _plots(plot_area) for ser in plot.findall(f"{_C}ser")]


def _rich_text(text: str) -> ET.Element:
    """``c:tx/c:rich`` holding one run, the shape of a title PowerPoint writes."""
    tx = _el("tx")
    rich = ET.SubElement(tx, _C + "rich")
    ET.SubElement(rich, f"{{{NS_A}}}bodyPr")
    ET.SubElement(rich, f"{{{NS_A}}}lstStyle")
    p = ET.SubElement(rich, f"{{{NS_A}}}p")
    r = ET.SubElement(p, f"{{{NS_A}}}r")
    t = ET.SubElement(r, f"{{{NS_A}}}t")
    t.text = text
    return tx


def _title_text(title: Optional[ET.Element]) -> Optional[str]:
    if title is None:
        return None
    return "".join(t.text or "" for t in title.iter(f"{{{NS_A}}}t"))


def _set_title(parent: ET.Element, text: str, order: List[str]) -> None:
    title = _el("title")
    title.append(_rich_text(text))
    title.append(_val("overlay", 0))
    _insert_ordered(parent, title, order)


# -- data --------------------------------------------------------------------

def _cache_el(numeric: bool, formula: str, values: Sequence) -> ET.Element:
    xml = _num_ref(formula, values) if numeric else _str_ref(formula, values)
    return ET.fromstring(f'<w xmlns:c="{NS_C}">{xml}</w>')[0]


def _set_ref(ser: ET.Element, name: str, numeric: bool, formula: str, values: Sequence) -> None:
    holder = _el(name)
    holder.append(_cache_el(numeric, formula, values))
    _insert_ordered(ser, holder, _SER_ORDER)


def set_data(chart_xml_bytes: bytes, categories: Sequence, series: Sequence[dict]) -> bytes:
    """Rewrite the series of the first plot, and nothing else.

    Each new series takes over the ``c:ser`` at its position (so a colour or a
    label setting on series 2 stays with series 2); series past the old count
    are cloned from the last one with its own formatting stripped, so they
    take the theme colour of their index the way a new series does. The
    ``c:externalData`` pointer is dropped, because the workbook it points at
    would no longer say what the chart says; the caller drops the part.
    """
    if not series:
        raise ValueError("a chart needs at least one series")
    n = len(categories)
    for s in series:
        if len(s.get("values", [])) != n:
            raise ValueError(
                f"series '{s.get('name', '')}' has {len(s.get('values', []))} "
                f"values for {n} categories"
            )
    root = load(chart_xml_bytes)
    plot_area = _plot_area(root)
    plots = _plots(plot_area)
    if not plots:
        raise PackageError("chart1.xml has no plot to put series into")
    plot = plots[0]
    old = plot.findall(f"{_C}ser")
    if not old:
        raise PackageError("the chart's first plot holds no series to take as a template")
    kind = _local(plot.tag)
    xy = kind in ("scatterChart", "bubbleChart")

    insert_at = list(plot).index(old[0])
    for ser in old:
        plot.remove(ser)
    template = _copy.deepcopy(old[-1])
    for name in ("spPr", "dPt", "dLbls", "marker", "trendline", "errBars"):
        _remove(template, name)

    for i, s in enumerate(series):
        ser = _copy.deepcopy(old[i]) if i < len(old) else _copy.deepcopy(template)
        col = _column(i + 1)
        _insert_ordered(ser, _val("idx", i), _SER_ORDER)
        _insert_ordered(ser, _val("order", i), _SER_ORDER)
        tx = _el("tx")
        tx.append(_cache_el(False, f"Sheet1!${col}$1", [s["name"]]))
        _insert_ordered(ser, tx, _SER_ORDER)
        if xy:
            _remove(ser, "cat")
            _remove(ser, "val")
            _set_ref(ser, "xVal", True, f"Sheet1!$A$2:$A${n + 1}", categories)
            _set_ref(ser, "yVal", True, f"Sheet1!${col}$2:${col}${n + 1}", s["values"])
            if kind == "bubbleChart":
                size_col = _column(i + 2)
                sizes = s.get("sizes") or [1.0] * n
                _set_ref(ser, "bubbleSize", True, f"Sheet1!${size_col}$2:${size_col}${n + 1}", sizes)
        else:
            _remove(ser, "xVal")
            _remove(ser, "yVal")
            _remove(ser, "bubbleSize")
            _set_ref(ser, "cat", False, f"Sheet1!$A$2:$A${n + 1}", categories)
            _set_ref(ser, "val", True, f"Sheet1!${col}$2:${col}${n + 1}", s["values"])
        plot.insert(insert_at + i, ser)

    _remove(root, "externalData")
    return dump(root)


def external_data_id(chart_xml_bytes: bytes) -> Optional[str]:
    """The ``r:id`` of the chart's workbook pointer, if it has one."""
    root = _parse(chart_xml_bytes)
    ext = root.find(f"{_C}externalData")
    return None if ext is None else ext.get(f"{{{NS_R}}}id")


# -- type --------------------------------------------------------------------

def _series_data(ser: ET.Element) -> dict:
    values = _cache_values(ser.find(f"{_C}val"), numeric=True)
    if not values:
        values = _cache_values(ser.find(f"{_C}yVal"), numeric=True)
    data = {"name": _series_name(ser), "values": values}
    sizes = _cache_values(ser.find(f"{_C}bubbleSize"), numeric=True)
    if sizes:
        data["sizes"] = sizes
    return data


def set_type(chart_xml_bytes: bytes, type_int: int) -> bytes:
    """Replace the plot with one of another kind, carrying the data across.

    The series' names and numbers move over; the plot level formatting (gap
    width, data labels, per series colours) is the new kind's default, which
    is also what Windows leaves after ``Chart.ChartType``. Axes that the new
    kind can keep (a category and a value axis) are kept with everything set
    on them; a kind that needs different axes (pie none, scatter two value
    axes, 3D line a series axis) gets fresh ones.

    A combo chart holds more than one plot group, a bar one and a line one
    over a secondary axis, say. Every group's series comes across, in the
    order the plot area holds them, which is the order ``read_chart``
    reports and the order the caller saw. The secondary axis itself goes,
    because one plot has one pair, but no series goes with it. What cannot
    come across is a group plotted against category labels of its own,
    since a single plot carries one set of them, and that is refused by
    name rather than quietly relabelled.
    """
    spec = SPECS.get(type_int)
    if spec is None:
        raise ChartTypeError(
            f"XlChartType {type_int} has no GVML template yet. Supported: "
            + ", ".join(str(t) for t in SPECS)
        )
    root = load(chart_xml_bytes)
    chart = root.find(f"{_C}chart")
    plot_area = _plot_area(root)
    plots = _plots(plot_area)
    if not plots:
        raise PackageError("chart1.xml has no plot to change the kind of")
    old_plot = plots[0]
    old_sers = _all_series(plot_area)
    if not old_sers:
        raise PackageError("the chart's plot holds no series")

    series = [_series_data(s) for s in old_sers]
    labels = [_cache_values(s.find(f"{_C}cat"), numeric=False) for s in old_sers]
    cats = next((c for c in labels if c), [])
    if len(plots) > 1:
        # One plot carries one set of category labels, so a group plotted
        # against its own cannot be carried over; say which series rather
        # than relabel it or leave it behind.
        odd = [s for s, c in zip(series, labels) if c and c != cats]
        if odd:
            raise PackageError(
                "the chart plots " + ", ".join(f"'{s['name']}'" for s in odd)
                + " against categories of its own, and one plot carries one "
                "set of category labels, so the series cannot be carried "
                "across without relabelling it"
            )
    if not cats:
        cats = _cache_values(old_sers[0].find(f"{_C}xVal"), numeric=True)
    n = max(len(s["values"]) for s in series)
    cats = list(cats) + [None] * (n - len(cats))
    if spec.kind in ("scatter", "bubble"):
        numeric_cats = []
        for i, c in enumerate(cats[:n]):
            try:
                numeric_cats.append(float(c))
            except (TypeError, ValueError):
                numeric_cats.append(float(i + 1))
        cats = numeric_cats
    else:
        cats = ["" if c is None else str(c) for c in cats[:n]]
    for s in series:
        s["values"] = [None if v is None else v for v in s["values"]]

    series_xml = "".join(_series_xml(spec, i, s, cats, n) for i, s in enumerate(series))
    new_plot = ET.fromstring(
        f'<w xmlns:c="{NS_C}" xmlns:a="{NS_A}">{_plot(spec, series_xml)}</w>'
    )[0]

    old_axis_ids = [a.get("val") for a in old_plot.findall(f"{_C}axId")]
    old_axes = [a for a in plot_area if _local(a.tag) in AXIS_TAGS]
    old_kinds = [_local(a.tag) for a in old_axes]
    wants = [_local(a.tag) for a in ET.fromstring(
        f'<w xmlns:c="{NS_C}" xmlns:a="{NS_A}">{_axes(spec)}</w>'
    )]
    keep_axes = wants == old_kinds[:len(wants)] and len(old_axis_ids) == len(wants) and wants

    # Everything that is not a plot or an axis of the first plot stays.
    first_axes = {a for a in old_axes if a.find(f"{_C}axId").get("val") in old_axis_ids}
    for child in list(plot_area):
        if child is old_plot or (child in first_axes and not keep_axes):
            plot_area.remove(child)
    for plot in plots[1:]:
        # A second plot group has no place on the new kind; its series are
        # already in the new plot and only the empty group goes.
        plot_area.remove(plot)
    for axis in old_axes:
        if axis not in first_axes:
            plot_area.remove(axis)

    layout = plot_area.find(f"{_C}layout")
    at = list(plot_area).index(layout) + 1 if layout is not None else 0
    plot_area.insert(at, new_plot)
    if keep_axes:
        for ax_ref, old_id in zip(new_plot.findall(f"{_C}axId"), old_axis_ids):
            ax_ref.set("val", old_id)
    else:
        for i, axis in enumerate(ET.fromstring(
            f'<w xmlns:c="{NS_C}" xmlns:a="{NS_A}">{_axes(spec)}</w>'
        )):
            plot_area.insert(at + 1 + i, axis)

    if spec.three_d:
        if chart.find(f"{_C}view3D") is None:
            view = ET.fromstring(
                f'<c:view3D xmlns:c="{NS_C}"><c:rotX val="15"/><c:rotY val="20"/>'
                '<c:rAngAx val="1"/></c:view3D>'
            )
            _insert_ordered(chart, view, _CHART_ORDER)
    else:
        for name in ("view3D", "floor", "sideWall", "backWall"):
            _remove(chart, name)
    return dump(root)


# -- title and legend ----------------------------------------------------------

def format_chart(
    chart_xml_bytes: bytes,
    title: Optional[str] = None,
    has_legend: Optional[bool] = None,
    legend_position: Optional[str] = None,
) -> bytes:
    """Title and legend, the way ``Chart.HasTitle``, ``HasLegend`` and
    ``Legend.Position`` set them. Raises ``ValueError`` with the Windows text
    for a legend position on a chart with no legend."""
    root = load(chart_xml_bytes)
    chart = root.find(f"{_C}chart")
    if chart is None:
        raise PackageError("chart1.xml has no c:chart")

    if title is not None:
        _set_title(chart, title, _CHART_ORDER)
        _insert_ordered(chart, _val("autoTitleDeleted", 0), _CHART_ORDER)

    if has_legend is not None:
        if has_legend:
            if chart.find(f"{_C}legend") is None:
                legend = _el("legend")
                legend.append(_val("legendPos", "r"))
                legend.append(_val("overlay", 0))
                _insert_ordered(chart, legend, _CHART_ORDER)
        else:
            _remove(chart, "legend")

    if legend_position is not None:
        legend = chart.find(f"{_C}legend")
        if legend is None:
            raise ValueError(
                "Cannot set legend position when chart has no legend. "
                "Set has_legend=true first."
            )
        key = legend_position.strip().lower()
        if key not in LEGEND_POS:
            raise ValueError(
                f"Unknown legend position '{legend_position}'. "
                f"PowerPoint presets: {', '.join(LEGEND_POS)}."
            )
        _insert_ordered(legend, _val("legendPos", LEGEND_POS[key]), _LEGEND_ORDER)
        # A legend that was dragged by hand keeps a manual layout, which
        # overrides the preset; PowerPoint drops it too when a preset is chosen.
        _remove(legend, "layout")
    return dump(root)


def chart_summary(chart_xml_bytes: bytes) -> dict:
    """What ``ppt_format_chart`` reports: has_title, has_legend, and the
    title text and legend position for a read-back check."""
    root = _parse(chart_xml_bytes)
    chart = root.find(f"{_C}chart")
    if chart is None:
        raise PackageError("chart1.xml has no c:chart")
    title = chart.find(f"{_C}title")
    legend = chart.find(f"{_C}legend")
    pos = legend.find(f"{_C}legendPos") if legend is not None else None
    return {
        "has_title": title is not None,
        "title": _title_text(title),
        "has_legend": legend is not None,
        "legend_position": pos.get("val") if pos is not None else None,
    }


# -- axes ----------------------------------------------------------------------

def _axis_for(root: ET.Element, axis_key: str) -> ET.Element:
    """The axis element Windows's ``Chart.Axes(type, group)`` would return.

    ``category`` is the category axis, or on a scatter the first value axis;
    ``value`` the value axis the first plot crosses it with;
    ``secondary_value`` a value axis no plot in the first group uses;
    ``series`` the series axis of a 3D chart.
    """
    plot_area = _plot_area(root)
    plots = _plots(plot_area)
    axes = [a for a in plot_area if _local(a.tag) in AXIS_TAGS]
    by_id = {a.find(f"{_C}axId").get("val"): a for a in axes if a.find(f"{_C}axId") is not None}
    first_ids = [a.get("val") for a in plots[0].findall(f"{_C}axId")] if plots else []

    if axis_key == "series":
        found = [a for a in axes if _local(a.tag) == "serAx"]
        return found[0] if found else None
    category = [a for a in axes if _local(a.tag) in ("catAx", "dateAx")]
    category_id = (
        category[0].find(f"{_C}axId").get("val") if category
        else (first_ids[0] if first_ids else None)
    )
    if axis_key == "category":
        return by_id.get(category_id) if category_id is not None else None
    if axis_key == "value":
        for axis_id in first_ids:
            axis = by_id.get(axis_id)
            if axis is not None and _local(axis.tag) == "valAx" and axis_id != category_id:
                return axis
        return None
    if axis_key == "secondary_value":
        for axis in axes:
            axis_id = axis.find(f"{_C}axId").get("val")
            if _local(axis.tag) == "valAx" and axis_id not in first_ids:
                return axis
        return None
    return None


def format_axis(
    chart_xml_bytes: bytes,
    axis_key: str,
    title: Optional[str] = None,
    min_scale: Optional[float] = None,
    max_scale: Optional[float] = None,
    major_unit: Optional[float] = None,
    minor_unit: Optional[float] = None,
    tick_label_spacing: Optional[int] = None,
    tick_mark_spacing: Optional[int] = None,
    major_tick_mark: Optional[str] = None,
    minor_tick_mark: Optional[str] = None,
    reverse_order: Optional[bool] = None,
    log_scale: Optional[bool] = None,
    log_base: Optional[float] = None,
    number_format: Optional[str] = None,
):
    """One axis, the fields ``ppt_format_chart_axis`` maps onto the XML.

    Returns ``(bytes, applied)``; ``applied`` lists the argument names in
    the order Windows applies them. Raises ``ValueError`` with the Windows
    text when the axis is not on this chart.
    """
    root = load(chart_xml_bytes)
    axis = _axis_for(root, axis_key)
    if axis is None:
        raise ValueError(
            f"Axis '{axis_key}' is not available on this chart type "
            "(e.g. pie/doughnut charts have no axes). Underlying error: "
            "no such axis element in chart1.xml"
        )
    applied: List[str] = []
    scaling = axis.find(f"{_C}scaling")
    if scaling is None:
        scaling = _insert_ordered(axis, _el("scaling"), _AXIS_ORDER)

    if title is not None:
        _set_title(axis, title, _AXIS_ORDER)
        applied.append("title")
    if min_scale is not None:
        _insert_ordered(scaling, _val("min", _num(min_scale)), _SCALING_ORDER)
        applied.append("min_scale")
    if max_scale is not None:
        _insert_ordered(scaling, _val("max", _num(max_scale)), _SCALING_ORDER)
        applied.append("max_scale")
    if major_unit is not None:
        _insert_ordered(axis, _val("majorUnit", _num(major_unit)), _AXIS_ORDER)
        applied.append("major_unit")
    if minor_unit is not None:
        _insert_ordered(axis, _val("minorUnit", _num(minor_unit)), _AXIS_ORDER)
        applied.append("minor_unit")
    if tick_label_spacing is not None:
        _insert_ordered(axis, _val("tickLblSkip", int(tick_label_spacing)), _AXIS_ORDER)
        applied.append("tick_label_spacing")
    if tick_mark_spacing is not None:
        _insert_ordered(axis, _val("tickMarkSkip", int(tick_mark_spacing)), _AXIS_ORDER)
        applied.append("tick_mark_spacing")
    if major_tick_mark is not None:
        _insert_ordered(axis, _val("majorTickMark", TICK_MARK[major_tick_mark]), _AXIS_ORDER)
        applied.append("major_tick_mark")
    if minor_tick_mark is not None:
        _insert_ordered(axis, _val("minorTickMark", TICK_MARK[minor_tick_mark]), _AXIS_ORDER)
        applied.append("minor_tick_mark")
    if reverse_order is not None:
        _insert_ordered(scaling, _val("orientation", "maxMin" if reverse_order else "minMax"), _SCALING_ORDER)
        applied.append("reverse_order")
    if log_scale is not None:
        if log_scale:
            _insert_ordered(scaling, _val("logBase", _num(log_base or 10)), _SCALING_ORDER)
        else:
            _remove(scaling, "logBase")
        applied.append("log_scale")
    if log_base is not None:
        _insert_ordered(scaling, _val("logBase", _num(log_base)), _SCALING_ORDER)
        applied.append("log_base")
    if number_format is not None:
        fmt = _el("numFmt", formatCode=number_format, sourceLinked="0")
        _insert_ordered(axis, fmt, _AXIS_ORDER)
        applied.append("number_format")
    return dump(root), applied


def read_axis(chart_xml_bytes: bytes, axis_key: str) -> Optional[dict]:
    """The same fields back out of an axis, for the read-back check."""
    root = _parse(chart_xml_bytes)
    axis = _axis_for(root, axis_key)
    if axis is None:
        return None

    def val(parent, name):
        e = parent.find(f"{_C}{name}") if parent is not None else None
        return None if e is None else e.get("val")

    def num(parent, name):
        v = val(parent, name)
        return None if v is None else float(v)

    scaling = axis.find(f"{_C}scaling")
    fmt = axis.find(f"{_C}numFmt")
    orientation = val(scaling, "orientation")
    tick_back = {v: k for k, v in TICK_MARK.items()}
    return {
        "title": _title_text(axis.find(f"{_C}title")),
        "min_scale": num(scaling, "min"),
        "max_scale": num(scaling, "max"),
        "major_unit": num(axis, "majorUnit"),
        "minor_unit": num(axis, "minorUnit"),
        "tick_label_spacing": None if val(axis, "tickLblSkip") is None else int(val(axis, "tickLblSkip")),
        "tick_mark_spacing": None if val(axis, "tickMarkSkip") is None else int(val(axis, "tickMarkSkip")),
        "major_tick_mark": tick_back.get(val(axis, "majorTickMark")),
        "minor_tick_mark": tick_back.get(val(axis, "minorTickMark")),
        "reverse_order": None if orientation is None else orientation == "maxMin",
        "log_scale": val(scaling, "logBase") is not None,
        "log_base": num(scaling, "logBase"),
        "number_format": None if fmt is None else fmt.get("formatCode"),
    }


# -- one series ------------------------------------------------------------------

def _sp_pr(ser: ET.Element) -> ET.Element:
    sp = ser.find(f"{_C}spPr")
    if sp is None:
        sp = _insert_ordered(ser, _el("spPr"), _SER_ORDER)
    return sp


def _solid_fill(color_hex: str) -> ET.Element:
    fill = ET.Element(f"{{{NS_A}}}solidFill")
    ET.SubElement(fill, f"{{{NS_A}}}srgbClr").set("val", color_hex)
    return fill


_SPPR_ORDER = ["xfrm", "custGeom", "prstGeom", "noFill", "solidFill", "gradFill",
               "blipFill", "pattFill", "grpFill", "ln", "effectLst", "effectDag",
               "scene3d", "sp3d", "extLst"]


_LN_ORDER = ["noFill", "solidFill", "gradFill", "pattFill", "prstDash", "custDash",
             "round", "bevel", "miter", "headEnd", "tailEnd", "extLst"]


def _insert_a(parent: ET.Element, child: ET.Element, order: List[str]) -> ET.Element:
    """``_insert_ordered`` for DrawingML children, whose fills are a choice."""
    name = _local(child.tag)
    fills = {"noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"}
    if name in fills:
        for c in list(parent):
            if _local(c.tag) in fills:
                parent.remove(c)
    return _insert_ordered(parent, child, order)


def set_series(
    chart_xml_bytes: bytes,
    series_index: int,
    color: Optional[str] = None,
    show_data_labels: Optional[bool] = None,
    line_weight: Optional[float] = None,
) -> bytes:
    """Colour, data labels and line weight of one series, 1-based.

    ``color`` is ``#RRGGBB``. On a line, scatter or radar series the colour
    goes on the line, which is the part of such a series that shows; on
    every other kind on the fill.
    """
    root = load(chart_xml_bytes)
    sers = _all_series(_plot_area(root))
    if series_index < 1 or series_index > len(sers):
        raise ValueError(
            f"series_index {series_index} out of range (chart has {len(sers)} series)."
        )
    ser = sers[series_index - 1]
    kind = next(
        _local(plot.tag) for plot in _plots(_plot_area(root)) if ser in list(plot)
    )
    liney = kind in ("lineChart", "line3DChart", "scatterChart", "radarChart", "stockChart")

    if color is not None:
        hex_part = color.lstrip("#").upper()
        if len(hex_part) != 6 or any(c not in "0123456789ABCDEF" for c in hex_part):
            raise ValueError(f"color must be '#RRGGBB', got '{color}'")
        sp = _sp_pr(ser)
        if liney:
            # PowerPoint keeps only the line's colour on a line series and
            # drops a fill written beside it (measured on a paste), so the
            # colour goes on the line alone.
            ln = sp.find(f"{{{NS_A}}}ln")
            if ln is None:
                ln = _insert_a(sp, ET.Element(f"{{{NS_A}}}ln"), _SPPR_ORDER)
            _insert_a(ln, _solid_fill(hex_part), _LN_ORDER)
        else:
            _insert_a(sp, _solid_fill(hex_part), _SPPR_ORDER)
    if line_weight is not None:
        sp = _sp_pr(ser)
        ln = sp.find(f"{{{NS_A}}}ln")
        if ln is None:
            ln = _insert_a(sp, ET.Element(f"{{{NS_A}}}ln"), _SPPR_ORDER)
        ln.set("w", str(int(round(float(line_weight) * 12700))))
    if show_data_labels is not None:
        d = _el("dLbls")
        for name, on in (("showLegendKey", False), ("showVal", show_data_labels),
                         ("showCatName", False), ("showSerName", False),
                         ("showPercent", False), ("showBubbleSize", False)):
            d.append(_val(name, 1 if on else 0))
        _insert_ordered(ser, d, _SER_ORDER)
    return dump(root)


def read_series(chart_xml_bytes: bytes, series_index: int) -> Optional[dict]:
    """Colour, data labels and line weight of one series, for the read-back."""
    root = _parse(chart_xml_bytes)
    sers = _all_series(_plot_area(root))
    if series_index < 1 or series_index > len(sers):
        return None
    ser = sers[series_index - 1]
    sp = ser.find(f"{_C}spPr")
    ln = sp.find(f"{{{NS_A}}}ln") if sp is not None else None
    clr = sp.find(f"{{{NS_A}}}solidFill/{{{NS_A}}}srgbClr") if sp is not None else None
    if clr is None and ln is not None:
        # A line series carries its colour on the line.
        clr = ln.find(f"{{{NS_A}}}solidFill/{{{NS_A}}}srgbClr")
    labels = ser.find(f"{_C}dLbls/{_C}showVal")
    return {
        "color": None if clr is None else "#" + clr.get("val", ""),
        "line_weight": None if ln is None or ln.get("w") is None else round(int(ln.get("w")) / 12700, 2),
        "show_data_labels": None if labels is None else labels.get("val") == "1",
    }


def series_count(chart_xml_bytes: bytes) -> int:
    return len(_all_series(_plot_area(_parse(chart_xml_bytes))))
