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
