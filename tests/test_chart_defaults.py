"""A new chart written by gvml, with what AddChart2 gives it (#299).

On Windows a new chart is pasted rather than added with AddChart2, so that no
Excel starts. Its XML comes from ``ppt_com/chart_defaults.py``. Nothing here
reaches PowerPoint; the live checks are in the PR.
"""

import datetime
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from gvml import charts as gvml_charts  # noqa: E402
from ppt_com import chart_defaults  # noqa: E402

C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"

# Every type pasted rather than added with AddChart2. A type goes on only
# once it is written and its look has been checked against AddChart2's.
PASTED = {
    1, 76, 77, 57, 58, 59, 51, 52, 53, 4, 63, 64, 65, 66, 67, -4151, 81, 82,
    5, 69, -4120, 80, -4169, 74, 75, 72, 73, 15, 87,
    54, -4102, -4101,
}


def chart(chart_type, lang="en"):
    return ET.fromstring(chart_defaults.chart_xml(chart_type, lang).encode())


def plot(root):
    area = root.find(f"{C}chart/{C}plotArea")
    return next(c for c in area if c.tag.endswith("Chart"))


def test_the_types_pasted():
    assert chart_defaults.CHART_TYPES == PASTED


def test_the_rest_still_go_through_addchart2():
    for chart_type in (-4100, 70, 88, 83, 99, 60, 102):
        assert chart_type not in chart_defaults.CHART_TYPES
        with pytest.raises(ValueError, match="not written without Excel"):
            chart_defaults.chart_xml(chart_type)


def test_every_type_has_its_style():
    assert set(chart_defaults.STYLES) == PASTED


@pytest.mark.parametrize("language_id, lang", [(1041, "ja"), (1033, "en"), (1036, "en"), (None, "en")])
def test_the_sample_follows_office_s_language(language_id, lang):
    assert chart_defaults.language(language_id) == lang


@pytest.mark.parametrize("chart_type", sorted(PASTED))
def test_each_type_is_written_with_its_sample(chart_type):
    root = chart(chart_type)
    assert root.find(f"{C}chart/{C}title") is None, "one series shows its name, as AddChart2's do"
    assert root.find(f"{C}chart/{C}autoTitleDeleted").get("val") == "0"
    assert root.find(f"{C}chart/{C}legend") is not None
    data = gvml_charts.read_chart(chart_defaults.chart_xml(chart_type).encode())
    sample = chart_defaults.sample(chart_type, "en")
    assert [s["name"] for s in data["series"]] == [name for name, _ in sample["series"]]
    if "categories" in sample:
        assert data["series"][0]["values"] == [float(v) for v in sample["series"][0][1]]


@pytest.mark.parametrize("chart_type", [51, 57, 4, -4151, 1, -4169, 15])
def test_a_series_keeps_one_colour(chart_type):
    assert plot(chart(chart_type)).find(f"{C}varyColors").get("val") == "0"


@pytest.mark.parametrize("chart_type", [5, -4120, -4102])
def test_a_pie_colours_each_slice(chart_type):
    assert plot(chart(chart_type)).find(f"{C}varyColors").get("val") == "1"


def test_the_japanese_sample_is_addchart2_s():
    data = gvml_charts.read_chart(chart_defaults.chart_xml(51, "ja").encode())
    assert data["categories"] == ["カテゴリ 1", "カテゴリ 2", "カテゴリ 3", "カテゴリ 4"]
    assert [s["name"] for s in data["series"]] == ["系列 1", "系列 2", "系列 3"]
    pie = gvml_charts.read_chart(chart_defaults.chart_xml(5, "ja").encode())
    assert pie["categories"][0] == "第 1 四半期"
    assert pie["series"][0]["name"] == "売上高"


@pytest.mark.parametrize("chart_type", [1, 76, 77, -4151, 81, 82])
def test_area_and_radar_charts_get_dates(chart_type):
    sample = chart_defaults.sample(chart_type)
    assert sample["categories"][0] == datetime.date(2002, 1, 5)
    assert len(sample["categories"]) == 5 and len(sample["series"]) == 2
    cache = plot(chart(chart_type, "ja")).find(f"{C}ser/{C}cat/{C}numRef/{C}numCache")
    assert cache.find(f"{C}formatCode").text == "yyyy/mm/dd"
    # Excel's day number for 5 January 2002.
    assert cache.find(f"{C}pt/{C}v").text == "37261"
    axis = chart(chart_type, "ja").find(f"{C}chart/{C}plotArea/{C}catAx/{C}numFmt")
    assert (axis.get("formatCode"), axis.get("sourceLinked")) == ("yyyy/mm/dd", "1"), "shown as dates"


def test_a_radar_chart_has_its_legend_on_top():
    assert chart(-4151).find(f"{C}chart/{C}legend/{C}legendPos").get("val") == "t"
    assert chart(51).find(f"{C}chart/{C}legend/{C}legendPos").get("val") == "b"


@pytest.mark.parametrize("chart_type, smooth", [(72, "1"), (73, "1"), (74, "0"), (-4169, "0")])
def test_smooth_scatter_lines_are_smooth(chart_type, smooth):
    p = plot(chart(chart_type))
    assert p.find(f"{C}ser/{C}smooth").get("val") == smooth
    assert p.find(f"{C}scatterStyle").get("val") == ("smoothMarker" if smooth == "1" else "lineMarker")


@pytest.mark.parametrize("chart_type, markers", [(72, True), (73, False), (74, True), (75, False)])
def test_scatter_lines_without_markers_say_so(chart_type, markers):
    symbol = plot(chart(chart_type)).find(f"{C}ser/{C}marker/{C}symbol")
    assert (symbol is None) is markers


@pytest.mark.parametrize("chart_type, tag, grouping", [
    (59, "barChart", "percentStacked"), (64, "lineChart", "percentStacked"),
    (66, "lineChart", "stacked"), (67, "lineChart", "percentStacked"),
    (77, "areaChart", "percentStacked"),
])
def test_stacked_kinds(chart_type, tag, grouping):
    p = plot(chart(chart_type))
    assert p.tag == C + tag
    assert p.find(f"{C}grouping").get("val") == grouping


def test_the_other_new_kinds():
    assert plot(chart(81)).find(f"{C}radarStyle").get("val") == "marker"
    filled = plot(chart(82))
    assert filled.find(f"{C}radarStyle").get("val") == "filled"
    assert filled.find(f"{C}ser/{C}marker") is None
    assert plot(chart(80)).find(f"{C}ser/{C}explosion").get("val") == "25"
    assert plot(chart(87)).find(f"{C}ser/{C}bubble3D").get("val") == "1"
    assert plot(chart(15)).find(f"{C}ser/{C}bubble3D") is None


def test_scatter_and_bubble_points():
    for chart_type in (-4169, 15):
        data = gvml_charts.read_chart(chart_defaults.chart_xml(chart_type).encode())
        assert data["series"][0]["values"] == [2.7, 3.2, 0.8]
        assert data["categories"] == ["0.7", "1.8", "2.6"]


A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _one_more_series(xml: bytes) -> ET.Element:
    count = len(gvml_charts.load(xml).findall(f".//{C}ser"))
    series = [{"name": f"S{i}", "values": [1, 2]} for i in range(count + 1)]
    return gvml_charts.load(gvml_charts.set_data(xml, [1, 2], series))


def test_a_series_added_does_not_take_the_last_one_s_colour():
    root = gvml_charts.load(chart_defaults.chart_xml(-4169).encode())
    ser = root.findall(f".//{C}ser")[-1]
    fill = ET.Element(f"{A}solidFill")
    ET.SubElement(fill, f"{A}srgbClr", {"val": "FF0000"})
    ser.find(f"{C}spPr").insert(0, fill)

    first, added = _one_more_series(gvml_charts.dump(root)).findall(f".//{C}ser")

    assert first.find(f"{C}spPr/{A}solidFill/{A}srgbClr").get("val") == "FF0000"
    assert added.find(f"{C}spPr/{A}solidFill/{A}srgbClr") is None
    assert added.find(f"{C}spPr/{A}solidFill/{A}schemeClr").get("val") == "accent2"


@pytest.mark.parametrize("chart_type, path", [
    (-4169, f"{C}spPr/{A}ln/{A}noFill"),    # a scatter with no line
    (4, f"{C}marker/{C}symbol"),            # a line with no markers
    (75, f"{C}marker/{C}symbol"),           # a scatter line with no markers
])
def test_an_added_series_keeps_what_the_last_one_leaves_out(chart_type, path):
    xml = chart_defaults.chart_xml(chart_type).encode()
    assert gvml_charts.load(xml).find(f".//{C}ser/{path}") is not None

    assert _one_more_series(xml).findall(f".//{C}ser")[-1].find(path) is not None


@pytest.mark.parametrize("sizes", [[], [1.0]])
def test_a_bubble_series_needs_one_size_per_point(sizes):
    xml = chart_defaults.chart_xml(15).encode()
    with pytest.raises(ValueError, match="sizes"):
        gvml_charts.set_data(xml, [1, 2], [{"name": "S", "values": [1, 2], "sizes": sizes}])


def test_only_the_creation_id_is_stripped():
    from gvml.shapes import strip_creation_ids

    element = ET.fromstring(f"""<sp xmlns:a="{A[1:-1]}"><a:cNvPr id="1" name="C"><a:extLst>
        <a:ext uri="{{FF2B}}"><x:creationId xmlns:x="urn:a16"/></a:ext>
        <a:ext uri="{{C183}}"><y:decorative xmlns:y="urn:adec"/></a:ext>
        </a:extLst></a:cNvPr></sp>""")

    strip_creation_ids(element)

    assert [e.get("uri") for e in element.iter(f"{A}ext")] == ["{C183}"]
