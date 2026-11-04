"""A new chart written with python-pptx, without Excel (#299)."""

import datetime
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from ppt_com import chart_file  # noqa: E402

pptx = pytest.importorskip("pptx", reason="python-pptx is installed on Windows only")

# Every type that is added without Excel. A type comes off this list only when
# it is shown to go wrong; one goes on only once python-pptx writes it and its
# look has been checked against AddChart2's.
WRITTEN = {
    1, 76, 77, 57, 58, 59, 51, 52, 53, 4, 63, 64, 65, 66, 67, -4151, 81, 82,
    5, 69, -4120, 80, -4169, 74, 75, 72, 73, 15, 87,
}


def test_the_types_written_without_excel():
    assert chart_file.CHART_TYPES == WRITTEN


def test_three_d_and_the_rest_still_go_through_addchart2():
    for chart_type in (54, -4100, -4102, -4101, 70, 88, 83, 99):
        assert chart_type not in chart_file.CHART_TYPES


def test_every_type_has_its_style():
    assert set(chart_file.STYLES) == WRITTEN


@pytest.mark.parametrize("language_id, lang", [(1041, "ja"), (1033, "en"), (1036, "en"), (None, "en")])
def test_the_sample_follows_office_s_language(language_id, lang):
    assert chart_file.language(language_id) == lang


def written(tmp_path, chart_type, lang="en"):
    path = str(tmp_path / f"chart{chart_type}.pptx")
    chart_file.write(path, chart_type, 960, 540, 100, 50, 400, 300, lang)
    deck = pptx.Presentation(path)
    [shape] = deck.slides[0].shapes
    return deck, shape


@pytest.mark.parametrize("chart_type", sorted(WRITTEN))
def test_each_type_is_written_where_asked_with_its_sample(tmp_path, chart_type):
    deck, shape = written(tmp_path, chart_type)
    assert (deck.slide_width, deck.slide_height) == (960 * 12700, 540 * 12700)
    assert (shape.left, shape.top, shape.width, shape.height) == (
        100 * 12700, 50 * 12700, 400 * 12700, 300 * 12700)
    chart = shape.chart
    assert int(chart.chart_type) == chart_type
    assert not chart.has_title, "one series shows its name, as AddChart2's do"
    assert chart._chartSpace.chart.autoTitleDeleted.val is False
    assert chart.has_legend
    sample = chart_file.sample(chart_type, "en")
    assert [s.name for s in chart.plots[0].series] == [name for name, _ in sample["series"]]
    if "categories" in sample:
        assert list(chart.plots[0].series[0].values) == list(sample["series"][0][1])


@pytest.mark.parametrize("chart_type", [51, 57, 4, -4151])
def test_a_series_keeps_one_colour(tmp_path, chart_type):
    _, shape = written(tmp_path, chart_type)
    assert shape.chart.plots[0].vary_by_categories is False


def test_a_pie_colours_each_slice(tmp_path):
    _, shape = written(tmp_path, 5)
    assert shape.chart.plots[0].vary_by_categories is True


def test_the_japanese_sample_is_addchart2_s(tmp_path):
    _, shape = written(tmp_path, 51, "ja")
    plot = shape.chart.plots[0]
    assert list(plot.categories) == ["カテゴリ 1", "カテゴリ 2", "カテゴリ 3", "カテゴリ 4"]
    assert [s.name for s in plot.series] == ["系列 1", "系列 2", "系列 3"]
    _, pie = written(tmp_path, 5, "ja")
    assert list(pie.chart.plots[0].categories)[0] == "第 1 四半期"
    assert pie.chart.plots[0].series[0].name == "売上高"


@pytest.mark.parametrize("chart_type", [1, 76, 77, -4151, 81, 82])
def test_area_and_radar_charts_get_dates(tmp_path, chart_type):
    sample = chart_file.sample(chart_type)
    assert sample["categories"][0] == datetime.date(2002, 1, 5)
    assert len(sample["categories"]) == 5 and len(sample["series"]) == 2
    _, shape = written(tmp_path, chart_type, "ja")
    xml = shape.chart._chartSpace.xml
    assert 'formatCode="yyyy/mm/dd"' in xml or "yyyy/mm/dd" in xml


def test_a_radar_chart_has_its_legend_on_top(tmp_path):
    from pptx.enum.chart import XL_LEGEND_POSITION

    _, shape = written(tmp_path, -4151)
    assert shape.chart.legend.position == XL_LEGEND_POSITION.TOP
    _, column = written(tmp_path, 51)
    assert column.chart.legend.position == XL_LEGEND_POSITION.BOTTOM


@pytest.mark.parametrize("chart_type, smooth", [(72, True), (73, True), (74, False), (-4169, False)])
def test_smooth_scatter_lines_are_smooth(tmp_path, chart_type, smooth):
    _, shape = written(tmp_path, chart_type)
    element = shape.chart.plots[0].series[0]._element.smooth
    assert (element is not None and element.val) is smooth


def test_scatter_and_bubble_points(tmp_path):
    _, scatter = written(tmp_path, -4169)
    assert list(scatter.chart.plots[0].series[0].values) == [2.7, 3.2, 0.8]
    _, bubble = written(tmp_path, 15)
    assert list(bubble.chart.plots[0].series[0].values) == [2.7, 3.2, 0.8]


def test_a_type_not_written_is_refused(tmp_path):
    with pytest.raises(ValueError, match="not written without Excel"):
        chart_file.write(str(tmp_path / "x.pptx"), 54, 960, 540, 0, 0, 100, 100)


@pytest.mark.parametrize("chart_type", [51, 5, 4, 1, -4151])
def test_replace_data_swaps_the_data_and_keeps_the_rest(tmp_path, chart_type):
    from pptx.dml.color import RGBColor

    path = str(tmp_path / "chart.pptx")
    chart_file.write(path, chart_type, 960, 540, 100, 50, 400, 300, "en")
    deck = pptx.Presentation(path)
    chart = deck.slides[0].shapes[0].chart
    fill = chart.plots[0].series[0].format.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0xFF, 0, 0)
    deck.save(path)

    chart_file.replace_data(path, ["A", "B", "C"], [
        {"name": "North", "values": [1, 2, 3]}, {"name": "South", "values": [4, 5, 6.5]}])
    chart = pptx.Presentation(path).slides[0].shapes[0].chart
    plot = chart.plots[0]
    assert list(plot.categories) == ["A", "B", "C"]
    assert [(s.name, list(s.values)) for s in plot.series] == [
        ("North", [1, 2, 3]), ("South", [4, 5, 6.5])]
    assert plot.series[0].format.fill.fore_color.rgb == RGBColor(0xFF, 0, 0), "formatting is kept"
    assert int(chart.chart_type) == chart_type


def test_only_category_charts_have_their_data_swapped():
    assert chart_file.REPLACEABLE_TYPES == chart_file.CATEGORY_TYPES | chart_file.PIE_TYPES
    assert not chart_file.REPLACEABLE_TYPES & (chart_file.XY_TYPES | chart_file.BUBBLE_TYPES)
