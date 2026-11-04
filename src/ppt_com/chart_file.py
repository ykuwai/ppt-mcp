"""A new chart, written to a file without Excel (#299).

``Shapes.AddChart2`` starts an Excel for the chart's data workbook, and an
Excel started while another is still ending crashes, often taking the chart's
data or PowerPoint with it. An Excel takes from two seconds to a minute to end,
so charts added one after another kept overlapping. python-pptx writes a chart
and its data workbook straight into a pptx, with no Excel at all, and
``ppt_add_chart`` brings the chart in from there.

``replace_data`` swaps the data of the chart in such a file, keeping its
formatting, for ``ppt_set_chart_data``.

``write`` makes a one slide pptx holding the chart, with the sample data and
the look ``AddChart2`` gives a chart of that type, read off PowerPoint for
each type in ``CHART_TYPES``. The types python-pptx cannot write, the 3-D,
cone, cylinder, pyramid, stock and surface ones among them, are not in it and
are still added with ``AddChart2``.

python-pptx is imported only when a chart is written.
"""

from __future__ import annotations

import datetime

# XlChartType values python-pptx writes, which are its XL_CHART_TYPE values.
CATEGORY_TYPES = frozenset({
    1, 76, 77,              # area, stacked, 100% stacked
    57, 58, 59,             # bar, stacked, 100% stacked
    51, 52, 53,             # column, stacked, 100% stacked
    4, 63, 64,              # line, stacked, 100% stacked
    65, 66, 67,             # line with markers, stacked, 100% stacked
    -4151, 81, 82,          # radar, with markers, filled
})
PIE_TYPES = frozenset({5, 69, -4120, 80})        # pie, exploded, doughnut, exploded
XY_TYPES = frozenset({-4169, 74, 75, 72, 73})     # scatter and its line and smooth kinds
SMOOTH_TYPES = frozenset({72, 73})
BUBBLE_TYPES = frozenset({15, 87})                # bubble, with a 3-D effect
CHART_TYPES = CATEGORY_TYPES | PIE_TYPES | XY_TYPES | BUBBLE_TYPES

# AddChart2 gives these five dates and two series, not four categories and
# three; a radar chart also puts its legend at the top.
AREA_TYPES = frozenset({1, 76, 77})
RADAR_TYPES = frozenset({-4151, 81, 82})

# The chart style AddChart2 gives each type, its first in the gallery, as
# PowerPoint 16 reports it.
STYLES: dict[int, int] = {
    1: 276, 76: 276, 77: 276,
    57: 216, 58: 297, 59: 297,
    51: 201, 52: 297, 53: 297,
    4: 227, 63: 227, 64: 227,
    65: 332, 66: 332, 67: 332,
    -4151: 317, 81: 317, 82: 317,
    5: 251, 69: 251, -4120: 251, 80: 251,
    -4169: 240, 74: 240, 75: 240, 72: 240, 73: 240,
    15: 269, 87: 269,
}

# The sample data AddChart2 puts in a new chart, by the language of Office.
# Area and radar charts get dates, as AddChart2 gives them.
_DATES = [datetime.date(2002, 1, 5) + datetime.timedelta(days=i) for i in range(5)]
_CATEGORY_VALUES = [(4.3, 2.5, 3.5, 4.5), (2.4, 4.4, 1.8, 2.8), (2, 2, 3, 5)]
_AREA_VALUES = [(32, 32, 28, 12, 15), (12, 12, 12, 21, 28)]
_PIE_VALUES = (8.2, 3.2, 1.4, 1.2)
_XY_POINTS = [(0.7, 2.7), (1.8, 3.2), (2.6, 0.8)]
_BUBBLE_POINTS = [(0.7, 2.7, 10), (1.8, 3.2, 4), (2.6, 0.8, 8)]


def _samples(series, category, quarter, sales, y_values):
    return {
        "category": {
            "categories": [f"{category} {i}" for i in range(1, 5)],
            "series": [(f"{series} {i}", v) for i, v in enumerate(_CATEGORY_VALUES, 1)],
        },
        "area": {
            "categories": _DATES,
            "series": [(f"{series} {i}", v) for i, v in enumerate(_AREA_VALUES, 1)],
        },
        "pie": {
            "categories": [quarter(i) for i in range(1, 5)],
            "series": [(sales, _PIE_VALUES)],
        },
        "xy": {"series": [(y_values, _XY_POINTS)]},
        "bubble": {"series": [(y_values, _BUBBLE_POINTS)]},
    }


SAMPLES: dict[str, dict] = {
    "en": _samples("Series", "Category", lambda i: f"{i}{['st', 'nd', 'rd', 'th'][i - 1]} Qtr",
                   "Sales", "Y-Values"),
    "ja": _samples("系列", "カテゴリ", lambda i: f"第 {i} 四半期", "売上高", "Y の値"),
}

_LANGUAGES = {1041: "ja"}

# How the dates of an area or radar sample are shown, as Excel shows a date.
DATE_FORMATS = {"en": "m/d/yyyy", "ja": "yyyy/mm/dd"}


def available() -> bool:
    """Whether python-pptx is there to write a chart with."""
    try:
        import pptx  # noqa: F401
    except ImportError:
        return False
    return True


def language(language_id) -> str:
    """The sample language for an Office UI language id; English otherwise."""
    return _LANGUAGES.get(language_id, "en")


def sample(chart_type: int, lang: str = "en") -> dict:
    """The sample data a new chart of ``chart_type`` gets."""
    table = SAMPLES.get(lang, SAMPLES["en"])
    if chart_type in AREA_TYPES | RADAR_TYPES:
        return table["area"]
    if chart_type in PIE_TYPES:
        return table["pie"]
    if chart_type in XY_TYPES:
        return table["xy"]
    if chart_type in BUBBLE_TYPES:
        return table["bubble"]
    return table["category"]


def chart_data(chart_type: int, lang: str = "en"):
    """python-pptx chart data holding the sample for ``chart_type``."""
    from pptx.chart.data import BubbleChartData, CategoryChartData, XyChartData

    data = sample(chart_type, lang)
    if chart_type in XY_TYPES or chart_type in BUBBLE_TYPES:
        chart = BubbleChartData() if chart_type in BUBBLE_TYPES else XyChartData()
        for name, points in data["series"]:
            series = chart.add_series(name)
            for point in points:
                series.add_data_point(*point)
        return chart
    chart = CategoryChartData()
    chart.categories = data["categories"]
    if isinstance(data["categories"][0], datetime.date):
        chart.categories.number_format = DATE_FORMATS.get(lang, DATE_FORMATS["en"])
    for name, values in data["series"]:
        chart.add_series(name, values)
    return chart


def write(path: str, chart_type: int, slide_width: float, slide_height: float,
          left: float, top: float, width: float, height: float, lang: str = "en") -> None:
    """Write a one slide pptx at ``path`` holding a new chart of ``chart_type``.

    Sizes are in points. The slide is as large as the deck's, so the chart
    keeps its place when it is brought in.
    """
    from pptx import Presentation
    from pptx.enum.chart import XL_LEGEND_POSITION
    from pptx.util import Emu

    def emu(points: float) -> Emu:
        return Emu(int(round(points * 12700)))

    if chart_type not in CHART_TYPES:
        raise ValueError(f"Chart type {chart_type} is not written without Excel")
    deck = Presentation()
    deck.slide_width, deck.slide_height = emu(slide_width), emu(slide_height)
    slide = deck.slides.add_slide(deck.slide_layouts[6])  # blank
    frame = slide.shapes.add_chart(
        chart_type, emu(left), emu(top), emu(width), emu(height),
        chart_data(chart_type, lang),
    )
    chart = frame.chart
    if chart_type not in PIE_TYPES | BUBBLE_TYPES:
        # python-pptx colours every point of a lone series on its own. A
        # bubble chart has no such setting.
        chart.plots[0].vary_by_categories = False
    # No title, and not deleted either: a chart of one series then shows that
    # series' name above it, as one made by AddChart2 does. python-pptx leaves
    # the flag out of a scatter chart, which PowerPoint takes for deleted.
    chart._chartSpace.chart.get_or_add_autoTitleDeleted().val = False
    if chart_type in SMOOTH_TYPES:
        # python-pptx marks the chart smooth but writes each series straight.
        for series in chart.plots[0].series:
            series._element.get_or_add_smooth().val = True
    chart.has_legend = True
    chart.legend.position = (
        XL_LEGEND_POSITION.TOP if chart_type in RADAR_TYPES else XL_LEGEND_POSITION.BOTTOM)
    chart.legend.include_in_layout = False
    deck.save(path)


# The types whose data replace_data swaps: categories and series of values.
REPLACEABLE_TYPES = CATEGORY_TYPES | PIE_TYPES


def replace_data(path: str, categories: list, series: list[dict]) -> None:
    """Swap the data of the one chart in the pptx at ``path``.

    python-pptx rewrites the series and the embedded workbook and keeps the
    chart's formatting, its series colours among it.
    """
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData

    deck = Presentation(path)
    [frame] = [shape for shape in deck.slides[0].shapes if shape.has_chart]
    data = CategoryChartData()
    data.categories = categories
    for entry in series:
        data.add_series(entry["name"], entry["values"])
    frame.chart.replace_data(data)
    deck.save(path)
