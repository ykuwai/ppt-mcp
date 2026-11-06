"""What a new chart gets on Windows, so it looks like one AddChart2 makes (#299).

``Shapes.AddChart2`` starts an Excel for the chart's data workbook, and an
Excel started while another is still ending crashes, often taking the chart's
data or PowerPoint with it. So a new chart is not made with ``AddChart2``:
``chart_xml`` writes it with ``gvml/charts.py``, the writer the Mac uses, and
``ppt_add_chart`` pastes it (``Art::GVML ClipFormat``). No Excel starts. The
chart has no workbook; PowerPoint builds one from the chart's numbers the
first time its data is opened.

The chart gets the sample data, legend and style ``AddChart2`` gives a chart
of its type, read off PowerPoint for each type. Types gvml has no template
for (cone, cylinder, pyramid, stock, surface and the other 3-D kinds) are
still added with ``AddChart2``.
"""

from __future__ import annotations

import datetime

from gvml import charts as _gvml

# XlChartType values, grouped by the sample data AddChart2 gives them.
CATEGORY_TYPES = frozenset({
    1, 76, 77,              # area, stacked, 100% stacked
    57, 58, 59,             # bar, stacked, 100% stacked
    51, 52, 53,             # column, stacked, 100% stacked
    4, 63, 64,              # line, stacked, 100% stacked
    65, 66, 67,             # line with markers, stacked, 100% stacked
    -4151, 81, 82,          # radar, with markers, filled
})
PIE_TYPES = frozenset({5, 69, -4120, 80})        # pie, exploded, doughnut, exploded
THREE_D_TYPES = frozenset({54, -4102, -4101})     # 3-D column, pie and line
XY_TYPES = frozenset({-4169, 74, 75, 72, 73})     # scatter and its line and smooth kinds
SMOOTH_TYPES = frozenset({72, 73})
BUBBLE_TYPES = frozenset({15, 87})                # bubble, with a 3-D effect
# Every type pasted rather than added with AddChart2: those gvml writes.
CHART_TYPES = frozenset(_gvml.supported_types())

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
    54: 286, -4102: 262, -4101: 307,
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


def language(language_id) -> str:
    """The sample language for an Office UI language id; English otherwise."""
    return _LANGUAGES.get(language_id, "en")


def sample(chart_type: int, lang: str = "en") -> dict:
    """The sample data a new chart of ``chart_type`` gets."""
    table = SAMPLES.get(lang, SAMPLES["en"])
    if chart_type in AREA_TYPES | RADAR_TYPES:
        return table["area"]
    if chart_type in PIE_TYPES or chart_type == -4102:
        return table["pie"]
    if chart_type in XY_TYPES:
        return table["xy"]
    if chart_type in BUBBLE_TYPES:
        return table["bubble"]
    return table["category"]


def chart_xml(chart_type: int, lang: str = "en") -> str:
    """``chart1.xml`` for a new chart of ``chart_type``, as AddChart2 makes it:
    its sample data in ``lang``, dates on an area or radar chart, the legend
    at the bottom, or at the top on a radar chart. The chart style is set
    once it is on the slide (``STYLES``)."""
    if chart_type not in CHART_TYPES:
        raise ValueError(f"Chart type {chart_type} is not written without Excel")
    data = sample(chart_type, lang)
    if chart_type in XY_TYPES or chart_type in BUBBLE_TYPES:
        name, points = data["series"][0]
        categories = [p[0] for p in points]
        series = [{"name": name, "values": [p[1] for p in points]}]
        if chart_type in BUBBLE_TYPES:
            series[0]["sizes"] = [p[2] for p in points]
    else:
        categories = list(data["categories"])
        series = [{"name": name, "values": list(values)} for name, values in data["series"]]
    dates = isinstance(categories[0], datetime.date)
    return _gvml.chart_xml(
        chart_type, categories, series,
        legend_pos="t" if chart_type in RADAR_TYPES else "b",
        date_format=DATE_FORMATS.get(lang, DATE_FORMATS["en"]) if dates else None,
    )
