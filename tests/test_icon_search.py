"""ppt_search_icons matches words, and ranks by how many match (#194).

"air conditioner heating" returned six aeroplanes in twelve, `airlines` and
`airplane_ticket` first and `heat_pump` last: "air" was found inside
`airlines`, and popularity decided the rest. A query word now matches a whole
word of an icon's name, tags or categories, or a word it begins, and an icon
matching more of the query's words comes first.

The icons here stand in for Google's list, with the tags the real ones carry
cut down to what matters; nothing is fetched.
"""

import sys
import time

import pytest

sys.path.insert(0, "src")

from ppt_com import advanced_ops  # noqa: E402

ICONS = [
    {"name": "airlines", "tags": ["airplane", "airport", "flight", "transportation"],
     "categories": ["travel"], "popularity": 9000},
    {"name": "airplane_ticket", "tags": ["airplane", "flight", "plane", "ticket"],
     "categories": ["travel"], "popularity": 8000},
    {"name": "chair", "tags": ["couch", "furniture", "seat"],
     "categories": ["household"], "popularity": 7000},
    {"name": "ac_unit", "tags": ["air", "cold", "conditioner", "freeze", "snowflake"],
     "categories": ["household"], "popularity": 6000},
    {"name": "mode_heat", "tags": ["fire", "heat", "temperature"],
     "categories": ["household"], "popularity": 3000},
    {"name": "heat_pump", "tags": ["air", "conditioner", "cool", "energy", "heat", "heating"],
     "categories": ["household"], "popularity": 500},
]


@pytest.fixture(autouse=True)
def _icons(monkeypatch):
    monkeypatch.setattr(advanced_ops, "_icon_cache", ICONS)
    monkeypatch.setattr(advanced_ops, "_icon_cache_time", time.time())
    monkeypatch.setattr(advanced_ops, "_icon_is_on_cdn", lambda name: True)


def names(query, max_results=20):
    return [icon["name"] for icon in advanced_ops._search_icons(query, max_results)]


def test_the_reported_search_puts_the_icons_matching_most_words_first():
    # heat_pump matches all three words, ac_unit two, and the aeroplanes
    # only begin with "air".
    assert names("air conditioner heating")[:2] == ["heat_pump", "ac_unit"]


def test_a_word_is_not_found_inside_another():
    assert "chair" not in names("air")


def test_a_whole_word_comes_before_one_it_begins():
    found = names("air")
    assert found.index("ac_unit") < found.index("airlines")


def test_a_word_that_begins_another_still_matches():
    assert "heat_pump" in names("pump")
    assert "airplane_ticket" in names("tick")
    assert "mode_heat" in names("heating")


def test_two_letters_begin_nothing():
    assert names("ai") == []


def test_more_words_matched_outrank_popularity():
    found = names("ticket flight")
    assert found[0] == "airplane_ticket"


def test_an_exact_name_comes_first():
    assert names("heat pump")[0] == "heat_pump"
    assert names("heat_pump")[0] == "heat_pump"


def test_max_results_still_applies():
    assert len(names("household", 2)) == 2
