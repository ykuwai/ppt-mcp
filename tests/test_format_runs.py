"""ppt_com.text.format_runs: bold and italic as Markdown that renders.

Emphasis that starts or ends with a space is not emphasis in CommonMark, so a
run's spaces at either end go outside its markers. Both platforms use it.
"""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from ppt_com.text import format_runs  # noqa: E402


def runs(*parts):
    return [{"text": t, "bold": "b" in f, "italic": "i" in f} for t, f in parts]


@pytest.mark.parametrize("parts, markdown", [
    ((("Costs ", ""), ("fell", "b"), (" sharply", "i")), "Costs **fell** *sharply*"),
    ((("Costs ", ""), ("fell ", "b"), ("sharply", "i")), "Costs **fell** *sharply*"),
    ((("A ", ""), (" both ", "bi"), ("end", "")), "A  ***both*** end"),
    ((("Plain", ""), ("   ", "b"), ("text", "")), "Plain   text"),
    ((("one ", "b"), ("two", "b")), "**one two**"),
    ((("Title", "b"),), "**Title**"),
    ((("no emphasis at all", ""),), "no emphasis at all"),
])
def test_spaces_stay_outside_the_markers(parts, markdown):
    assert format_runs(runs(*parts)) == markdown


def test_it_renders_as_meant():
    markdown_it = pytest.importorskip("markdown_it")
    html = markdown_it.MarkdownIt("commonmark").renderInline(
        format_runs(runs(("Costs ", ""), ("fell", "b"), (" sharply", "i"))))
    assert html == "Costs <strong>fell</strong> <em>sharply</em>"


def test_the_mac_reads_through_the_same_one():
    source = (Path(__file__).resolve().parent.parent / "src/ppt_mac/text.py").read_text(encoding="utf-8")
    body = source[source.index("def _format_runs("):]
    assert "format_runs(raw)" in body[:400]
