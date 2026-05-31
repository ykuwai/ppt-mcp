"""Tests for slide layout/design resolution helpers (issue #161).

Pure Python tests — a fake COM presentation mimics the Designs / SlideMaster /
CustomLayouts object graph so the name-matching logic can be verified without
PowerPoint.
"""

import sys

sys.path.insert(0, "src")

from ppt_com.slides import AddSlideInput, _find_layout_matches


# --- Fake COM object graph -------------------------------------------------

class _FakeLayout:
    def __init__(self, name):
        self.Name = name


class _FakeCollection:
    """1-based collection callable like a COM collection: coll(i)."""

    def __init__(self, items):
        self._items = items

    @property
    def Count(self):
        return len(self._items)

    def __call__(self, i):
        return self._items[i - 1]


class _FakeMaster:
    def __init__(self, layout_names):
        self.CustomLayouts = _FakeCollection([_FakeLayout(n) for n in layout_names])


class _FakeDesign:
    def __init__(self, name, layout_names):
        self.Name = name
        self.SlideMaster = _FakeMaster(layout_names)


class _FakePres:
    def __init__(self, designs):
        # designs: list of (design_name, [layout_names])
        self.Designs = _FakeCollection(
            [_FakeDesign(n, ls) for n, ls in designs]
        )


# --- AddSlideInput model ---------------------------------------------------

def test_like_slide_index_accepted():
    m = AddSlideInput(like_slide_index=3)
    assert m.like_slide_index == 3


def test_like_slide_index_defaults_none():
    m = AddSlideInput()
    assert m.like_slide_index is None


def test_like_slide_index_zero_accepted_by_model():
    # The model intentionally has no ge=1 constraint — the 1-based bounds
    # check is enforced at runtime in _add_slide_impl, not at the model level.
    m = AddSlideInput(like_slide_index=0)
    assert m.like_slide_index == 0


# --- _find_layout_matches --------------------------------------------------

def _pres():
    return _FakePres([
        ("Design A", ["Title", "Title Only", "Content"]),
        ("Design B", ["Title Only", "Two Content"]),
        ("Design C", ["Section Header"]),
    ])


def test_single_match_returns_one():
    matches = _find_layout_matches(_pres(), "Content", None)
    assert [(m.design_index, m.design_name) for m in matches] == [(1, "Design A")]


def test_ambiguous_match_across_designs():
    matches = _find_layout_matches(_pres(), "Title Only", None)
    # Present in Design A (1) and Design B (2)
    assert [(m.design_index, m.design_name) for m in matches] == [
        (1, "Design A"), (2, "Design B")
    ]
    assert len(matches) > 1  # caller flags this as ambiguous


def test_no_match_returns_empty():
    assert _find_layout_matches(_pres(), "Nonexistent", None) == []


def test_design_index_restricts_search():
    # "Title Only" exists in both A and B, but searching only design 2 -> B
    matches = _find_layout_matches(_pres(), "Title Only", 2)
    assert [(m.design_index, m.design_name) for m in matches] == [(2, "Design B")]


def test_design_index_no_match_in_that_design():
    # "Content" only exists in design 1; restricting to design 3 -> none
    assert _find_layout_matches(_pres(), "Content", 3) == []


def test_at_most_one_match_per_design():
    # A design with the same layout name twice should still yield one match.
    pres = _FakePres([("Dup", ["X", "X", "Y"])])
    matches = _find_layout_matches(pres, "X", None)
    assert len(matches) == 1
    assert matches[0].design_name == "Dup"
