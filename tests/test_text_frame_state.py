"""Tests for the text frame state ppt_get_shape_info reports.

Pure Python. The COM shape is a stand in that answers the same properties, so
this runs anywhere. It pins the two things worth pinning, that the words read
back are the words ppt_set_textframe accepts, and that a property PowerPoint
refuses to answer costs its own key rather than the whole report.
"""

import sys

sys.path.insert(0, "src")

import pytest

from ppt_com.constants import (
    msoAnchorBottom,
    msoAnchorMiddle,
    msoAnchorTop,
    msoFalse,
    msoTextOrientationHorizontal,
    msoTextOrientationVertical,
    msoTriStateMixed,
    msoTrue,
    ppAutoSizeNone,
    ppAutoSizeShapeToFitText,
    ppAutoSizeTextToFitShape,
)
from ppt_com.shapes import _text_frame_state
from ppt_com.text import (
    AUTO_SIZE_MAP,
    AUTO_SIZE_NAMES,
    ORIENTATION_MAP,
    ORIENTATION_NAMES,
    VERTICAL_ANCHOR_MAP,
    VERTICAL_ANCHOR_NAMES,
)


class Raises:
    """A property PowerPoint will not answer."""


class FakeTextFrame2:
    def __init__(self, auto_size):
        self._auto_size = auto_size

    @property
    def AutoSize(self):
        if self._auto_size is Raises:
            raise RuntimeError("no")
        return self._auto_size


class FakeTextFrame:
    def __init__(self, word_wrap, anchor, orientation, margins):
        self.WordWrap = word_wrap
        self.VerticalAnchor = anchor
        self.Orientation = orientation
        if margins is Raises:
            self._margins = None
        else:
            self._margins = margins

    def _margin(self, side):
        if self._margins is None:
            raise RuntimeError("no")
        return self._margins[side]

    @property
    def MarginLeft(self):
        return self._margin("left")

    @property
    def MarginRight(self):
        return self._margin("right")

    @property
    def MarginTop(self):
        return self._margin("top")

    @property
    def MarginBottom(self):
        return self._margin("bottom")


class FakeShape:
    def __init__(self, has_text_frame=True, auto_size=ppAutoSizeNone,
                 word_wrap=msoTrue, anchor=msoAnchorTop,
                 orientation=msoTextOrientationHorizontal,
                 margins=None):
        self.HasTextFrame = has_text_frame
        self.TextFrame2 = FakeTextFrame2(auto_size)
        self.TextFrame = FakeTextFrame(
            word_wrap, anchor, orientation,
            {"left": 7.2, "right": 7.2, "top": 3.6, "bottom": 3.6}
            if margins is None else margins,
        )


class TestTheWordsMatchWhatSetTextframeAccepts:
    """A caller reads a frame and writes it back. One vocabulary, not two."""

    def test_every_auto_size_the_writer_takes_can_be_read_back(self):
        for name, value in AUTO_SIZE_MAP.items():
            assert AUTO_SIZE_NAMES[value] == name

    def test_every_anchor_the_writer_takes_can_be_read_back(self):
        for name, value in VERTICAL_ANCHOR_MAP.items():
            assert VERTICAL_ANCHOR_NAMES[value] == name

    def test_every_orientation_the_writer_takes_can_be_read_back(self):
        for name, value in ORIENTATION_MAP.items():
            assert ORIENTATION_NAMES[value] == name

    def test_the_two_baseline_anchors_are_named_and_not_writable(self):
        assert set(VERTICAL_ANCHOR_NAMES.values()) - set(VERTICAL_ANCHOR_MAP) == {
            "top_baseline", "bottom_baseline",
        }


class TestReadingAFrame:
    def test_a_shape_with_no_text_frame_has_no_state_at_all(self):
        assert _text_frame_state(FakeShape(has_text_frame=False)) is None

    def test_a_default_box(self):
        assert _text_frame_state(FakeShape()) == {
            "autofit": "none",
            "word_wrap": True,
            "vertical_anchor": "top",
            "orientation": "horizontal",
            "margins": {"left": 7.2, "right": 7.2, "top": 3.6, "bottom": 3.6},
        }

    @pytest.mark.parametrize("value,word", [
        (ppAutoSizeNone, "none"),
        (ppAutoSizeShapeToFitText, "shape_to_fit"),
        (ppAutoSizeTextToFitShape, "shrink_to_fit"),
    ])
    def test_autofit_is_reported_by_name(self, value, word):
        assert _text_frame_state(FakeShape(auto_size=value))["autofit"] == word

    @pytest.mark.parametrize("value,word", [
        (msoAnchorTop, "top"),
        (msoAnchorMiddle, "middle"),
        (msoAnchorBottom, "bottom"),
    ])
    def test_the_anchor_is_reported_by_name(self, value, word):
        state = _text_frame_state(FakeShape(anchor=value))
        assert state["vertical_anchor"] == word

    def test_a_vertical_box_says_so(self):
        state = _text_frame_state(
            FakeShape(orientation=msoTextOrientationVertical)
        )
        assert state["orientation"] == "vertical"

    def test_word_wrap_off_is_false_and_not_none(self):
        assert _text_frame_state(FakeShape(word_wrap=msoFalse))["word_wrap"] is False

    def test_a_mixed_wrap_is_neither_true_nor_false(self):
        state = _text_frame_state(FakeShape(word_wrap=msoTriStateMixed))
        assert state["word_wrap"] is None

    def test_an_autofit_powerpoint_will_not_answer_costs_only_its_own_key(self):
        state = _text_frame_state(FakeShape(auto_size=Raises))
        assert state["autofit"] is None
        assert state["word_wrap"] is True
        assert state["margins"]["left"] == 7.2

    def test_margins_powerpoint_will_not_answer_cost_only_the_margins(self):
        state = _text_frame_state(FakeShape(margins=Raises))
        assert state["margins"] is None
        assert state["autofit"] == "none"

    def test_an_unknown_value_is_reported_as_nothing_rather_than_a_near_miss(self):
        state = _text_frame_state(FakeShape(auto_size=-2, anchor=99))
        assert state["autofit"] is None
        assert state["vertical_anchor"] is None
