"""Tests for formatting several spans of one shape in one call.

Pure Python. The span arithmetic is a function of the text, and the batch
itself runs against a stand in text frame that records what was written.
"""

import sys
from unittest.mock import patch

sys.path.insert(0, "src")

import pytest
from pydantic import ValidationError

from ppt_com.text import (
    FormatTextRangeInput,
    TextRangeSpec,
    _format_text_ranges_impl,
    _resolve_span,
)

SENTENCE = "吹奏楽部で「べろだして」に聞こえて爆笑してしまいました"


class TestFindingASpan:
    def test_start_and_length_pass_straight_through(self):
        assert _resolve_span(SENTENCE, "TextBox 19", 4, 6, None, 1) == (4, 6)

    def test_search_text_answers_a_one_based_start(self):
        # 「べろだして」 is the 6th character onwards, and 1-based start says 6
        # where Python's own find says 5.
        start, length = _resolve_span(SENTENCE, "TextBox 19", None, None, "「べろだして」", 1)
        assert SENTENCE.find("「べろだして」") == 5
        assert (start, length) == (6, 7)
        assert SENTENCE[start - 1:start - 1 + length] == "「べろだして」"

    def test_the_second_occurrence_is_the_second_one(self):
        text = "はいはい"
        assert _resolve_span(text, "T", None, None, "はい", 2) == (3, 2)

    def test_text_that_is_not_there_says_so_with_the_shape_name(self):
        with pytest.raises(ValueError, match="not found in shape 'TextBox 19'"):
            _resolve_span(SENTENCE, "TextBox 19", None, None, "ありません", 1)

    def test_asking_for_more_occurrences_than_exist_says_how_many_there_are(self):
        with pytest.raises(ValueError, match="only 1 occurrence"):
            _resolve_span("はい", "T", None, None, "はい", 2)


class Characters:
    def __init__(self, frame, start, length):
        self._frame = frame
        self.Start = start
        self.Length = length
        self.Font = object()

    @property
    def Text(self):
        return self._frame.Text[self.Start - 1:self.Start - 1 + self.Length]


class FakeTextRange:
    def __init__(self, text):
        self.Text = text
        self.written = []

    def Characters(self, Start, Length):
        return Characters(self, Start, Length)


class FakeShape:
    def __init__(self, name, text):
        self.Name = name
        self.HasTextFrame = True
        self.range = FakeTextRange(text)


def run(ranges, base=None, text=SENTENCE):
    """Call the batch against a stand in shape, and report what it wrote."""
    shape = FakeShape("TextBox 19", text)
    written = []

    def fake_apply(font, *props):
        written.append(("font", props))

    def fake_frame(slide_index, shape_name_or_index):
        return shape, shape.range

    with patch("ppt_com.text._text_frame_of", fake_frame), \
            patch("ppt_com.text._apply_font_props", fake_apply), \
            patch("ppt_com.text._apply_highlight") as highlight:
        result = _format_text_ranges_impl(1, "TextBox 19", base, ranges)
    return result, written, highlight


class TestFormattingSeveralSpans:
    def test_every_span_is_reported_with_the_text_it_landed_on(self):
        result, _, _ = run([
            {"search_text": "吹奏楽部", "font_size": 36},
            {"search_text": "「べろだして」", "font_size": 50, "color": "#FF0066"},
        ])
        assert result["count"] == 2
        assert [r["formatted_text"] for r in result["ranges"]] == [
            "吹奏楽部", "「べろだして」"]

    def test_the_spans_are_written_in_the_order_they_were_given(self):
        # Later entries win where they overlap, which only holds if the order
        # is the caller's.
        _, written, _ = run([
            {"search_text": "吹奏楽部", "font_size": 36},
            {"search_text": "楽部で", "font_size": 50},
        ])
        assert [props[2] for _, props in written] == [36, 50]

    def test_base_is_written_before_any_span(self):
        _, written, _ = run(
            [{"search_text": "吹奏楽部", "font_size": 36}],
            base={"font_size": 28},
        )
        assert [props[2] for _, props in written] == [28, 36]

    def test_base_covers_the_whole_frame(self):
        result, _, _ = run([{"search_text": "吹奏楽部", "font_size": 36}],
                           base={"font_size": 28})
        # The base is not reported as a range of its own.
        assert result["count"] == 1

    def test_no_base_writes_nothing_extra(self):
        _, written, _ = run([{"search_text": "吹奏楽部", "font_size": 36}])
        assert len(written) == 1

    def test_a_highlight_goes_through_for_the_span_that_asked_for_it(self):
        _, _, highlight = run([
            {"search_text": "吹奏楽部"},
            {"search_text": "爆笑", "highlight_color": "#FFFF00"},
        ])
        assert highlight.call_count == 1
        assert highlight.call_args[0][1] == "#FFFF00"

    def test_start_and_length_entries_work_beside_search_text_ones(self):
        result, _, _ = run([
            {"start": 1, "length": 4},
            {"search_text": "爆笑"},
        ])
        assert [r["start"] for r in result["ranges"]] == [1, 18]


class TestWhenAColourIsWrong:
    """A colour is only refused when it is written, so a bad value in the
    fourth entry used to leave the first three applied.
    """

    def test_a_malformed_colour_anywhere_fails_the_call(self):
        with pytest.raises(Exception):
            run([
                {"search_text": "吹奏楽部", "font_size": 36},
                {"search_text": "爆笑", "color": "magenta"},
            ])

    def test_and_nothing_was_written(self):
        shape = FakeShape("TextBox 19", SENTENCE)
        written = []
        with patch("ppt_com.text._text_frame_of", lambda *a: (shape, shape.range)),                 patch("ppt_com.text._apply_font_props",
                      lambda *a: written.append(a)),                 patch("ppt_com.text._apply_highlight"):
            with pytest.raises(Exception):
                _format_text_ranges_impl(1, "TextBox 19", None, [
                    {"search_text": "吹奏楽部", "font_size": 36},
                    {"search_text": "爆笑", "color": "magenta"},
                ])
        assert written == []

    def test_an_unknown_theme_colour_is_caught_the_same_way(self):
        with pytest.raises(Exception):
            run([{"search_text": "吹奏楽部", "font_color_theme": "accent99"}])


class TestWhenOneSpanCannotBeFound:
    """Every span is resolved before anything is written, so the shape is
    left alone rather than half restyled.
    """

    def test_the_call_fails(self):
        with pytest.raises(ValueError, match="not found in shape"):
            run([
                {"search_text": "吹奏楽部", "font_size": 36},
                {"search_text": "ありません", "font_size": 50},
            ])

    def test_and_nothing_at_all_was_written(self):
        shape = FakeShape("TextBox 19", SENTENCE)
        written = []
        with patch("ppt_com.text._text_frame_of", lambda *a: (shape, shape.range)), \
                patch("ppt_com.text._apply_font_props",
                      lambda *a: written.append(a)), \
                patch("ppt_com.text._apply_highlight"):
            with pytest.raises(ValueError):
                _format_text_ranges_impl(1, "TextBox 19", {"font_size": 28}, [
                    {"search_text": "吹奏楽部", "font_size": 36},
                    {"search_text": "ありません", "font_size": 50},
                ])
        # Not even the base, which would otherwise be written first.
        assert written == []


def rejected(**kwargs):
    with pytest.raises(ValidationError) as caught:
        FormatTextRangeInput(slide_index=1, shape_name_or_index="T", **kwargs)
    return str(caught.value)


class TestWhatTheToolAccepts:
    def test_one_span_the_old_way_still_works(self):
        params = FormatTextRangeInput(
            slide_index=1, shape_name_or_index="T", search_text="a", font_size=36)
        assert params.ranges is None

    def test_a_batch_with_a_base(self):
        params = FormatTextRangeInput(
            slide_index=1, shape_name_or_index="T",
            base={"font_size": 28},
            ranges=[{"search_text": "a", "font_size": 50}],
        )
        assert params.base.font_size == 28
        assert params.ranges[0].font_size == 50

    def test_a_span_beside_ranges_is_refused(self):
        assert "cannot be used with ranges" in rejected(
            ranges=[{"search_text": "a"}], search_text="b")

    def test_and_so_is_formatting_beside_ranges(self):
        # It would be ambiguous whether it meant the frame or every span.
        assert "cannot be used with ranges" in rejected(
            ranges=[{"search_text": "a"}], font_size=28)

    def test_an_empty_ranges_list_is_refused(self):
        assert "must not be empty" in rejected(ranges=[])

    def test_base_without_ranges_is_refused(self):
        assert "only means something with ranges" in rejected(
            base={"font_size": 28})

    def test_saying_no_span_at_all_is_still_refused(self):
        assert "Either search_text or both start and length" in rejected(font_size=28)


class TestWhatARangeEntryAccepts:
    def test_search_text_and_start_together_are_refused(self):
        with pytest.raises(ValidationError, match="mutually exclusive"):
            TextRangeSpec(search_text="a", start=1, length=1)

    def test_a_start_without_a_length_is_refused(self):
        with pytest.raises(ValidationError, match="Either search_text or both"):
            TextRangeSpec(start=1)

    def test_an_occurrence_without_search_text_is_refused(self):
        with pytest.raises(ValidationError, match="only valid with search_text"):
            TextRangeSpec(start=1, length=2, occurrence=2)

    def test_an_empty_search_text_is_refused(self):
        with pytest.raises(ValidationError, match="must not be empty"):
            TextRangeSpec(search_text="")

    def test_a_bad_highlight_colour_is_refused(self):
        with pytest.raises(ValidationError, match="RRGGBB"):
            TextRangeSpec(search_text="a", highlight_color="yellow")
