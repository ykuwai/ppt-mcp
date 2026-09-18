"""Tests for editing text without losing the formatting it already has.

Pure Python. run_offsets is the arithmetic that decides where each run lands,
and the impl runs against a stand in text frame that records what was written.
"""

import sys
from unittest.mock import patch

sys.path.insert(0, "src")

import pytest
from pydantic import ValidationError

from ppt_com.text import RunSpec, SetTextInput, _set_text_impl, run_offsets

CR = chr(13)
LF = chr(10)
VT = chr(11)


class TestWhereEachRunLands:
    def test_runs_follow_one_another(self):
        assert run_offsets([{"text": "abc"}, {"text": "de"}]) == [(1, 3), (4, 2)]

    def test_a_single_run_starts_at_one(self):
        assert run_offsets([{"text": "abc"}]) == [(1, 3)]

    def test_no_runs_is_no_spans(self):
        assert run_offsets([]) == []

    def test_a_paragraph_break_counts_as_the_one_character_it_becomes(self):
        # The caller types \\n, PowerPoint stores \\r, and it is one character
        # either way. Counting the caller's two would shift every later run.
        assert run_offsets([{"text": "a" + LF + "b"}, {"text": "c"}]) == [(1, 3), (4, 1)]

    def test_a_line_break_counts_as_one_character_too(self):
        assert run_offsets([{"text": "a" + VT + "b"}, {"text": "c"}]) == [(1, 3), (4, 1)]

    def test_an_empty_run_takes_no_room_and_does_not_move_the_next_one(self):
        assert run_offsets([{"text": "ab"}, {"text": ""}, {"text": "c"}]) == [
            (1, 2), (3, 0), (3, 1)]


class FakeRun:
    def __init__(self, text, start):
        self.Text = text
        self.Start = start
        self.Length = len(text)


class Characters:
    def __init__(self, frame, start, length):
        self._frame = frame
        self.Start = start
        self.Length = length
        self.Font = object()

    @property
    def Text(self):
        return self._frame.Text[self.Start - 1:self.Start - 1 + self.Length]

    @Text.setter
    def Text(self, value):
        whole = self._frame.Text
        # Straight at the backing text, so this does not look like a write of
        # the whole frame, which is the thing the span path must not do.
        self._frame._text = (whole[:self.Start - 1] + value
                             + whole[self.Start - 1 + self.Length:])
        self._frame.writes.append(("span", self.Start, self.Length, value))


class FakeTextRange:
    def __init__(self, text=""):
        self._text = text
        self.writes = []

    @property
    def Text(self):
        return self._text

    @Text.setter
    def Text(self, value):
        self._text = value
        self.writes.append(("whole", value))

    @property
    def Length(self):
        return len(self._text)

    def Characters(self, Start, Length):
        return Characters(self, Start, Length)

    def Paragraphs(self):
        return type("C", (), {"Count": self._text.count(CR) + 1})()

    def Runs(self, index=None):
        pieces, at = [], 1
        for piece in self._text.split(CR):
            pieces.append(FakeRun(piece, at))
            at += len(piece) + 1
        if index is None:
            return type("C", (), {"Count": len(pieces)})()
        return pieces[index - 1]


class FakeShape:
    def __init__(self, name, text):
        self.Name = name
        self.HasTextFrame = True
        self.range = FakeTextRange(text)


def run(text=None, existing="", **kwargs):
    """Call the impl against a stand in shape, and report what it wrote."""
    shape = FakeShape("TextBox 19", existing)
    formatted = []
    with patch("ppt_com.text._text_frame_of", lambda *a: (shape, shape.range)), \
            patch("ppt_com.text._format_span",
                  lambda s, tr, start, length, spec: formatted.append(
                      (start, length, spec.get("font_size"))) or {}), \
            patch("ppt_com.text._apply_font_props"), \
            patch("ppt_com.text._apply_highlight"):
        result = _set_text_impl(1, "TextBox 19", text, **kwargs)
    return result, shape.range, formatted


SENTENCE = "吹奏楽部で爆笑してしまいました"


class TestReplacingTheWholeFrame:
    def test_it_still_does_what_it_always_did(self):
        result, tr, _ = run("あたらしい文", existing=SENTENCE)
        assert tr.Text == "あたらしい文"
        assert result["text_length"] == 6

    def test_a_paragraph_break_is_written_as_powerpoint_spells_it(self):
        _, tr, _ = run("上" + LF + "下")
        assert tr.Text == "上" + CR + "下"

    def test_and_a_line_break_is_left_alone(self):
        _, tr, _ = run("上" + VT + "下")
        assert tr.Text == "上" + VT + "下"


class TestReplacingASpan:
    def test_search_text_replaces_just_the_match(self):
        result, tr, _ = run("大笑い", existing=SENTENCE, search_text="爆笑")
        assert tr.Text == "吹奏楽部で大笑いしてしまいました"
        assert (result["start"], result["replaced_length"]) == (6, 2)
        assert result["written_length"] == 3

    def test_start_and_length_replace_that_much(self):
        _, tr, _ = run("XX", existing="ABCDEF", start=3, length=2)
        assert tr.Text == "ABXXEF"

    def test_a_length_of_zero_inserts_and_removes_nothing(self):
        # The caption box case: one character at the end of a line.
        _, tr, _ = run("。", existing=SENTENCE, start=len(SENTENCE) + 1, length=0)
        assert tr.Text == SENTENCE + "。"

    def test_inserting_in_the_middle_keeps_both_sides(self):
        _, tr, _ = run("XX", existing="ABCDEF", start=4, length=0)
        assert tr.Text == "ABCXXDEF"

    def test_only_the_span_is_written_never_the_whole_frame(self):
        # Writing the frame is what flattens the runs, so the span path must
        # not touch it.
        _, tr, _ = run("XX", existing="ABCDEF", start=3, length=2)
        assert [kind for kind, *_ in tr.writes] == ["span"]

    def test_text_that_is_not_there_fails_and_writes_nothing(self):
        with pytest.raises(ValueError, match="not found in shape"):
            run("X", existing=SENTENCE, search_text="ありません")


class TestWritingRuns:
    RUNS = [
        {"text": "吹奏楽部で", "font_size": 36},
        {"text": "「べろだして」", "font_size": 54, "color": "#FF0066"},
        {"text": "に聞こえた", "font_size": 36},
    ]

    def test_the_text_is_written_once(self):
        _, tr, _ = run(None, runs=self.RUNS)
        assert [kind for kind, *_ in tr.writes] == ["whole"]
        assert tr.Text == "吹奏楽部で「べろだして」に聞こえた"

    def test_each_run_is_formatted_over_its_own_span(self):
        _, _, formatted = run(None, runs=self.RUNS)
        assert formatted == [(1, 5, 36), (6, 7, 54), (13, 5, 36)]

    def test_the_spans_are_reported_back(self):
        result, _, _ = run(None, runs=self.RUNS)
        assert [r["start"] for r in result["runs"]] == [1, 6, 13]
        assert result["runs"][1]["text"] == "「べろだして」"

    def test_a_paragraph_break_inside_a_run_does_not_shift_the_next_one(self):
        _, _, formatted = run(None, runs=[
            {"text": "上" + LF + "下", "font_size": 36},
            {"text": "次", "font_size": 20},
        ])
        assert formatted == [(1, 3, 36), (4, 1, 20)]

    def test_a_bad_colour_in_a_later_run_writes_nothing_at_all(self):
        with pytest.raises(Exception):
            run(None, runs=[
                {"text": "あ", "font_size": 36},
                {"text": "い", "color": "magenta"},
            ])


def rejected(**kwargs):
    with pytest.raises(ValidationError) as caught:
        SetTextInput(slide_index=1, shape_name_or_index="T", **kwargs)
    return str(caught.value)


class TestWhatTheToolAccepts:
    def test_text_alone(self):
        assert SetTextInput(slide_index=1, shape_name_or_index="T", text="a").text == "a"

    def test_a_span_with_text(self):
        params = SetTextInput(
            slide_index=1, shape_name_or_index="T", text="。", start=5, length=0)
        assert (params.start, params.length) == (5, 0)

    def test_runs_alone(self):
        params = SetTextInput(
            slide_index=1, shape_name_or_index="T", runs=[{"text": "a"}])
        assert params.runs[0].text == "a"

    def test_neither_text_nor_runs_is_refused(self):
        assert "either text or runs" in rejected(start=1, length=0)

    def test_both_text_and_runs_is_refused(self):
        assert "either text or runs" in rejected(text="a", runs=[{"text": "b"}])

    def test_a_span_with_runs_is_refused(self):
        assert "cannot take a span" in rejected(runs=[{"text": "a"}], start=1, length=0)

    def test_a_start_without_a_length_is_refused(self):
        assert "Either search_text or both" in rejected(text="a", start=1)

    def test_a_negative_length_is_refused(self):
        assert "greater than or equal to 0" in rejected(text="a", start=1, length=-1)

    def test_a_start_of_zero_is_refused(self):
        # Positions are 1-based, and 0 would silently mean the first character.
        assert "greater than or equal to 1" in rejected(text="a", start=0, length=1)

    def test_search_text_beside_start_is_refused(self):
        assert "mutually exclusive" in rejected(
            text="a", search_text="b", start=1, length=1)

    def test_an_empty_search_text_is_refused(self):
        # It matches at position 1 with no length, so it would insert at the
        # front of the shape instead of reporting the mistake.
        assert "must not be empty" in rejected(text="a", search_text="")

    def test_an_empty_runs_list_is_refused(self):
        assert "at least 1 item" in rejected(runs=[])

    def test_an_occurrence_without_search_text_is_refused(self):
        assert "only valid with search_text" in rejected(text="a", occurrence=2)

    def test_a_run_keeps_its_leading_and_trailing_spaces(self):
        # A run is a fragment, so its edges are the gaps between words. The
        # other strings on these models are stripped; this one must not be.
        assert RunSpec(text=" word ").text == " word "

    def test_and_its_trailing_newline(self):
        # Stripping it would drop the caller's paragraph break in silence.
        assert RunSpec(text="line" + LF).text == "line" + LF

    def test_a_run_of_only_spaces_is_still_a_run(self):
        assert RunSpec(text="  ").text == "  "

    def test_an_empty_run_is_refused(self):
        with pytest.raises(ValidationError, match="at least 1 character"):
            RunSpec(text="")
