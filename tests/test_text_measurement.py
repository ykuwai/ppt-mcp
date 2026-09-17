"""Tests for the overflow arithmetic behind ppt_get_text(measure=True).

Pure Python. `build_measurement` is the part both platforms share, so this
pins the answer rather than the walk that gathers it.
"""

import sys

sys.path.insert(0, "src")

import pytest

from ppt_com.text import build_measurement

MARGINS = {"left": 7.2, "right": 7.2, "top": 3.6, "bottom": 3.6}


def line(text, width, height=34.0):
    return {"text": text, "width_pt": width, "height_pt": height}


def measure(lines=None, text_width=392.4, text_height=68.0,
            shape_width=427.0, shape_height=200.0, margins=MARGINS,
            word_wrap=True, autofit="none", orientation="horizontal"):
    if lines is None:
        lines = [line("一行目", 392.4), line("二行目", 210.0)]
    return build_measurement(
        lines, text_width, text_height,
        shape_width, shape_height, margins, word_wrap, autofit, orientation,
    )


class TestTheUsableSizeIsTheShapeLessItsMargins:
    def test_a_default_box_loses_about_fourteen_points_of_width(self):
        assert measure()["usable_width_pt"] == 412.6

    def test_and_about_seven_of_height(self):
        assert measure()["usable_height_pt"] == 192.8

    def test_without_margins_there_is_no_usable_size_and_no_verdict(self):
        state = measure(margins=None)
        assert state["usable_width_pt"] is None
        assert state["usable_height_pt"] is None
        assert state["overflows"] is None

    def test_and_the_measured_numbers_still_come_back(self):
        state = measure(margins=None)
        assert state["text_height_pt"] == 68.0
        assert state["line_count"] == 2


class TestOverflow:
    def test_text_inside_the_box_does_not_overflow(self):
        assert measure()["overflows"] is False

    def test_text_taller_than_the_box_overflows(self):
        assert measure(text_height=260.0)["overflows"] is True

    def test_a_line_exactly_on_the_boundary_fits(self):
        assert measure(text_height=192.8)["overflows"] is False

    def test_a_hundredth_of_a_point_over_is_still_a_fit(self):
        # Bounds are computed in EMU and come back rounded, so exact
        # equality is not something PowerPoint promises.
        assert measure(text_height=192.81)["overflows"] is False

    def test_a_point_over_is_not(self):
        assert measure(text_height=193.8)["overflows"] is True

    def test_width_is_not_an_overflow_while_wrapping_is_on(self):
        # A long line becomes two lines, which is a height problem.
        assert measure(text_width=900.0)["overflows"] is False

    def test_width_is_an_overflow_once_wrapping_is_off(self):
        state = measure(text_width=900.0, word_wrap=False)
        assert state["overflows"] is True

    def test_an_unknown_wrap_setting_is_treated_as_wrapping(self):
        state = measure(text_width=900.0, word_wrap=None)
        assert state["overflows"] is False

    def test_height_wins_over_width_when_both_are_out(self):
        state = measure(text_width=900.0, text_height=260.0, word_wrap=False)
        assert state["overflows"] is True


class TestAnEmptyOrUnmeasurableFrame:
    def test_no_lines_is_a_count_of_zero_rather_than_a_failure(self):
        state = measure(lines=[], text_width=None, text_height=None)
        assert state["line_count"] == 0
        assert state["lines"] == []

    def test_and_what_was_never_measured_gets_no_verdict(self):
        state = measure(lines=[], text_width=None, text_height=None)
        assert state["overflows"] is None

    def test_an_unreadable_height_costs_the_verdict_on_a_wrapping_box(self):
        assert measure(text_height=None)["overflows"] is None

    def test_an_unreadable_width_does_not_while_wrapping_is_on(self):
        assert measure(text_width=None)["overflows"] is False

    def test_but_it_does_once_wrapping_is_off(self):
        assert measure(text_width=None, word_wrap=False)["overflows"] is None

    def test_a_line_powerpoint_would_not_measure_keeps_its_text(self):
        state = measure(lines=[line("一行目", None, None)])
        assert state["lines"][0]["text"] == "一行目"
        assert state["lines"][0]["width_pt"] is None


class TestWhatTheBlockCarries:
    def test_the_line_count_is_the_lines_it_was_given(self):
        assert measure()["line_count"] == 2

    def test_autofit_travels_with_the_numbers(self):
        # "overflows": false under shrink_to_fit only reads correctly next to
        # the setting that made it false.
        assert measure(autofit="shrink_to_fit")["autofit"] == "shrink_to_fit"

    def test_the_keys_are_the_same_whatever_was_readable(self):
        assert set(measure()) == set(measure(margins=None))


class TestVerticalText:
    """A vertical line is a column. Wrapping ends it at the bottom of the box,
    so the block grows sideways and that is where overflow shows.

    Measured on a real 200 by 300pt box: three columns of 38.4 by 288, a block
    115.2 wide and 294.5 tall, in 185.6 by 292.8 of usable space. Read as
    horizontal text that is an overflow, and the rendered slide shows the text
    sitting comfortably inside its box.
    """

    VERTICAL = dict(
        lines=[line("吹奏楽部で「べろだ", 38.4, 288.0),
               line("して」に聞こえて爆", 38.4, 288.0),
               line("笑してしまいました", 38.4, 288.01)],
        text_width=115.2, text_height=294.5,
        shape_width=200.0, shape_height=300.0,
        orientation="vertical",
    )

    def test_a_column_a_little_taller_than_the_box_is_not_an_overflow(self):
        assert measure(**self.VERTICAL)["overflows"] is False

    def test_the_same_numbers_read_as_horizontal_are(self):
        assert measure(**dict(self.VERTICAL, orientation="horizontal"))[
            "overflows"] is True

    def test_more_columns_than_fit_across_is_an_overflow(self):
        # The same text in an 80pt wide box.
        assert measure(**dict(self.VERTICAL, shape_width=80.0))[
            "overflows"] is True

    def test_upward_and_downward_are_laid_out_the_same_way(self):
        # PowerPoint answers the same bounds for all three, the glyphs turn
        # and the geometry does not.
        for orientation in ("upward", "downward"):
            state = measure(**dict(self.VERTICAL, orientation=orientation))
            assert state["overflows"] is False, orientation

    def test_with_wrapping_off_a_column_past_the_bottom_overflows(self):
        state = measure(**dict(self.VERTICAL, word_wrap=False))
        assert state["overflows"] is True

    def test_an_unknown_orientation_is_treated_as_horizontal(self):
        assert measure(**dict(self.VERTICAL, orientation=None))[
            "overflows"] is True
