"""Tests for the layout checks behind ppt_check_layout and layout_issues.

Pure Python for the checks themselves: a slide is a list of plain dicts, the
shape ``_slide_facts_impl`` returns on both platforms, so these pin the
answers rather than the walk that gathers them. The Windows reader is then
run over fake COM objects, since it is the half CI exercises.
"""

import json
import re
import sys
from pathlib import Path
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import pytest

from ppt_com import layout_check
from ppt_com.constants import (
    msoAutoShape, msoFalse, msoFillSolid, msoPicture, msoPlaceholder,
    msoTable, msoTextBox, msoTrue,
)
from ppt_com.layout_check import (
    deck_findings, find_title, ink_box, paints_box, report_layout,
    slide_findings, text_box, touched_names,
)

W, H = 960.0, 540.0
MARGINS = {"left": 7.2, "right": 7.2, "top": 3.6, "bottom": 3.6}


def text(width=150.0, height=30.0, left=None, top=None, chars=12,
         runs=None, alignment=None, anchor="top", word_wrap=True,
         autofit="none", orientation="horizontal", margins=MARGINS):
    return {
        "chars": chars, "left": left, "top": top, "width": width,
        "height": height, "alignment": alignment, "vertical_anchor": anchor,
        "runs": runs if runs is not None else [
            {"color_hex": "#222222", "size": 20.0, "bold": False}],
        "margins": margins, "word_wrap": word_wrap, "autofit": autofit,
        "orientation": orientation,
    }


def shape(name, left, top, width, height, z=1, type=msoAutoShape,
          fill=None, line=None, words=None, **extra):
    facts = {
        "name": name, "z": z, "type": type,
        "left": left, "top": top, "width": width, "height": height,
        "rotation": 0.0, "visible": True, "connector": False,
        "placeholder_type": None, "fill": fill, "line_visible": line,
        "text": words,
    }
    facts.update(extra)
    return facts


def textbox(name, left, top, width, height, z=1, text_width=150.0,
            text_height=30.0, **words):
    """A text box with no fill or line, its text at the top left inset."""
    return shape(
        name, left, top, width, height, z=z, type=msoTextBox,
        fill={"visible": False, "solid": True, "color_hex": "#FFFFFF", "transparency": 0.0},
        line=False,
        words=text(width=text_width, height=text_height,
                   left=left + 7.2, top=top + 3.6, **words),
    )


def card(name, left, top, width, height, z=1, colour="#FFFFFF", words=None):
    return shape(
        name, left, top, width, height, z=z,
        fill={"visible": True, "solid": True, "color_hex": colour, "transparency": 0.0},
        line=False, words=words,
    )


def slide(*shapes, index=1, background="#FFFFFF", layout="Blank"):
    return {"slide_index": index, "layout": layout,
            "background_hex": background, "shapes": list(shapes)}


def kinds(findings):
    return [f["type"] for f in findings]


def check(*shapes, **kwargs):
    return slide_findings(slide(*shapes, **kwargs), W, H)


def titled_row(lefts=(40, 340, 640), heights=(180, 180, 180)):
    """A title across the slide and a row of three cards under it, 40 pt in
    from both sides, as the live test slide was built."""
    title = textbox("title", 40, 30, 880, 55.74, z=1, text_width=300, text_height=48)
    return [title] + [card(f"card{i}", left, 140, 280, height, z=2 + i)
                      for i, (left, height) in enumerate(zip(lefts, heights, strict=True))]


def spill():
    """The paragraph on the first card, where PowerPoint for Mac drew it:
    its text from 153.6 to 345.6 pt, 25.6 pt past a 180 pt card."""
    return textbox("words", 52, 150, 256, 201.15, z=9,
                   text_width=245.07, text_height=192)


# ---------------------------------------------------------------------------
# What a shape puts on the slide
# ---------------------------------------------------------------------------
class TestWhatAShapeDraws:
    def test_a_filled_shape_draws_its_box(self):
        assert ink_box(card("c", 10, 20, 100, 50)) == (10, 20, 110, 70)

    def test_a_text_box_draws_only_its_text(self):
        box = textbox("t", 100, 100, 400, 200, text_width=150, text_height=30)
        assert ink_box(box) == pytest.approx((107.2, 103.6, 257.2, 133.6))

    def test_an_empty_text_box_draws_nothing(self):
        box = textbox("t", 100, 100, 400, 200)
        box["text"]["chars"] = 0
        assert ink_box(box) is None

    def test_a_hidden_shape_draws_nothing(self):
        assert ink_box(card("c", 10, 20, 100, 50) | {"visible": False}) is None

    def test_a_picture_draws_its_box_whatever_its_fill_says(self):
        assert paints_box(shape("p", 0, 0, 10, 10, type=msoPicture, fill={"visible": False}))

    def test_an_autoshape_whose_line_cannot_be_read_is_taken_as_outlined(self):
        # macOS cannot say whether a line is drawn, and a new autoshape has one.
        assert paints_box(shape("s", 0, 0, 10, 10, fill={"visible": False}, line=None))

    def test_a_text_box_whose_line_cannot_be_read_is_taken_as_not(self):
        assert not paints_box(shape("t", 0, 0, 10, 10, type=msoTextBox,
                                    fill={"visible": False}, line=None))


class TestWhereTextIsPlacedWhenPowerPointWillNotSay:
    def frame(self, **words):
        return shape("t", 100, 100, 400, 200, type=msoTextBox,
                     fill={"visible": False}, line=False,
                     words=text(**{"width": 100, "height": 40, **words}))

    def test_top_left_is_the_inset_corner(self):
        assert text_box(self.frame()) == pytest.approx((107.2, 103.6, 207.2, 143.6))

    def test_centred_text_is_centred_in_the_usable_width(self):
        left = text_box(self.frame(alignment=2))[0]
        assert left == pytest.approx(107.2 + (385.6 - 100) / 2)

    def test_right_aligned_text_ends_at_the_right_inset(self):
        assert text_box(self.frame(alignment=3))[2] == pytest.approx(500 - 7.2)

    def test_middle_anchored_text_is_centred_in_the_usable_height(self):
        top = text_box(self.frame(anchor="middle"))[1]
        assert top == pytest.approx(103.6 + (192.8 - 40) / 2)

    def test_bottom_anchored_text_ends_at_the_bottom_inset(self):
        assert text_box(self.frame(anchor="bottom"))[3] == pytest.approx(300 - 3.6)

    def test_reported_bounds_win_over_the_arithmetic(self):
        assert text_box(self.frame(left=5, top=6)) == (5, 6, 105, 46)

    def test_without_a_measured_size_there_is_no_answer(self):
        assert text_box(self.frame(width=None)) is None


# ---------------------------------------------------------------------------
# Off the slide
# ---------------------------------------------------------------------------
class TestOffSlide:
    def test_a_shape_past_the_right_edge_says_by_how_much(self):
        [finding] = check(card("c", 900, 100, 100, 50))
        assert finding["type"] == "off_slide"
        assert finding["edges"] == {"right": 40.0}
        # With nothing else on the slide there is no margin to keep.
        assert "left by at least 40 pt (ppt_update_shape left=860)" in finding["fix"]

    def test_it_comes_back_inside_the_margin_not_onto_the_edge(self):
        # Live: moved by the overshoot alone, the box ended flush against the
        # slide's right edge.
        [finding] = check(*titled_row(), card("box", 900, 470, 120, 50, z=9))
        assert finding["edges"] == {"right": 60.0}
        assert "left by 100 pt, to keep the 40 pt margin" in finding["fix"]
        assert "ppt_update_shape left=800" in finding["fix"]

    def test_a_frame_hanging_over_the_edge_with_its_text_inside_is_fine(self):
        assert check(textbox("t", 800, 100, 300, 50, text_width=100)) == []

    def test_text_spilling_off_the_bottom_is_not(self):
        box = textbox("t", 100, 480, 300, 50, text_height=90,
                      autofit="shape_to_fit")
        assert "off_slide" in kinds(check(box))

    def test_half_a_point_is_rounding(self):
        assert check(card("c", 860.4, 100, 100, 50)) == []

    def test_a_rotated_shape_is_measured_by_the_box_around_it(self):
        bar = card("bar", 380, 200, 200, 20) | {"rotation": 90.0}
        assert check(bar) == []
        bar["top"] = 5
        assert kinds(check(bar)) == ["off_slide"]


# ---------------------------------------------------------------------------
# Overflow
# ---------------------------------------------------------------------------
class TestOverflow:
    def test_text_taller_than_its_box_says_how_much_more_it_needs(self):
        box = textbox("t", 100, 100, 300, 40, text_height=60)
        [finding] = [f for f in check(box) if f["type"] == "overflow"]
        assert finding["needed_pt"] == 60
        assert finding["available_pt"] == pytest.approx(32.8)
        assert "27.2 pt taller" in finding["fix"]

    def test_text_that_fits_is_not_reported(self):
        assert check(textbox("t", 100, 100, 300, 40, text_height=30)) == []

    def test_a_line_too_long_for_an_unwrapped_box_is_a_width_problem(self):
        box = textbox("t", 100, 100, 200, 40, text_width=300, text_height=30,
                      word_wrap=False)
        [finding] = [f for f in check(box) if f["type"] == "overflow"]
        assert "wider" in finding["fix"]
        assert finding["needed_pt"] == 300

    def test_an_unwrapped_box_sized_to_its_text_cannot_run_out_of_its_side(self):
        # Live, on the Mac: "1516 עד 1917" at 44 pt, in a box PowerPoint had
        # sized to it at 283.3 pt, was read as a 276.1 pt line, 7.3 pt past
        # the insets. It drew inside the box.
        box = textbox("row1", 516.7, 40, 283.3, 70, text_width=276.1, text_height=60,
                      word_wrap=False, autofit="shape_to_fit")
        assert [f for f in check(box) if f["type"] == "overflow"] == []


# ---------------------------------------------------------------------------
# Overlap
# ---------------------------------------------------------------------------
class TestOverlap:
    def test_two_texts_running_into_each_other_are_reported_once(self):
        found = check(textbox("a", 100, 100, 200, 40, z=1),
                      textbox("b", 100, 120, 200, 40, z=2))
        overlaps = [f for f in found if f["type"] == "overlap"]
        assert len(overlaps) == 1
        assert set(overlaps[0]["shapes"]) == {"a", "b"}
        assert overlaps[0]["overlap_pt"] == {"width": 150.0, "height": 10.0}

    def test_text_on_a_card_is_on_it_not_in_its_way(self):
        found = check(card("card", 80, 80, 300, 200, z=1),
                      textbox("words", 100, 100, 260, 60, z=2))
        assert "overlap" not in kinds(found)

    def test_text_running_past_the_bottom_of_its_card_says_how_far(self):
        found = check(card("card", 80, 80, 300, 100, z=1),
                      textbox("words", 100, 100, 260, 60, z=2, text_height=110))
        [finding] = [f for f in found if f["type"] == "overlap"]
        assert finding["past_edges"] == {"bottom": pytest.approx(33.6)}
        assert "sits on 'card'" in finding["message"]

    def test_an_icon_dropped_on_a_paragraph_is_in_the_way(self):
        found = check(textbox("para", 100, 100, 400, 100, z=1,
                              text_width=380, text_height=90),
                      shape("icon", 200, 120, 40, 40, z=2, type=msoPicture))
        [finding] = [f for f in found if f["type"] == "overlap"]
        assert "runs into 'icon'" in finding["message"]

    def test_the_card_grows_with_padding_and_its_row_with_it(self):
        # Live: "larger by that much" left the last line on the card's edge
        # and the card taller than the other two.
        [finding] = [f for f in check(*titled_row(), spill()) if f["type"] == "overlap"]
        assert finding["past_edges"] == {"bottom": 25.6}
        assert "make 'card0' 39.2 pt taller" in finding["fix"].lower()
        assert "shape_names=['card0', 'card1', 'card2'] height=219.2" in finding["fix"]
        assert "13.6 pt below, as it has above" in finding["fix"]

    def test_two_filled_shapes_crossing_are_left_alone(self):
        # A Venn diagram looks exactly like this.
        assert check(card("a", 100, 100, 200, 200, z=1),
                     card("b", 250, 100, 200, 200, z=2)) == []

    def test_touching_is_not_overlapping(self):
        found = check(textbox("a", 100, 100, 200, 40, z=1),
                      textbox("b", 100, 132, 200, 40, z=2))
        assert "overlap" not in kinds(found)


# ---------------------------------------------------------------------------
# Contrast
# ---------------------------------------------------------------------------
WHITE_TEXT = [{"color_hex": "#FFFFFF", "size": 20.0, "bold": False}]


class TestContrast:
    def test_white_text_on_a_white_slide_is_reported(self):
        [finding] = check(textbox("t", 100, 100, 300, 50, runs=WHITE_TEXT))
        assert finding["type"] == "low_contrast"
        assert finding["contrast"] == 1.0
        assert "#000000" in finding["fix"]

    def test_dark_text_on_a_white_slide_is_fine(self):
        assert check(textbox("t", 100, 100, 300, 50)) == []

    def test_the_card_underneath_is_what_the_text_is_read_against(self):
        found = check(card("card", 80, 80, 400, 200, z=1, colour="#1F3864"),
                      textbox("t", 100, 100, 300, 50, z=2, runs=WHITE_TEXT))
        assert "low_contrast" not in kinds(found)

    def test_text_on_a_picture_gets_no_verdict(self):
        found = check(shape("photo", 0, 0, 960, 540, z=1, type=msoPicture),
                      textbox("t", 100, 100, 300, 50, z=2, runs=WHITE_TEXT))
        assert "low_contrast" not in kinds(found)

    def test_a_see_through_fill_gets_no_verdict(self):
        glass = card("glass", 80, 80, 400, 200, z=1, colour="#FFFFFF")
        glass["fill"]["transparency"] = 0.6
        found = check(glass, textbox("t", 100, 100, 300, 50, z=2, runs=WHITE_TEXT))
        assert "low_contrast" not in kinds(found)

    def test_the_fix_names_the_call(self):
        [finding] = check(textbox("t", 100, 100, 300, 50, runs=WHITE_TEXT))
        assert "ppt_format_text color='#000000'" in finding["fix"]

    def test_the_fix_prefers_the_colour_the_other_text_uses(self):
        found = check(textbox("title", 40, 30, 880, 60, z=1),
                      textbox("note", 40, 440, 400, 40, z=2, runs=WHITE_TEXT))
        [finding] = found
        assert "#222222, the colour of the other text on the slide" in finding["fix"]
        assert "ppt_format_text color='#222222'" in finding["fix"]

    def test_large_text_is_held_to_the_lower_minimum(self):
        grey = [{"color_hex": "#949494", "size": 24.0, "bold": False}]
        assert check(textbox("t", 100, 100, 300, 50, runs=grey)) == []
        small = [{"color_hex": "#949494", "size": 12.0, "bold": False}]
        assert kinds(check(textbox("t", 100, 100, 300, 50, runs=small))) == ["low_contrast"]


# ---------------------------------------------------------------------------
# Alignment
# ---------------------------------------------------------------------------
def three_cards(second_top=200.0, third_left=560.0):
    return (card("one", 100, 200, 200, 100, z=1),
            card("two", 330, second_top, 200, 100, z=2),
            card("three", third_left, 200, 200, 100, z=3))


class TestMisaligned:
    def test_a_card_three_points_low_is_the_one_named(self):
        found = check(*three_cards(second_top=203))
        [finding] = [f for f in found if f["type"] == "misaligned"]
        assert finding["shapes"][0] == "two"
        assert set(finding["shapes"][1:]) == {"one", "three"}
        assert finding["edge"] == "top"
        assert "top=200" in finding["fix"]

    def test_ten_points_is_a_decision_not_a_slip(self):
        assert "misaligned" not in kinds(check(*three_cards(second_top=210)))

    def test_aligned_cards_are_fine(self):
        assert check(*three_cards()) == []

    def test_shapes_in_neither_a_row_nor_a_column_are_not_compared(self):
        found = check(card("a", 100, 100, 100, 50, z=1),
                      card("b", 400, 302, 100, 50, z=2))
        assert "misaligned" not in kinds(found)

    def test_between_two_the_later_one_moves(self):
        found = check(card("title", 40, 40, 880, 60, z=1),
                      card("body", 43, 120, 400, 300, z=2))
        [finding] = [f for f in found if f["type"] == "misaligned"]
        assert finding["shapes"] == ["body", "title"]
        assert "left=40" in finding["fix"]

    def test_text_on_a_card_is_not_lined_up_with_the_row(self):
        # Live on PowerPoint for Mac, once the cards had grown to hold this
        # text: the middle of its frame was 0.6 pt off the middle of the
        # cards, and the text was reported as out of line with them.
        found = check(card("one", 40, 140, 280, 220, z=1),
                      card("two", 340, 140, 280, 220, z=2),
                      card("three", 650, 140, 280, 220, z=3),
                      textbox("words", 52, 150, 256, 201.15, z=4,
                              text_width=245.07, text_height=192))
        assert "misaligned" not in kinds(found)

    def test_the_text_on_one_card_is_lined_up_with_the_next(self):
        lefts = (40, 340, 640)
        cards = [card(f"card{i}", left, 140, 280, 220, z=i + 1)
                 for i, left in enumerate(lefts)]
        heads = [textbox(f"head{i}", left + 20, 163 if i == 1 else 160, 240, 40,
                         z=4 + i, text_width=100)
                 for i, left in enumerate(lefts)]
        [finding] = [f for f in check(*cards, *heads) if f["type"] == "misaligned"]
        assert finding["shapes"][0] == "head1"
        assert set(finding["shapes"][1:]) == {"head0", "head2"}

    def test_only_the_shapes_on_the_line_are_named(self):
        found = check(card("one", 100, 200, 100, 100, z=1),
                      card("two", 230, 203, 100, 100, z=2),
                      card("three", 360, 200, 100, 100, z=3),
                      card("four", 490, 201.5, 100, 100, z=4))
        by_mover = {f["shapes"][0]: f for f in found if f["type"] == "misaligned"}
        assert set(by_mover) == {"two", "four"}
        assert set(by_mover["four"]["shapes"][1:]) == {"one", "three"}
        assert "'two'" not in by_mover["four"]["message"]


# ---------------------------------------------------------------------------
# Spacing
# ---------------------------------------------------------------------------
class TestUnevenSpacing:
    def test_gaps_of_thirty_and_twenty_three_are_reported(self):
        found = check(*three_cards(third_left=560))
        assert "uneven_spacing" not in kinds(found)
        found = check(card("one", 100, 200, 200, 100, z=1),
                      card("two", 324, 200, 200, 100, z=2),
                      card("three", 555, 200, 200, 100, z=3))
        [finding] = [f for f in found if f["type"] == "uneven_spacing"]
        assert finding["shapes"] == ["one", "two", "three"]
        assert finding["gaps_pt"] == [24.0, 31.0]
        assert "direction='horizontal'" in finding["fix"]

    def test_the_text_on_the_cards_is_not_a_second_row(self):
        cards = [card("one", 100, 200, 200, 100, z=1),
                 card("two", 324, 200, 200, 100, z=2),
                 card("three", 555, 200, 200, 100, z=3)]
        words = [textbox(f"t{i}", c["left"] + 10, 210, 180, 40, z=4 + i,
                         text_width=100)
                 for i, c in enumerate(cards)]
        found = check(*cards, *words)
        assert kinds(found).count("uneven_spacing") == 1

    def test_text_spilling_off_a_card_still_belongs_to_the_card(self):
        # Live on PowerPoint for Mac this hid the row: the spilt text box was
        # counted as a member, overlapped its card, and the row was dropped.
        found = check(card("one", 40, 140, 280, 180, z=1),
                      card("two", 340, 140, 280, 180, z=2),
                      card("three", 650, 140, 280, 180, z=3),
                      textbox("spill", 52, 150, 256, 201.15, z=4,
                              text_width=245.07, text_height=192))
        [finding] = [f for f in found if f["type"] == "uneven_spacing"]
        assert finding["gaps_pt"] == [20.0, 30.0]
        # The frame grew with its text until it was larger than the card,
        # 256 by 201 against 280 by 180, which is what PowerPoint did.

    def test_shapes_of_different_sizes_are_not_a_group(self):
        found = check(card("one", 100, 200, 100, 100, z=1),
                      card("two", 224, 200, 300, 100, z=2),
                      card("three", 555, 200, 50, 100, z=3))
        assert "uneven_spacing" not in kinds(found)

    def test_a_row_across_the_slide_is_spread_between_the_margins(self):
        # Live: the end card was the one 10 pt out, and distributing kept it
        # there, with the gaps even and the row off centre.
        found = check(*titled_row(lefts=(40, 340, 650)))
        [finding] = [f for f in found if f["type"] == "uneven_spacing"]
        assert finding["gaps_pt"] == [20.0, 30.0]
        assert "Set left of 'card2' to 640 pt" in finding["fix"]
        assert "Every gap is then 20 pt" in finding["fix"]
        # The same card is not reported a second time for the margin.
        assert "narrow_margin" not in kinds(found)


# ---------------------------------------------------------------------------
# Sizes, room on a card, margins
# ---------------------------------------------------------------------------
class TestUnevenSizes:
    def test_one_card_taller_than_its_row_is_reported(self):
        found = check(*titled_row(heights=(205.6, 180, 180)))
        [finding] = [f for f in found if f["type"] == "uneven_sizes"]
        assert finding["sizes_pt"] == {"card0": 205.6, "card1": 180.0, "card2": 180.0}
        assert "205.6 pt for 'card0'; 180 pt for 'card1', 'card2'" in finding["message"]
        assert "shape_names=['card0', 'card1', 'card2'] height=205.6" in finding["fix"]

    def test_a_point_or_two_is_the_same_size(self):
        assert "uneven_sizes" not in kinds(check(*titled_row(heights=(181.5, 180, 180))))

    def test_text_boxes_of_different_heights_are_not_judged(self):
        # Their frames cannot be seen, and text runs as long as it runs.
        found = check(textbox("a", 40, 140, 280, 100, z=1, text_height=90),
                      textbox("b", 340, 140, 280, 90, z=2, text_height=80),
                      textbox("c", 640, 140, 280, 110, z=3, text_height=100))
        assert "uneven_sizes" not in kinds(found)

    def test_a_write_to_one_card_hears_about_the_row(self):
        facts = slide(*titled_row(heights=(205.6, 180, 180)))
        found = slide_findings(facts, W, H, only={"card0"})
        assert kinds(found) == ["uneven_sizes"]


class TestCrampedText:
    def test_text_on_the_edge_of_its_card_is_reported(self):
        # Live, after the card grew by the overshoot alone.
        found = check(*titled_row(heights=(205.6, 205.6, 205.6)), spill())
        [finding] = [f for f in found if f["type"] == "cramped_text"]
        assert finding["shapes"] == ["words", "card0"]
        assert "0 pt from the bottom edge of 'card0', with 13.6 pt above" in finding["message"]
        assert "shape_names=['card0', 'card1', 'card2'] height=219.2" in finding["fix"]
        assert "overlap" not in kinds(found)

    def test_evenly_padded_text_is_fine(self):
        assert check(*titled_row(heights=(220, 220, 220)), spill()) == []

    def test_text_at_the_top_edge_moves_down_by_the_side_padding(self):
        found = check(card("card", 100, 100, 300, 200, z=1),
                      textbox("words", 112.8, 97.4, 280, 40, z=2,
                              text_width=200, text_height=40))
        [finding] = [f for f in found if f["type"] == "cramped_text"]
        assert "Move 'words' down by 19 pt (ppt_update_shape dtop=19)" in finding["fix"]

    def test_a_quote_and_its_attribution_are_padded_as_one_block(self):
        # Live: a band with a quote and its source centred together under it.
        found = check(card("band", 60, 400, 840, 64, z=1),
                      textbox("quote", 80, 409, 800, 28, z=2, text_width=786, text_height=22),
                      textbox("source", 80, 437, 800, 20, z=3, text_width=786, text_height=14.4))
        assert "cramped_text" not in kinds(found)

    def test_a_block_pressed_against_the_bottom_names_its_last_line(self):
        found = check(card("card", 100, 100, 300, 200, z=1),
                      textbox("head", 112.8, 116.4, 280, 40, z=2, text_width=270, text_height=30),
                      textbox("body", 112.8, 150, 280, 150, z=3, text_width=270, text_height=145.4))
        [finding] = [f for f in found if f["type"] == "cramped_text"]
        assert finding["shapes"] == ["body", "head", "card"]
        assert "'body', set with 'head' as one block, is 1 pt from the bottom edge" in finding["message"]
        assert "with 20 pt above" in finding["message"]

    def test_a_block_at_the_top_edge_moves_down_together(self):
        found = check(card("card", 100, 100, 300, 200, z=1),
                      textbox("head", 112.8, 97.4, 280, 30, z=2, text_width=200, text_height=24),
                      textbox("body", 112.8, 124, 280, 40, z=3, text_width=200, text_height=30))
        [finding] = [f for f in found if f["type"] == "cramped_text"]
        assert "Move the text on 'card' ('head', 'body') down by 19 pt " \
               "(ppt_update_shape shape_names=['head', 'body'] dtop=19)" in finding["fix"]

    def test_little_room_all_round_is_a_badge(self):
        found = check(card("pill", 100, 100, 80, 24, z=1),
                      textbox("label", 95.8, 99.4, 88.4, 25.2, z=2,
                              text_width=74, text_height=18))
        assert "cramped_text" not in kinds(found)


class TestNarrowMargins:
    def test_a_card_inside_the_margin_moves_the_row(self):
        # Live, after ppt_distribute_shapes: gaps of 25 pt, the row from 40
        # to 930 pt, so 10 pt short of the margin on the right.
        found = check(*titled_row(lefts=(40, 345, 650)))
        [finding] = found
        assert finding["type"] == "narrow_margin"
        assert finding["shapes"] == ["card2"]
        assert (finding["edge"], finding["distance_pt"], finding["margin_pt"]) == ("right", 30.0, 40.0)
        assert "Set left of 'card1' to 340 pt and of 'card2' to 640 pt" in finding["fix"]

    def test_a_shape_on_its_own_moves_in(self):
        found = check(*titled_row(), card("box", 808, 470, 120, 50, z=9))
        [finding] = found
        assert finding["distance_pt"] == 32.0
        assert "Move 'box' left by 8 pt (ppt_update_shape left=800)" in finding["fix"]

    def test_a_corner_element_well_inside_the_margin_is_placed_there(self):
        # A logo or a page number often keeps a narrower margin than the
        # content, and 20 pt against 40 is a decision rather than a slip.
        assert check(*titled_row(), card("logo", 820, 480, 120, 40, z=9)) == []

    def test_a_shape_bleeding_off_the_edge_is_meant(self):
        photo = shape("photo", 760, 0, 200, 540, z=9, type=msoPicture)
        found = check(textbox("title", 40, 30, 680, 60, z=1),
                      card("card", 40, 140, 280, 180, z=2), photo)
        assert "narrow_margin" not in kinds(found)

    def test_one_shape_is_not_a_margin(self):
        assert check(card("card", 10, 100, 300, 100)) == []

    def test_between_two_margins_the_narrower_is_kept(self):
        # Two edges keep 40 pt and two keep 30: nothing says which is meant.
        a, b = card("a", 40, 100, 880, 100, z=1), card("b", 30, 300, 900, 100, z=2)
        layout = layout_check._layout([a, b], W, H)
        assert layout["margins"]["horizontal"] == 30.0
        assert "narrow_margin" not in kinds(check(a, b))


# ---------------------------------------------------------------------------
# Placeholders
# ---------------------------------------------------------------------------
class TestEmptyPlaceholders:
    def placeholder(self, kind):
        return shape("Text Placeholder 2", 50, 150, 860, 300, type=msoPlaceholder,
                     fill={"visible": False}, line=False,
                     words=text(chars=0), placeholder_type=kind)

    def test_an_empty_body_placeholder_is_reported(self):
        [finding] = check(self.placeholder(2))
        assert finding["type"] == "empty_placeholder"
        assert "ppt_delete_shape" in finding["fix"]

    def test_an_empty_footer_belongs_to_the_layout(self):
        assert check(self.placeholder(15)) == []


# ---------------------------------------------------------------------------
# Titles across the deck
# ---------------------------------------------------------------------------
def titled(index, size=40.0, top=30.0, left=40.0, layout="Blank"):
    words = [{"color_hex": "#222222", "size": size, "bold": True}]
    title = textbox(f"Title {index}", left, top, 880, 60, runs=words)
    return slide(title, index=index, layout=layout)


class TestTitles:
    def test_one_title_a_size_smaller_is_reported(self):
        slides = [titled(2), titled(3), titled(4, size=36), titled(5)]
        [finding] = deck_findings(slides, H)
        assert finding["type"] == "inconsistent_titles"
        assert finding["slides"] == [4]
        assert "40 pt" in finding["fix"]

    def test_one_title_six_points_lower_is_reported(self):
        slides = [titled(2), titled(3, top=36), titled(4)]
        [finding] = deck_findings(slides, H)
        assert "top edge" in finding["message"]
        assert finding["slides"] == [3]

    def test_a_cover_title_in_the_middle_of_the_slide_is_not_a_title(self):
        cover = titled(1, size=60, top=220)
        assert find_title(cover, H) is None
        assert deck_findings([cover, titled(2), titled(3), titled(4)], H) == []

    def test_slides_on_different_layouts_are_compared(self):
        # Live: one slide on the Title and Content layout, its 44 pt title
        # placeholder among slides with 32 pt title boxes of their own.
        slides = [titled(2), titled(3), titled(4, size=36, layout="Title and Content")]
        [finding] = deck_findings(slides, H)
        assert finding["slides"] == [4]

    def test_a_title_placeholder_is_the_title_whatever_its_size(self):
        heading = shape("Title 1", 40, 30, 880, 60, type=msoPlaceholder,
                        placeholder_type=1, words=text())
        assert find_title(slide(heading), H) is heading

    def test_a_section_headers_title_lower_on_the_slide_is_not_compared(self):
        section = shape("Title 1", 40, 220, 880, 80, type=msoPlaceholder,
                        placeholder_type=1, words=text())
        assert find_title(slide(section, layout="Section Header"), H) is None
        assert deck_findings([titled(2), slide(section, index=3), titled(4), titled(5)], H) == []


# ---------------------------------------------------------------------------
# What a write tool reports
# ---------------------------------------------------------------------------
class TestOnlyTheTouchedShape:
    def test_findings_are_limited_to_the_shapes_named(self):
        shapes = (card("off", 900, 10, 100, 50, z=1), *three_cards(second_top=203))
        found = slide_findings(slide(*shapes), W, H, only={"two"})
        assert kinds(found) == ["misaligned"]

    def test_and_to_the_kinds_a_shape_can_cause(self):
        shapes = (card("one", 100, 200, 200, 100, z=1),
                  card("two", 324, 200, 200, 100, z=2),
                  card("three", 555, 200, 200, 100, z=3))
        assert slide_findings(slide(*shapes), W, H, only={"three"}) == []


class FakePpt:
    def __init__(self, answer):
        self.answer = answer
        self.calls = []

    def execute(self, func, *args):
        self.calls.append((func.__name__, args))
        if isinstance(self.answer, Exception):
            raise self.answer
        return self.answer


def facts(*shapes):
    return {"slide_width": W, "slide_height": H, "slides": [slide(*shapes)]}


class FakeDeck:
    """PowerPoint answering a deck check: a slide count, then one slide per call."""

    def __init__(self, slides):
        self.slides = slides
        self.calls = []

    def execute(self, func, *args):
        self.calls.append((func.__name__, args))
        if func.__name__ == "_slide_count_impl":
            return len(self.slides)
        return {"slide_width": W, "slide_height": H,
                "slides": [self.slides[i - 1] for i in args[0]]}


class TestCheckingADeck:
    def test_a_deck_is_read_a_slide_per_call_with_one_template_cache(self, monkeypatch):
        pages = [slide(card("c", 10, 10, 50, 50), index=i) for i in (1, 2, 3)]
        fake = FakeDeck(pages)
        monkeypatch.setattr(layout_check, "ppt", fake)
        result = json.loads(layout_check.check_layout(layout_check.CheckLayoutInput()))
        assert result["slides_checked"] == 3
        reads = [args for name, args in fake.calls if name == "_slide_facts_impl"]
        assert [args[0] for args in reads] == [[1], [2], [3]]
        # Not navigated, read in full, and every call shares one cache.
        assert all(args[1] is False and args[2] is None for args in reads)
        assert len({id(args[3]) for args in reads}) == 1

    def test_one_slide_is_navigated_to(self, monkeypatch):
        fake = FakeDeck([slide(card("c", 10, 10, 50, 50))])
        monkeypatch.setattr(layout_check, "ppt", fake)
        layout_check.check_layout(layout_check.CheckLayoutInput(slide_index=1))
        [read] = [args for name, args in fake.calls if name == "_slide_facts_impl"]
        assert read[0] == [1] and read[1] is True


class TestReportLayout:
    def test_a_problem_is_added_to_the_result(self, monkeypatch):
        fake = FakePpt(facts(card("c", 900, 100, 100, 50)))
        monkeypatch.setattr(layout_check, "ppt", fake)
        result = report_layout({"shape_name": "c"}, 1, ["c"])
        assert kinds(result["layout_issues"]) == ["off_slide"]
        assert fake.calls == [("_slide_facts_impl", ([1], False, ["c"], None))]

    def test_nothing_is_added_when_nothing_is_wrong(self, monkeypatch):
        monkeypatch.setattr(layout_check, "ppt", FakePpt(facts(card("c", 10, 10, 50, 50))))
        assert report_layout({"shape_name": "c"}, 1, ["c"]) == {"shape_name": "c"}

    def test_a_check_that_could_not_run_says_so(self, monkeypatch):
        monkeypatch.setattr(layout_check, "ppt", FakePpt(RuntimeError("busy")))
        result = report_layout({"shape_name": "c"}, 1, ["c"])
        assert result["layout_check_failed"] == "busy"
        assert "layout_issues" not in result

    def test_a_failed_write_is_left_as_it_is(self, monkeypatch):
        fake = FakePpt(facts())
        monkeypatch.setattr(layout_check, "ppt", fake)
        assert report_layout({"error": "nope"}, 1, ["c"]) == {"error": "nope"}
        assert fake.calls == []

    def test_the_names_come_from_one_shape_or_several(self):
        assert touched_names({"shape_name": "a"}) == ["a"]
        assert touched_names({"updated": [{"shape_name": "a"}, {"shape_name": "b"}]}) == ["a", "b"]
        assert touched_names("not a dict") == []


# ---------------------------------------------------------------------------
# Reading a slide on Windows, over fake COM objects
# ---------------------------------------------------------------------------
class FakeRange:
    def __init__(self, content, runs, bounds=(107.2, 103.6, 150.0, 30.0), alignment=1):
        self.Text = content
        self._runs = runs
        if bounds is not None:
            self.BoundLeft, self.BoundTop, self.BoundWidth, self.BoundHeight = bounds
        self._alignment = alignment

    def Runs(self, i=None):
        if i is None:
            return SimpleNamespace(Count=len(self._runs))
        text, rgb, size, bold, *theme = self._runs[i - 1]
        return SimpleNamespace(Text=text, Font=SimpleNamespace(
            Color=SimpleNamespace(RGB=rgb, ObjectThemeColor=theme[0] if theme else 0),
            Size=size, Bold=bold, Name="Calibri Light" if size >= 30 else "Calibri"))

    def Paragraphs(self, i=None):
        paragraphs = self.Text.split("\r")
        if i is None:
            return SimpleNamespace(Count=len(paragraphs))
        text = paragraphs[i - 1]
        direction = 2 if text.startswith("\u05d0") else 1
        language = 1037 if text.startswith("\u05d1") else 1033   # he-IL, en-US
        return SimpleNamespace(ParagraphFormat=SimpleNamespace(
            Alignment=self._alignment, TextDirection=direction), LanguageID=language)


def fake_shape(name="TextBox 1", content="Hello there", runs=None, bounds=(107.2, 103.6, 150.0, 30.0)):
    runs = runs if runs is not None else [("Hello there", 0x222222, 20.0, msoFalse)]
    tr = FakeRange(content, runs, bounds)
    frame = SimpleNamespace(
        TextRange=tr, WordWrap=msoTrue, VerticalAnchor=1, Orientation=1,
        MarginLeft=7.2, MarginRight=7.2, MarginTop=3.6, MarginBottom=3.6,
        HasText=msoTrue,
    )
    return SimpleNamespace(
        Name=name, Type=msoTextBox, Left=100.0, Top=100.0, Width=300.0, Height=50.0,
        Rotation=0.0, Visible=msoTrue, Connector=msoFalse, HasTextFrame=msoTrue,
        TextFrame=frame, TextFrame2=SimpleNamespace(AutoSize=0),
        Fill=SimpleNamespace(Visible=msoFalse, Type=msoFillSolid,
                             ForeColor=SimpleNamespace(RGB=0xFFFFFF), Transparency=0.0),
        Line=SimpleNamespace(Visible=msoFalse),
    )


class TestTheWindowsReader:
    def test_text_bounds_come_from_powerpoint(self):
        facts = layout_check._shape_facts(fake_shape(), 1, None)
        words = facts["text"]
        assert (words["left"], words["top"], words["width"], words["height"]) == (
            107.2, 103.6, 150.0, 30.0)
        assert words["margins"] == MARGINS
        assert facts["line_visible"] is False
        assert facts["fill"]["visible"] is False

    def test_runs_are_folded_into_distinct_styles(self):
        runs = [("Big ", 0x000000, 40.0, msoTrue), ("small", 0x000000, 12.0, msoFalse),
                ("   ", 0xFFFFFF, 99.0, msoFalse), ("again", 0x000000, 12.0, msoFalse)]
        words = layout_check._shape_facts(fake_shape(runs=runs), 1, None)["text"]
        assert words["runs"] == [
            {"color_hex": "#000000", "size": 12.0, "bold": False, "font": "Calibri", "theme": False},
            {"color_hex": "#000000", "size": 40.0, "bold": True, "font": "Calibri Light",
             "theme": False},
        ]

    def test_a_theme_colour_is_told_from_one_the_caller_set(self):
        # msoThemeColorText1 is 13; a colour set by RGB reads msoNotThemeColor, 0.
        runs = [("Mine ", 0x000000, 20.0, msoFalse, 0), ("template's", 0x000000, 20.0, msoFalse, 13)]
        words = layout_check._shape_facts(fake_shape(runs=runs), 1, None)["text"]
        assert sorted(run["theme"] for run in words["runs"]) == [False, True]

    def test_a_bound_powerpoint_will_not_give_stays_unknown(self):
        words = layout_check._shape_facts(fake_shape(bounds=None), 1, None)["text"]
        assert words["width"] is None and words["left"] is None

    def test_an_empty_frame_is_not_measured(self):
        words = layout_check._shape_facts(fake_shape(content="  ", runs=[]), 1, None)["text"]
        assert words["chars"] == 0
        assert words["width"] is None

    def test_a_shape_not_asked_about_is_read_only_for_where_its_text_is(self):
        words = layout_check._shape_facts(fake_shape(), 1, ["Another"])["text"]
        assert (words["left"], words["width"]) == (107.2, 150.0)
        assert words["runs"] == [] and words["margins"] is None

    def test_the_shape_asked_about_is_read_in_full(self):
        words = layout_check._shape_facts(fake_shape(), 1, ["TextBox 1"])["text"]
        assert words["runs"] and words["margins"] == MARGINS

    def test_paragraph_directions_are_read_for_hebrew_text(self):
        # Set right to left; tagged as Hebrew, which PowerPoint draws from the
        # right as well; neither.
        content = "\u05d0 1\r\u05d1 2\r\u05d2 3"
        shape = fake_shape(content=content, runs=[(content, 0x222222, 20.0, msoFalse)])
        assert layout_check._shape_facts(shape, 1, None)["text"]["directions"] == ["rtl", "rtl", "ltr"]

    def test_and_not_for_text_without_it(self):
        assert "directions" not in layout_check._shape_facts(fake_shape(), 1, None)["text"]

    def test_alignment_is_read_only_when_powerpoint_will_not_place_the_text(self):
        assert layout_check._shape_facts(fake_shape(), 1, None)["text"]["alignment"] is None
        unplaced = layout_check._shape_facts(fake_shape(bounds=None), 1, None)["text"]
        assert unplaced["alignment"] == 1


class TestWhatAWriteReads:
    def test_a_shape_meeting_the_touched_one_is_read(self):
        touched = card("moved", 100, 100, 200, 100)
        assert layout_check.within_reach(card("near", 250, 150, 100, 100), [touched])

    def test_a_shape_far_from_it_is_not(self):
        touched = card("moved", 100, 100, 200, 100)
        assert not layout_check.within_reach(card("far", 600, 300, 100, 100), [touched])

    def test_the_touched_text_reaches_as_far_as_it_is_drawn(self):
        spilling = textbox("words", 100, 100, 200, 40, text_height=200)
        below = card("below", 100, 250, 200, 50)
        assert layout_check.within_reach(below, [spilling])

    def test_a_like_shape_in_its_row_is_read(self):
        touched = card("moved", 40, 140, 280, 180)
        assert layout_check.within_reach(card("third", 640, 140, 280, 180), [touched])
        assert not layout_check.within_reach(card("below", 640, 400, 280, 100), [touched])

    def test_a_placeholder_is_always_read(self):
        far = shape("Body", 600, 300, 100, 100, type=msoPlaceholder, placeholder_type=2)
        assert layout_check.within_reach(far, [card("moved", 0, 0, 10, 10)])

    def test_the_reading_follows_that(self):
        refs = ["moved", "near", "far"]
        boxes = {"moved": (100, 100), "near": (250, 150), "far": (600, 300)}
        read = []

        def geometry(ref, z):
            left, top = boxes[ref]
            return card(ref, left, top, 200, 100, z=z) | {"fill": None}

        def appearance(facts, ref, full):
            read.append((ref, full))
            return facts

        layout_check.read_shapes(refs, geometry, appearance, ["moved"])
        assert read == [("moved", True), ("near", False)]
        read.clear()
        layout_check.read_shapes(refs, geometry, appearance, None)
        assert read == [("moved", True), ("near", True), ("far", True)]

    def test_the_background_is_read_only_when_touched_text_needs_it(self):
        words = textbox("words", 100, 100, 200, 40)
        assert layout_check.needs_background([words], None)
        assert layout_check.needs_background([words], ["words"])
        assert not layout_check.needs_background([words, card("c", 0, 0, 5, 5)], ["c"])


# ---------------------------------------------------------------------------
# A real deck
# ---------------------------------------------------------------------------
# Fourteen slides built through the MCP server on PowerPoint for Mac the way
# a model builds them, read back as the checks read them, before and after
# their findings were fixed. Every finding below was checked against a
# screenshot of the slide, and so was every slide reported clean.
FIXTURES = Path(__file__).parent / "fixtures" / "layout"


def recorded(name):
    return json.loads((FIXTURES / f"{name}.json").read_text(encoding="utf-8"))


@pytest.fixture
def without_title_edge(monkeypatch):
    """The recorded decks were judged before titles were held to the body's
    edge, and their tests are about the other checks. Which of their titles
    are off the body's edge is pinned in TestTitleEdge."""
    monkeypatch.setattr(layout_check, "_title_body_edge", lambda *args: [])


def findings_by_slide(deck):
    return {
        s["slide_index"]: [(f["type"], f["shapes"][0])
                           for f in slide_findings(s, deck["slide_width"], deck["slide_height"])]
        for s in deck["slides"]
    }


@pytest.mark.usefixtures("without_title_edge")
class TestARealDeck:
    SEEN = {
        1: [],                                   # cover: centred placeholders, corner logo
        2: [("misaligned", "TextBox 7")],         # agenda item 3, 2 pt low
        3: [("cramped_text", "TextBox 13")],      # the long KPI label on the card's edge
        4: [("misaligned", "Shape_234")],         # one card 2 pt low, with its contents
        5: [("off_slide", "TextBox 17"),          # the last timeline label, off the edge
            ("misaligned", "TextBox 17")],        # and 20 pt off its dot
        6: [],                                   # section band bleeding off both sides
        7: [],                                   # two columns
        8: [("overlap", "TextBox 3")],           # the caption into the table
        9: [],                                   # chart and takeaways
        10: [("misaligned", "TextBox 6")],       # a name 3 pt off its photo
        11: [],                                  # centred quote, corner logo
        12: [("low_contrast", "TextBox 10")],    # dark text on a dark card
        13: [],                                  # its title differs from the deck's
        14: [],                                  # closing slide, corner logo
    }

    def test_each_slide_reports_what_its_screenshot_shows(self):
        assert findings_by_slide(recorded("realdeck_broken_mac")) == self.SEEN

    def test_the_odd_title_is_reported_in_every_way_it_differs(self):
        deck = recorded("realdeck_broken_mac")
        found = deck_findings(deck["slides"], deck["slide_height"])
        assert {f["message"].split(":")[0] for f in found} == {
            f"Slide titles differ in {what}"
            for what in ("size", "left edge", "top edge", "weight", "colour")
        }
        assert all(f["slides"] == [13] for f in found)
        # Its frame was in place; its line was anchored at the bottom.
        [top] = [f for f in found if "top edge" in f["message"]]
        assert "vertical_anchor='top'" in top["fix"]

    def test_the_fixes_say_what_was_done(self):
        deck = recorded("realdeck_broken_mac")
        fixes = {}
        for s in deck["slides"]:
            for f in slide_findings(s, deck["slide_width"], deck["slide_height"]):
                fixes[(s["slide_index"], f["type"])] = f["fix"]
        # The KPI row grows as one.
        assert "shape_names=['Shape_225', 'Shape_222', 'Shape_223', 'Shape_224']" in fixes[(3, "cramped_text")]
        # The card moves with what is on it.
        assert "'Shape_235', 'TextBox 20', 'TextBox 21'" in fixes[(4, "misaligned")]
        # Centred on its dot, the label is shortened, never moved off it.
        assert fixes[(5, "off_slide")].startswith("Shorten the text of 'TextBox 17'")
        # The caption moves, not the table.
        assert fixes[(8, "overlap")].startswith("Move 'TextBox 3' down")
        # Body text takes the other body text's colour, not the headings'.
        assert "#CBD5E1" in fixes[(12, "low_contrast")]

    def test_the_fixed_deck_is_clean(self):
        deck = recorded("realdeck_fixed_mac")
        assert all(found == [] for found in findings_by_slide(deck).values())
        assert deck_findings(deck["slides"], deck["slide_height"]) == []


@pytest.mark.usefixtures("without_title_edge")
class TestASecondRealDeck:
    # Eleven more slides, the patterns the first deck did not have: a photo
    # filling half the slide, a pricing row with its middle plan featured,
    # process chevrons, a 2x2 matrix, an org chart on connectors, a column of
    # big figures, grouped icons, a full-bleed photo with text on it.
    SEEN = {
        1: [],                                          # photo left, text right
        2: [("misaligned", "TextBox 4")],               # one price 2 pt low on its card
        3: [("overflow", "Shape_283"),                  # the last step's text in its chevron
            ("narrow_margin", "Shape_283")],            # and the row 5 pt short of the margin
        4: [("misaligned", "TextBox 7")],               # one quadrant label 3 pt off
        5: [("misaligned", "Shape_287")],               # one report 3 pt low
        6: [("uneven_spacing", "TextBox 2"),            # the figures 30 and 42 pt apart
            ("uneven_spacing", "TextBox 3")],           # and their labels with them
        7: [],                                          # bullets that fit their box
        8: [],                                          # white text on a full-bleed photo
        9: [("uneven_spacing", "Group 4")],             # grouped icons 100 and 108 pt apart
        10: [("sample_text", "TextBox 2"),              # lorem ipsum
             ("sample_text", "TextBox 4")],             # and XX%
        11: [("overlap", "TextBox 2")],                 # a subtitle wrapped into the cards
    }

    def test_each_slide_reports_what_its_screenshot_shows(self):
        assert findings_by_slide(recorded("realdeck2_broken_mac")) == self.SEEN

    def test_no_title_is_out_of_place(self):
        # The photo slide's title sits beside the photo, where it has to; the
        # figures slide's big numbers are not its title.
        deck = recorded("realdeck2_broken_mac")
        assert deck_findings(deck["slides"], deck["slide_height"], deck["slide_width"]) == []

    def test_the_fixes_say_what_was_done(self):
        deck = recorded("realdeck2_broken_mac")
        fixes = {}
        for s in deck["slides"]:
            for f in slide_findings(s, deck["slide_width"], deck["slide_height"]):
                fixes.setdefault((s["slide_index"], f["type"]), f["fix"])
        # The price that sits differently on its card moves, not the other.
        assert fixes[(2, "misaligned")].startswith("Set top of 'TextBox 4'")
        # A step of a process is shortened, or the whole row grows.
        assert fixes[(3, "overflow")].startswith("Cut the text or make the font smaller")
        # Grouped icons set in from the sides keep their ends.
        assert fixes[(9, "uneven_spacing")].startswith("ppt_distribute_shapes")
        # Moving the subtitle up would jam it under the title: the cards move.
        assert fixes[(11, "overlap")].startswith("Move 'Shape_294', 'Shape_295', 'Shape_296' down")

    def test_the_fixed_deck_is_clean(self):
        deck = recorded("realdeck2_fixed_mac")
        assert all(found == [] for found in findings_by_slide(deck).values())
        assert deck_findings(deck["slides"], deck["slide_height"], deck["slide_width"]) == []


def fixes_of(deck):
    fixes = {}
    for s in deck["slides"]:
        for f in slide_findings(s, deck["slide_width"], deck["slide_height"]):
            fixes.setdefault((s["slide_index"], f["type"]), f["fix"])
    return fixes


@pytest.mark.usefixtures("without_title_edge")
class TestAThemedDeck:
    # Ten slides on PowerPoint's Badge theme, built through its layouts and
    # placeholders: a dark band down the left edge and a yellow strip down the
    # right, both on the slide master, and a title font of its own. The title,
    # section and caption layouts draw a design of their own, which PowerPoint
    # for Mac cannot read.
    SEEN = {
        1: [],                                          # title slide
        2: [],                                          # agenda in the body placeholder
        3: [("empty_placeholder", "Text Placeholder 4")],  # section header, text left empty
        4: [("overlap", "TextBox 4")],                  # a source note on the dark band
        5: [],                                          # comparison layout
        6: [],                                          # ten bullets, shrunk to fit
        7: [("overlap", "TextBox 1")],                  # a hand-drawn title on the band
        8: [],                                          # caption layout: light text on its own panel
        9: [("overlap", "Shape_2"),                     # the last card over the yellow strip
            ("uneven_sizes", "Shape_0")],               # and wider than the other two
        10: [],                                         # closing section header
    }

    def test_each_slide_reports_what_its_screenshot_shows(self):
        assert findings_by_slide(recorded("themedeck_broken_mac")) == self.SEEN

    def test_the_hand_drawn_title_is_the_only_odd_one(self):
        # The caption layout's small yellow title in its side column is the
        # theme's, and is not held to the others.
        deck = recorded("themedeck_broken_mac")
        found = deck_findings(deck["slides"], deck["slide_height"], deck["slide_width"])
        assert {f["message"].split(":")[0] for f in found} == {
            f"Slide titles differ in {what}"
            for what in ("size", "left edge", "weight", "font", "colour")
        }
        assert all(f["slides"] == [7] for f in found)

    def test_the_fixes_say_what_was_done(self):
        fixes = fixes_of(recorded("themedeck_broken_mac"))
        # Clear of the band and in line with the title, not just clear of it.
        assert fixes[(4, "overlap")] == (
            "Move 'TextBox 4' right by 78.6 pt, in line with 'Title 2' "
            "(ppt_update_shape dleft=78.56), clear of 'Freeform 6'.")
        # A card in a row takes the row's width back rather than moving.
        assert "width=255" in fixes[(9, "overlap")] and "like the rest of its row" in fixes[(9, "overlap")]

    def test_the_fixed_deck_is_clean(self):
        deck = recorded("themedeck_fixed_mac")
        assert all(found == [] for found in findings_by_slide(deck).values())
        assert deck_findings(deck["slides"], deck["slide_height"], deck["slide_width"]) == []


@pytest.mark.usefixtures("without_title_edge")
class TestACompanyTemplateDeck:
    # Ten slides on a company-style template: a navy bar across the top, an
    # ACME logo bottom right and a confidentiality line bottom left, on the
    # master; Georgia titles, Arial text, layouts of its own.
    SEEN = {
        1: [], 2: [], 3: [],
        4: [("overlap", "Chart 3")],                    # a chart over the logo and the line
        5: [("overlap", "TextBox 12")],                 # a footnote on the confidentiality line
        6: [], 7: [],                                   # (7: its title is the deck's finding)
        8: [],                                          # caption layout, its own small title
        9: [("overlap", "TextBox 4")],                  # a contact line on the logo
        10: [],
    }

    def test_each_slide_reports_what_its_screenshot_shows(self):
        assert findings_by_slide(recorded("companydeck_broken_mac")) == self.SEEN

    def test_only_the_hand_drawn_title_is_odd(self):
        deck = recorded("companydeck_broken_mac")
        found = deck_findings(deck["slides"], deck["slide_height"], deck["slide_width"])
        assert found and all(f["slides"] == [7] for f in found)

    def test_the_fixes_say_what_was_done(self):
        fixes = fixes_of(recorded("companydeck_broken_mac"))
        # The chart is trimmed above the footer rather than pushed into the title.
        assert fixes[(4, "overlap")].startswith("Make 'Chart 3' 394 pt tall (ppt_update_shape height=394)")
        # A footnote steps up over the line, not across the slide.
        assert fixes[(5, "overlap")].startswith("Move 'TextBox 12' up by 13 pt")
        assert fixes[(9, "overlap")].startswith("Move 'TextBox 4' up by 22 pt")

    def test_the_fixed_deck_is_clean(self):
        # Fixed, the hand-drawn title matches the others, and the caption
        # layout's own title is still not held to them.
        deck = recorded("companydeck_fixed_mac")
        assert all(found == [] for found in findings_by_slide(deck).values())
        assert deck_findings(deck["slides"], deck["slide_height"], deck["slide_width"]) == []


@pytest.mark.usefixtures("without_title_edge")
class TestACompanyTemplateDeckOnWindows:
    # The company-style template again, built on Windows through layouts
    # picked by name ("Acme Content", "Acme Section"), with a photo, icons and
    # one more mistake: a mail icon dropped on the logo.
    SEEN = {
        1: [], 2: [], 3: [],
        4: [("overlap", "Chart 2")],                    # a chart over the logo and the line
        5: [("overlap", "TextBox 11")],                 # a footnote on the confidentiality line
        6: [], 7: [], 8: [],
        9: [("overlap", "TextBox 3")],                  # a contact line on the logo
        10: [],
        11: [("overlap", "Graphic 15")],                # an icon on the logo; white text on the photo is fine
    }

    def test_each_slide_reports_what_its_preview_shows(self):
        assert findings_by_slide(recorded("companydeck_broken_win")) == self.SEEN

    def test_the_fixes_say_what_was_done(self):
        fixes = fixes_of(recorded("companydeck_broken_win"))
        assert fixes[(4, "overlap")].startswith("Make 'Chart 2' 394 pt tall")
        assert fixes[(5, "overlap")].startswith("Move 'TextBox 11' up by 13 pt")
        assert fixes[(11, "overlap")].startswith("Move 'Graphic 15' up by 16 pt")

    def test_only_the_hand_drawn_title_is_odd_and_the_fixed_deck_is_clean(self):
        deck = recorded("companydeck_broken_win")
        found = deck_findings(deck["slides"], deck["slide_height"], deck["slide_width"])
        assert found and all(f["slides"] == [7] for f in found)
        fixed = recorded("companydeck_fixed_win")
        assert all(found == [] for found in findings_by_slide(fixed).values())
        assert deck_findings(fixed["slides"], fixed["slide_height"], fixed["slide_width"]) == []


@pytest.mark.usefixtures("without_title_edge")
class TestAThemedDeckOnWindows:
    # PowerPoint's Retrospect theme on Windows, where a layout's own graphics
    # can be read: a footer band and a line above it on the master, and a dark
    # panel down the left of the caption layout, a band under the picture one.
    SEEN = {
        1: [], 2: [], 3: [],
        4: [("overlap", "Chart 2")],                    # a chart into the footer band
        5: [],                                          # white caption text on the layout's panel
        6: [("low_contrast", "TextBox 3")],             # a grey note on the footer band
        7: [],                                          # picture layout, its own band
        8: [("overlap", "Rounded Rectangle 2"),         # a card into the band
            ("narrow_margin", "Rounded Rectangle 2")],  # and 5 pt past the margin
    }

    def test_each_slide_reports_what_its_preview_shows(self):
        assert findings_by_slide(recorded("retrodeck_broken_win")) == self.SEEN

    def test_the_layouts_own_panel_is_read(self):
        deck = recorded("retrodeck_broken_win")
        caption = deck["slides"][4]
        assert any(g["source"] == "layout" and g["width"] > 300 for g in caption["template"])

    def test_text_put_on_the_band_is_moved_off_it_before_it_is_recoloured(self):
        deck = recorded("retrodeck_broken_win")
        [found] = slide_findings(deck["slides"][5], deck["slide_width"], deck["slide_height"])
        assert "It sits on 'Rectangle 6', part of the template." in found["message"]
        # The thin line on top of the band is part of it: clear of both.
        assert found["fix"].startswith("Move 'TextBox 3' up by 33.2 pt")
        assert "'Rectangle 6' and 'Rectangle 8'" in found["fix"]
        assert "Or make the text #000000" in found["fix"]

    def test_the_hand_drawn_title_is_odd_and_the_fixed_deck_is_clean(self):
        deck = recorded("retrodeck_broken_win")
        found = deck_findings(deck["slides"], deck["slide_height"], deck["slide_width"])
        assert len(found) == 5 and all(f["slides"] == [8] for f in found)
        fixed = recorded("retrodeck_fixed_win")
        assert all(found == [] for found in findings_by_slide(fixed).values())
        assert deck_findings(fixed["slides"], fixed["slide_height"], fixed["slide_width"]) == []


@pytest.mark.usefixtures("without_title_edge")
class TestARightToLeftDeckOnWindows:
    # Eight Hebrew slides built on Windows, titles and text set right, as a
    # model writing Hebrew sets them: a title's right edge stays put from slide
    # to slide while its left edge moves with the length of the words.
    SEEN = {
        1: [], 2: [], 3: [], 4: [], 5: [],
        6: [],                                          # its title is set left: the deck's finding
        7: [("uneven_spacing", "Rounded Rectangle 6")],  # cards 32 and 20 pt apart
        8: [],
    }

    def test_each_slide_reports_what_its_preview_shows(self):
        assert findings_by_slide(recorded("rtldeck_broken_win")) == self.SEEN

    def test_right_set_titles_are_compared_on_their_right_edge(self):
        deck = recorded("rtldeck_broken_win")
        [found] = deck_findings(deck["slides"], deck["slide_height"], deck["slide_width"])
        assert found["message"] == ("Slide titles differ in alignment: right on slides 2, 3, 4, 5, 7, 8; "
                                    "left on slide 6.")
        assert found["fix"].endswith("(ppt_set_paragraph_format alignment='right').")

    def test_the_cards_move_with_their_labels(self):
        fixes = fixes_of(recorded("rtldeck_broken_win"))
        assert "shape_names=['Rounded Rectangle 6', 'TextBox 7'] dleft=-18" in fixes[(7, "uneven_spacing")]

    def test_the_fixed_deck_is_clean(self):
        fixed = recorded("rtldeck_fixed_win")
        assert all(found == [] for found in findings_by_slide(fixed).values())
        assert deck_findings(fixed["slides"], fixed["slide_height"], fixed["slide_width"]) == []


@pytest.mark.usefixtures("without_title_edge")
class TestARightToLeftDeckOnMac:
    # The Hebrew deck again, on PowerPoint for Mac, with a ninth slide on the
    # layout's own title and body: a two-line title in the title placeholder
    # among titles drawn as boxes, and fourteen bullets PowerPoint shrank to
    # fit. PowerPoint for Mac reports that title's lines 11 pt wider than it
    # draws them, past both of the frame's insets.
    SEEN = {
        1: [], 2: [], 3: [], 4: [], 5: [], 6: [],
        7: [("uneven_spacing", "Shape_5")],             # cards 32 and 20 pt apart, as on Windows
        8: [],
        9: [],                                          # the title fits; the bullets were shrunk
    }

    def test_each_slide_reports_what_its_screenshot_shows(self):
        assert findings_by_slide(recorded("rtldeck_broken_mac")) == self.SEEN

    def test_the_first_eight_slides_agree_with_windows(self):
        def found(name):
            deck = recorded(name)
            return {s["slide_index"]: kinds(slide_findings(s, deck["slide_width"], deck["slide_height"]))
                    for s in deck["slides"] if s["slide_index"] <= 8}
        assert found("rtldeck_broken_mac") == found("rtldeck_broken_win")

    def test_the_placeholder_title_is_held_to_the_drawn_ones(self):
        deck = recorded("rtldeck_broken_mac")
        found = deck_findings(deck["slides"], deck["slide_height"], deck["slide_width"])
        assert {f["message"].split(":")[0] for f in found} == {
            f"Slide titles differ in {what}"
            for what in ("size", "right edge", "alignment", "weight", "font", "colour")
        }
        assert all(f["slides"] == ([6] if "alignment" in f["message"] else [9]) for f in found)

    def test_the_fixed_deck_is_clean(self):
        fixed = recorded("rtldeck_fixed_mac")
        assert all(found == [] for found in findings_by_slide(fixed).values())
        assert deck_findings(fixed["slides"], fixed["slide_height"], fixed["slide_width"]) == []


class TestGradientsAndThemesChangedByHand:
    # Text on gradients: on Windows, Slice's background at 102 degrees and a
    # card filled black to white with ppt_set_fill; on the Mac, a dark
    # background gradient and a pale card, read without an angle. And a deck
    # whose light1 was set dark with ppt_set_theme_colors, on both.
    def test_each_slide_reports_what_its_preview_shows(self):
        assert findings_by_slide(recorded("gradientdeck_broken_win")) == {
            1: [("low_contrast", "TextBox 1"),          # a grey note on the deep blue
                ("low_contrast", "TextBox 3")],         # white on the pale top
            2: [("low_contrast", "TextBox 2")],         # dark words on the card's black end
        }
        assert findings_by_slide(recorded("gradientdeck_broken_mac")) == {
            1: [("low_contrast", "TextBox 1")],         # dark grey on a dark gradient
            2: [("low_contrast", "TextBox 2")],         # white on a pale one
        }

    def test_windows_reads_the_angle_and_the_mac_cannot(self):
        win = recorded("gradientdeck_broken_win")["slides"][0]["background_gradient"]
        mac = recorded("gradientdeck_broken_mac")["slides"][0]["background_gradient"]
        assert win["angle"] == 102.0 and mac["angle"] is None

    def test_the_title_of_a_theme_changed_by_hand_is_reported_on_both(self):
        for name in ("handthemedeck_broken_win", "handthemedeck_broken_mac"):
            [found] = slide_findings(recorded(name)["slides"][0], W, H)
            assert found["type"] == "low_contrast" and "the theme's dark1" in found["fix"]

    def test_the_fixed_decks_are_clean(self):
        for name in ("gradientdeck_fixed_win", "gradientdeck_fixed_mac",
                     "handthemedeck_fixed_win", "handthemedeck_fixed_mac"):
            fixed = recorded(name)
            assert all(found == [] for found in findings_by_slide(fixed).values()), name
            assert deck_findings(fixed["slides"], fixed["slide_height"], fixed["slide_width"]) == []


@pytest.mark.usefixtures("without_title_edge")
class TestADeckOnTwoDesigns:
    # Office Theme on slides 1-3 and Retrospect on 4-5, on Windows. Each design
    # has a "Title Only" layout of its own, and each styles its title its own
    # way: 44 pt black at 73 pt in, and 48 pt grey at 94 pt in, lower.
    SEEN = {
        1: [], 2: [], 3: [], 4: [],
        5: [("low_contrast", "TextBox 3")],             # a grey note on Retrospect's band
    }

    def test_each_slide_reports_what_its_preview_shows(self):
        assert findings_by_slide(recorded("mmdeck_broken_win")) == self.SEEN

    def test_each_designs_titles_are_held_only_to_their_own(self):
        deck = recorded("mmdeck_broken_win")
        assert {s["design"] for s in deck["slides"]} == {"Office Theme", "Retrospect"}
        assert deck_findings(deck["slides"], deck["slide_height"], deck["slide_width"]) == []

    def test_the_fixed_deck_is_clean(self):
        fixed = recorded("mmdeck_fixed_win")
        assert all(found == [] for found in findings_by_slide(fixed).values())


@pytest.mark.usefixtures("without_title_edge")
class TestThemesDrawnInLinesOnWindows:
    # Slice, a fan of thin lines in the bottom right corner, and Wisp, curves
    # down the left: both on the master, as groups whose boxes are mostly
    # empty, and both themes start their title and body clear of them.
    SLICE = {
        1: [], 2: [],
        3: [("overlap", "Rounded Rectangle 6")],        # a row of cards past the content area
        4: [("overlap", "TextBox 4"),                   # a source note under the lines' tips
            ("misaligned", "TextBox 3")],               # a note 1.2 pt off the text beside it
        5: [],
    }
    WISP = {
        1: [], 2: [], 3: [],
        4: [("overlap", "TextBox 2")],                  # text set over the curves
        5: [],
    }

    def test_each_slide_reports_what_its_preview_shows(self):
        assert findings_by_slide(recorded("slicedeck_broken_win")) == self.SLICE
        assert findings_by_slide(recorded("wispdeck_broken_win")) == self.WISP

    def test_text_over_the_curves_moves_in_line_with_the_title(self):
        fixes = fixes_of(recorded("wispdeck_broken_win"))
        assert fixes[(4, "overlap")].startswith("Move 'TextBox 2' right by 164.2 pt, in line with 'Title 1'")

    def test_the_row_narrows_evenly_to_keep_inside_the_content_area(self):
        fixes = fixes_of(recorded("slicedeck_broken_win"))
        assert fixes[(3, "overlap")].startswith("Make the row fit before 'Group 6'")
        assert fixes[(3, "overlap")].endswith("The gaps stay 20 pt.")

    def test_the_fixed_decks_are_clean(self):
        for name in ("slicedeck_fixed_win", "wispdeck_fixed_win"):
            fixed = recorded(name)
            assert all(found == [] for found in findings_by_slide(fixed).values()), name
            assert deck_findings(fixed["slides"], fixed["slide_height"], fixed["slide_width"]) == []


# ---------------------------------------------------------------------------
# What text keeps to
# ---------------------------------------------------------------------------
def placed(name, left, top, width, height, text_left, text_top, text_width, text_height, z=1):
    """A text box whose text PowerPoint drew at the given place."""
    box = textbox(name, left, top, width, height, z=z)
    box["text"].update(left=text_left, top=text_top, width=text_width, height=text_height)
    return box


class TestWhereTextKeeps:
    def test_left_set_text_hugs_the_left_inset(self):
        title = placed("t", 40, 30, 880, 46, 47.2, 33.6, 195, 38.4)
        assert layout_check._text_anchor(title) == ("left", "full")

    def test_centred_text_is_centred_even_a_few_points_off(self):
        # PowerPoint for Mac reported this label 3 pt off the centre it was set to.
        label = placed("l", 0, 300, 160, 23.02, 33.47, 303.6, 96.07, 15.6)
        assert layout_check._text_anchor(label) == ("center", "full")

    def test_right_set_text_hugs_the_right_inset(self):
        note = placed("n", 160, 320, 640, 30, 600, 323.6, 192.8, 22.8)
        assert layout_check._text_anchor(note)[0] == "right"

    def test_text_at_the_top_of_a_tall_frame(self):
        body = placed("b", 40, 100, 400, 300, 47.2, 103.6, 200, 60)
        assert layout_check._text_anchor(body) == ("left", "top")

    def test_a_reader_sees_only_the_edges_text_keeps_to(self):
        label = placed("l", 0, 300, 160, 23.02, 33.47, 303.6, 96.07, 15.6)
        assert layout_check._seen_edges(label) == {"center", "top", "middle", "bottom"}
        assert layout_check._seen_edges(card("c", 0, 0, 10, 10)) is None


# ---------------------------------------------------------------------------
# Margins, from more kinds of slide
# ---------------------------------------------------------------------------
class TestWhatAMarginIs:
    def test_centred_text_keeps_no_margin(self):
        # Live: a centred quote 160 pt in from both sides made a "160 pt
        # margin", and the corner logo was reported against it.
        quote = placed("quote", 160, 150, 640, 160, 230, 153.6, 500, 70)
        attribution = placed("by", 160, 320, 640, 30, 600, 323.6, 192.8, 22.8, z=2)
        logo = card("logo", 840, 468, 80, 32, z=3)
        assert check(quote, attribution, logo) == []

    def test_a_cover_on_its_placeholders_has_no_margin_to_break(self):
        title = placed("title", 120, 88, 720, 188, 290, 150, 380, 70)
        title.update(type=msoPlaceholder, placeholder_type=3)
        logo = card("logo", 840, 468, 80, 32, z=2)
        assert check(title, logo) == []

    def test_a_margin_is_kept_on_both_sides(self):
        # Live, the agenda: four circles 60 pt in on the left made 60 pt the
        # most common distance, and the title, 40 pt in, was reported.
        title = placed("title", 40, 30, 880, 46, 47.2, 33.6, 120, 38.4)
        dots = [card(f"dot{i}", 60, 130 + 80 * i, 44, 44, z=2 + i) for i in range(4)]
        layout = layout_check._layout([title, *dots], W, H)
        assert layout["margins"]["horizontal"] == 40.0
        assert "narrow_margin" not in kinds(check(title, *dots))


# ---------------------------------------------------------------------------
# Labels
# ---------------------------------------------------------------------------
class TestLabels:
    def test_a_label_off_the_edge_is_narrowed_around_its_centre(self):
        # Centred on its dot at 820 pt, 300 pt of text runs 10 pt off the edge.
        dot = card("dot", 810, 270, 20, 20, z=8)
        label = placed("label", 620, 300, 400, 23, 670, 303.6, 300, 15.6, z=9)
        found = check(*titled_row(heights=(100, 100, 100)), dot, label)
        [finding] = [f for f in found if f["type"] == "off_slide"]
        assert finding["fix"].startswith("Narrow 'label' to ")
        assert "around its centre" in finding["fix"]

    def test_a_label_well_off_its_dot_is_reported(self):
        dot = card("dot", 870, 270, 20, 20, z=1)
        label = placed("label", 780 + 20, 300, 200, 23, 821.5, 303.6, 160, 15.6, z=2)
        [finding] = [f for f in check(dot, label) if f["type"] == "misaligned"]
        assert finding["shapes"] == ["label", "dot"]
        assert "left=780" in finding["fix"]

    def test_a_centred_name_is_compared_by_its_centre_only(self):
        # Live: a name 3 pt off its photo was also reported against the
        # left edge of the frame under it, which no one can see.
        photo = card("photo", 310, 120, 110, 110, z=1)
        name = placed("name", 288, 245, 160, 29, 326.2, 248.6, 87.8, 21.6, z=2)
        role = placed("role", 285, 275, 160, 24, 343, 278.6, 47.1, 16.8, z=3)
        [finding] = [f for f in check(photo, name, role) if f["type"] == "misaligned"]
        assert finding["edge"] == "center"
        assert set(finding["shapes"]) == {"name", "photo", "role"}


# ---------------------------------------------------------------------------
# What else a real deck gets wrong
# ---------------------------------------------------------------------------
class TestSpacingIsForShapesThatPaint:
    def test_text_beside_a_row_of_icons_is_not_spaced(self):
        dots = [card(f"dot{i}", 60, 130 + 80 * i, 44, 44, z=1 + i) for i in range(4)]
        items = [textbox(f"item{i}", 124, 134 + 80 * i + (3 if i == 2 else 0), 600, 34,
                         z=5 + i, text_width=200, text_height=26.8) for i in range(4)]
        assert "uneven_spacing" not in kinds(check(*dots, *items))


class TestCollisions:
    def test_the_smaller_shape_moves_and_clears(self):
        table = shape("table", 40, 110, 880, 250, z=1, type=msoTable)
        caption = textbox("caption", 40, 350, 880, 30, z=2, text_width=160, text_height=15)
        [finding] = [f for f in check(table, caption) if f["type"] == "overlap"]
        assert finding["fix"].startswith("Move 'caption' down by ")
        assert "6 pt between it and 'table'" in finding["fix"]


class TestRightToLeftText:
    # PowerPoint reports the lines of Hebrew text wider than it draws them: a
    # right-set title to the frame's edge rather than its inset on Windows, a
    # two-line placeholder title 11 pt past both insets on the Mac.
    def test_wrapped_text_is_held_to_its_frames_insets(self):
        title = placed("Title 3", 66, 28.75, 828, 104.38, 62.19, 33.42, 830.26, 95.04)
        title["text"]["word_wrap"] = True
        assert layout_check.text_box(title) == (73.2, 33.42, 886.8, 128.46)

    def test_so_a_placeholder_title_keeps_the_margin_it_is_drawn_at(self):
        title = placed("Title 3", 66, 28.75, 828, 104.38, 62.19, 33.42, 830.26, 95.04)
        body = placed("Content 4", 66, 143.75, 828, 342.63, 91.2, 147.35, 290.51, 307.1, z=2)
        assert [f for f in check(title, body) if f["type"] == "narrow_margin"] == []

    def test_text_that_does_not_wrap_is_where_it_is_reported(self):
        line = placed("line", 66, 30, 300, 40, 60, 33.6, 320, 30)
        line["text"]["word_wrap"] = False
        assert layout_check.text_box(line)[0] == 60

    def test_a_margin_fix_moves_the_frame_by_what_its_text_needs(self):
        # Text poking 4 pt out of the left of a frame set at the margin.
        words = placed("words", 66, 200, 300, 40, 62, 203.6, 200, 30)
        words["text"]["word_wrap"] = False
        other = placed("other", 66, 100, 828, 40, 73.2, 103.6, 300, 30, z=2)
        [found] = [f for f in check(words, other) if f["type"] == "narrow_margin"]
        assert found["fix"].startswith("Move 'words' right by")
        assert "by 0 pt" not in found["fix"]


class TestATableThatGrows:
    # Live: two rows added to a table took it down over its footnote, which
    # then lay on its cells. The footnote was reported as pressed against the
    # table's edge and as grey on black, the fill a table reports for itself.
    def slide(self):
        table = shape("Table 1", 60, 300, 600, 231, z=1, type=msoTable,
                      fill={"visible": True, "solid": True, "color_hex": "#000000",
                            "transparency": 0.0})
        note = textbox("TextBox 2", 60, 470, 400, 24, z=2, text_width=150, text_height=15,
                       runs=[{"color_hex": "#595959", "size": 12.0, "bold": False}])
        return table, note

    def test_text_over_its_cells_runs_into_it(self):
        found = check(*self.slide())
        assert [(f["type"], f["shapes"]) for f in found] == [("overlap", ["TextBox 2", "Table 1"])]

    def test_its_write_reports_it(self):
        found = slide_findings(slide(*self.slide()), W, H, only={"Table 1"})
        assert [f["type"] for f in found] == ["overlap"]

    def test_the_fix_clears_the_whole_table(self):
        [found] = check(*self.slide())
        assert found["message"] == "The text of 'TextBox 2' lies on 'Table 1', over its cells."
        # Down past the table would take it off the slide, so the table moves
        # up, far enough to clear the footnote, and nothing says to shorten it.
        assert found["fix"] == ("Move 'Table 1' up by 63.4 pt (ppt_update_shape shape_names=['Table 1'] "
                                "dtop=-63.4), which leaves 6 pt between it and the text of 'TextBox 2'.")


class TestTextUnderACard:
    def test_a_card_added_after_its_words_hides_them(self):
        words = textbox("label", 76, 136, 228, 31, z=1, text_height=22)
        cover = card("card", 60, 120, 260, 140, z=2, colour="#1F2937")
        [found] = check(words, cover)
        assert found["message"] == "The text of 'label' is hidden under 'card', which is in front of it."
        assert found["fix"].startswith("Bring 'label' to the front (ppt_set_shape_zorder")

    def test_words_on_a_card_are_fine(self):
        assert check(card("card", 60, 120, 260, 140, z=1),
                     textbox("label", 76, 136, 228, 31, z=2, text_height=22)) == []

    def test_a_see_through_shape_in_front_hides_nothing(self):
        words = textbox("label", 76, 136, 228, 31, z=1, text_height=22)
        frame = shape("frame", 60, 120, 260, 140, z=2, line=True,
                      fill={"visible": False, "solid": True, "color_hex": "#FFFFFF",
                            "transparency": 0.0})
        assert [f for f in check(words, frame) if f["type"] == "overlap"] == []


class TestAFillUnderText:
    def slide(self, colour):
        words = textbox("label", 76, 136, 228, 31, z=2,
                        runs=[{"color_hex": "#1F2937", "size": 20.0, "bold": False}])
        return slide(card("card", 60, 120, 260, 140, z=1, colour=colour), words)

    def test_a_card_filled_dark_under_dark_words_reports_them(self):
        found = slide_findings(self.slide("#1F2937"), W, H, only={"card"})
        assert [(f["type"], f["shapes"]) for f in found] == [("low_contrast", ["label"])]
        assert slide_findings(self.slide("#F1F5F9"), W, H, only={"card"}) == []

    def test_the_words_on_a_touched_card_are_read_in_full(self):
        shapes = {"card": card("card", 60, 120, 260, 140, z=1),
                  "label": textbox("label", 76, 136, 228, 31, z=2),
                  "aside": textbox("aside", 400, 136, 228, 31, z=3)}
        read = {}

        def geometry(ref, z):
            return {k: v for k, v in shapes[ref].items() if k not in ("fill", "text")} | {
                "fill": None, "text": None}

        def appearance(facts, ref, full):
            read[ref] = full
            facts.update(fill=shapes[ref]["fill"], text=shapes[ref]["text"])
            return facts

        layout_check.read_shapes(["card", "label", "aside"], geometry, appearance, ["card"])
        assert read["card"] is True and read["label"] is True
        assert read.get("aside") in (None, False)


class TestAFontChange:
    # Live: after ppt_replace_font, a text box set to fit its text was read
    # before PowerPoint had grown it, and reported as overflowing.
    def test_a_box_that_fits_its_text_is_taken_at_its_new_height(self):
        box = textbox("words", 76, 136, 228, 24.3, text_width=200, text_height=48,
                      autofit="shape_to_fit")
        assert check(box) == []

    def test_what_it_grows_into_is_reported(self):
        box = textbox("words", 60, 40, 420, 87.3, text_width=400, text_height=129.6,
                      autofit="shape_to_fit")
        below = card("panel", 60, 110, 840, 300, z=2)
        found = check(box, below)
        assert [(f["type"], f["shapes"]) for f in found] == [("overlap", ["words", "panel"])]

    def test_a_box_that_does_not_grow_still_overflows(self):
        box = textbox("words", 76, 136, 228, 24.3, text_width=200, text_height=48)
        assert kinds(check(box)) == ["overflow"]


class TestTextInsideOtherText:
    def test_a_label_inside_a_paragraphs_lines_is_moved_clear_of_all_of_them(self):
        body = placed("body", 60, 120, 600, 200, 67.2, 123.6, 500, 150)
        label = placed("label", 200, 160, 200, 30, 207.2, 163.6, 120, 22, z=2)
        for order in ((body, label), (label, body)):
            [found] = [f for f in check(*order) if f["type"] == "overlap"]
            assert found["shapes"] == ["body", "label"]
            # 6 pt above the paragraph's first line, not 6 pt past the label's own height.
            assert "ppt_update_shape shape_names=['label'] dtop=-68" in found["fix"]


# Slice's background as PowerPoint for Windows reports it. Its render matched
# these colours, blended straight across in sRGB, to the unit.
SLICE = {"angle": 102.0, "stops": [[0.1, "#64D4EF", 0.0], [1.0, "#06588E", 0.0]]}
# Wisp's, a gradient from the centre: it has no angle.
WISP = {"angle": None, "stops": [[0.0, "#FFFFFF", 0.0], [1.0, "#DFE8C4", 0.0]]}


def on_gradient(*shapes, gradient=SLICE):
    page = slide(*shapes, background=None)
    page["background_gradient"] = gradient
    return page


def contrast(page):
    return [f for f in slide_findings(page, W, H) if f["type"] == "low_contrast"]


def words(name, left, top, colour, size=12.0, z=1):
    return textbox(name, left, top, 200, 24, z=z, text_width=150, text_height=15,
                   runs=[{"color_hex": colour, "size": size, "bold": False}])


class TestTextOnAGradient:
    def test_a_grey_note_on_the_deep_blue_is_reported(self):
        # Live: the source note at the foot of a Slice slide, all but unreadable.
        [found] = contrast(on_gradient(words("note", 740, 485, "#595959")))
        assert found["message"].startswith("Text in 'note' is #595959 on the background's gradient, ")
        assert "behind it, a contrast of as little as" in found["message"]
        assert "ppt_format_text color='#FFFFFF'" in found["fix"]

    def test_white_on_the_deep_blue_reads(self):
        assert contrast(on_gradient(words("title", 60, 470, "#FFFFFF", size=28.0))) == []

    def test_white_on_the_pale_top_does_not(self):
        [found] = contrast(on_gradient(words("title", 60, 30, "#FFFFFF", size=28.0)))
        assert found["shapes"] == ["title"]

    def test_a_card_filled_with_a_gradient_is_judged_where_the_words_are(self):
        # Black to white, left to right, as ppt_set_fill makes a two-colour fill.
        fill = {"visible": True, "solid": False, "color_hex": "#000000", "transparency": 0.0,
                "gradient": {"angle": 0.0, "stops": [[0.0, "#000000", 0.0], [1.0, "#FFFFFF", 0.0]]}}
        panel = shape("panel", 100, 100, 600, 200, z=1, fill=fill, line=False)
        dark_end = words("dark end", 110, 150, "#333333", size=20.0, z=2)
        light_end = words("light end", 490, 150, "#333333", size=20.0, z=3)
        assert [f["shapes"] for f in contrast(slide(panel, dark_end, light_end))] == [["dark end"]]

    def test_where_the_two_blends_disagree_nothing_is_reported(self):
        # A quarter of the way along black to white: about #404040 blended in
        # sRGB, about #707070 along the curve. White reads on the one, not the
        # other, so neither verdict can be stood behind.
        fill = {"visible": True, "solid": False, "color_hex": "#000000", "transparency": 0.0,
                "gradient": {"angle": 0.0, "stops": [[0.0, "#000000", 0.0], [1.0, "#FFFFFF", 0.0]]}}
        panel = shape("panel", 0, 100, 800, 200, z=1, fill=fill, line=False)
        assert contrast(slide(panel, words("w", 190, 150, "#FFFFFF", z=2))) == []

    def test_without_an_angle_only_text_unreadable_anywhere_on_it_is_reported(self):
        page = on_gradient(words("white", 300, 200, "#FFFFFF", size=28.0),
                           words("grey", 300, 300, "#595959", z=2), gradient=WISP)
        [found] = contrast(page)
        assert found["shapes"] == ["white"] and "anywhere on it" in found["message"]

    def test_a_see_through_gradient_is_no_verdict(self):
        clear = {"angle": 90.0, "stops": [[0.0, "#000000", 0.6], [1.0, "#000000", 0.6]]}
        assert contrast(on_gradient(words("w", 300, 200, "#111111"), gradient=clear)) == []


OFFICE = {"dark1": "#000000", "light1": "#FFFFFF", "dark2": "#0E2841", "light2": "#E8E8E8"}


class TestAThemeChangedByHand:
    # ppt_set_theme_colors light1='#111827': the background goes dark, the
    # titles stay black in dark1, and a check run later should still see it.
    def page(self, theme):
        title = title_placeholder(colour="#000000")
        page = templated(title, background="#111827")
        page["theme_colors"] = theme
        return page

    def test_a_theme_whose_text_colour_does_not_read_on_its_background_is_not_trusted(self):
        [found] = contrast(self.page({**OFFICE, "light1": "#111827"}))
        assert found["shapes"] == ["Title 1"]
        assert found["fix"].startswith(
            "Its colour is the theme's, and the theme's dark1, #000000, is 1.2:1 on its "
            "light1, #111827: choose theme colours that read on each other (ppt_set_theme_colors)")

    def test_a_theme_as_shipped_is_still_trusted(self):
        # Whatever is drawn behind the title, the template designed it.
        assert contrast(self.page(OFFICE)) == []
        assert contrast(self.page(None)) == []


class TestContrastSuggestions:
    def test_body_text_takes_the_colour_of_body_text(self):
        headings = [textbox(f"h{i}", 60 + 300 * i, 150, 240, 36, z=1 + i,
                            runs=[{"color_hex": "#F8FAFC", "size": 20.0, "bold": True}])
                    for i in range(3)]
        bodies = [textbox(f"b{i}", 60 + 300 * i, 196, 240, 100, z=4 + i,
                          runs=[{"color_hex": "#CBD5E1", "size": 16.0, "bold": False}])
                  for i in range(2)]
        dim = textbox("dim", 660, 196, 240, 100, z=9,
                      runs=[{"color_hex": "#475569", "size": 16.0, "bold": False}])
        found = check(*headings, *bodies, dim, background="#0F172A")
        [finding] = [f for f in found if f["type"] == "low_contrast"]
        assert "color='#CBD5E1'" in finding["fix"]


class TestDeckTitles:
    @staticmethod
    def titled_with(index, runs, background="#FFFFFF"):
        title = textbox(f"Title {index}", 40, 30, 880, 60, runs=runs)
        return slide(title, index=index, background=background)

    def test_a_title_in_another_typeface_is_reported(self):
        body = [{"color_hex": "#1F2937", "size": 32.0, "bold": True, "font": "Calibri"}]
        light = [{"color_hex": "#1F2937", "size": 32.0, "bold": True, "font": "Calibri Light"}]
        slides = [self.titled_with(i, body) for i in (2, 3, 4)] + [self.titled_with(5, light)]
        [finding] = deck_findings(slides, H)
        assert "font_name='Calibri'" in finding["fix"]

    def test_a_light_title_on_a_dark_slide_is_meant(self):
        dark = [{"color_hex": "#1F2937", "size": 32.0, "bold": True}]
        light = [{"color_hex": "#F8FAFC", "size": 32.0, "bold": True}]
        slides = [self.titled_with(i, dark) for i in (2, 3, 4)]
        slides.append(self.titled_with(5, light, background="#0F172A"))
        assert deck_findings(slides, H) == []

    def test_a_title_moving_takes_the_body_lined_up_with_it(self):
        def page(index, left):
            words = [{"color_hex": "#222222", "size": 32.0, "bold": True}]
            title = textbox(f"Title {index}", left, 30, 840, 60, runs=words)
            body = textbox(f"Body {index}", left, 140, 840, 300, z=2)
            return slide(title, body, index=index)
        slides = [page(2, 40), page(3, 40), page(4, 66)]
        [finding] = deck_findings(slides, H)
        assert "shape_names=['Title 4', 'Body 4'] dleft=-26" in finding["fix"]


class TestSampleText:
    @pytest.mark.parametrize("words, sample", [
        ("Lorem ipsum dolor sit amet", "Lorem ipsum"),
        ("Click to add text", "Click to add"),
        ("Your title here", "Your title"),
        ("Prepared for [Client name]", "[Client name]"),
        ("Intro\nYour company here", "Your company"),
        ("[Insert date]", "[Insert date]"),
        ("Revenue grew XX% this year", "XX%"),
        ("Margin of XX.X points", "XX.X"),
    ])
    def test_stand_in_text_is_reported(self, words, sample):
        box = textbox("t", 100, 100, 400, 40)
        box["text"]["content"] = words
        [finding] = check(box)
        assert finding["type"] == "sample_text"
        assert finding["sample"] == sample

    @pytest.mark.parametrize("words", [
        "Super Bowl XX", "Sizes up to XXL", "See note [1]", "Our text here is final",
        "The name of the game", "Grow your company", "Protect your image online",
        "Keep your content fresh", "[Number of users]", "[Name 2020]",
        "Your company is growing",
    ])
    def test_real_text_is_not(self, words):
        box = textbox("t", 100, 100, 400, 40)
        box["text"]["content"] = words
        assert check(box) == []


def worded(name, top, content):
    box = textbox(name, 100, top, 600, 40)
    box["text"]["content"] = content
    return box


class TestNumberOrder:
    """Found on the Hebrew history deck: "1516 – 1917" under a Hebrew heading
    was drawn with 1516 on the left, so it read 1917 to 1516. Live on both
    platforms, a line of numbers only draws left to right even when its
    paragraph is set right to left."""

    def test_a_range_of_numbers_on_a_hebrew_slide_is_reported(self):
        [finding] = check(worded("heading", 40, "השלטון העות'מאני"),
                          worded("years", 100, "1516 – 1917"))
        assert finding["type"] == "number_order"
        assert finding["shapes"] == ["years"]
        assert finding["line"] == "1516 – 1917"
        assert "1516 עד 1917" in finding["fix"]
        assert 'find_text="1516 – 1917"' in finding["fix"]
        assert "ppt_set_paragraph_format direction='rtl'" in finding["fix"]

    def test_a_line_already_set_right_to_left_needs_only_the_word(self):
        years = worded("years", 100, "1516 – 1917")
        years["text"]["directions"] = ["rtl"]
        [finding] = check(worded("heading", 40, "השלטון העות'מאני"), years)
        assert "direction" not in finding["fix"]

    def test_a_list_of_years_is_reported(self):
        [finding] = check(worded("heading", 40, "מלחמות"),
                          worded("years", 100, "1948, 1967, 1973"))
        assert "Hebrew word" in finding["fix"]

    def test_an_arabic_slide_gets_an_arabic_word(self):
        [finding] = check(worded("heading", 40, "الحكم العثماني"),
                          worded("years", 100, "1516 – 1917"))
        assert "من 1516 إلى 1917" in finding["fix"]

    @staticmethod
    def fix_for(line):
        found = check(worded("heading", 40, "השלטון העות'מאני"), worded("years", 100, line))
        return [f["fix"] for f in found if f["type"] == "number_order"]

    def test_the_replacement_is_only_on_this_slide_and_shape(self):
        # Without them, ppt_find_replace_text rewrites the line on every slide.
        [fix] = self.fix_for("1516 – 1917")
        assert "ppt_find_replace_text slide_indices=[1] shape_name='years' find_text=" in fix

    def test_a_range_keeps_the_signs_of_its_ends(self):
        assert 'replace_text="10% עד 20%"' in self.fix_for("10% – 20%")[0]
        assert 'replace_text="$5 עד $10"' in self.fix_for("$5 – $10")[0]
        assert 'replace_text="-5 עד 10"' in self.fix_for("-5 – 10")[0]
        assert 'replace_text="-10 עד 5"' in self.fix_for("-10 – 5")[0]

    def test_two_numbers_that_are_not_a_range_get_a_word_not_an_until(self):
        for line in ("3 - 1", "5 - 5", "2020: 5"):
            [fix] = self.fix_for(line)
            assert "עד" not in fix and "Hebrew word" in fix

    def test_a_phone_number_is_not_reported(self):
        assert self.fix_for("(03) 123-4567") == []

    def test_arabic_indic_digits_are_drawn_from_the_right_already(self):
        assert check(worded("heading", 40, "الحكم العثماني"),
                     worded("years", 100, "١٩٤٨–١٩٦٧")) == []

    def test_one_line_of_a_shape_is_judged_on_its_own(self):
        [finding] = check(worded("card", 40, "השלטון העות'מאני\r1516 – 1917"))
        assert finding["line"] == "1516 – 1917"
        assert "paragraph_index=2 direction='rtl'" in finding["fix"]

    @pytest.mark.parametrize("content", [
        "1516 עד 1917",        # a word sets the line right to left
        "586 לפנה״ס",
        "1948-1967",           # one hyphen keeps it one number
        "19.10.2026",
        "03 123 4567",         # spaces alone: a phone number, read left to right
        "1917",
    ])
    def test_lines_that_read_right_are_not(self, content):
        assert check(worded("heading", 40, "ציר הזמן"), worded("years", 100, content)) == []

    def test_an_english_slide_is_not(self):
        assert check(worded("heading", 40, "The Ottoman years"),
                     worded("years", 100, "1516 – 1917")) == []

    def test_a_slide_of_numbers_alone_is_not_guessed_at(self):
        assert check(worded("years", 100, "1516 – 1917")) == []


def directed(name, top, content, directions):
    box = worded(name, top, content)
    box["text"]["directions"] = directions
    return box


class TestTextDirection:
    """Live on Windows, "1516 עד 1917" in a paragraph not set right to left
    drew 1516 on the left; set right to left, it drew on the right. The Mac
    drew both from the right."""

    def test_a_hebrew_paragraph_with_a_number_left_to_right_is_reported(self):
        [finding] = check(directed("years", 100, "1516 עד 1917", ["ltr"]))
        assert finding["type"] == "text_direction"
        assert finding["paragraphs"] == [1]
        assert finding["fix"] == "Set them right to left (ppt_set_paragraph_format direction='rtl')."

    def test_punctuation_is_enough(self):
        [finding] = check(directed("t", 100, "שלום, עולם.", ["ltr"]))
        assert finding["type"] == "text_direction"

    def test_words_alone_draw_the_same_and_are_not(self):
        assert check(directed("t", 100, "ציר הזמן", ["ltr"])) == []

    def test_an_apostrophe_inside_a_word_draws_the_same_and_is_not(self):
        # Live: the heading "השלטון העות'מאני" was reported until the check
        # drew the line both ways instead of looking for punctuation.
        assert check(directed("t", 100, "השלטון העות'מאני", ["ltr"])) == []

    def test_a_number_after_the_words_is_reported(self):
        # The Unicode algorithm draws "עד 1917" the same both ways; PowerPoint
        # for Windows draws digits as Latin letters, and then it does not.
        [finding] = check(directed("t", 100, "עד 1917", ["ltr"]))
        assert finding["type"] == "text_direction"

    def test_a_paragraph_set_right_to_left_is_not(self):
        assert check(directed("t", 100, "1516 עד 1917", ["rtl"])) == []

    def test_an_english_paragraph_beside_it_is_left_alone(self):
        [finding] = check(directed("t", 100, "Ottoman rule\r1516 עד 1917", ["ltr", "ltr"]))
        assert finding["paragraphs"] == [2]
        assert "paragraph_index=2 direction='rtl'" in finding["fix"]

    def test_a_direction_never_read_is_not_guessed(self):
        assert check(worded("t", 100, "1516 עד 1917")) == []


class TestWhatAWriteCannotSee:
    def test_a_text_frame_the_write_did_not_read_is_not_lined_up_against(self):
        # Live: distributing three groups, the write compared the middle one's
        # centre with the centre of the unread title frame, 4 pt away.
        title = shape("title", 40, 30, 880, 60, type=msoTextBox, line=False,
                      fill={"visible": False, "solid": True, "color_hex": "#FFFFFF",
                            "transparency": 0.0})
        groups = [card(f"g{i}", left, 170, 180, 136, z=2 + i)
                  for i, left in enumerate((110, 394, 678))]
        found = slide_findings(slide(title, *groups), W, H, only={"g1"})
        assert found == []

    def test_an_unread_title_frame_still_marks_the_margins(self):
        title = shape("title", 40, 30, 880, 60, type=msoTextBox, line=False,
                      fill={"visible": False, "solid": True, "color_hex": "#FFFFFF",
                            "transparency": 0.0})
        cards = [card(f"c{i}", left, 140, 280, 180, z=2 + i)
                 for i, left in enumerate((40, 340, 650))]
        layout = layout_check._layout([title, *cards], W, H)
        assert layout["margins"]["horizontal"] == 40.0


class TestEmphasisAndRoles:
    def test_a_featured_card_of_its_own_colour_is_meant(self):
        cards = [card("a", 50, 140, 260, 300, z=1), card("b", 350, 125, 260, 330, z=2, colour="#DBEAFE"),
                 card("c", 650, 140, 260, 300, z=3)]
        assert "uneven_sizes" not in kinds(check(*cards))

    def test_a_card_that_grew_by_accident_kept_its_colour(self):
        cards = [card("a", 50, 140, 260, 300, z=1), card("b", 350, 140, 260, 330, z=2),
                 card("c", 650, 140, 260, 300, z=3)]
        assert "uneven_sizes" in kinds(check(*cards))

    def test_the_shape_placed_differently_on_its_card_is_the_one_that_moves(self):
        cards = [card("a", 50, 140, 260, 300, z=1), card("b", 350, 125, 260, 330, z=2, colour="#DBEAFE"),
                 card("c", 650, 140, 260, 300, z=3)]
        prices = [textbox("pa", 70, 202, 220, 50, z=4), textbox("pb", 370, 185, 220, 50, z=5),
                  textbox("pc", 670, 200, 220, 50, z=6)]
        [finding] = [f for f in check(*cards, *prices) if f["type"] == "misaligned"]
        assert finding["shapes"][0] == "pa"


class TestColumnsOfText:
    def test_figures_that_label_nothing_are_spaced(self):
        figures = [textbox(f"f{i}", 80, top, 220, 70, z=1 + i, text_height=58)
                   for i, top in enumerate((130, 230, 342))]
        [finding] = [f for f in check(*figures) if f["type"] == "uneven_spacing"]
        assert finding["gaps_pt"] == [30.0, 42.0]

    def test_a_big_figure_and_its_label_are_not_held_to_a_point(self):
        figure = placed("figure", 80, 130, 220, 65.4, 87.2, 133.6, 100, 58.2)
        label = placed("label", 320, 148, 500, 33.6, 327.2, 151.6, 220, 26.4, z=2)
        assert "misaligned" not in kinds(check(figure, label))


class TestRowsAndMargins:
    def test_a_row_set_in_from_both_sides_keeps_its_ends(self):
        title = placed("title", 40, 30, 880, 46, 47.2, 33.6, 200, 38.4)
        groups = [card(f"g{i}", left, 170, 180, 136, z=2 + i)
                  for i, left in enumerate((110, 390, 678))]
        [finding] = [f for f in check(title, *groups) if f["type"] == "uneven_spacing"]
        assert finding["fix"].startswith("ppt_distribute_shapes")


class TestCollisionsInARow:
    def test_a_subtitle_into_a_row_of_cards_is_one_finding_and_the_cards_move(self):
        title = placed("title", 40, 30, 880, 46, 47.2, 33.6, 400, 38.4)
        subtitle = placed("subtitle", 40, 82, 880, 55, 47.2, 85.6, 860, 48, z=2)
        cards = [card(f"c{i}", 40 + 300 * i, 125, 280, 200, z=3 + i) for i in range(3)]
        [finding] = [f for f in check(title, subtitle, *cards) if f["type"] == "overlap"]
        assert finding["shapes"] == ["subtitle", "c0", "c1", "c2"]
        assert finding["fix"].startswith("Move 'c0', 'c1', 'c2' down")


class TestFindingTheTitle:
    def test_the_topmost_large_text_is_the_title_not_the_biggest(self):
        title = textbox("title", 40, 30, 880, 60, runs=[{"color_hex": "#222222", "size": 32.0, "bold": True}])
        figure = textbox("figure", 80, 130, 220, 70, z=2,
                         runs=[{"color_hex": "#2563EB", "size": 48.0, "bold": True}])
        assert find_title(slide(title, figure), H) is title

    def test_a_title_beside_a_photo_is_where_it_has_to_be(self):
        def page(index):
            words = [{"color_hex": "#222222", "size": 32.0, "bold": True}]
            return slide(textbox(f"Title {index}", 40, 30, 880, 60, runs=words), index=index)
        photo = card("photo", 0, 0, 440, 540, z=1, colour="#6B7280")
        beside = textbox("Title 5", 480, 60, 440, 60, z=2,
                         runs=[{"color_hex": "#222222", "size": 32.0, "bold": True}])
        slides = [page(2), page(3), page(4), slide(photo, beside, index=5)]
        assert deck_findings(slides, H, W) == []


# ---------------------------------------------------------------------------
# A template's own graphics, colours and titles
# ---------------------------------------------------------------------------
def graphic(name, left, top, width, height, colour="#0B2545", source="master", z=-10_000, **extra):
    """A graphic the slide master or layout draws: a band, a logo, a bar."""
    facts = card(name, left, top, width, height, z=z, colour=colour)
    facts["source"] = source
    facts.update(extra)
    return facts


def template_text(name, left, top, width, height, text_width, text_height):
    """A line of text the template draws, a confidentiality notice, with no fill."""
    facts = textbox(name, left, top, width, height, z=-9_000,
                    text_width=text_width, text_height=text_height)
    facts["text"].update(left=left, top=top)
    facts["source"] = "master"
    return facts


def templated(*shapes, template=(), index=1, background="#FFFFFF", layout="title only",
              follows=True):
    page = slide(*shapes, index=index, background=background, layout=layout)
    page["template"] = list(template)
    page["follows_background"] = follows
    return page


def check_on(template, *shapes, **kwargs):
    return slide_findings(templated(*shapes, template=template, **kwargs), W, H)


def overlaps(findings):
    return [f for f in findings if f["type"] == "overlap"]


BAND = graphic("band", 0, 0, 70, 540)
STRIP = graphic("strip", 937.68, 0, 22.32, 540, colour="#F8B323")
LOGO = graphic("logo", 858, 500, 72, 24, colour="#13A89E")
LINE = template_text("notice", 40, 503, 220, 18, text_width=140, text_height=15)


def title_placeholder(name="Title 1", left=98.56, top=30, size=28.0, colour="#0B2545",
                      theme=True, layout_type=layout_check.ppPlaceholderTitle):
    title = textbox(name, left, top, 800, 56, runs=[
        {"color_hex": colour, "size": size, "bold": True, "font": "Georgia", "theme": theme}])
    title["type"] = msoPlaceholder
    title["placeholder_type"] = layout_type
    return title


def set_to(shape, side):
    """Draw a text box's text at the right of its frame, or at its centre."""
    words, inset = shape["text"], 7.2
    room = shape["width"] - 2 * inset - words["width"]
    words["left"] = shape["left"] + inset + (room if side == "right" else room / 2)
    return shape


class TestTheTemplatesOwnGraphics:
    def test_text_on_a_band_down_the_side_is_moved_clear_of_it(self):
        note = textbox("note", 20, 478, 320, 24)
        [found] = overlaps(check_on([BAND], note))
        assert found["shapes"] == ["note", "band"]
        assert "part of the template drawn by the slide master" in found["message"]
        assert found["fix"].startswith("Move 'note' right by")
        assert found["fix"].endswith("leaves 6 pt between it and 'band'.")

    def test_it_goes_on_to_line_up_with_the_content_when_that_is_near(self):
        note = textbox("note", 20, 478, 320, 24)
        found = overlaps(check_on([BAND], title_placeholder(), note))
        assert found[0]["fix"].startswith("Move 'note' right by 78.6 pt, in line with 'Title 1'")

    def test_on_a_right_to_left_slide_it_lines_up_the_right_edges(self):
        title = title_placeholder()
        title["width"] = 300
        note = textbox("note", 20, 478, 320, 24)
        found = overlaps(check_on([BAND], title, note))
        assert found[0]["fix"].startswith("Move 'note' right by 78.6 pt, in line with 'Title 1'")
        title["text"]["content"] = "ההיסטוריה של ישראל"
        set_to(title, "right")
        note["text"]["content"] = "מקור: הלשכה המרכזית לסטטיסטיקה"
        found = overlaps(check_on([BAND], title, note))
        assert found[0]["fix"].startswith("Move 'note' right by 58.6 pt, in line with 'Title 1'")

    def test_a_centred_title_is_not_an_edge_to_line_up_with(self):
        # Its frame starts at 98.56, but its text 325 pt further in: a note
        # put at the frame's edge is in line with nothing a reader sees.
        title = set_to(title_placeholder(), "center")
        note = textbox("note", 20, 478, 320, 24)
        [found] = overlaps(check_on([BAND], title, note))
        assert "in line with 'Title 1'" not in found["fix"]
        assert found["fix"].startswith("Move 'note' right by 48.8 pt (ppt_update_shape dleft=48.8)")

    def test_only_the_side_text_is_set_to_counts(self):
        # A right-set title keeps its right edge, not its left.
        title = set_to(title_placeholder(), "right")
        note = textbox("note", 20, 478, 320, 24)
        [found] = overlaps(check_on([BAND], title, note))
        assert "in line with 'Title 1'" not in found["fix"]
        # A card keeps both, whatever its text does.
        card_ = card("panel", 98.56, 150, 300, 200, z=2)
        [found] = overlaps(check_on([BAND], title, card_, note))
        assert found["fix"].startswith("Move 'note' right by 78.6 pt, in line with 'panel'")

    def test_a_footnote_steps_up_rather_than_across_the_slide(self):
        cards = [card(f"c{i}", 40 + 305 * i, 150, 270, 200, z=2 + i) for i in range(3)]
        footnote = textbox("footnote", 40, 498, 400, 20, z=9, text_width=280, text_height=14)
        [found] = overlaps(check_on([LINE], *cards, footnote))
        assert found["fix"].startswith("Move 'footnote' up by")

    def test_a_card_in_a_row_takes_the_rows_width_back(self):
        cards = [card("c0", 99, 180, 255, 200, z=2), card("c1", 374, 180, 255, 200, z=3),
                 card("c2", 649, 180, 300, 200, z=4)]
        found = check_on([BAND, STRIP], *cards)
        [hit] = overlaps(found)
        assert hit["shapes"] == ["c2", "strip"]
        assert hit["fix"].startswith("Make 'c2' 255 pt wide like the rest of its row (ppt_update_shape width=255)")

    def test_a_chart_over_the_footer_is_trimmed_not_pushed_into_the_title(self):
        title = title_placeholder(left=40)
        chart = card("Chart 3", 40, 100, 880, 420, z=2)
        chart["type"] = layout_check.msoChart
        [found] = overlaps(check_on([LOGO, LINE], title, chart))
        assert found["fix"].startswith("Make 'Chart 3' 394 pt tall (ppt_update_shape height=394)")

    def test_covering_a_graphic_whole_is_reported(self):
        badge = card("badge", 800, 440, 150, 100, z=2)
        [found] = overlaps(check_on([LOGO], badge))
        assert found["message"].startswith("'badge' covers 'logo'")

    def test_a_placeholder_is_where_its_layout_put_it(self):
        body = textbox("Content 2", 60, 72, 485, 392)
        body["type"], body["placeholder_type"] = msoPlaceholder, 7
        assert overlaps(check_on([BAND], body)) == []

    def test_text_wholly_on_a_panel_is_judged_against_the_panel(self):
        panel = graphic("panel", 600, 0, 360, 540, colour="#2A1A00", source="layout")
        light = textbox("light", 640, 100, 280, 40, runs=[{"color_hex": "#F3F3F2", "size": 20.0, "bold": False}])
        dark = textbox("dark", 640, 200, 280, 40, runs=[{"color_hex": "#222222", "size": 20.0, "bold": False}])
        found = check_on([panel], light, dark)
        assert overlaps(found) == []
        [contrast] = [f for f in found if f["type"] == "low_contrast"]
        assert contrast["shapes"] == ["dark"] and contrast["backdrop_color"] == "#2A1A00"

    def test_a_backdrop_picture_is_no_obstacle_and_no_colour(self):
        texture = graphic("texture", 0, 0, 960, 540, type=msoPicture)
        pale = textbox("pale", 100, 100, 300, 40, runs=[{"color_hex": "#FFFFFF", "size": 20.0, "bold": False}])
        assert check_on([texture], pale) == []

    def test_a_slide_without_template_facts_is_checked_as_before(self):
        note = textbox("note", 20, 478, 320, 24)
        assert overlaps(check(note)) == []


class TestTheTemplatesOwnColours:
    def light_title(self, theme=True):
        return title_placeholder(colour="#F3F3F2", theme=theme)

    def test_a_placeholder_in_theme_colours_is_taken_as_meant(self):
        # A caption layout's light title on a dark panel the Mac cannot read.
        found = check_on([], self.light_title(), background="#F3F3F2")
        assert [f for f in found if f["type"] == "low_contrast"] == []

    def test_the_same_placeholder_coloured_by_the_caller_is_checked(self):
        found = check_on([], self.light_title(theme=False), background="#F3F3F2")
        assert [f["shapes"] for f in found if f["type"] == "low_contrast"] == [["Title 1"]]

    def test_a_text_box_in_theme_colours_is_still_the_callers(self):
        box = textbox("box", 100, 100, 300, 40,
                      runs=[{"color_hex": "#F3F3F2", "size": 20.0, "bold": False, "theme": True}])
        found = check_on([], box, background="#F3F3F2")
        assert [f["shapes"] for f in found if f["type"] == "low_contrast"] == [["box"]]

    def test_a_background_the_caller_changed_is_the_callers(self):
        found = check_on([], self.light_title(), background="#F3F3F2", follows=False)
        assert [f["shapes"] for f in found if f["type"] == "low_contrast"] == [["Title 1"]]


def themed_title_slide(index, layout, size=28.0, left=98.56, theme=True, name=None,
                       placeholder=True, colour="#0B2545"):
    if placeholder:
        title = title_placeholder(name or f"Title {index}", left=left, size=size,
                                  theme=theme, colour=colour)
    else:
        title = textbox(name or f"TextBox {index}", left, 30, 800, 56, runs=[
            {"color_hex": colour, "size": size, "bold": True, "font": "Georgia", "theme": theme}])
    return templated(title, index=index, layout=layout)


class TestATemplatesTitles:
    def test_a_caption_layouts_title_in_its_side_column_is_not_compared(self):
        slides = [themed_title_slide(i, "title only") for i in (2, 3, 4)]
        slides.append(themed_title_slide(5, "content with caption", size=19.0, left=656.5))
        assert deck_findings(slides, H, W) == []

    def test_a_layouts_own_smaller_title_is_not_compared_when_the_template_leads(self):
        slides = [themed_title_slide(i, "title only") for i in (2, 3, 4)]
        slides.append(themed_title_slide(5, "content with caption", size=20.0))
        assert deck_findings(slides, H, W) == []

    def test_a_title_box_of_the_callers_own_is_held_to_the_templates(self):
        slides = [themed_title_slide(i, "title only") for i in (2, 3, 4)]
        slides.append(themed_title_slide(5, "blank", size=36.0, theme=False, placeholder=False))
        found = deck_findings(slides, H, W)
        assert found and all(f["slides"] == [5] for f in found)

    def test_when_the_callers_own_titles_lead_a_template_title_is_held_to_them(self):
        # The first real deck: titles drawn as boxes on blank slides, and one
        # slide on the layout's own placeholder, which jumps between them.
        slides = [themed_title_slide(i, "blank", size=32.0, theme=False, placeholder=False,
                                     colour="#1F2937") for i in (2, 3, 4)]
        slides.append(themed_title_slide(5, "title and content", size=44.0))
        found = deck_findings(slides, H, W)
        assert found and all(f["slides"] == [5] for f in found)

    def test_titles_on_one_layout_are_still_held_to_each_other(self):
        slides = [themed_title_slide(i, "title only") for i in (2, 3, 4)]
        slides.append(themed_title_slide(5, "title only", size=32.0))
        [found] = deck_findings(slides, H, W)
        assert found["slides"] == [5] and "size" in found["message"]

    def test_a_layout_of_the_same_name_on_another_design_styles_its_own_title(self):
        # Office Theme and Retrospect in one deck: each has a "Title and
        # Content", with a title of its own size, place and colour.
        slides = []
        for i in (1, 2, 3):
            page = themed_title_slide(i, "Title and Content", size=44.0, left=73.2,
                                      colour="#000000")
            page["design"] = "Office Theme"
            slides.append(page)
        for i in (4, 5):
            page = themed_title_slide(i, "Title and Content", size=48.0, left=93.6,
                                      colour="#404040")
            page["design"] = "Retrospect"
            slides.append(page)
        assert deck_findings(slides, H, W) == []
        for page in slides:
            page["design"] = None
        assert deck_findings(slides, H, W) != []

    def test_on_a_mac_a_master_is_told_by_where_it_puts_its_title_and_body(self):
        office = [[66, 28.75, 894, 133.13], [66, 143.75, 894, 486.38]]
        retro = [[93.6, 39.6, 866.4, 136.8], [93.6, 162, 866.4, 478]]
        key = layout_check.design_key
        assert key(office) == key(list(reversed(office))) != key(retro)
        assert key([]) is None and key(None) is None

    def test_a_title_low_in_a_tall_frame_is_anchored_not_moved_off_the_slide(self):
        def boxed(index, frame_top, text_top):
            title = placed(f"Title {index}", 98.56, frame_top, 801.44, 117.6, 105.8, text_top, 300, 38)
            title["text"]["runs"] = [{"color_hex": "#2A1A00", "size": 40.0, "bold": False}]
            return slide(title, index=index)
        slides = [boxed(i, 30.0, 33.7) for i in (2, 3, 4)] + [boxed(5, 36.0, 103.8)]
        [top] = [f for f in deck_findings(slides, H, W) if "top edge" in f["message"]]
        assert "vertical_anchor='top'" in top["fix"] and "top=30.1" in top["fix"]
        assert "top=-" not in top["fix"]


class TestBarsOfAChart:
    # Live, the beer deck: a bar chart drawn from rounded rectangles, one per
    # country, each as long as its value. Their widths are the data.
    def bars(self, *values):
        return [card(f"bar{i}", 246, 144 + 52 * i, v / 150 * 540, 30, z=1 + i)
                for i, v in enumerate(values)]

    def test_bars_that_shrink_in_order_are_a_chart(self):
        assert "uneven_sizes" not in kinds(check(*self.bars(110.6, 104.6, 99.0, 95.1, 93.2)))

    def test_bars_that_grow_in_order_are_a_chart(self):
        assert "uneven_sizes" not in kinds(check(*self.bars(93.2, 95.1, 99.0, 104.6)))

    def test_one_card_off_the_rest_is_still_a_slip(self):
        assert "uneven_sizes" in kinds(check(*self.bars(100, 100, 104, 100)))

    def test_sizes_out_of_order_are_still_a_slip(self):
        assert "uneven_sizes" in kinds(check(*self.bars(100, 104, 99, 103)))


@pytest.mark.usefixtures("without_title_edge")
class TestWhatTheCoffeeDeckTaught:
    # Live, a deck on coffee built a slide at a time: three kinds of finding
    # that a model told to act on every finding would have acted on, and
    # been wrong to.
    def test_a_photo_the_height_of_the_slide_has_no_middle_to_meet(self):
        photo = shape("photo", 0, 0, 330, 540, z=1, type=layout_check.msoPicture)
        step = card("step 3", 387, 248, 38, 38, z=2)
        assert "misaligned" not in kinds(check(photo, step))

    def test_a_shorter_photo_is_still_lined_up_by_its_middle(self):
        photo = shape("photo", 0, 100, 330, 340, z=1, type=layout_check.msoPicture)
        step = card("step 3", 387, 248, 38, 38, z=2)
        assert "misaligned" in kinds(check(photo, step))

    def figures(self, gaps_from=146):
        title = placed("title", 450, 38, 450, 52, 457.2, 41.6, 380, 44)
        title["text"]["runs"] = [{"color_hex": "#F3E9DC", "size": 34.0, "bold": True}]
        sub = placed("sub", 450, 92, 450, 26, 457.2, 95.6, 260, 19, z=2)
        shapes = [title, sub]
        for i in range(3):
            y = gaps_from + 100 * i
            big = placed(f"big{i}", 450, y, 450, 50, 457.2, y + 3.6, 300, 44, z=3 + 2 * i)
            big["text"]["runs"] = [{"color_hex": "#C67C3A", "size": 40.0, "bold": True}]
            label = placed(f"label{i}", 450, y + 52, 450, 28, 457.2, y + 55.6, 300, 20, z=4 + 2 * i)
            shapes += [big, label]
        return shapes

    def test_a_heading_is_not_the_first_of_the_figures_under_it(self):
        assert "uneven_spacing" not in kinds(check(*self.figures()))

    def test_figures_unevenly_spaced_among_themselves_still_are(self):
        shapes = self.figures()
        for s in shapes:
            if s["name"] in ("big2", "label2"):
                s["top"] += 8
                s["text"]["top"] += 8
        assert "uneven_spacing" in kinds(check(*shapes))

    def test_a_write_does_not_judge_two_bars_of_a_chart_still_being_drawn(self):
        bars = [card("Colombia", 246, 266, 96.5, 32, z=1), card("Ethiopia", 246, 324, 87.1, 32, z=2)]
        page = slide(*bars)
        assert "uneven_sizes" not in kinds(slide_findings(page, W, H, only={"Ethiopia"}))
        assert "uneven_sizes" in kinds(slide_findings(page, W, H))

    def test_the_finished_deck_is_clean(self):
        deck = recorded("coffeedeck_mac")
        assert all(found == [] for found in findings_by_slide(deck).values())
        assert deck_findings(deck["slides"], deck["slide_height"], deck["slide_width"]) == []


@pytest.mark.usefixtures("without_title_edge")
class TestADesignedDeckOnTheMac:
    # Twelve slides on beer, built on PowerPoint for Mac through the server:
    # photos, icons, cards with shadows, a timeline, a colour scale, and a bar
    # chart drawn from shapes. Every finding the checks made while it was
    # built was fixed, so it must report nothing.
    def test_the_finished_deck_is_clean(self):
        deck = recorded("beerdeck_mac")
        assert all(found == [] for found in findings_by_slide(deck).values())
        assert deck_findings(deck["slides"], deck["slide_height"], deck["slide_width"]) == []


class TestUnevenWidths:
    def row(self, *widths, colours=None):
        shapes, left = [], 99.0
        for i, width in enumerate(widths):
            shapes.append(card(f"c{i}", left, 180, width, 200, z=2 + i,
                               colour=(colours or ["#FFFFFF"] * len(widths))[i]))
            left += width + 20
        return shapes

    def sizes(self, *shapes):
        return [f for f in check(*shapes) if f["type"] == "uneven_sizes"]

    def test_a_card_wider_than_the_two_beside_it_is_reported(self):
        [found] = self.sizes(*self.row(255, 255, 300))
        assert "not one width" in found["message"]
        assert found["fix"].startswith("Make 'c2' 255 pt wide like the rest")
        assert "ppt_distribute_shapes" not in found["fix"]

    def test_a_middle_card_made_narrow_again_needs_the_gaps_evened(self):
        [found] = self.sizes(*self.row(255, 300, 255))
        assert "ppt_distribute_shapes" in found["fix"]

    def test_two_cards_of_two_widths_are_a_design(self):
        assert self.sizes(*self.row(560, 255)) == []

    def test_a_much_wider_card_is_a_design(self):
        assert self.sizes(*self.row(200, 200, 300)) == []

    def test_a_featured_card_is_meant(self):
        assert self.sizes(*self.row(255, 300, 255, colours=["#FFFFFF", "#2563EB", "#FFFFFF"])) == []


class FakeCollection:
    def __init__(self, shapes):
        self._shapes = shapes
        self.Count = len(shapes)

    def __call__(self, i):
        return self._shapes[i - 1]


class TestReadingATemplatesGraphics:
    def fake(self, shows=msoTrue, layout_shows=msoTrue, design=1, layout_index=2):
        band = fake_shape(name="Band")
        band.Type, band.Left, band.Width, band.HasTextFrame = msoAutoShape, 0.0, 70.0, msoFalse
        band.Fill = SimpleNamespace(Visible=msoTrue, Type=msoFillSolid,
                                    ForeColor=SimpleNamespace(RGB=0x00), Transparency=0.0)
        title = fake_shape(name="Title Placeholder 1")
        title.Type = msoPlaceholder
        title.PlaceholderFormat = SimpleNamespace(Type=1)
        panel = fake_shape(name="Panel")
        panel.Type, panel.HasTextFrame = msoAutoShape, msoFalse
        layout = SimpleNamespace(DisplayMasterShapes=layout_shows, Shapes=FakeCollection([panel]),
                                 Index=layout_index)
        master = SimpleNamespace(Shapes=FakeCollection([band, title]))
        return SimpleNamespace(DisplayMasterShapes=shows, CustomLayout=layout, Master=master,
                               Design=SimpleNamespace(Index=design))

    def test_the_master_and_layout_graphics_are_read_without_placeholders(self):
        graphics = layout_check._template_facts(self.fake())
        assert [(g["name"], g["source"]) for g in graphics] == [("Band", "master"), ("Panel", "layout")]
        assert all(g["z"] < 0 for g in graphics)
        assert graphics[0]["fill"]["visible"] is True

    def test_a_layout_hiding_the_masters_graphics_shows_only_its_own(self):
        graphics = layout_check._template_facts(self.fake(layout_shows=msoFalse))
        assert [g["name"] for g in graphics] == ["Panel"]

    def test_a_slide_hiding_background_graphics_has_none(self):
        assert layout_check._template_facts(self.fake(shows=msoFalse)) == []

    def test_a_write_reads_only_the_graphics_near_what_it_touched(self):
        # The band runs down the left edge; the touched text box is far right.
        far = [textbox("note", 700, 100, 200, 30)]
        graphics = layout_check._template_facts(self.fake(), touched=far)
        band = next(g for g in graphics if g["name"] == "Band")
        assert band["fill"] is None and band["left"] == 0.0
        near = [textbox("note", 20, 100, 200, 30)]
        graphics = layout_check._template_facts(self.fake(), touched=near)
        assert next(g for g in graphics if g["name"] == "Band")["fill"]["visible"] is True

    def test_a_whole_check_reads_each_layout_once(self):
        cache = {}
        first = layout_check._template_facts(self.fake(), cache=cache)
        again = self.fake()
        again.Master = None   # reading it again would fail
        assert layout_check._template_facts(again, cache=cache) == first
        other = layout_check._template_facts(self.fake(layout_index=3), cache=cache)
        assert len(cache) == 2 and other is not first


class TestAWritesUnreadGraphics:
    def test_a_graphic_left_unread_still_blocks_a_move(self):
        # A band across the slide; the note's text runs into its bottom edge,
        # and the shortest way clear is down, onto a line of template text the
        # write did not read, being far from what it touched.
        band = graphic("band", 0, 280, 300, 30)
        note = textbox("note", 100, 300, 200, 24)
        unread = template_text("notice", 100, 330, 220, 18, text_width=140, text_height=15)
        unread["text"] = None
        [found] = overlaps(check_on([band, unread], note))
        assert found["fix"].startswith("Move 'note' up by")
        [found] = overlaps(check_on([band], note))
        assert found["fix"].startswith("Move 'note' down by")


# ---------------------------------------------------------------------------
# Which writes report, read with ast so it runs anywhere
# ---------------------------------------------------------------------------
import ast  # noqa: E402

REPORTING = {
    "shapes.py": ["add_shape", "add_textbox", "add_picture", "update_shape", "duplicate_shape"],
    "text.py": ["set_text", "format_text", "format_text_range", "set_textframe",
                "find_replace_text", "set_paragraph_format"],
    "advanced_ops.py": ["add_svg_icon", "add_picture_from_url", "replace_font",
                        "set_default_fonts"],
    "layout.py": ["align_shapes", "distribute_shapes", "set_slide_background", "set_slide_size"],
    "placeholders.py": ["set_placeholder_text"],
    "tables.py": ["add_table", "set_table_cell", "set_table_data", "add_table_row",
                  "add_table_column", "merge_table_cells", "set_table_layout"],
    "charts.py": ["add_chart"],
    "edit_ops.py": ["copy_shape_to_slide"],
    "formatting.py": ["set_fill"],
    "batch_apply.py": ["batch_apply_formatting"],
    "themes.py": ["apply_theme", "set_theme_colors"],
}
REPORTERS = {"report_layout", "report_layouts", "report_contrast", "report_deck_change"}


class TestWhichWritesReport:
    @pytest.mark.parametrize("module, function", [
        (module, function) for module, functions in REPORTING.items() for function in functions])
    def test_the_write_reports_what_it_left_wrong(self, module, function):
        path = Path(__file__).resolve().parents[1] / "src" / "ppt_com" / module
        tree = ast.parse(path.read_text(encoding="utf-8"))
        [node] = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == function]
        called = {c.func.id for c in ast.walk(node)
                  if isinstance(c, ast.Call) and isinstance(c.func, ast.Name)}
        assert called & REPORTERS

    def test_a_copy_or_a_placeholder_is_named_by_its_own_key(self):
        assert touched_names({"new_shape_name": "Card copy"}) == ["Card copy"]
        assert touched_names({"placeholder_name": "Title 1"}) == ["Title 1"]

    def test_a_new_background_reports_only_the_text_it_left_hard_to_read(self, monkeypatch):
        grey = textbox("grey", 100, 100, 300, 40,
                       runs=[{"color_hex": "#555555", "size": 20.0, "bold": False}])
        stray = card("stray", 900, 100, 200, 50, z=2)   # off the slide, but not this call's business
        fake = FakeDeck([slide(grey, stray, background="#333333")])
        monkeypatch.setattr(layout_check, "ppt", fake)
        result = layout_check.report_contrast({"success": True, "slide_indices": [1]}, [1])
        assert [(f["type"], f["shapes"]) for f in result["layout_issues"]] == [("low_contrast", ["grey"])]


def set_title(index, text_left, text_width=200.0, frame_left=40.0, frame_width=880.0, top=30.0):
    """A title box whose text PowerPoint drew at ``text_left``: set left, right or centred."""
    title = placed(f"Title {index}", frame_left, top, frame_width, 60,
                   text_left, top + 3.6, text_width, 40)
    title["text"]["runs"] = [{"color_hex": "#1F2937", "size": 32.0, "bold": True}]
    return slide(title, index=index)


class TestTitlesSetRightOrCentred:
    RIGHT = 40 + 880 - 7.2     # where a right-set title's text ends

    def test_right_set_titles_of_any_length_agree(self):
        slides = [set_title(i, self.RIGHT - width, width) for i, width in ((2, 120), (3, 420), (4, 250))]
        assert deck_findings(slides, H, W) == []

    def test_a_right_set_title_off_the_right_edge_is_moved_by_it(self):
        slides = [set_title(i, self.RIGHT - width, width) for i, width in ((2, 120), (3, 420), (4, 250))]
        # Its frame 8 pt left of the others', its text set against the frame's right side.
        slides.append(set_title(5, self.RIGHT - 8 - 300, 300, frame_left=32.0))
        [found] = deck_findings(slides, H, W)
        assert found["message"].startswith("Slide titles differ in right edge")
        assert found["slides"] == [5] and "ppt_update_shape left=40 on 'Title 5'" in found["fix"]

    def test_centred_titles_are_compared_on_their_centres(self):
        middle = 40 + 440
        slides = [set_title(i, middle - width / 2, width) for i, width in ((2, 120), (3, 420), (4, 250))]
        assert deck_findings(slides, H, W) == []
        slides.append(set_title(5, middle + 10 - 150, 300))
        [found] = deck_findings(slides, H, W)
        assert found["message"].startswith("Slide titles differ in centre") and found["slides"] == [5]

    def test_a_title_filling_its_frame_is_not_a_different_alignment(self):
        slides = [set_title(i, self.RIGHT - width, width) for i, width in ((2, 120), (3, 420))]
        slides.append(set_title(4, 47.2, 865.6))
        assert deck_findings(slides, H, W) == []

    def test_a_title_set_left_among_right_set_ones_is_reported_as_its_alignment(self):
        slides = [set_title(i, self.RIGHT - width, width) for i, width in ((2, 120), (3, 420), (4, 250))]
        slides.append(set_title(5, 47.2, 200))
        [found] = deck_findings(slides, H, W)
        assert "differ in alignment" in found["message"] and found["slides"] == [5]


class TestWritesThatNameShapesOtherwise:
    def test_a_shape_given_by_position_is_checked_by_its_name(self, monkeypatch):
        off = card("Rectangle 2", 900, 100, 100, 50, z=2)
        fake = FakeDeck([slide(card("Rectangle 1", 10, 10, 50, 50), off)])
        monkeypatch.setattr(layout_check, "ppt", fake)
        result = report_layout({"success": True}, 1, [2])
        assert [f["shapes"] for f in result["layout_issues"]] == [["Rectangle 2"]]

    def test_a_write_across_slides_reports_each(self, monkeypatch):
        first = slide(card("a", 900, 100, 100, 50), index=1)
        second = slide(card("b", -20, 100, 100, 50), index=2)
        fake = FakeDeck([first, second])
        monkeypatch.setattr(layout_check, "ppt", fake)
        result = layout_check.report_layouts({"mode": "replace"}, {1: ["a"], 2: ["b"]})
        assert sorted(f["slide_index"] for f in result["layout_issues"]) == [1, 2]


class TestAChangeToTheWholeDeck:
    def test_only_what_the_change_broke_is_reported(self, monkeypatch):
        already = card("already", 900, 100, 100, 50)                      # off the slide before
        words = textbox("words", 100, 100, 300, 40,
                        runs=[{"color_hex": "#222222", "size": 20.0, "bold": False}])
        fake = FakeDeck([slide(already, words, background="#FFFFFF")])
        monkeypatch.setattr(layout_check, "ppt", fake)
        before = layout_check.deck_snapshot()
        # The new theme darkens the background behind the words.
        fake.slides = [slide(already, words, background="#1A1A1A")]
        result = layout_check.report_deck_change({"success": True}, before)
        assert [(f["type"], f["shapes"]) for f in result["layout_issues"]] == [("low_contrast", ["words"])]

    def test_a_deck_that_could_not_be_read_first_blames_nothing(self, monkeypatch):
        monkeypatch.setattr(layout_check, "ppt", FakePpt(RuntimeError("busy")))
        assert layout_check.deck_snapshot() is None
        assert layout_check.report_deck_change({"success": True}, None) == {"success": True}

    def test_titles_in_theme_colours_the_caller_changed_are_judged(self, monkeypatch):
        # A dark background in light1 and the titles left in a dark dark1: the
        # placeholders still take the theme's colours, which are no longer
        # the template's design but the caller's.
        def deck(text, background):
            title = title_placeholder(colour=text)
            title["text"]["chars"] = 12
            return [templated(title, index=i, background=background) for i in (1, 2)]
        fake = FakeDeck(deck("#0B2545", "#FFFFFF"))
        monkeypatch.setattr(layout_check, "ppt", fake)
        before = layout_check.deck_snapshot(trust_template=False)
        fake.slides = deck("#0B2545", "#111827")
        result = layout_check.report_deck_change({"success": True}, before, trust_template=False)
        found = result["layout_issues"]
        assert [(f["slide_index"], f["type"]) for f in found] == [(1, "low_contrast"), (2, "low_contrast")]
        assert found[0]["fix"].startswith("Its colour is the theme's: choose theme colours that "
                                          "read on each other (ppt_set_theme_colors), or make the text")
        # A new theme, colours and graphics chosen together, is still trusted.
        fake.slides = deck("#0B2545", "#FFFFFF")
        before = layout_check.deck_snapshot()
        fake.slides = deck("#0B2545", "#111827")
        assert "layout_issues" not in layout_check.report_deck_change({"success": True}, before)

    def test_setting_theme_colours_judges_the_placeholders(self):
        source = (Path(__file__).parent.parent / "src" / "ppt_com" / "themes.py").read_text()
        body = source[source.index("def set_theme_colors("):]
        body = body[:body.index("\ndef ")]
        assert "deck_snapshot(trust_template=False)" in body
        assert "report_deck_change(result, before, trust_template=False)" in body


class TestTurnedOff:
    """PPT_LAYOUT_CHECKS=off: the writes read nothing back; ppt_check_layout still runs."""

    @pytest.fixture(autouse=True)
    def off(self, monkeypatch):
        monkeypatch.setattr(layout_check, "CHECK_EACH_WRITE", False)

    def test_a_write_reports_nothing_and_reads_nothing(self, monkeypatch):
        fake = FakePpt(facts(card("c", 900, 100, 100, 50)))
        monkeypatch.setattr(layout_check, "ppt", fake)
        assert report_layout({"shape_name": "c"}, 1, ["c"]) == {"shape_name": "c"}
        assert fake.calls == []

    def test_a_picture_keeps_its_result_clean(self, monkeypatch):
        # The proportions ppt_add_picture hands the check are not for the caller.
        fake = FakePpt(facts(card("c", 900, 100, 100, 50)))
        monkeypatch.setattr(layout_check, "ppt", fake)
        assert report_layout({"shape_name": "c", "_natural_aspect": 1.5}, 1, ["c"]) == {
            "shape_name": "c"}

    def test_nor_a_new_background(self, monkeypatch):
        fake = FakeDeck([slide(card("c", 10, 10, 50, 50), background="#333333")])
        monkeypatch.setattr(layout_check, "ppt", fake)
        assert layout_check.report_contrast({"success": True}, [1]) == {"success": True}
        assert fake.calls == []

    def test_nor_a_change_to_the_whole_deck(self, monkeypatch):
        fake = FakeDeck([slide(card("c", 10, 10, 50, 50))])
        monkeypatch.setattr(layout_check, "ppt", fake)
        before = layout_check.deck_snapshot()
        assert before is None
        assert layout_check.report_deck_change({"success": True}, before) == {"success": True}
        assert fake.calls == []

    def test_the_check_on_request_still_runs(self, monkeypatch):
        fake = FakeDeck([slide(card("c", 900, 100, 100, 50))])
        monkeypatch.setattr(layout_check, "ppt", fake)
        result = json.loads(layout_check.check_layout(layout_check.CheckLayoutInput(slide_index=1)))
        assert kinds(result["issues"]) == ["off_slide"]


class TestTheSetting:
    @pytest.mark.parametrize("value", ["on", "ON", "true", "1", "yes", ""])
    def test_on(self, value):
        assert layout_check._check_each_write(value) is True

    @pytest.mark.parametrize("value", ["off", "Off", "false", "0", "no", " off "])
    def test_off(self, value):
        assert layout_check._check_each_write(value) is False

    def test_anything_else_is_on_and_says_so(self, caplog):
        # Read at import: a typo must not stop the server from starting.
        with caplog.at_level("WARNING"):
            assert layout_check._check_each_write("of") is True
        assert "PPT_LAYOUT_CHECKS='of' is not a setting; use 'on'" in caplog.text

    def test_the_instructions_follow_it(self):
        source = (Path(__file__).parent.parent / "src" / "server.py").read_text(encoding="utf-8")
        assert '7. """ + _WRITE_CHECKS + """' in source
        assert "if CHECK_EACH_WRITE else" in source


class TestATitleTooLongForItsBox:
    def test_growing_it_into_the_body_is_not_the_fix(self):
        # A title placeholder that PowerPoint did not shrink, its body right below.
        title = placed("Title 1", 66, 28.75, 828, 97.2, 73.2, 12.6, 800, 129.6)
        body = placed("Content 2", 66, 143.75, 828, 342.6, 73.2, 147.35, 600, 300, z=2)
        [found] = [f for f in check(title, body) if f["type"] == "overflow"]
        assert found["fix"] == ("Cut the text or make the font smaller. Making 'Title 1' 39.6 pt "
                                "taller would run it into the text of 'Content 2'.")

    def test_with_room_below_the_box_grows(self):
        title = placed("Title 1", 66, 28.75, 828, 97.2, 73.2, 12.6, 800, 129.6)
        [found] = [f for f in check(title) if f["type"] == "overflow"]
        assert found["fix"].startswith("Make 'Title 1' 39.6 pt taller (ppt_update_shape height=136.8)")


class TestDecorationsDrawnInLines:
    # Wisp: thin curves down the left whose box reaches 225 pt in, and a
    # title and body that the theme itself starts at 204 pt.
    WISP_AREA = [[204, 49, 906, 150], [204, 168, 906, 474]]

    def curves(self):
        return graphic("Group 22", 0, 18, 225, 523, type=layout_check.msoGroup, fill=None)

    def test_content_lined_up_with_the_themes_title_is_clear_of_them(self):
        cards = [card(f"c{i}", 204 + 240 * i, 170, 220, 220, z=2 + i) for i in range(3)]
        page = templated(*cards, template=[self.curves()])
        page["template_area"] = self.WISP_AREA
        assert overlaps(slide_findings(page, W, H)) == []

    def test_text_set_in_front_of_them_is_reported_and_moved_in_line_with_the_title(self):
        title = title_placeholder(left=204.2, top=49)
        words = textbox("words", 40, 170, 600, 100, z=3, text_width=440, text_height=58)
        page = templated(title, words, template=[self.curves()])
        page["template_area"] = self.WISP_AREA
        [found] = overlaps(slide_findings(page, W, H))
        assert found["fix"].startswith("Move 'words' right by 164.2 pt, in line with 'Title 1'")

    def test_a_solid_band_is_never_carved_out(self):
        band = graphic("band", 0, 0, 960, 90, colour="#0B2545")
        words = textbox("words", 100, 70, 400, 40, text_width=300, text_height=30)
        page = templated(words, template=[band])
        page["template_area"] = [[60, 20, 900, 80]]
        assert [f["shapes"] for f in overlaps(slide_findings(page, W, H))] == [["words", "band"]]

    def test_a_card_reaching_past_the_themes_margin_still_is(self):
        # Slice: its fan of lines starts where its title and body end, at 726 pt.
        lines = graphic("Group 6", 725, 233, 235, 253, type=layout_check.msoGroup, fill=None)
        cards = [card(f"c{i}", 54 + 246 * i, 60, 226, 250, z=2 + i) for i in range(3)]
        page = templated(*cards, template=[lines])
        page["template_area"] = [[54, 353, 726, 472], [54, 54, 726, 339]]
        [found] = overlaps(slide_findings(page, W, H))
        assert found["shapes"] == ["c2", "Group 6"]
        # Its lines may be nowhere near the card: what it did is leave the area.
        assert found["message"].startswith(
            "'c2' runs out of the template's content area into the space of 'Group 6', "
            "a decoration drawn by the slide master: an area")

    def test_without_the_themes_area_a_decoration_is_its_box(self):
        lines = graphic("Group 6", 725, 233, 235, 253, type=layout_check.msoGroup, fill=None)
        cards = [card(f"c{i}", 54 + 246 * i, 60, 226, 250, z=2 + i) for i in range(3)]
        [found] = overlaps(check_on([lines], *cards))
        assert found["message"].startswith("'c2' runs into 'Group 6', part of the template")


class TestARowIntoADecoration:
    def test_the_row_narrows_evenly_and_the_cards_take_their_labels(self):
        lines = graphic("Group 6", 725, 233, 235, 253, type=layout_check.msoGroup, fill=None)
        shapes = []
        for i in range(3):
            left = 54 + 246 * i
            shapes.append(card(f"c{i}", left, 60, 226, 250, z=2 + 2 * i))
            shapes.append(textbox(f"label{i}", left + 16, 76, 194, 36, z=3 + 2 * i))
        page = templated(*shapes, template=[lines])
        page["template_area"] = [[54, 353, 726, 472], [54, 54, 726, 339]]
        [found] = overlaps(slide_findings(page, W, H))
        assert found["fix"].startswith("Make the row fit before 'Group 6': ppt_update_shape "
                                       "shape_names=['c0', 'c1', 'c2'] width=208.67")
        assert "'c2' with what is on it left by 34.7 pt (ppt_update_shape shape_names=['c2', 'label2']" in found["fix"]
        assert found["fix"].endswith("The gaps stay 20 pt.")


class TestWhatCannotBeReadDoesNotFailTheSlide:
    def test_a_shape_that_cannot_be_described_is_left_out(self):
        def geometry(ref, z):
            if ref == "odd":
                raise RuntimeError("PowerPoint will not say")
            return {"name": ref, "z": z}

        kept = layout_check.read_shapes(["a", "odd", "b"], geometry,
                                        lambda facts, ref, full: facts, None)
        assert [s["name"] for s in kept] == ["a", "b"]
        assert [s["z"] for s in kept] == [1, 3]

    def test_the_commonest_font_is_the_same_every_time(self):
        assert layout_check._commonest(["B", "A", "B", "A"]) == "B"
        assert layout_check._commonest(["A", "B", "B", "A"]) == "A"
        assert layout_check._commonest([]) is None


# ---------------------------------------------------------------------------
# Stretched pictures
# ---------------------------------------------------------------------------
import struct  # noqa: E402
import zipfile  # noqa: E402

from utils.package import image_size, picture_aspects  # noqa: E402


def png(width, height):
    return b"\x89PNG\r\n\x1a\n" + struct.pack(">I", 13) + b"IHDR" + struct.pack(">II", width, height) + b"\x08\x02\x00\x00\x00"


def jpeg(width, height, orientation=None):
    data = b"\xff\xd8"
    if orientation is not None:
        tiff = b"MM\x00\x2a" + struct.pack(">I", 8) + struct.pack(">H", 1) + struct.pack(
            ">HHIHH", 0x0112, 3, 1, orientation, 0) + b"\x00\x00\x00\x00"
        exif = b"Exif\x00\x00" + tiff
        data += b"\xff\xe1" + struct.pack(">H", len(exif) + 2) + exif
    sof = b"\x08" + struct.pack(">HH", height, width) + b"\x03\x01\x11\x00\x02\x11\x01\x03\x11\x01"
    return data + b"\xff\xc0" + struct.pack(">H", len(sof) + 2) + sof + b"\xff\xd9"


def picture(name, left, top, width, height, aspect=None, z=1):
    facts = shape(name, left, top, width, height, z=z, type=msoPicture)
    if aspect is not None:
        facts["natural_aspect"] = aspect
    return facts


PRESENTATION = (
    '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>')
RELS = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '{}</Relationships>')


def pic_xml(name, embed, crop=""):
    return (f'<p:pic><p:nvPicPr><p:cNvPr id="4" name="{name}"/></p:nvPicPr>'
            f'<p:blipFill><a:blip r:embed="{embed}"/>{crop}<a:stretch/></p:blipFill></p:pic>')


def deck_file(tmp_path, pics, media):
    slide_xml = (
        '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<p:cSld><p:spTree>{"".join(pics)}</p:spTree></p:cSld></p:sld>')
    path = tmp_path / "deck.pptx"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("ppt/presentation.xml", PRESENTATION)
        archive.writestr("ppt/_rels/presentation.xml.rels", RELS.format(
            '<Relationship Id="rId2" Target="slides/slide1.xml" Type="slide"/>'))
        archive.writestr("ppt/slides/slide1.xml", slide_xml)
        archive.writestr("ppt/slides/_rels/slide1.xml.rels", RELS.format("".join(
            f'<Relationship Id="{rid}" Target="../media/{file}" Type="image"/>'
            for rid, (file, _) in media.items())))
        for file, data in media.values():
            archive.writestr(f"ppt/media/{file}", data)
    return path


class TestImageSizes:
    def test_png_and_jpeg_are_measured(self):
        assert image_size(png(1600, 900)) == (1600, 900)
        assert image_size(jpeg(1200, 800)) == (1200, 800)

    def test_a_jpeg_marked_as_turned_is_not_guessed_at(self):
        assert image_size(jpeg(1200, 800, orientation=6)) is None
        assert image_size(jpeg(1200, 800, orientation=1)) == (1200, 800)

    def test_something_else_is_unknown(self):
        assert image_size(b"<svg/>") is None


class TestPictureProportionsFromTheFile:
    def test_a_picture_takes_its_images_proportions(self, tmp_path):
        path = deck_file(tmp_path, [pic_xml("Photo", "rId1")], {"rId1": ("a.png", png(1600, 900))})
        assert picture_aspects(path, [1]) == {1: {"Photo": pytest.approx(16 / 9)}}

    def test_a_crop_is_taken_off_first(self, tmp_path):
        crop = '<a:srcRect l="25000" r="25000"/>'
        path = deck_file(tmp_path, [pic_xml("Photo", "rId1", crop)], {"rId1": ("a.png", png(1600, 900))})
        assert picture_aspects(path, [1])[1]["Photo"] == pytest.approx(800 / 900)

    def test_two_pictures_of_one_name_are_left_unknown(self, tmp_path):
        path = deck_file(tmp_path, [pic_xml("Photo", "rId1"), pic_xml("Photo", "rId1")],
                         {"rId1": ("a.png", png(1600, 900))})
        assert picture_aspects(path, [1]) == {1: {}}


class TestStretchedPictures:
    def test_a_picture_stretched_sideways_is_narrowed_about_its_centre(self):
        # A 4:3 photo given 400 by 200.
        [finding] = check(picture("Photo", 100, 100, 400, 200, aspect=4 / 3))
        assert finding["type"] == "stretched_picture"
        assert finding["stretch_pct"] == 50
        assert finding["fix"].startswith(
            "Crop the image to the proportions it is drawn at, keeping its size "
            "(ppt_crop_picture shape_name_or_index='Photo' crop_fit='shape'), or make it")
        assert "ppt_update_shape width=266.67 left=166.67" in finding["fix"]

    def test_a_logo_squeezed_is_made_shorter_before_it_is_cropped(self):
        [finding] = check(picture("Logo", 100, 100, 80, 80, aspect=2.0))
        assert finding["fix"].startswith("Make it 40 pt shorter")
        assert "ppt_update_shape height=40 top=120" in finding["fix"]

    def test_a_point_of_rounding_is_not(self):
        assert check(picture("Photo", 100, 100, 400, 301, aspect=4 / 3)) == []

    def test_a_picture_whose_image_is_unknown_is_not_judged(self):
        assert check(picture("Photo", 100, 100, 400, 200)) == []

    def test_a_new_picture_is_judged_by_what_the_write_saw(self, monkeypatch):
        fake = FakeDeck([slide(picture("Picture 1", 100, 100, 400, 200))])
        monkeypatch.setattr(layout_check, "ppt", fake)
        result = report_layout({"success": True, "shape_name": "Picture 1",
                                "_natural_aspect": 4 / 3}, 1, ["Picture 1"])
        assert "_natural_aspect" not in result
        assert [f["type"] for f in result["layout_issues"]] == ["stretched_picture"]


def test_a_locked_stretched_picture_is_unlocked_first():
    photo = picture("Photo", 100, 100, 400, 200, aspect=4 / 3)
    photo["lock_aspect"] = True
    [finding] = check(photo)
    assert "or unlock its proportions first (ppt_lock_aspect_ratio locked=false), then make it" in finding["fix"]


# ---------------------------------------------------------------------------
# Body text sizes across the deck
# ---------------------------------------------------------------------------
def body_slide(index, size, chars=120, autofit="none"):
    title = textbox(f"Title {index}", 40, 30, 880, 60, runs=[
        {"color_hex": "#111111", "size": 32.0, "bold": True}])
    body = textbox(f"Body {index}", 40, 120, 880, 300, z=2, chars=chars, autofit=autofit,
                   runs=[{"color_hex": "#333333", "size": size, "bold": False}])
    body["text"]["content"] = "Our first point, said plainly.\rAnd the second one, said as plainly."
    return slide(title, body, index=index)


class TestBodyTextSizes:
    def test_a_slide_whose_copy_was_shrunk_to_fit_is_reported(self):
        slides = [body_slide(2, 18.0), body_slide(3, 18.0), body_slide(4, 14.0), body_slide(5, 18.0)]
        [finding] = deck_findings(slides, H, W)
        assert finding["type"] == "inconsistent_body_text"
        assert finding["slides"] == [4] and finding["shapes"] == ["Body 4"]
        assert finding["message"] == "Body text differs in size between slides: 18 pt on slides 2, 3, 5; 14 pt on slide 4."
        assert "ppt_format_text font_size=18" in finding["fix"] and "split the slide" in finding["fix"]

    def test_text_set_far_larger_is_another_kind_of_text(self):
        slides = [body_slide(2, 18.0), body_slide(3, 18.0), body_slide(4, 32.0)]
        assert deck_findings(slides, H, W) == []

    def test_a_point_apart_is_the_same_size(self):
        slides = [body_slide(2, 18.0), body_slide(3, 18.0), body_slide(4, 17.0)]
        assert deck_findings(slides, H, W) == []

    def test_text_powerpoint_shrinks_to_fit_is_left_to_it(self):
        slides = [body_slide(2, 18.0), body_slide(3, 18.0), body_slide(4, 14.0, autofit="shrink_to_fit")]
        assert deck_findings(slides, H, W) == []

    def test_a_label_is_not_body_text(self):
        slides = [body_slide(2, 18.0), body_slide(3, 18.0), body_slide(4, 14.0, chars=12)]
        assert deck_findings(slides, H, W) == []

    def test_a_body_placeholder_is_held_to_its_own_layout(self):
        def placeholder_slide(index, size, layout):
            page = body_slide(index, size)
            page["shapes"][1]["placeholder_type"] = 2   # ppPlaceholderBody
            page["shapes"][1]["type"] = msoPlaceholder
            page["layout"] = layout
            return page
        slides = [placeholder_slide(2, 20.0, "Text"), placeholder_slide(3, 20.0, "Text"),
                  placeholder_slide(4, 16.0, "Text"), placeholder_slide(5, 24.0, "Comparison")]
        [finding] = deck_findings(slides, H, W)
        assert finding["slides"] == [4]

    def test_none_of_the_recorded_decks_is_reported(self):
        for path in sorted(FIXTURES.glob("*.json")):
            deck = json.loads(path.read_text(encoding="utf-8"))
            found = layout_check._body_sizes(deck["slides"], deck["slide_width"], deck["slide_height"])
            assert found == [], path.name


# ---------------------------------------------------------------------------
# Charts left with sample data
# ---------------------------------------------------------------------------
from ppt_com.constants import msoChart  # noqa: E402


def chart_shape(categories, series):
    facts = shape("Chart 3", 100, 100, 500, 300, type=msoChart)
    facts["chart"] = {"categories": categories, "series": series}
    return facts


class TestChartSampleData:
    def test_a_new_chart_is_reported(self):
        [finding] = check(chart_shape(
            ["Category 1", "Category 2", "Category 3", "Category 4"],
            [{"name": "Series 1", "values": [4.3, 2.5, 3.5, 4.5]}]))
        assert finding["type"] == "chart_sample_data"
        assert "categories named 'Category 1'" in finding["message"]
        assert "series named 'Series 1'" in finding["message"]
        assert "sample numbers" in finding["message"]

    def test_real_numbers_under_sample_names_are_still_reported(self):
        [finding] = check(chart_shape(
            ["North", "South"], [{"name": "Series 1", "values": [12.0, 30.0]}]))
        assert finding["message"].endswith("series named 'Series 1'.")

    def test_a_finished_chart_is_not(self):
        assert check(chart_shape(["Q1", "Q2"], [{"name": "Revenue", "values": [12.0, 30.0]}])) == []

    QUARTERS = ("1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr")

    def test_a_real_chart_by_quarter_is_not(self):
        # The pie's sample quarters are what real quarterly data is called too.
        assert check(chart_shape(self.QUARTERS, [{"name": "Revenue",
                                                  "values": [12.0, 30.0, 18.0, 25.0]}])) == []

    def test_a_new_pie_is_reported_with_its_quarters(self):
        [finding] = check(chart_shape(self.QUARTERS, [{"name": "Sales",
                                                       "values": [8.2, 3.2, 1.4, 1.2]}]))
        assert "categories named '1st Qtr'" in finding["message"]
        assert "sample numbers" in finding["message"]

    def test_a_chart_whose_data_is_unknown_is_not_judged(self):
        assert check(shape("Chart 3", 100, 100, 500, 300, type=msoChart)) == []


def test_a_chart_in_the_file_is_read(tmp_path):
    chart_xml = (
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea>'
        '<c:barChart><c:ser><c:tx><c:strRef><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Series 1</c:v></c:pt>'
        '</c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Category 1</c:v></c:pt>'
        '</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>4.3</c:v></c:pt>'
        '</c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>')
    frame = ('<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Chart 3"/></p:nvGraphicFramePr>'
             '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
             '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId9"/>'
             '</a:graphicData></a:graphic></p:graphicFrame>')
    path = deck_file(tmp_path, [frame], {})
    with zipfile.ZipFile(path, "a") as archive:
        archive.writestr("ppt/charts/chart1.xml", chart_xml)
        rels = archive.read("ppt/slides/_rels/slide1.xml.rels").decode()
    # Rewrite the slide's relationships with the chart in them.
    entries = {n: zipfile.ZipFile(path).read(n) for n in zipfile.ZipFile(path).namelist()}
    entries["ppt/slides/_rels/slide1.xml.rels"] = rels.replace(
        "</Relationships>", '<Relationship Id="rId9" Target="../charts/chart1.xml" Type="chart"/></Relationships>').encode()
    with zipfile.ZipFile(path, "w") as archive:
        for n, data in entries.items():
            archive.writestr(n, data)
    from utils.package import package_facts
    chart = package_facts(path, [1])["charts"][1]["Chart 3"]
    assert chart["categories"] == ["Category 1"] and chart["series"][0]["name"] == "Series 1"



# ---------------------------------------------------------------------------
# A title's words held to the body's edge
# ---------------------------------------------------------------------------
class TestTitleEdge:
    @staticmethod
    def big(shape):
        shape["text"]["runs"] = [{"color_hex": "#222222", "size": 32.0, "bold": True}]
        return shape

    def test_a_title_box_set_on_the_cards_edge_starts_its_words_inside_it(self):
        # The title frame at 40 like the cards, its words at 47.2.
        shapes = titled_row()
        self.big(shapes[0])
        [finding] = [f for f in check(*shapes) if f["type"] == "title_edge"]
        assert finding["shapes"] == ["title", "card0"]
        assert finding["off_by_pt"] == 7.2
        # The row is fitted again from the title's words to where it ended.
        assert finding["fix"].startswith(
            "Bring the body's left edge to the title's words: ppt_update_shape "
            "shape_names=['card0', 'card1', 'card2'] width=277.6")
        assert "The gaps stay 20 pt." in finding["fix"]

    def test_a_single_shape_is_moved_and_narrowed(self):
        title = self.big(textbox("title", 40, 30, 880, 55.74, z=1, text_width=300, text_height=48))
        chart = card("chart", 40, 140, 880, 300, z=2)
        [finding] = [f for f in check(title, chart) if f["type"] == "title_edge"]
        assert finding["fix"].endswith(
            "ppt_update_shape shape_name='chart' dleft=7.2 dwidth=-7.2.")

    def test_words_on_the_edge_are_not(self):
        shapes = titled_row()
        shapes[0] = self.big(textbox("title", 32.8, 30, 880, 55.74, z=1, text_width=300, text_height=48))
        assert [f for f in check(*shapes) if f["type"] == "title_edge"] == []

    def test_a_title_set_well_apart_is_a_design(self):
        shapes = titled_row()
        shapes[0] = self.big(textbox("title", 80, 30, 800, 55.74, z=1, text_width=300, text_height=48))
        assert [f for f in check(*shapes) if f["type"] == "title_edge"] == []

    def test_a_row_reaching_left_to_the_title_widens_as_one(self):
        # Widening only the first card made it the odd one out, and the
        # uneven-sizes fix took the change back.
        title = self.big(textbox("title", 40, 30, 880, 56, text_width=400, text_height=40))
        cards = [card(f"card{i}", left, 140, 268, 180, z=2 + i)
                 for i, left in enumerate((52, 352, 652))]
        [finding] = [f for f in check(title, *cards) if f["type"] == "title_edge"]
        assert "shape_names=['card0', 'card1', 'card2'] width=269.6" in finding["fix"]
        for c, by in zip(cards, (-4.8, -3.2, -1.6), strict=True):
            c["left"], c["width"] = c["left"] + by, 269.6
        assert [f for f in check(title, *cards)
                if f["type"] in ("title_edge", "uneven_sizes", "uneven_gaps")] == []

    def test_centred_text_keeps_no_edge_to_bring_in(self):
        # Its words start wherever their length puts them, and move half as
        # far as its frame does: the fix only ever halved the gap.
        title = self.big(textbox("title", 40, 30, 880, 56, text_width=400, text_height=40))
        caption = set_to(textbox("caption", 40, 100, 880, 30, text_width=856, text_height=20),
                         "center")
        assert [f for f in check(title, caption) if f["type"] == "title_edge"] == []

    def test_left_set_text_under_a_right_set_title_is_not_held_to_its_right(self):
        title = self.big(placed("title", 40, 30, 880, 55.74, 920 - 7.2 - 300, 33.6, 300, 48))
        note = textbox("note", 40, 140, 880, 30, text_width=858, text_height=20)
        assert [f for f in check(title, note) if f["type"] == "title_edge"] == []

    def test_a_write_does_not_take_the_line_it_touched_for_the_title(self):
        # Live on the Mac: a write reads only what it touched, so the title's
        # words were unknown and the touched line under it was taken for it.
        title = shape("title", 40, 30, 880, 50, unread=True)
        line = self.big(placed("line", 40, 120, 880, 40, 920 - 7.2 - 150, 123.6, 150, 30))
        below = placed("below", 40, 200, 874.6, 40, 914.6 - 7.2 - 140, 203.6, 140, 30)
        page = slide(title, line, below)
        assert [f for f in slide_findings(page, W, H, only={"line"})
                if f["type"] == "title_edge"] == []
        # Read in full, the slide has its real title and the line is body.
        del title["unread"]
        assert layout_check.find_title({"shapes": [line, below]}, H) is line

    def test_shapes_a_write_leaves_unread_are_marked(self):
        at = {"a": (0, 0), "b": (5, 5), "far": (500, 400)}
        shapes = layout_check.read_shapes(
            ["a", "b", "far"], lambda ref, z: shape(ref, *at[ref], 10, 10, z=z),
            lambda facts, ref, full: facts, {"a"})
        assert [s.get("unread") for s in shapes] == [None, None, True]

    def test_a_right_set_title_is_held_to_the_right_edge(self):
        shapes = titled_row()
        shapes[0] = self.big(placed("title", 40, 30, 880, 55.74, 920 - 7.2 - 300, 33.6, 300, 48))
        [finding] = [f for f in check(*shapes) if f["type"] == "title_edge"]
        assert finding["edge"] == "right"
        assert finding["fix"].startswith("Bring the body's right edge to the title's words")

    def test_the_recorded_decks(self):
        found = {}
        for path in sorted(FIXTURES.glob("*fixed*.json")):
            deck = json.loads(path.read_text(encoding="utf-8"))
            slides = [s["slide_index"] for s in deck["slides"]
                      for f in slide_findings(s, deck["slide_width"], deck["slide_height"])
                      if f["type"] == "title_edge"]
            if slides:
                found[path.stem] = slides
        assert found == self.SEEN

    # Each a title box or placeholder over cards, a chart, a table or a
    # picture: its words 7.2 pt in from the frame, 8.6 pt where its frame's
    # insets are wider.
    SEEN = {
        "companydeck_fixed_mac": [4, 5], "companydeck_fixed_win": [4, 5, 11],
        "mmdeck_fixed_win": [4, 5], "realdeck2_fixed_mac": [2, 3, 11],
        "realdeck_fixed_mac": [3, 4, 7, 8, 9, 12], "retrodeck_fixed_win": [6, 8],
        "rtldeck_fixed_mac": [3, 7], "rtldeck_fixed_win": [3, 7],
        "themedeck_fixed_mac": [9], "wispdeck_fixed_win": [3],
    }


# ---------------------------------------------------------------------------
# Fonts this computer does not have, or does not draw with
# ---------------------------------------------------------------------------
def styled(name, top, font, content="Quarterly results", **text_extra):
    box = textbox(name, 100, top, 600, 40, runs=[
        {"color_hex": "#222222", "size": 20.0, "bold": False, "font": font}])
    box["text"]["content"] = content
    box["text"].update(text_extra)
    return box


class TestFonts:
    def test_a_font_not_installed_is_reported_with_the_slides_own_font(self):
        page = slide(styled("body", 100, "Calibri"), styled("title", 30, "Frutiger Next Pro"))
        layout_check._with_fonts([page], lambda: frozenset({"calibri"}))
        [finding] = slide_findings(page, W, H)
        assert finding["type"] == "font_missing" and finding["shapes"] == ["title"]
        assert "ppt_format_text font_name='Calibri'" in finding["fix"]

    def test_a_theme_font_is_not_judged(self):
        page = slide(styled("body", 100, "+mn-lt"))
        layout_check._with_fonts([page], lambda: frozenset({"calibri"}))
        assert slide_findings(page, W, H) == []

    def test_the_list_is_not_read_when_no_text_names_a_font(self):
        page = slide(card("c", 10, 10, 50, 50))
        layout_check._with_fonts([page], lambda: pytest.fail("read the fonts"))

    def test_hebrew_drawn_in_another_font_than_asked_is_reported(self):
        page = slide(styled("title", 30, "David", content="תולדות ישראל", script_font="Arial"))
        [finding] = slide_findings(page, W, H)
        assert finding["message"].startswith("The Hebrew in 'title' is drawn in 'Arial', not 'David'")
        assert finding["fix"].startswith("Set the font again (ppt_format_text font_name='David')")

    def test_hebrew_in_the_font_asked_for_is_not(self):
        page = slide(styled("title", 30, "David", content="תולדות ישראל", script_font="David"))
        assert slide_findings(page, W, H) == []


class TestOtherFontNames:
    def test_a_name_the_system_list_leaves_out_is_found_in_the_font_files(self, monkeypatch):
        # A deck made in Japan asks for 游ゴシック, which an English Windows
        # lists as Yu Gothic; Japanese Macs name Hiragino the same way.
        from utils import fonts

        folders = []

        def names_in_folder(folder):
            folders.append(folder)
            return {"游ゴシック", "Yu Gothic"} if folder == "fonts" else set()

        monkeypatch.setattr(fonts, "names_in_folder", names_in_folder)
        installed = layout_check.FontNames({"Yu Gothic"}, ("fonts", "more fonts"))
        assert "yu gothic" in installed and folders == []
        assert "游ゴシック" in installed
        assert "no such font" not in installed
        assert folders == ["fonts", "more fonts"]

    def test_only_the_name_table_is_read(self, tmp_path):
        from utils.fonts import file_names

        class Counted:
            def __init__(self, data):
                self.data, self.at, self.read_bytes = data, 0, 0

            def seek(self, at):
                self.at = at

            def read(self, length):
                self.read_bytes += length
                return self.data[self.at:self.at + length]

        data = TestFontNames.font_file()
        padded = data + b"\0" * 1_000_000
        handle = Counted(padded)
        assert file_names(handle) == {"Calibri Light", "Calibri"}
        assert handle.read_bytes < 1000


class TestFontNames:
    @staticmethod
    def font_file():
        def record(name_id, text):
            return struct.pack(">HHHHHH", 3, 1, 0x409, name_id, len(text.encode("utf-16-be")), 0), text.encode("utf-16-be")

        entries, strings = [], b""
        for name_id, text in ((1, "Calibri Light"), (16, "Calibri")):
            head, raw = record(name_id, text)
            entries.append(head[:10] + struct.pack(">H", len(strings)))
            strings += raw
        table = struct.pack(">HHH", 0, len(entries), 6 + 12 * len(entries)) + b"".join(entries) + strings
        return struct.pack(">IHHHH", 0x00010000, 1, 0, 0, 0) + b"name" + struct.pack(">III", 0, 28, len(table)) + table

    def test_a_font_files_names_are_read(self):
        from utils.fonts import font_names

        assert font_names(self.font_file()) == {"Calibri Light", "Calibri"}


class TestTitleCapitals:
    @staticmethod
    def worded_title(index, words, **text_extra):
        page = titled(index)
        page["shapes"][0]["text"]["content"] = words
        page["shapes"][0]["text"].update(text_extra)
        return page

    def test_a_title_typed_in_capitals_among_normal_ones_is_reported(self):
        slides = [self.worded_title(2, "Where we stand"), self.worded_title(3, "What moved"),
                  self.worded_title(4, "NEXT STEPS")]
        [finding] = deck_findings(slides, H, W)
        assert finding["message"] == ("Slide titles differ in capitals: normal on slides 2, 3; "
                                      "capitals on slide 4.")
        assert finding["fix"] == ("Set the titles on slide 4 in normal case (ppt_set_text "
                                  "with the title retyped in normal case).")

    def test_capitals_the_template_sets_are_its_style(self):
        slides = [self.worded_title(2, "Where we stand", all_caps=True),
                  self.worded_title(3, "What moved", all_caps=True),
                  self.worded_title(4, "NEXT STEPS")]
        assert deck_findings(slides, H, W) == []

    def test_hebrew_has_no_capitals(self):
        slides = [self.worded_title(2, "איפה אנחנו"), self.worded_title(3, "מה השתנה"),
                  self.worded_title(4, "NEXT STEPS")]
        assert deck_findings(slides, H, W) == []


class TestContrastSuggestsTheTheme:
    THEME = {"dark1": "#1B2A41", "light1": "#FAF7F2", "dark2": "#44546A", "light2": "#E7E6E6"}

    def grey_on(self, background):
        grey = textbox("grey", 100, 100, 300, 40,
                       runs=[{"color_hex": "#777777", "size": 14.0, "bold": False}])
        page = slide(grey, background=background)
        page["theme_colors"] = self.THEME
        return page

    def test_a_write_that_knows_no_other_text_suggests_the_themes_colour(self):
        [finding] = slide_findings(self.grey_on("#FFFFFF"), W, H, only={"grey"})
        assert "make the text #1B2A41, the theme's text colour" in finding["fix"].lower().replace("#1b2a41", "#1B2A41")

    def test_on_a_dark_background_the_themes_light_colour(self):
        [finding] = slide_findings(self.grey_on("#101010"), W, H, only={"grey"})
        assert "#FAF7F2, the theme's text colour" in finding["fix"]

    def test_the_slides_own_text_colour_comes_first(self):
        page = self.grey_on("#FFFFFF")
        page["shapes"].append(textbox("body", 100, 200, 300, 40, z=2,
                                      runs=[{"color_hex": "#0B2545", "size": 14.0, "bold": False}]))
        [finding] = [f for f in slide_findings(page, W, H) if f["type"] == "low_contrast"]
        assert "#0B2545, the colour of the other text on the slide" in finding["fix"]


class TestTopAndBottomMargins:
    def row_at(self, bottom_gap):
        shapes = titled_row()
        for c in shapes[1:]:
            c["height"] = H - bottom_gap - c["top"]
        return shapes

    def test_cards_run_almost_to_the_bottom_are_reported_as_a_row(self):
        [found] = [f for f in check(*self.row_at(6)) if f["type"] == "narrow_margin"]
        assert found["edge"] == "bottom" and found["distance_pt"] == 6
        assert found["shapes"] == ["card0", "card1", "card2"]
        # 54 pt between the title and the cards: room to move up.
        assert found["fix"] == ("Move 'card0', 'card1' and 'card2' up by 24 pt "
                                "(ppt_update_shape shape_names=['card0', 'card1', 'card2'] dtop=-24).")

    def test_with_no_room_above_the_cards_are_shortened(self):
        shapes = self.row_at(6)
        for c in shapes[1:]:
            c["top"], c["height"] = 100, H - 6 - 100
        [found] = [f for f in check(*shapes) if f["type"] == "narrow_margin"]
        assert "24 pt shorter, keeping their top" in found["fix"]
        assert "dheight=-24" in found["fix"]

    def test_shortened_cards_keep_their_margin(self):
        shapes = self.row_at(6)
        for c in shapes[1:]:
            c["top"], c["height"] = 100, H - 6 - 100 - 24
        assert [f for f in check(*shapes) if f["type"] == "narrow_margin"] == []

    def text_under_a_card(self, words):
        title = textbox("Title", 40, 30, 880, 56, text_width=400, text_height=40)
        return title, card("Card", 40, 100, 880, 300), textbox(
            "Notes", 40, 410, 880, 130, text_width=600, text_height=words)

    def test_text_with_no_room_is_cut_not_its_box_shortened(self):
        # A shorter box leaves top-set words where they are and spills them,
        # and the same finding came back after every shortening.
        [found] = [f for f in check(*self.text_under_a_card(120)) if f["type"] == "narrow_margin"]
        assert found["fix"].startswith("Cut the text in 'Notes' or make the font smaller, "
                                       "so it ends 23.6 pt further from the edge")
        assert "dheight" not in found["fix"]

    def bottom_set(self, height):
        title, other, _ = self.text_under_a_card(0)
        other["height"] = 340
        notes = textbox("Notes", 40, 410, 880, height, text_width=600, text_height=80)
        notes["text"]["top"] = 410 + height - 3.6 - 80
        return title, other, notes

    def test_words_set_to_the_bottom_come_up_with_a_shorter_box(self):
        # Cutting them would leave their last line where it is.
        [found] = [f for f in check(*self.bottom_set(124)) if f["type"] == "narrow_margin"]
        assert "20.4 pt shorter, keeping its top" in found["fix"]
        assert [f for f in check(*self.bottom_set(124 - 20.4))
                if f["type"] in ("narrow_margin", "overflow")] == []

    def test_cut_text_keeps_its_margin(self):
        assert [f for f in check(*self.text_under_a_card(120 - 23.6))
                if f["type"] in ("narrow_margin", "overflow")] == []

    def test_a_card_whose_words_would_spill_is_not_shortened(self):
        title, other, _ = self.text_under_a_card(0)
        full = card("Full", 40, 410, 880, 124, z=3,
                    words=text(width=600, height=110, left=47.2, top=413.6))
        [found] = [f for f in check(title, other, full) if f["type"] == "narrow_margin"]
        assert found["fix"].startswith("Cut the text in 'Full'")

    def test_cards_that_keep_room_are_not(self):
        assert [f for f in check(*self.row_at(40)) if f["type"] == "narrow_margin"] == []

    def test_a_bleed_off_the_bottom_is_not(self):
        assert [f for f in check(*self.row_at(0)) if f["type"] == "narrow_margin"] == []

    def test_a_footnote_set_low_is_placed_there(self):
        shapes = titled_row() + [textbox("note", 40, H - 22, 400, 18, z=9, text_height=14)]
        assert [f for f in check(*shapes) if f["type"] == "narrow_margin"] == []


class TestTemplateRules:
    RULE = graphic("Title rule", 40, 100, 880, 1.5, source="layout")

    def title(self, lines):
        title = placed("Title 1", 40, 30, 880, 70, 47.2, 40, 600, 40 * lines)
        title["placeholder_type"] = 1
        title["type"] = msoPlaceholder
        title["text"]["runs"] = [{"color_hex": "#0B2545", "size": 32.0, "bold": True}]
        return title

    def test_a_title_that_wrapped_across_the_rule_is_reported(self):
        [found] = [f for f in check_on([self.RULE], self.title(2)) if f["type"] == "overlap"]
        assert found["shapes"] == ["Title 1", "Title rule"]
        assert "a line the template draws, drawn by the slide's layout" in found["message"]
        assert found["fix"].startswith("Keep 'Title 1' to the room above the line")

    def test_one_line_above_the_rule_is_not(self):
        assert [f for f in check_on([self.RULE], self.title(1)) if f["type"] == "overlap"] == []

    def test_a_card_set_over_the_rule_takes_back_its_top(self):
        over = card("card", 40, 90, 300, 100, z=2)
        [found] = [f for f in check_on([self.RULE], over) if f["type"] == "overlap"]
        assert found["fix"].startswith("Make 'card' 82.5 pt tall, keeping its far edge "
                                       "(ppt_update_shape top=107.5 height=82.5)")

    FOOTER = graphic("Footer rule", 40, 500, 880, 1.5, source="master")

    def test_a_card_over_a_footer_rule_stays_on_the_slide(self):
        # It was told to move down 387.5 pt, to 507-892 on a 540 pt slide.
        over = card("card", 40, 120, 880, 385, z=2)
        [found] = [f for f in check_on([self.FOOTER], over) if f["type"] == "overlap"]
        assert found["fix"].startswith("Make 'card' 374 pt tall (ppt_update_shape height=374)")

    def test_a_card_with_words_is_resized_not_its_words_cut(self):
        over = card("card", 40, 120, 880, 385, z=2,
                    words=text(width=400, height=30, left=47.2, top=123.6))
        [found] = [f for f in check_on([self.FOOTER], over) if f["type"] == "overlap"]
        assert found["message"].startswith("'card' runs across")
        assert found["fix"].startswith("Make 'card' 374 pt tall")

    def test_a_card_is_not_narrowed_past_its_words(self):
        # Narrower than its words, the card wraps them onto more lines than
        # it was measured with, and they spill.
        rule = graphic("Side rule", 280, 50, 1.5, 400, source="layout")
        over = card("card", 100, 150, 200, 100, z=2,
                    words=text(width=180, height=80, left=107.2, top=153.6))
        [found] = [f for f in check_on([rule], over) if f["type"] == "overlap"]
        assert "174 pt wide" not in found["fix"]
        assert found["fix"].startswith("Move 'card' left")

    def test_a_card_its_words_fill_is_moved_rather_than_shortened(self):
        over = card("card", 40, 120, 880, 385, z=2,
                    words=text(width=800, height=370, left=47.2, top=123.6))
        [found] = [f for f in check_on([self.FOOTER], over) if f["type"] == "overlap"]
        assert found["fix"].startswith("Move 'card' up by 11 pt (ppt_update_shape dtop=-11)")
        over["top"] -= 11
        assert [f for f in check_on([self.FOOTER], over) if f["type"] == "overlap"] == []

    def test_a_slanting_line_is_not_a_rule(self):
        slant = graphic("Slant", 600, 0, 200, 540, source="layout")
        slant["type"] = 9   # msoLine
        assert [f for f in check_on([slant], self.title(2)) if f["type"] == "overlap"] == []


class TestATitleOffItsLayout:
    def page(self, index, layout, left=40.0, width=880.0, home=(40.0, 30.0, 880.0, 60.0)):
        title = textbox("Title 1", left, 30, width, 60)
        title["placeholder_type"] = 1
        title["type"] = msoPlaceholder
        page = slide(title, index=index, layout=layout)
        page["layout_title"] = list(home)
        return page

    def test_a_title_nudged_on_a_layout_used_once_is_reported(self):
        slides = [self.page(2, "Title and Content"), self.page(3, "Title and Content"),
                  self.page(4, "Quote", left=44.0, width=876.0)]
        [finding] = layout_check._titles_off_their_layout(slides)
        assert finding["slides"] == [4]
        assert "left 44 pt where the layout has 40 pt" in finding["message"]
        assert finding["fix"] == ("Put it back (ppt_update_shape shape_name='Title 1' left=40 "
                                  "top=30 width=880 height=60).")

    def test_a_layout_used_twice_is_left_to_the_titles_comparison(self):
        slides = [self.page(2, "Quote", left=44.0), self.page(3, "Quote")]
        assert layout_check._titles_off_their_layout(slides) == []

    def test_a_title_moved_far_is_a_decision(self):
        slides = [self.page(4, "Cover", left=200.0)]
        assert layout_check._titles_off_their_layout(slides) == []

    def grown(self, words, left=40.0):
        page = self.page(4, "Section", left=left, home=(40.0, 30.0, 880.0, 56.0))
        title = page["shapes"][0]
        title["height"], title["text"]["height"] = 65.2, words
        return page

    def test_a_title_grown_to_hold_its_words_is_left_taller(self):
        # The overflow fix grew it from 56 to 65.2 pt; putting it back spilled
        # the words again, and the two findings took turns.
        page = self.grown(58)
        assert [f["type"] for f in slide_findings(page, W, H) if f["type"] == "overflow"] == []
        assert layout_check._titles_off_their_layout([page]) == []

    def test_a_grown_title_that_also_moved_keeps_its_height(self):
        [finding] = layout_check._titles_off_their_layout([self.grown(58, left=44.0)])
        assert "height" not in finding["message"]
        assert finding["fix"].endswith("left=40 top=30 width=880 height=65.2).")

    def test_a_title_grown_past_its_words_is_reported(self):
        [finding] = layout_check._titles_off_their_layout([self.grown(30)])
        assert finding["fix"].endswith("height=56).")


def test_a_layouts_own_graphics_and_title_are_read_from_the_file(tmp_path):
    from utils.package import layout_facts

    ns = ('xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
          'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
    layout = (f'<p:sldLayout {ns} showMasterSp="0"><p:cSld><p:spTree>'
              '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>'
              '<p:spPr><a:xfrm><a:off x="508000" y="381000"/><a:ext cx="11176000" cy="762000"/></a:xfrm></p:spPr></p:sp>'
              '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Rule"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
              '<p:spPr><a:xfrm><a:off x="508000" y="1270000"/><a:ext cx="11176000" cy="19050"/></a:xfrm>'
              '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill></p:spPr></p:sp>'
              '</p:spTree></p:cSld></p:sldLayout>')
    master = (f'<p:sldMaster {ns}><p:cSld><p:spTree/></p:cSld>'
              '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"/></p:sldMaster>')
    theme = ('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements>'
             '<a:clrScheme name="x"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>'
             '<a:accent1><a:srgbClr val="C0504D"/></a:accent1></a:clrScheme>'
             '<a:fontScheme name="x"><a:majorFont><a:latin typeface="Calibri Light"/><a:cs typeface=""/>'
             '</a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:cs typeface="Arial"/>'
             '</a:minorFont></a:fontScheme></a:themeElements></a:theme>')
    path = deck_file(tmp_path, [], {})
    entries = {n: zipfile.ZipFile(path).read(n) for n in zipfile.ZipFile(path).namelist()}
    entries["ppt/slides/_rels/slide1.xml.rels"] = RELS.format(
        '<Relationship Id="rId1" Target="../slideLayouts/slideLayout1.xml" Type="slideLayout"/>').encode()
    entries["ppt/slideLayouts/slideLayout1.xml"] = layout.encode()
    entries["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = RELS.format(
        '<Relationship Id="rId1" Target="../slideMasters/slideMaster1.xml" Type="slideMaster"/>').encode()
    entries["ppt/slideMasters/slideMaster1.xml"] = master.encode()
    entries["ppt/slideMasters/_rels/slideMaster1.xml.rels"] = RELS.format(
        '<Relationship Id="rId1" Target="../theme/theme1.xml" Type="theme"/>').encode()
    entries["ppt/theme/theme1.xml"] = theme.encode()
    with zipfile.ZipFile(path, "w") as archive:
        for n, data in entries.items():
            archive.writestr(n, data)
    found = layout_facts(path, [1])[1]
    assert found["title"] == [40.0, 30.0, 880.0, 60.0]
    assert found["shows_master"] is False
    [rule] = found["graphics"]
    assert (rule["name"], rule["top"], rule["height"]) == ("Rule", 100.0, 1.5)
    assert rule["fill"]["color_hex"] == "#C0504D"
    assert rule["visible"] is True
    # The theme's fonts, which the Mac holds a chosen right-to-left font to.
    assert found["theme_fonts"] == {"latin": ["Calibri Light", "Calibri"], "script": ["Arial"]}
    # A rule the layout hides is not drawn, and nothing runs across it.
    entries["ppt/slideLayouts/slideLayout1.xml"] = layout.replace(
        'name="Rule"/>', 'name="Rule" hidden="1"/>').encode()
    with zipfile.ZipFile(path, "w") as archive:
        for n, data in entries.items():
            archive.writestr(n, data)
    [rule] = layout_facts(path, [1])[1]["graphics"]
    assert rule["visible"] is False


class TestInsideGroups:
    def grouped(self, *children, under=()):
        group = shape("Group 4", 100, 100, 400, 200, z=2, type=6)   # msoGroup
        page = slide(*under, group)
        for order, child in enumerate(children, start=1):
            child["name"] = "Group 4/" + child["name"]
            child["in_group"] = "Group 4"
            child["z"] = 2 + order / 1000
        page["inner"] = list(children)
        return page

    def test_stand_in_words_inside_a_group_are_reported_by_their_path(self):
        words = textbox("TextBox 2", 110, 110, 300, 40)
        words["text"]["content"] = "Lorem ipsum dolor"
        [finding] = slide_findings(self.grouped(words), W, H)
        assert finding["type"] == "sample_text" and finding["shapes"] == ["Group 4/TextBox 2"]

    def test_text_spilling_out_of_its_box_inside_a_group(self):
        words = textbox("TextBox 2", 110, 110, 300, 40, text_height=80)
        [finding] = slide_findings(self.grouped(words), W, H)
        assert finding["type"] == "overflow"

    def test_white_words_on_a_dark_card_in_the_group_read(self):
        dark = card("Card", 100, 100, 400, 200, colour="#1F2937")
        words = textbox("TextBox 2", 110, 110, 300, 40,
                        runs=[{"color_hex": "#FFFFFF", "size": 14.0, "bold": False}])
        assert slide_findings(self.grouped(dark, words), W, H) == []

    def test_white_words_in_a_group_on_a_dark_slide_card_read(self):
        dark = card("Backing", 90, 90, 420, 220, z=1, colour="#1F2937")
        words = textbox("TextBox 2", 110, 110, 300, 40,
                        runs=[{"color_hex": "#FFFFFF", "size": 14.0, "bold": False}])
        found = [f for f in slide_findings(self.grouped(words, under=(dark,)), W, H)
                 if f["type"] == "low_contrast"]
        assert found == []

    def test_white_words_in_a_group_on_white_do_not(self):
        words = textbox("TextBox 2", 110, 110, 300, 40,
                        runs=[{"color_hex": "#FFFFFF", "size": 14.0, "bold": False}])
        [finding] = slide_findings(self.grouped(words), W, H)
        assert finding["type"] == "low_contrast" and finding["shapes"][0] == "Group 4/TextBox 2"


def test_the_text_inside_a_group_is_read_from_the_file(tmp_path):
    from utils.package import group_texts

    group = ('<p:grpSp><p:nvGrpSpPr><p:cNvPr id="4" name="Group 4"/></p:nvGrpSpPr><p:grpSpPr/>'
             '<p:sp><p:nvSpPr><p:cNvPr id="5" name="TextBox 2"/></p:nvSpPr><p:spPr/>'
             '<p:txBody><a:p><a:pPr rtl="1"/><a:r><a:rPr sz="2000"><a:latin typeface="David"/>'
             '<a:cs typeface="Arial"/></a:rPr><a:t>1516 עד 1917</a:t></a:r></a:p></p:txBody></p:sp>'
             '</p:grpSp>')
    path = deck_file(tmp_path, [group], {})
    [inner] = group_texts(path, [1])[1]
    assert inner["name"] == "Group 4/TextBox 2" and inner["in_group"] == "Group 4"
    text = inner["text"]
    assert text["content"] == "1516 עד 1917" and text["directions"] == ["rtl"]
    assert text["runs"][0]["font"] == "David" and text["script_font"] == "Arial"


def group_file(tmp_path, body, outer="", inner="", member=""):
    group = (f'<p:grpSp><p:nvGrpSpPr><p:cNvPr id="4" name="Group 1"{outer}/></p:nvGrpSpPr><p:grpSpPr/>'
             f'<p:grpSp><p:nvGrpSpPr><p:cNvPr id="6" name="Group 2"{inner}/></p:nvGrpSpPr><p:grpSpPr/>'
             f'<p:sp><p:nvSpPr><p:cNvPr id="5" name="TextBox 3"{member}/></p:nvSpPr><p:spPr/>'
             f'<p:txBody>{body}</p:txBody></p:sp></p:grpSp></p:grpSp>')
    return deck_file(tmp_path, [group], {})


ENGLISH = ('<a:p><a:r><a:rPr><a:latin typeface="Helvetica"/><a:cs typeface="Times New Roman"/>'
           '</a:rPr><a:t>Quarterly results</a:t></a:r></a:p>')


class TestGroupTextFromTheFile:
    def test_english_has_no_right_to_left_font(self, tmp_path):
        # It was reported as drawn in its unused right-to-left font.
        from utils.package import group_texts

        [inner] = group_texts(group_file(tmp_path, ENGLISH), [1])[1]
        assert inner["text"]["script_font"] is None
        assert [f for f in layout_check._fonts(1, [inner]) if f["type"] == "font_missing"] == []

    @pytest.mark.parametrize("where", ["outer", "inner", "member"])
    def test_hidden_text_is_not_visible(self, tmp_path, where):
        from utils.package import group_texts

        [inner] = group_texts(group_file(tmp_path, ENGLISH, **{where: ' hidden="1"'}), [1])[1]
        assert inner["visible"] is False

    def test_a_group_in_a_group_is_ungrouped_in_turn(self, tmp_path):
        from utils.package import group_texts

        body = '<a:p><a:r><a:t>Lorem ipsum dolor</a:t></a:r></a:p>'
        [inner] = group_texts(group_file(tmp_path, body), [1])[1]
        assert inner["groups"] == ["Group 1", "Group 2"]
        page = slide()
        page["inner"] = [inner]
        [finding] = slide_findings(page, W, H)
        assert finding["fix"].startswith(
            "Ungroup 'Group 1' and then 'Group 2' first (ppt_ungroup_shapes "
            "shape_name_or_index='Group 1', then ppt_ungroup_shapes shape_name_or_index='Group 2')")

    def test_a_line_break_and_a_field_are_kept(self, tmp_path):
        from utils.package import group_texts

        body = ('<a:p><a:r><a:t>Hello</a:t></a:r><a:br/><a:r><a:t>World </a:t></a:r>'
                '<a:fld id="{1}" type="slidenum"><a:t>3</a:t></a:fld></a:p>')
        [inner] = group_texts(group_file(tmp_path, body), [1])[1]
        assert inner["text"]["content"] == "Hello\vWorld 3"

    def test_a_direction_set_by_the_shapes_list_style_is_read(self, tmp_path):
        from utils.package import group_texts

        body = ('<a:lstStyle><a:lvl1pPr rtl="1"/></a:lstStyle>'
                '<a:p><a:r><a:t>שלום</a:t></a:r></a:p><a:p><a:pPr rtl="0"/><a:r><a:t>שלום</a:t></a:r></a:p>')
        [inner] = group_texts(group_file(tmp_path, body), [1])[1]
        assert inner["text"]["directions"] == ["rtl", "ltr"]

    def test_a_direction_set_for_the_whole_deck_is_read(self, tmp_path):
        from utils.package import group_texts

        path = group_file(tmp_path, '<a:p><a:r><a:t>שלום</a:t></a:r></a:p>')
        entries = {n: zipfile.ZipFile(path).read(n) for n in zipfile.ZipFile(path).namelist()}
        entries["ppt/presentation.xml"] = PRESENTATION.replace(
            "</p:presentation>",
            '<p:defaultTextStyle><a:lvl1pPr xmlns:a="http://schemas.openxmlformats.org/'
            'drawingml/2006/main" rtl="1"/></p:defaultTextStyle></p:presentation>').encode()
        with zipfile.ZipFile(path, "w") as archive:
            for n, data in entries.items():
                archive.writestr(n, data)
        [inner] = group_texts(path, [1])[1]
        assert inner["text"]["directions"] == ["rtl"]


# ---------------------------------------------------------------------------
# Text on a photo
# ---------------------------------------------------------------------------
from utils.bitmap import sample_bmp  # noqa: E402


def bmp(width, height, pixel):
    """A 24-bit BMP whose pixel (x, y), from the top, is pixel(x, y)."""
    stride = (width * 3 + 3) & ~3
    rows = b""
    for y in range(height - 1, -1, -1):
        row = b"".join(bytes(reversed(pixel(x, y))) for x in range(width))
        rows += row + b"\x00" * (stride - len(row))
    header = struct.pack("<2sIHHI", b"BM", 54 + len(rows), 0, 0, 54)
    info = struct.pack("<IiiHHIIiiII", 40, width, height, 1, 24, 0, len(rows), 2835, 2835, 0, 0)
    return header + info + rows


class TestTextOnAPhoto:
    def photo_slide(self, top_half, bottom_half, text_colour="#FFFFFF", text_top=110):
        photo = picture("Photo", 0, 0, 960, 540)
        photo["photo"] = sample_bmp(bmp(96, 54, lambda x, y: top_half if y < 27 else bottom_half))
        words = textbox("Caption", 100, text_top, 400, 60, z=2, runs=[
            {"color_hex": text_colour, "size": 28.0, "bold": True}])
        return slide(photo, words)

    def test_the_sampler_reads_rows_from_the_top(self):
        grid = sample_bmp(bmp(4, 2, lambda x, y: (255, 0, 0) if y == 0 else (0, 0, 255)), columns=4)
        assert grid["colours"][:4] == ["#FF0000"] * 4 and grid["colours"][4:] == ["#0000FF"] * 4

    def test_white_text_on_the_sky_is_reported_with_a_colour_that_reads(self):
        [finding] = slide_findings(self.photo_slide((200, 220, 240), (40, 40, 40)), W, H)
        assert finding["type"] == "low_contrast" and finding["on_photo"]
        assert finding["shapes"] == ["Caption", "Photo"]
        assert finding["fix"].startswith("Make the text #000000")

    def test_white_text_on_the_dark_ground_reads(self):
        page = self.photo_slide((200, 220, 240), (40, 40, 40), text_top=380)
        assert slide_findings(page, W, H) == []

    def test_text_across_light_and_dark_gets_a_card_behind_it(self):
        page = self.photo_slide((250, 250, 250), (10, 10, 10), text_top=240)
        page["shapes"][1]["height"] = 120
        page["shapes"][1]["text"]["height"] = 100
        [finding] = slide_findings(page, W, H)
        assert "ppt_add_shape shape_type='rectangle'" in finding["fix"]
        # Added at the front and the text brought over it: 'behind_text' goes
        # under the lowest text on the slide, often the title, under the photo.
        assert "zorder" not in finding["fix"].split("ppt_set_shape_zorder")[0]
        assert ("ppt_set_shape_zorder shape_name='Caption' command='bring_to_front'"
                in finding["fix"])

    def test_following_the_card_fix_leaves_the_slide_clean(self):
        # The card sat 0.8 pt off the text box, which read as a slip.
        page = self.photo_slide((250, 250, 250), (10, 10, 10), text_top=240)
        caption = page["shapes"][1]
        caption["height"], caption["text"]["height"] = 120, 100
        [finding] = slide_findings(page, W, H)
        at = {k: float(v) for k, v in re.findall(r"\b(left|top|width|height)=([\d.]+)",
                                                  finding["fix"])}
        assert (at["left"], at["top"]) == (caption["left"], caption["top"])
        backing = card("Scrim", at["left"], at["top"], at["width"], at["height"], z=3,
                       colour="#000000")
        backing["fill"]["transparency"] = 0.35
        caption["z"] = 4
        caption["text"]["runs"][0]["color_hex"] = "#FFFFFF"
        page["shapes"].append(backing)
        assert slide_findings(page, W, H) == []

    def test_a_see_through_part_of_a_picture_is_not_judged(self):
        page = self.photo_slide((200, 220, 240), (40, 40, 40))
        photo = page["shapes"][0]["photo"]
        photo["colours"] = [None] * len(photo["colours"])
        assert slide_findings(page, W, H) == []

    def test_a_photo_never_sampled_is_not_judged(self):
        page = self.photo_slide((200, 220, 240), (40, 40, 40))
        del page["shapes"][0]["photo"]
        assert slide_findings(page, W, H) == []


def test_hebrew_in_the_themes_own_pair_of_fonts_is_its_design():
    page = slide(styled("body", 30, "Calibri", content="שלום, עולם.", script_font="Arial"))
    page["theme_fonts"] = {"latin": ["Calibri Light", "Calibri"], "script": ["Arial", "Arial"]}
    assert [f for f in slide_findings(page, W, H) if f["type"] == "font_missing"] == []


def test_a_theme_naming_no_right_to_left_font_leaves_it_to_powerpoint():
    # Live on Windows: the Office theme's complex script fonts are empty, and
    # Hebrew in a text box with no font asked for drew in Arial.
    page = slide(styled("body", 30, "Calibri", content="שלום, עולם.", script_font="Arial"))
    page["theme_fonts"] = {"latin": ["Calibri Light", "Calibri"], "script": []}
    assert [f for f in slide_findings(page, W, H) if f["type"] == "font_missing"] == []


THEME_FONTS = {"latin": ["Calibri Light", "Calibri"], "script": ["Arial", "Arial"]}


@pytest.mark.parametrize("latin", ["Calibri", "+mn-lt"])
def test_a_hebrew_font_chosen_beside_the_themes_latin_one_is_a_choice(latin):
    # The fix set the Hebrew to Calibri, which has no Hebrew letters, or to
    # font_name='+mn-lt'.
    page = slide(styled("body", 30, latin, content="שלום, עולם.", script_font="David"))
    page["theme_fonts"] = THEME_FONTS
    assert [f for f in slide_findings(page, W, H) if f["type"] == "font_missing"] == []


def test_a_hebrew_font_chosen_beside_a_named_latin_one_is_a_choice():
    page = slide(styled("body", 30, "Georgia", content="שלום, עולם.", script_font="David"))
    page["theme_fonts"] = THEME_FONTS
    assert [f for f in slide_findings(page, W, H) if f["type"] == "font_missing"] == []


def test_hebrew_left_in_the_themes_font_under_a_named_one_is_reported():
    page = slide(styled("title", 30, "David", content="תולדות ישראל", script_font="Arial"))
    page["theme_fonts"] = THEME_FONTS
    [finding] = [f for f in slide_findings(page, W, H) if f["type"] == "font_missing"]
    assert finding["fix"].startswith("Set the font again (ppt_format_text font_name='David')")


class TestCropToFit:
    def test_shape_keeps_the_box_and_crops_the_image(self):
        from ppt_com.advanced_ops import _fit_aspect, _fit_crops, _fit_size
        # A 3:2 image in a 16:9 box.
        aspect = _fit_aspect("shape", 960, 540)
        crops = _fit_crops(300, 200, aspect, 0.5)
        assert crops["left"] == crops["right"] == 0
        assert crops["top"] == pytest.approx(crops["bottom"])
        assert (200 - crops["top"] - crops["bottom"]) * aspect == pytest.approx(300)
        assert _fit_size(960, 540, aspect, "shape") == (960, 540)

    def test_a_ratio_fits_inside_the_box(self):
        from ppt_com.advanced_ops import _fit_aspect, _fit_size
        assert _fit_size(400, 400, _fit_aspect("16:9"), "16:9") == (400, 225)

    def test_square_is_as_before(self):
        from ppt_com.advanced_ops import _fit_aspect, _fit_crops
        assert _fit_crops(400, 300, _fit_aspect("square"), 0.5) == {
            "left": 50.0, "right": 50.0, "top": 0.0, "bottom": 0.0}

    def test_anything_else_is_refused(self):
        from ppt_com.advanced_ops import _fit_aspect
        with pytest.raises(ValueError, match="Supported values"):
            _fit_aspect("wide")
