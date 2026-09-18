"""Tests for where a newly added shape lands in the stack.

Pure Python over a stand in slide that implements PowerPoint's four z-order
commands, so the walk can be checked without PowerPoint.
"""

import sys

sys.path.insert(0, "src")

import pytest

from ppt_com.constants import (
    msoBringForward,
    msoBringToFront,
    msoSendBackward,
    msoSendToBack,
)
from ppt_com.shapes import place_in_zorder


class Collection:
    def __init__(self, items):
        self._items = items

    def __call__(self, index):
        return self._items[index - 1]

    @property
    def Count(self):
        return len(self._items)


class FakeShape:
    def __init__(self, name, text=None, children=None):
        self.Name = name
        self._text = text
        self.slide = None
        self.Type = 6 if children is not None else 1  # msoGroup
        self.GroupItems = Collection(list(children or []))

    # --- text frame ---
    @property
    def HasTextFrame(self):
        if self.Type == 6:
            # A group answers nothing here, the way COM does.
            raise RuntimeError("groups have no text frame")
        return self._text is not None

    @property
    def TextFrame(self):
        shape = self

        class Frame:
            HasText = shape._text is not None and shape._text != ""

        return Frame()

    # --- z order ---
    @property
    def ZOrderPosition(self):
        return self.slide.order.index(self) + 1

    def ZOrder(self, command):
        order = self.slide.order
        i = order.index(self)
        order.pop(i)
        if command == msoSendToBack:
            order.insert(0, self)
        elif command == msoBringToFront:
            order.append(self)
        elif command == msoBringForward:
            order.insert(min(i + 1, len(order)), self)
        elif command == msoSendBackward:
            order.insert(max(i - 1, 0), self)
        else:
            raise AssertionError(f"unexpected command {command}")


class FakeSlide:
    def __init__(self, *shapes):
        self.order = list(shapes)
        for shape in self.order:
            shape.slide = self
            for child in shape.GroupItems._items:
                child.slide = self

    @property
    def Shapes(self):
        return Collection(self.order)

    def names(self):
        return [shape.Name for shape in self.order]


def deck_slide():
    """The slide the issue describes: a full bleed background, then text."""
    return FakeSlide(
        FakeShape("Background"),
        FakeShape("Caption", text="ここが本文"),
        FakeShape("Heading", text="見出し"),
    )


def add(slide, name, text=None):
    """A newly added shape, which PowerPoint always puts at the front."""
    shape = FakeShape(name, text=text)
    shape.slide = slide
    slide.order.append(shape)
    return shape


class TestFront:
    def test_front_leaves_it_where_powerpoint_put_it(self):
        slide = deck_slide()
        art = add(slide, "Art")
        assert place_in_zorder(slide, art, "front") == {}
        assert slide.names()[-1] == "Art"

    def test_and_so_does_no_placement_at_all(self):
        slide = deck_slide()
        art = add(slide, "Art")
        assert place_in_zorder(slide, art, None) == {}
        assert art.ZOrderPosition == 4


class TestBack:
    def test_back_is_the_very_bottom(self):
        slide = deck_slide()
        art = add(slide, "Art")
        placed = place_in_zorder(slide, art, "back")
        assert slide.names() == ["Art", "Background", "Caption", "Heading"]
        assert placed == {"zorder": "back", "z_position": 1}

    def test_which_is_under_the_background_and_therefore_invisible(self):
        # Exactly the confusing intermediate state the issue describes.
        slide = deck_slide()
        art = add(slide, "Art")
        place_in_zorder(slide, art, "back")
        assert art.ZOrderPosition < slide.order[1].ZOrderPosition


class TestBehindText:
    def test_it_lands_above_the_background_and_below_the_text(self):
        slide = deck_slide()
        art = add(slide, "Art")
        placed = place_in_zorder(slide, art, "behind_text")
        assert slide.names() == ["Background", "Art", "Caption", "Heading"]
        assert placed == {"zorder": "behind_text", "z_position": 2}

    def test_with_no_background_it_is_simply_at_the_bottom(self):
        slide = FakeSlide(FakeShape("Caption", text="本文"))
        art = add(slide, "Art")
        place_in_zorder(slide, art, "behind_text")
        assert slide.names() == ["Art", "Caption"]

    def test_several_shapes_without_text_stay_below_it(self):
        slide = FakeSlide(
            FakeShape("Background"),
            FakeShape("Texture"),
            FakeShape("Caption", text="本文"),
        )
        art = add(slide, "Art")
        place_in_zorder(slide, art, "behind_text")
        assert slide.names() == ["Background", "Texture", "Art", "Caption"]

    def test_an_empty_text_frame_does_not_count_as_text(self):
        slide = FakeSlide(FakeShape("Background"), FakeShape("Empty", text=""))
        art = add(slide, "Art")
        placed = place_in_zorder(slide, art, "behind_text")
        assert placed["zorder"] == "front"
        assert "note" in placed

    def test_a_slide_with_no_text_leaves_it_at_the_front_and_says_so(self):
        slide = FakeSlide(FakeShape("Background"))
        art = add(slide, "Art")
        placed = place_in_zorder(slide, art, "behind_text")
        assert placed["zorder"] == "front"
        assert "behind_text" in placed["note"]
        # Not sent to the bottom, where it would be invisible.
        assert slide.names() == ["Background", "Art"]

    def test_the_shape_being_placed_does_not_count_as_text_itself(self):
        slide = FakeSlide(FakeShape("Background"))
        art = add(slide, "Art", text="わたしにも文字がある")
        placed = place_in_zorder(slide, art, "behind_text")
        assert placed["zorder"] == "front"

    def test_a_caption_inside_a_group_counts_as_text(self):
        # The badge group case. COM raises on a group's HasTextFrame, so the
        # question has to move inward or the art lands on top of the caption.
        slide = FakeSlide(
            FakeShape("Background"),
            FakeShape("Badge", children=[FakeShape("Label", text="テーマ")]),
        )
        art = add(slide, "Art")
        place_in_zorder(slide, art, "behind_text")
        assert slide.names() == ["Background", "Art", "Badge"]

    def test_a_group_with_nothing_written_in_it_does_not(self):
        slide = FakeSlide(
            FakeShape("Background"),
            FakeShape("Icons", children=[FakeShape("Star")]),
        )
        art = add(slide, "Art")
        placed = place_in_zorder(slide, art, "behind_text")
        assert placed["zorder"] == "front"

    def test_it_works_on_a_shape_that_was_already_on_the_slide(self):
        # ppt_set_shape_zorder send_behind_text, rather than a fresh add.
        slide = FakeSlide(
            FakeShape("Background"),
            FakeShape("Art"),
            FakeShape("Caption", text="本文"),
        )
        art = slide.order[1]
        place_in_zorder(slide, art, "behind_text")
        assert slide.names() == ["Background", "Art", "Caption"]

    def test_including_one_that_starts_above_the_text(self):
        slide = FakeSlide(
            FakeShape("Background"),
            FakeShape("Caption", text="本文"),
            FakeShape("Art"),
        )
        art = slide.order[2]
        place_in_zorder(slide, art, "behind_text")
        assert slide.names() == ["Background", "Art", "Caption"]


class TestAnUnknownPlacement:
    def test_is_refused_by_name(self):
        slide = deck_slide()
        art = add(slide, "Art")
        with pytest.raises(ValueError, match="Unknown zorder 'middle'"):
            place_in_zorder(slide, art, "middle")

    def test_and_the_message_lists_what_is_accepted(self):
        slide = deck_slide()
        art = add(slide, "Art")
        with pytest.raises(ValueError, match="behind_text"):
            place_in_zorder(slide, art, "middle")
