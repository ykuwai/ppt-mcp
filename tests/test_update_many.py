"""Tests for updating several shapes in one call.

Pure Python. select_targets is the name arithmetic both platforms share, and
UpdateShapeInput is where a request that cannot mean anything is refused.
"""

import sys

sys.path.insert(0, "src")

import pytest
from pydantic import ValidationError

from ppt_com.shapes import UpdateShapeInput, select_targets

SLIDE = ["Picture 2", "Title 1", "Body 3", "Badge 4"]


class TestChoosingWhatMoves:
    """select_targets answers with positions on the slide, not names, and
    hands back a name it could not place so the caller can look in the groups.
    """

    def test_all_is_every_shape_in_slide_order(self):
        assert select_targets(SLIDE, None, True, None) == [0, 1, 2, 3]

    def test_all_with_exclude_leaves_the_background_alone(self):
        assert select_targets(SLIDE, None, True, ["Picture 2"]) == [1, 2, 3]

    def test_a_list_comes_back_in_the_order_it_was_given(self):
        assert select_targets(SLIDE, ["Badge 4", "Title 1"], False, None) == [3, 1]

    def test_a_name_that_is_not_at_the_top_level_comes_back_as_itself(self):
        # The caller then tries the groups before calling it missing.
        assert select_targets(SLIDE, ["Title 1", "Nope"], False, None) == [1, "Nope"]

    def test_every_such_name_comes_back_not_just_the_first(self):
        assert select_targets(SLIDE, ["Nope", "Also nope"], False, None) == [
            "Nope", "Also nope"]

    def test_exclude_works_on_a_list_too(self):
        assert select_targets(
            SLIDE, ["Title 1", "Body 3"], False, ["Body 3"]) == [1]

    def test_excluding_something_absent_asks_for_nothing(self):
        # The caller said to leave it alone, and it is already alone.
        assert select_targets(SLIDE, ["Title 1", "Nope"], False, ["Nope"]) == [1]

    def test_excluding_everything_moves_nothing_and_is_not_an_error(self):
        assert select_targets(SLIDE, None, True, SLIDE) == []

    def test_an_empty_slide_with_all_moves_nothing(self):
        assert select_targets([], None, True, None) == []

    def test_a_name_asked_for_twice_is_taken_twice(self):
        # Applying the same offset twice to one shape is the caller's
        # business; silently collapsing it would be a different surprise.
        assert select_targets(SLIDE, ["Title 1", "Title 1"], False, None) == [1, 1]


class TestWhenTwoShapesShareAName:
    """PowerPoint does not stop a slide holding two shapes called the same
    thing. Going back through the name would move the first of them twice and
    leave the second where it was, reporting both as done.
    """

    TWINS = ["Picture 2", "Rectangle 5", "Body 3", "Rectangle 5"]

    def test_all_touches_both_of_them_once_each(self):
        assert select_targets(self.TWINS, None, True, None) == [0, 1, 2, 3]

    def test_and_still_leaves_out_what_was_excluded(self):
        assert select_targets(self.TWINS, None, True, ["Body 3"]) == [0, 1, 3]

    def test_excluding_the_shared_name_leaves_out_both(self):
        assert select_targets(self.TWINS, None, True, ["Rectangle 5"]) == [0, 2]

    def test_naming_it_means_the_first_one(self):
        # One name cannot pick between them, and the first is what every
        # other tool means by it.
        assert select_targets(self.TWINS, ["Rectangle 5"], False, None) == [1]


def rejected(**kwargs):
    with pytest.raises(ValidationError) as caught:
        UpdateShapeInput(slide_index=1, **kwargs)
    return str(caught.value)


class TestSayingWhichShapes:
    def test_one_shape_by_name_still_works(self):
        assert UpdateShapeInput(slide_index=1, shape_name="A", top=10).shape_name == "A"

    def test_several_by_name(self):
        params = UpdateShapeInput(slide_index=1, shape_names=["A", "B"], dtop=-26)
        assert params.shape_names == ["A", "B"]

    def test_the_whole_slide(self):
        assert UpdateShapeInput(slide_index=1, all=True, dtop=-26).all is True

    def test_saying_nothing_at_all_is_refused(self):
        assert "Say which shapes to update" in rejected(top=10)

    def test_two_ways_at_once_are_refused(self):
        assert "not shape_name and all" in rejected(shape_name="A", all=True, top=10)

    def test_an_empty_list_is_refused(self):
        assert "must not be empty" in rejected(shape_names=[], dtop=-26)

    def test_exclude_without_a_set_to_exclude_from_is_refused(self):
        assert "exclude only means something" in rejected(
            shape_name="A", exclude=["B"], dtop=-26)


class TestPositionsAndOffsets:
    @pytest.mark.parametrize("absolute,relative", [
        ("left", "dleft"), ("top", "dtop"),
        ("width", "dwidth"), ("height", "dheight"),
    ])
    def test_an_axis_takes_a_position_or_an_offset_not_both(self, absolute, relative):
        message = rejected(shape_name="A", **{absolute: 10, relative: 5})
        assert f"{absolute} and {relative} are mutually exclusive" in message

    def test_a_position_on_one_axis_and_an_offset_on_another_is_fine(self):
        params = UpdateShapeInput(slide_index=1, shape_name="A", left=10, dtop=-26)
        assert (params.left, params.dtop) == (10, -26)


class TestWhatOnlyMakesSenseForOneShape:
    def test_renaming_several_shapes_is_refused(self):
        # It would make duplicates.
        assert "name applies to one shape" in rejected(
            shape_names=["A", "B"], name="C")

    def test_and_renaming_the_whole_slide(self):
        assert "name applies to one shape" in rejected(all=True, name="C")

    def test_adjustment_handles_across_several_shapes_are_refused(self):
        # Handle 1 means a different thing on a triangle and on a callout.
        assert "adjustments applies to one shape" in rejected(
            all=True, adjustments={1: 0.5})

    def test_but_both_are_fine_on_one_shape(self):
        params = UpdateShapeInput(
            slide_index=1, shape_name="A", name="B", adjustments={1: 0.5})
        assert params.name == "B"
