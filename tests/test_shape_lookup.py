"""Tests for the shared shape resolver, including the way into a group.

Pure Python over a stand in for the COM collections, so it runs anywhere.
"""

import sys

sys.path.insert(0, "src")

import pytest

from ppt_com.constants import msoGroup
from ppt_com.shape_lookup import (
    require_top_level,
    resolve_shape,
    resolve_with_path,
    walk_group_children,
)


class Collection:
    """A COM collection: called with a 1-based index, and has a Count."""

    def __init__(self, items):
        self._items = list(items)

    def __call__(self, index):
        return self._items[index - 1]

    @property
    def Count(self):
        return len(self._items)


class FakeShape:
    def __init__(self, name, children=None, shape_type=1):
        self.Name = name
        self.Type = msoGroup if children is not None else shape_type
        self.GroupItems = Collection(children or [])


class FakeSlide:
    def __init__(self, *shapes):
        self.Shapes = Collection(shapes)


def group(name, *children):
    return FakeShape(name, children=list(children))


def badge_slide():
    """The template slide from the issue.

    A badge group of a rounded rectangle carrying the label and a triangle
    forming the speech bubble tail.
    """
    return FakeSlide(
        FakeShape("TextBox 19"),
        group("Group 20", FakeShape("二等辺三角形 23"),
              FakeShape("Rounded Rectangle 22")),
    )


class TestTheOldBehaviourIsUntouched:
    def test_a_shape_at_the_top_level_is_found_by_name(self):
        slide = badge_slide()
        assert resolve_shape(slide, "TextBox 19").Name == "TextBox 19"

    def test_an_index_still_means_the_top_level(self):
        slide = badge_slide()
        assert resolve_shape(slide, 2).Name == "Group 20"

    def test_an_index_past_the_end_says_the_range(self):
        with pytest.raises(ValueError, match=r"out of range \(1-2\)"):
            resolve_shape(badge_slide(), 3)

    def test_an_index_below_one_does_too(self):
        with pytest.raises(ValueError, match="out of range"):
            resolve_shape(badge_slide(), 0)

    def test_a_name_that_is_nowhere_keeps_the_old_message(self):
        with pytest.raises(ValueError, match="Shape 'Nope' not found on slide"):
            resolve_shape(badge_slide(), "Nope")


class TestReachingIntoAGroup:
    def test_a_child_is_found_by_its_own_name(self):
        slide = badge_slide()
        assert resolve_shape(slide, "Rounded Rectangle 22").Name == (
            "Rounded Rectangle 22")

    def test_and_by_its_path(self):
        slide = badge_slide()
        shape = resolve_shape(slide, "Group 20/Rounded Rectangle 22")
        assert shape.Name == "Rounded Rectangle 22"

    def test_a_child_of_a_nested_group_is_reachable(self):
        slide = FakeSlide(group("Outer", group("Inner", FakeShape("Deep"))))
        assert resolve_shape(slide, "Deep").Name == "Deep"

    def test_and_by_its_full_path(self):
        slide = FakeSlide(group("Outer", group("Inner", FakeShape("Deep"))))
        assert resolve_shape(slide, "Outer/Inner/Deep").Name == "Deep"

    def test_a_path_with_a_level_missing_finds_nothing(self):
        # A stale path must not quietly edit the shape two levels down.
        slide = FakeSlide(group("Outer", group("Inner", FakeShape("Deep"))))
        with pytest.raises(ValueError, match="not found on slide"):
            resolve_shape(slide, "Outer/Deep")

    def test_a_path_whose_last_segment_is_wrong_is_not_found(self):
        slide = badge_slide()
        with pytest.raises(ValueError, match="not found on slide"):
            resolve_shape(slide, "Group 20/Nope")

    def test_a_path_whose_group_is_wrong_is_not_found_either(self):
        slide = badge_slide()
        with pytest.raises(ValueError, match="not found on slide"):
            resolve_shape(slide, "Nope/Rounded Rectangle 22")

    def test_the_message_names_the_groups_worth_looking_in(self):
        with pytest.raises(ValueError, match="ppt_get_group_items"):
            resolve_shape(badge_slide(), "Nope")

    def test_and_does_not_when_the_slide_has_no_group(self):
        slide = FakeSlide(FakeShape("TextBox 19"))
        with pytest.raises(ValueError) as caught:
            resolve_shape(slide, "Nope")
        assert str(caught.value) == "Shape 'Nope' not found on slide"


class TestWhenTwoShapesShareAName:
    """Generated names are unique in practice, and a duplicated slide or a
    pasted group can still produce two of the same.
    """

    def test_the_one_at_the_top_level_wins(self):
        slide = FakeSlide(
            FakeShape("Rectangle 5"),
            group("Group 20", FakeShape("Rectangle 5")),
        )
        # The caller who has not thought about groups meant this one.
        assert resolve_shape(slide, "Rectangle 5") is slide.Shapes(1)

    def test_two_children_with_one_name_ask_for_a_path(self):
        slide = FakeSlide(
            group("Group A", FakeShape("Rectangle 5")),
            group("Group B", FakeShape("Rectangle 5")),
        )
        with pytest.raises(ValueError) as caught:
            resolve_shape(slide, "Rectangle 5")
        assert "Group A/Rectangle 5" in str(caught.value)
        assert "Group B/Rectangle 5" in str(caught.value)

    def test_and_the_path_then_resolves(self):
        slide = FakeSlide(
            group("Group A", FakeShape("Rectangle 5")),
            group("Group B", FakeShape("Rectangle 5")),
        )
        shape = resolve_shape(slide, "Group B/Rectangle 5")
        assert shape is slide.Shapes(2).GroupItems(1)


class TestANameWithASlashInIt:
    """A person can call a shape "A/B". The whole string is tried as a name
    before it is read as a path, so such a shape is found as itself.
    """

    def test_a_top_level_shape_named_with_a_slash(self):
        slide = FakeSlide(FakeShape("A/B"))
        assert resolve_shape(slide, "A/B").Name == "A/B"

    def test_a_child_named_with_a_slash(self):
        slide = FakeSlide(group("G", FakeShape("A/B")))
        assert resolve_shape(slide, "A/B").Name == "A/B"

    def test_and_a_real_path_still_wins_when_there_is_no_such_name(self):
        slide = FakeSlide(group("A", FakeShape("B")))
        assert resolve_shape(slide, "A/B").Name == "B"


class TestWalkingAGroup:
    def test_every_child_comes_back_with_its_path(self):
        slide = FakeSlide(group("Outer", FakeShape("One"),
                                group("Inner", FakeShape("Two"))))
        walked = [(shape.Name, path) for shape, path
                  in walk_group_children(slide.Shapes(1), "Outer/")]
        assert walked == [
            ("One", "Outer/One"),
            ("Inner", "Outer/Inner"),
            ("Two", "Outer/Inner/Two"),
        ]

    def test_a_shape_that_is_not_a_group_yields_nothing(self):
        assert list(walk_group_children(FakeShape("TextBox 19"))) == []


class TestTheToolsThatCannotTakeAChild:
    """Aligning, distributing, merging and grouping go through Shapes.Range,
    which only addresses the top level.
    """

    def test_top_level_names_pass(self):
        require_top_level(badge_slide(), ["TextBox 19", "Group 20"], 1,
                          "ppt_align_shapes")

    def test_a_child_is_told_where_it_is_and_why_it_cannot_be_used(self):
        with pytest.raises(ValueError) as caught:
            require_top_level(badge_slide(), ["Rounded Rectangle 22"], 7,
                              "ppt_align_shapes")
        message = str(caught.value)
        assert "slide 7" in message
        assert "Group 20" in message
        assert "ppt_align_shapes" in message
        assert "ppt_ungroup_shapes" in message

    def test_a_name_that_is_nowhere_keeps_the_old_message(self):
        with pytest.raises(ValueError) as caught:
            require_top_level(badge_slide(), ["Nope"], 7, "ppt_group_shapes")
        assert str(caught.value) == "Shape 'Nope' not found on slide 7"

    def test_a_path_is_recognised_as_a_child_too(self):
        # ppt_get_group_items hands out paths; passing one here used to be
        # reported as a shape that does not exist.
        with pytest.raises(ValueError) as caught:
            require_top_level(badge_slide(),
                              ["Group 20/Rounded Rectangle 22"], 7,
                              "ppt_merge_shapes")
        assert "inside group 'Group 20'" in str(caught.value)


class TestWhatATextSearchLooksAt:
    """ppt_find_replace_text walked past a group without saying so, because a
    group has no text frame of its own.
    """

    @staticmethod
    def _walk(include_groups):
        from ppt_com.text import _searchable_shapes

        return [(shape.Name, path) for shape, path
                in _searchable_shapes(badge_slide(), include_groups)]

    def test_by_default_only_the_top_level(self):
        assert self._walk(False) == [
            ("TextBox 19", "TextBox 19"),
            ("Group 20", "Group 20"),
        ]

    def test_with_include_groups_the_children_come_too(self):
        assert self._walk(True) == [
            ("TextBox 19", "TextBox 19"),
            ("Group 20", "Group 20"),
            ("二等辺三角形 23", "Group 20/二等辺三角形 23"),
            ("Rounded Rectangle 22", "Group 20/Rounded Rectangle 22"),
        ]

    def test_the_group_itself_is_still_offered(self):
        # It has no text frame, so the caller skips it; leaving it out of the
        # walk would be a second place to get that decision wrong.
        assert ("Group 20", "Group 20") in self._walk(True)


class TestThePathAShapeReportsBack:
    """resolve_with_path answers where the shape turned out to live, which is
    what ppt_get_group_items puts in front of its members.
    """

    def test_a_top_level_shape_is_its_own_path(self):
        _, path = resolve_with_path(badge_slide(), "TextBox 19")
        assert path == "TextBox 19"

    def test_an_index_answers_the_name(self):
        _, path = resolve_with_path(badge_slide(), 2)
        assert path == "Group 20"

    def test_a_child_found_by_bare_name_answers_its_full_path(self):
        _, path = resolve_with_path(badge_slide(), "Rounded Rectangle 22")
        assert path == "Group 20/Rounded Rectangle 22"

    def test_a_nested_group_answers_every_level(self):
        slide = FakeSlide(group("Outer", group("Inner", FakeShape("Deep"))))
        _, path = resolve_with_path(slide, "Inner")
        assert path == "Outer/Inner"

    def test_which_is_what_a_nested_member_path_has_to_be_built_on(self):
        slide = FakeSlide(group("Outer", group("Inner", FakeShape("Deep"))))
        _, path = resolve_with_path(slide, "Outer/Inner")
        assert path + "/Deep" == "Outer/Inner/Deep"
        assert resolve_shape(slide, path + "/Deep").Name == "Deep"


class TestAGroupOrChildNamedWithASlash:
    """A segment may contain the separator, so every way of reading the
    string is tried rather than splitting it once and hoping.
    """

    def test_a_child_named_with_a_slash_is_reachable_by_its_path(self):
        slide = FakeSlide(group("G", FakeShape("A/B")))
        assert resolve_shape(slide, "G/A/B").Name == "A/B"

    def test_a_group_named_with_a_slash_is_too(self):
        slide = FakeSlide(group("G/H", FakeShape("Child")))
        assert resolve_shape(slide, "G/H/Child").Name == "Child"

    def test_two_children_named_with_a_slash_are_still_separable(self):
        slide = FakeSlide(group("G1", FakeShape("A/B")),
                          group("G2", FakeShape("A/B")))
        with pytest.raises(ValueError, match="more than one group"):
            resolve_shape(slide, "A/B")
        assert resolve_shape(slide, "G2/A/B") is slide.Shapes(2).GroupItems(1)

    def test_the_path_ppt_get_group_items_reports_resolves(self):
        slide = FakeSlide(group("G", FakeShape("A/B")))
        path = "G" + "/" + slide.Shapes(1).GroupItems(1).Name
        assert resolve_shape(slide, path).Name == "A/B"
