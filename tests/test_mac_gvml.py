"""Tests for the clipboard route to PowerPoint, over a fake object graph.

Nothing here launches PowerPoint or touches the real pasteboard. The fake deck
below models exactly what ``ppt_mac/gvml_paste.py`` and the three tool
modules ask of PowerPoint, and the fake pasteboard is a dict with a change
count. appscript only installs on macOS, so the file is skipped elsewhere;
the XML itself is tested in ``test_gvml.py``, which runs everywhere.

What is pinned here is the procedure and its refusals, the list in
docs/gvml-design.md section 7: the order of the steps, that a paste which
added nothing is refused rather than reported, that a clipboard rewritten
between the write and the paste stops the paste, that a navigation that failed
stops it too, that what landed wrong is removed before the refusal, that the
originals of a group are deleted only after the group is verified, and that
the user's clipboard comes back only when it is still ours to overwrite.

Two of the fakes' behaviours are taken from the machine rather than invented.
``paste_object`` on a deaf deck reports nothing and adds nothing, which is what
PowerPoint does with a package it cannot use. And the shape a paste adds is
named and typed from the package that was written, because that is what
PowerPoint was measured to do (design section 0).
"""

import json
import os
import sys
from contextlib import contextmanager
from unittest import mock

import pytest

sys.path.insert(0, "src")

macos_only = pytest.mark.skipif(
    sys.platform != "darwin", reason="the Apple Event backend needs appscript"
)

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures", "gvml")


def _fixture_bytes(name):
    with open(os.path.join(FIXTURES, f"{name}.gvml.zip"), "rb") as fh:
        return fh.read()


# ---------------------------------------------------------------------------
# The swap
# ---------------------------------------------------------------------------
@macos_only
class TestTheFifteenToolsAreTheClipboardOnes:
    @pytest.mark.parametrize(
        "module_name,impl_name",
        [
            ("ppt_com.charts", "_add_chart_impl"),
            ("ppt_com.charts", "_get_chart_data_impl"),
            ("ppt_com.freeform", "_build_freeform_impl"),
            ("ppt_com.freeform", "_get_shape_nodes_impl"),
            ("ppt_com.groups", "_group_shapes_impl"),
            ("ppt_com.groups", "_get_group_items_impl"),
            ("ppt_com.charts", "_set_chart_data_impl"),
            ("ppt_com.charts", "_change_chart_type_impl"),
            ("ppt_com.charts", "_format_chart_impl"),
            ("ppt_com.charts", "_format_chart_axis_impl"),
            ("ppt_com.charts", "_set_chart_series_impl"),
            ("ppt_com.freeform", "_set_node_position_impl"),
            ("ppt_com.freeform", "_insert_node_impl"),
            ("ppt_com.freeform", "_delete_node_impl"),
            ("ppt_com.freeform", "_set_segment_type_impl"),
        ],
    )
    def test_the_com_module_answers_with_the_apple_event_one(self, module_name, impl_name):
        import importlib

        module = importlib.import_module(module_name)
        assert getattr(module, impl_name).__module__ == module_name.replace("ppt_com", "ppt_mac")

    def test_gvml_paste_defines_no_impl_of_its_own(self):
        """MACOS_PORT's tool count walks every `_impl` in ppt_mac."""
        from ppt_mac import gvml_paste

        assert not [n for n in dir(gvml_paste) if n.endswith("_impl")]


# ---------------------------------------------------------------------------
# The procedure
# ---------------------------------------------------------------------------
@macos_only
class TestThePasteProcedure:
    """One order of steps, and every place it stops."""

    def test_the_steps_run_in_order_and_the_clipboard_comes_back(self):
        from ppt_mac.freeform import _build_freeform_impl

        with _fake_deck(shapes=[("Title", "auto")]) as deck:
            payload = json.loads(_build_freeform_impl(1, 1, 10.0, 20.0, _LINE_NODES, True, "Zig"))

        assert payload["success"] is True
        assert payload["shape_name"] == "Zig"
        assert payload["left"] == 10.0 and payload["top"] == 20.0
        assert deck.order == [
            "snapshot", "write", "goto", "unselect", "paste", "position", "restore",
        ]
        assert deck.board.contents == {"public.utf8-plain-text": b"kept"}
        assert "warnings" not in payload

    def test_a_paste_that_added_nothing_is_a_refusal_without_success(self):
        from ppt_mac.freeform import _build_freeform_impl

        with _fake_deck(shapes=[("Title", "auto")], deaf=True) as deck:
            payload = json.loads(_build_freeform_impl(1, 1, 0.0, 0.0, _LINE_NODES, False, None))

        assert payload["error"] == "ppt_build_freeform is not available on macOS"
        assert "silent no-op" in payload["reason"]
        assert "success" not in payload
        assert [s.name() for s in deck.shapes] == ["Title"]
        assert deck.order[-1] == "restore"

    def test_a_clipboard_rewritten_after_the_write_stops_the_paste(self):
        from ppt_mac.freeform import _build_freeform_impl

        with _fake_deck(shapes=[], rewrite_after_write=True) as deck:
            payload = json.loads(_build_freeform_impl(1, 1, 0.0, 0.0, _LINE_NODES, False, None))

        assert "paste" not in deck.order
        assert "written to by something else" in payload["reason"]
        assert "success" not in payload
        # Not ours any more, so not overwritten with the old contents either.
        assert deck.board.contents != {"public.utf8-plain-text": b"kept"}

    def test_a_navigation_that_failed_stops_the_paste(self):
        from ppt_mac.freeform import _build_freeform_impl

        with _fake_deck(shapes=[], goto_error=-1728) as deck:
            payload = json.loads(_build_freeform_impl(1, 1, 0.0, 0.0, _LINE_NODES, False, None))

        assert "paste" not in deck.order
        assert "-1728" in payload["reason"]

    def test_what_landed_as_the_wrong_type_is_removed_before_the_refusal(self):
        from ppt_mac.charts import _add_chart_impl

        with _fake_deck(shapes=[("Title", "auto")], lands_as="auto") as deck:
            result = _add_chart_impl(1, "column", 50, 50, 500, 350)

        assert "shape_type_auto" in result["reason"]
        assert "removed again" in result["reason"]
        assert [s.name() for s in deck.shapes] == ["Title"]
        assert "success" not in result

    def test_two_shapes_where_one_was_expected_are_both_removed(self):
        from ppt_mac.charts import _add_chart_impl

        with _fake_deck(shapes=[], lands_twice=True) as deck:
            result = _add_chart_impl(1, "column", 50, 50, 500, 350)

        assert "put 2 shapes" in result["reason"]
        assert deck.shapes == []

    def test_a_position_that_did_not_stick_is_a_warning_not_a_refusal(self):
        from ppt_mac.charts import _add_chart_impl

        with _fake_deck(shapes=[], stuck_position=(380.0, 210.0)):
            result = _add_chart_impl(1, "column", 50, 50, 500, 350)

        assert result["success"] is True
        assert "(380.0, 210.0)" in result["warnings"][0]

    def test_the_clipboard_is_left_alone_when_someone_else_wrote_meanwhile(self):
        from ppt_mac.charts import _add_chart_impl

        with _fake_deck(shapes=[], rewrite_after_paste=True) as deck:
            result = _add_chart_impl(1, "column", 50, 50, 500, 350)

        assert result["success"] is True
        assert "restore" not in deck.order
        assert "not put back" in result["warnings"][0]

    def test_a_clipboard_too_large_to_save_is_said_so(self):
        from ppt_mac.charts import _add_chart_impl

        with _fake_deck(shapes=[], too_large=True) as deck:
            result = _add_chart_impl(1, "column", 50, 50, 500, 350)

        assert result["success"] is True
        assert "restore" not in deck.order
        assert "could not be put back" in result["warnings"][0]

    def test_the_selection_is_cleared_before_every_paste(self):
        """The finding that a chart selection swallows the next paste."""
        from ppt_mac.charts import _add_chart_impl

        with _fake_deck(shapes=[]) as deck:
            _add_chart_impl(1, "column", 50, 50, 500, 350)
            _add_chart_impl(1, "pie", 50, 400, 300, 200)

        pastes = [i for i, step in enumerate(deck.order) if step == "paste"]
        assert len(pastes) == 2
        for i in pastes:
            assert deck.order[i - 1] == "unselect"
        assert [s.name() for s in deck.shapes] == ["Chart 1", "Chart 2"]


# ---------------------------------------------------------------------------
# Charts
# ---------------------------------------------------------------------------
@macos_only
class TestAddChart:
    def test_it_returns_what_windows_returns(self):
        from ppt_mac.charts import _add_chart_impl

        with _fake_deck(shapes=[("Title", "auto")]) as deck:
            result = _add_chart_impl(1, "pie", 50, 60, 500, 350)

        assert result == {
            "success": True, "shape_name": "Chart 1", "shape_index": 2,
            "chart_type": "pie", "chart_type_int": 5,
        }
        assert deck.shapes[-1].shape_type().AS_name == "shape_type_chart"
        assert (deck.shapes[-1].left_position(), deck.shapes[-1].top()) == (50, 60)

    def test_a_chart_type_without_a_template_refuses_by_argument(self):
        """An XlChartType the writer cannot draw names the argument, not the tool."""
        from ppt_mac.charts import _add_chart_impl

        with _no_powerpoint():
            result = _add_chart_impl(1, -4100, 50, 50, 500, 350)

        assert result["error"] == "ppt_add_chart cannot draw chart_type -4100 on macOS"
        assert "column" in result["reason"] and "pie" in result["reason"]
        assert "success" not in result

    def test_a_misspelled_type_is_still_heard_first(self):
        from ppt_mac.charts import _add_chart_impl

        with _no_powerpoint():
            with pytest.raises(ValueError, match="Unknown chart type 'colunm'"):
                _add_chart_impl(1, "colunm", 50, 50, 500, 350)


@macos_only
class TestGetChartData:
    def test_it_reads_the_caches_out_of_the_copied_package(self):
        from ppt_mac.charts import _get_chart_data_impl

        with _fake_deck(shapes=[("Sales", "chart", _fixture_bytes("chart"))]) as deck:
            result = _get_chart_data_impl(1, "Sales")

        assert result["success"] is True
        assert result["shape_name"] == "Sales"
        assert result["categories"] == ["Category 1", "Category 2", "Category 3", "Category 4"]
        assert result["series"][0] == {"name": "Series 1", "values": [4.3, 2.5, 3.5, 4.5]}
        assert deck.order == ["snapshot", "copy", "claim", "restore"]
        assert deck.board.contents == {"public.utf8-plain-text": b"kept"}

    def test_a_chart_whose_package_holds_no_chart_part_is_refused(self):
        from ppt_mac.charts import _get_chart_data_impl

        with _fake_deck(shapes=[("Sales", "chart", _fixture_bytes("rect"))]):
            result = _get_chart_data_impl(1, "Sales")

        assert "holds no chart part" in result["reason"]
        assert "success" not in result

    def test_a_copy_that_left_no_package_is_refused(self):
        from ppt_mac.charts import _get_chart_data_impl

        with _fake_deck(shapes=[("Sales", "chart", None)]):
            result = _get_chart_data_impl(1, "Sales")

        assert "put nothing of type" in result["reason"]
        assert "success" not in result

    def test_a_shape_that_is_not_a_chart_says_so_before_the_clipboard_is_touched(self):
        from ppt_mac.charts import _get_chart_data_impl

        with _fake_deck(shapes=[("Title", "auto")]) as deck:
            with pytest.raises(ValueError, match="'Title' is not a chart"):
                _get_chart_data_impl(1, "Title")

        assert deck.order == []


# ---------------------------------------------------------------------------
# Freeforms
# ---------------------------------------------------------------------------
@macos_only
class TestFreeformTools:
    def test_building_returns_a_json_string_in_the_windows_shape(self):
        from ppt_mac.freeform import _build_freeform_impl

        with _fake_deck(shapes=[]):
            raw = _build_freeform_impl(1, 1, 100.0, 150.0, _CURVE_NODES, True, None)

        assert isinstance(raw, str)
        payload = json.loads(raw)
        assert set(payload) == {"success", "shape_name", "shape_index", "left", "top", "width", "height"}
        assert payload["shape_name"] == "Freeform 1"
        assert (payload["left"], payload["top"]) == (100.0, 100.0)
        # The auto curve's control points reach past x=250, and the box holds them.
        assert (payload["width"], payload["height"]) == (163.33, 160.0)

    def test_reading_nodes_numbers_them_as_windows_does(self):
        from ppt_mac.freeform import _get_shape_nodes_impl

        with _fake_deck(shapes=[("Blob", "freeform", _fixture_bytes("freeform"))]) as deck:
            payload = json.loads(_get_shape_nodes_impl(1, "Blob", None))

        assert payload["shape_name"] == "Blob"
        assert payload["node_count"] == 9
        assert payload["nodes"][2]["segment_type"] == "inaccessible"
        assert payload["nodes"][2]["note"].startswith("Metadata not accessible")
        assert deck.order == ["snapshot", "copy", "claim", "restore"]

    def test_reading_a_shape_that_is_not_a_freeform_says_so_first(self):
        from ppt_mac.freeform import _get_shape_nodes_impl

        with _fake_deck(shapes=[("Title", "auto")]) as deck:
            with pytest.raises(ValueError, match="is not a freeform"):
                _get_shape_nodes_impl(1, "Title", None)

        assert deck.order == []

# ---------------------------------------------------------------------------
# Groups
# ---------------------------------------------------------------------------
@macos_only
class TestGroupShapes:
    def test_the_members_are_copied_pasted_as_one_group_then_deleted(self):
        from ppt_mac.groups import _group_shapes_impl

        with _fake_deck(shapes=[("Box", "auto"), ("Circle", "auto"), ("Other", "auto")]) as deck:
            result = _group_shapes_impl(1, ["Box", "Circle"])

        assert result == {"success": True, "group_name": "Group 1", "shape_index": 2}
        assert [s.name() for s in deck.shapes] == ["Other", "Group 1"]
        assert deck.order == [
            "snapshot", "copy", "claim", "copy", "claim", "write", "goto",
            "unselect", "paste", "position", "delete", "delete", "restore",
        ]
        # The group sits where its members were.
        assert (deck.shapes[-1].left_position(), deck.shapes[-1].top()) == (10.0, 20.0)

    def test_a_dropped_paste_leaves_the_originals_untouched(self):
        from ppt_mac.groups import _group_shapes_impl

        with _fake_deck(shapes=[("Box", "auto"), ("Circle", "auto")], deaf=True) as deck:
            result = _group_shapes_impl(1, ["Box", "Circle"])

        assert "success" not in result
        assert "delete" not in deck.order
        assert [s.name() for s in deck.shapes] == ["Box", "Circle"]

    def test_originals_that_would_not_delete_are_named_beside_the_group(self):
        from ppt_mac.groups import _group_shapes_impl

        with _fake_deck(shapes=[("Box", "auto"), ("Circle", "auto")], undeletable={"Circle"}) as deck:
            result = _group_shapes_impl(1, ["Box", "Circle"])

        assert result["error"] == "ppt_group_shapes left both the group and its originals on the slide"
        assert "Group 1" in result["reason"] and "['Circle']" in result["reason"]
        assert [s.name() for s in deck.shapes] == ["Circle", "Group 1"]

    def test_a_missing_member_is_refused_before_anything_is_copied(self):
        from ppt_mac.groups import _group_shapes_impl

        with _fake_deck(shapes=[("Box", "auto")]) as deck:
            with pytest.raises(ValueError, match="'Circle' not found on slide 1"):
                _group_shapes_impl(1, ["Box", "Circle"])

        assert deck.order == []

    def test_reading_members_comes_from_the_copied_package(self):
        from ppt_mac.groups import _get_group_items_impl

        with _fake_deck(shapes=[("Diagram", "group", _fixture_bytes("group"))]) as deck:
            result = _get_group_items_impl(1, "Diagram")

        assert result["success"] is True
        assert result["group_name"] == "Diagram"
        assert [i["name"] for i in result["items"]] == ["KidA", "KidB"]
        assert result["items"][0]["type_name"] == "AutoShape"
        assert result["items"][0]["width"] == 80.0
        assert deck.order == ["snapshot", "copy", "claim", "restore"]

    def test_a_group_whose_package_is_not_a_group_is_refused(self):
        from ppt_mac.groups import _get_group_items_impl

        with _fake_deck(shapes=[("Diagram", "group", _fixture_bytes("rect"))]):
            result = _get_group_items_impl(1, "Diagram")

        assert "expected a group" in result["reason"]
        assert "items" not in result


# ---------------------------------------------------------------------------
# Replacing a shape with a rewritten copy of itself
# ---------------------------------------------------------------------------
_CHART_STEPS = [
    "snapshot", "copy", "claim", "effects", "write", "goto", "unselect", "paste",
    "position", "copy", "claim", "delete",
]


@macos_only
class TestTheReplaceProcedure:
    """Copy, rewrite, paste, read back, then delete the original, then walk
    the new shape back down the z order. Every stop on that road."""

    def test_the_steps_run_in_order_and_the_chart_keeps_its_place(self):
        from ppt_mac.charts import _set_chart_data_impl

        deck_shapes = [("A", "auto"), ("Sales", "chart", _fixture_bytes("chart")), ("B", "auto")]
        with _fake_deck(shapes=deck_shapes) as deck:
            result = _set_chart_data_impl(1, "Sales", ["Q1"], [{"name": "S", "values": [1]}])

        assert result["success"] is True
        assert result["shape_name"] == "Sales"
        # Pasted last, so one step back puts it where the original was.
        assert deck.order == _CHART_STEPS + ["backward", "restore"]
        assert [s.name() for s in deck.shapes] == ["A", "Sales", "B"]
        assert deck.shapes[1].shape_type().AS_name == "shape_type_chart"
        assert (deck.shapes[1].left_position(), deck.shapes[1].top()) == (10.0, 20.0)
        assert deck.board.contents == {"public.utf8-plain-text": b"kept"}

    def test_the_walk_back_is_one_send_backward_per_step(self):
        from ppt_mac.charts import _set_chart_data_impl

        deck_shapes = [("Sales", "chart", _fixture_bytes("chart")), ("A", "auto"), ("B", "auto"), ("C", "auto")]
        with _fake_deck(shapes=deck_shapes) as deck:
            _set_chart_data_impl(1, "Sales", ["Q1"], [{"name": "S", "values": [1]}])

        assert deck.order.count("backward") == 3
        assert [s.name() for s in deck.shapes] == ["Sales", "A", "B", "C"]

    def test_the_caller_is_told_the_shape_was_recreated_in_full_then_briefly(self):
        from ppt_mac.charts import _set_chart_data_impl

        deck_shapes = [("Sales", "chart", _fixture_bytes("chart"))]
        with _fake_deck(shapes=deck_shapes):
            first = _set_chart_data_impl(1, "Sales", ["Q1"], [{"name": "S", "values": [1]}])
            second = _set_chart_data_impl(1, "Sales", ["Q2"], [{"name": "S", "values": [2]}])

        assert first["warnings"] == [
            "'Sales' was recreated rather than edited in place. PowerPoint for Mac "
            "cannot reach inside a chart from a script, so the shape was copied, its "
            "XML rewritten and pasted back, and the original deleted. Its name, "
            "position and z order were put back. Its animations were not, because "
            "the clipboard package does not carry them, and anything else that "
            "pointed at the old object (a comment anchor, an animation trigger on "
            "another shape) now points at nothing."
        ]
        assert second["warnings"] == [
            "'Sales' was recreated by paste; name, position and z order kept, animations not."
        ]

    def test_lost_animations_are_counted_before_the_original_goes(self):
        from ppt_mac.charts import _set_chart_data_impl

        deck_shapes = [("Sales", "chart", _fixture_bytes("chart")), ("A", "auto")]
        with _fake_deck(shapes=deck_shapes, effects=["Sales", "A", "Sales"]) as deck:
            result = _set_chart_data_impl(1, "Sales", ["Q1"], [{"name": "S", "values": [1]}])

        assert result["warnings"][1] == (
            "2 animation effect(s) on 'Sales' were lost with the original; the "
            "clipboard package does not carry them."
        )
        assert deck.order.index("effects") < deck.order.index("delete")

    def test_a_sequence_that_cannot_be_read_is_said_to_be_unknown_not_empty(self):
        from ppt_mac.charts import _set_chart_data_impl

        deck_shapes = [("Sales", "chart", _fixture_bytes("chart"))]
        with _fake_deck(shapes=deck_shapes, no_timeline=True):
            result = _set_chart_data_impl(1, "Sales", ["Q1"], [{"name": "S", "values": [1]}])

        assert "could not be read, so any it had are gone" in result["warnings"][1]

    def test_a_dropped_paste_leaves_the_original_untouched(self):
        from ppt_mac.charts import _set_chart_data_impl

        deck_shapes = [("Sales", "chart", _fixture_bytes("chart"))]
        with _fake_deck(shapes=deck_shapes, deaf=True) as deck:
            result = _set_chart_data_impl(1, "Sales", ["Q1"], [{"name": "S", "values": [1]}])

        assert "success" not in result
        assert "silent no-op" in result["reason"]
        assert "delete" not in deck.order
        assert [s.name() for s in deck.shapes] == ["Sales"]
        assert deck.shapes[0].package == _fixture_bytes("chart")

    def test_a_paste_that_reads_back_without_the_edit_is_removed_and_refused(self):
        """The read-back is the evidence. When it does not carry the change,
        the new shape goes and the original stays, and the refusal says both."""
        from ppt_mac.charts import _set_chart_data_impl

        deck_shapes = [("Sales", "chart", _fixture_bytes("chart"))]
        with _fake_deck(shapes=deck_shapes, stale_package=_fixture_bytes("chart")) as deck:
            result = _set_chart_data_impl(1, "Sales", ["Q1"], [{"name": "S", "values": [1]}])

        assert result["error"] == "ppt_set_chart_data pasted a chart that did not carry the edit"
        assert "categories read back as" in result["reason"]
        assert "The pasted copy was removed again. The original is untouched." in result["reason"]
        assert "success" not in result
        assert [s.name() for s in deck.shapes] == ["Sales"]
        assert deck.order[-1] == "restore" and "backward" not in deck.order

    def test_an_original_that_would_not_delete_is_named_beside_its_replacement(self):
        from ppt_mac.charts import _set_chart_data_impl

        deck_shapes = [("A", "auto"), ("Sales", "chart", _fixture_bytes("chart"))]
        with _fake_deck(shapes=deck_shapes, undeletable={"Sales"}) as deck:
            result = _set_chart_data_impl(1, "Sales", ["Q1"], [{"name": "S", "values": [1]}])

        assert result["error"] == "ppt_set_chart_data left both the original and its replacement on the slide"
        assert "the original at position 2 and the new one at position 3" in result["reason"]
        assert [s.name() for s in deck.shapes] == ["A", "Sales", "Sales"]

    def test_a_z_order_that_would_not_move_is_a_warning_with_the_position(self):
        from ppt_mac.charts import _set_chart_data_impl

        deck_shapes = [("Sales", "chart", _fixture_bytes("chart")), ("A", "auto")]
        with _fake_deck(shapes=deck_shapes, stuck_z=True) as deck:
            result = _set_chart_data_impl(1, "Sales", ["Q1"], [{"name": "S", "values": [1]}])

        assert result["success"] is True
        assert result["warnings"][-1] == (
            "'Sales' could not be put back at z order position 1; it is at position 2 now."
        )
        assert [s.name() for s in deck.shapes] == ["A", "Sales"]

    def test_a_shape_that_is_not_a_chart_says_so_before_the_clipboard_is_touched(self):
        from ppt_mac.charts import _set_chart_data_impl

        with _fake_deck(shapes=[("Title", "auto")]) as deck:
            with pytest.raises(ValueError, match="'Title' is not a chart"):
                _set_chart_data_impl(1, "Title", ["Q1"], [{"name": "S", "values": [1]}])

        assert deck.order == []


# ---------------------------------------------------------------------------
# The five chart editors
# ---------------------------------------------------------------------------
@macos_only
class TestChartEditors:
    """Each returns what Windows returns, plus the warnings, and what it
    wrote is what a second copy reads back."""

    def test_set_chart_data(self):
        from ppt_mac.charts import _get_chart_data_impl, _set_chart_data_impl

        with _fake_deck(shapes=[("Sales", "chart", _fixture_bytes("chart"))]):
            result = _set_chart_data_impl(1, "Sales", ["Q1", "Q2"], [
                {"name": "North", "values": [1, 2]}, {"name": "South", "values": [3, 4]},
            ])
            data = _get_chart_data_impl(1, "Sales")

        assert {k: v for k, v in result.items() if k != "warnings"} == {
            "success": True, "shape_name": "Sales", "categories_count": 2, "series_count": 2,
        }
        assert data["categories"] == ["Q1", "Q2"]
        assert data["series"] == [
            {"name": "North", "values": [1.0, 2.0]}, {"name": "South", "values": [3.0, 4.0]},
        ]

    def test_a_series_without_values_is_refused_before_anything_is_copied(self):
        from ppt_mac.charts import _set_chart_data_impl

        with _fake_deck(shapes=[("Sales", "chart", _fixture_bytes("chart"))]) as deck:
            with pytest.raises(ValueError, match="'name' and a 'values'"):
                _set_chart_data_impl(1, "Sales", ["Q1"], [{"name": "S"}])

        assert deck.order == []

    def test_change_chart_type(self):
        from ppt_mac.charts import _change_chart_type_impl, _get_chart_data_impl

        with _fake_deck(shapes=[("Sales", "chart", _fixture_bytes("chart"))]) as deck:
            result = _change_chart_type_impl(1, "Sales", "pie")
            data = _get_chart_data_impl(1, "Sales")

        assert {k: v for k, v in result.items() if k != "warnings"} == {
            "success": True, "shape_name": "Sales", "new_chart_type": "pie", "new_chart_type_int": 5,
        }
        from gvml import Package, charts

        assert charts.kind_of(Package.from_bytes(deck.shapes[0].package).chart()) == "pieChart"
        assert data["series"][0]["values"] == [4.3, 2.5, 3.5, 4.5]

    def test_a_combo_chart_changes_kind_with_all_of_its_series(self):
        """Two plot groups, one over a secondary axis. The series of both
        have to arrive in the new plot, or the read-back check sees series
        go missing, takes the pasted chart away again and refuses."""
        from conftest import combo_chart_xml
        from gvml import Package, charts

        from ppt_mac.charts import _change_chart_type_impl, _get_chart_data_impl

        package = Package.from_bytes(_fixture_bytes("chart"))
        package.parts[package.chart_part()] = combo_chart_xml(package.chart())
        with _fake_deck(shapes=[("Sales", "chart", package.to_bytes())]) as deck:
            result = _change_chart_type_impl(1, "Sales", "line")
            data = _get_chart_data_impl(1, "Sales")

        assert result["success"] is True and result["new_chart_type"] == "line"
        assert [s["name"] for s in data["series"]] == ["Series 1", "Series 2", "Series 3"]
        assert charts.kind_of(Package.from_bytes(deck.shapes[0].package).chart()) == "lineChart"

    def test_a_combo_chart_whose_groups_disagree_is_refused_by_name(self):
        from conftest import combo_chart_xml
        from gvml import Package, charts

        from ppt_mac.charts import _change_chart_type_impl

        package = Package.from_bytes(_fixture_bytes("chart"))
        root = charts.load(combo_chart_xml(package.chart()))
        c = f"{{{charts.NS_C}}}"
        cats = root.find(f"{c}chart/{c}plotArea/{c}lineChart/{c}ser/{c}cat")
        for i, pt in enumerate(cats.iter(f"{c}pt")):
            pt.find(f"{c}v").text = f"Week {i + 1}"
        package.parts[package.chart_part()] = charts.dump(root)
        before = package.to_bytes()
        with _fake_deck(shapes=[("Sales", "chart", before)]) as deck:
            payload = _change_chart_type_impl(1, "Sales", "line")

        assert "success" not in payload
        assert "'Series 3' against categories of its own" in payload["reason"]
        assert "paste" not in deck.order
        assert deck.shapes[0].package == before

    def test_format_chart(self):
        from ppt_mac.charts import _format_chart_impl

        with _fake_deck(shapes=[("Sales", "chart", _fixture_bytes("chart"))]):
            result = _format_chart_impl(
                1, "Sales", "Revenue", None, "top", None, None, None, None, None, None, None,
            )
            gone = _format_chart_impl(
                1, "Sales", None, False, None, None, None, None, None, None, None, None,
            )

        assert {k: v for k, v in result.items() if k != "warnings"} == {
            "success": True, "shape_name": "Sales", "has_title": True, "has_legend": True,
        }
        assert gone["has_legend"] is False and gone["has_title"] is True

    def test_a_legend_position_with_no_legend_is_the_windows_error_and_changes_nothing(self):
        from ppt_mac.charts import _format_chart_impl

        with _fake_deck(shapes=[("Sales", "chart", _fixture_bytes("chart"))]) as deck:
            _format_chart_impl(1, "Sales", None, False, None, None, None, None, None, None, None, None)
            with pytest.raises(ValueError, match="Cannot set legend position when chart has no legend"):
                _format_chart_impl(1, "Sales", None, None, "left", None, None, None, None, None, None, None)

        assert deck.order.count("paste") == 1

    def test_format_chart_axis(self):
        from ppt_mac.charts import _format_chart_axis_impl

        with _fake_deck(shapes=[("Sales", "chart", _fixture_bytes("chart"))]):
            result = _format_chart_axis_impl(
                1, "Sales", "value", "Yen",
                0.0, 100.0, 25.0, None, None, None, "cross", None,
                True, False, None, None, "0.0",
            )

        assert {k: v for k, v in result.items() if k != "warnings"} == {
            "success": True, "shape_name": "Sales", "axis": "value",
            "applied": ["title", "min_scale", "max_scale", "major_unit", "major_tick_mark",
                        "reverse_order", "log_scale", "number_format"],
        }

    def test_an_axis_the_chart_does_not_have_is_the_windows_error(self):
        from ppt_mac.charts import _change_chart_type_impl, _format_chart_axis_impl

        with _fake_deck(shapes=[("Sales", "chart", _fixture_bytes("chart"))]):
            _change_chart_type_impl(1, "Sales", "pie")
            with pytest.raises(ValueError, match="Axis 'value' is not available on this chart type"):
                _format_chart_axis_impl(1, "Sales", "value", "x", *([None] * 13))

    def test_set_chart_series(self):
        from ppt_mac.charts import _set_chart_series_impl

        with _fake_deck(shapes=[("Sales", "chart", _fixture_bytes("chart"))]) as deck:
            result = _set_chart_series_impl(1, "Sales", 2, "#FF0000", True, 2.0)

        assert {k: v for k, v in result.items() if k != "warnings"} == {
            "success": True, "shape_name": "Sales", "series_index": 2,
        }
        from gvml import Package, charts

        assert charts.read_series(Package.from_bytes(deck.shapes[0].package).chart(), 2) == {
            "color": "#FF0000", "line_weight": 2.0, "show_data_labels": True,
        }

    def test_a_series_past_the_end_is_the_windows_error(self):
        from ppt_mac.charts import _set_chart_series_impl

        with _fake_deck(shapes=[("Sales", "chart", _fixture_bytes("chart"))]) as deck:
            with pytest.raises(ValueError, match="series_index 7 out of range"):
                _set_chart_series_impl(1, "Sales", 7, "#FF0000", None, None)

        assert "paste" not in deck.order


# ---------------------------------------------------------------------------
# A call that asks for nothing
# ---------------------------------------------------------------------------
@macos_only
class TestACallThatAsksForNothing:
    """Three of the editors have nothing but optional arguments. A call that
    leaves them all out is a success on Windows, where no argument means no
    property set, and it must cost the chart nothing here either. Replacing
    the chart is how macOS edits one, and it loses the animations on it, so a
    call with nothing to write stops after the copy."""

    def test_format_chart_with_no_argument_leaves_the_chart_standing(self):
        from ppt_mac.charts import _NOTHING_ASKED, _format_chart_impl

        with _fake_deck(shapes=[("Sales", "chart", _fixture_bytes("chart"))], effects=["Sales"]) as deck:
            before = deck.shapes[0]
            result = _format_chart_impl(1, "Sales", *([None] * 10))

        assert {k: v for k, v in result.items() if k != "warnings"} == {
            "success": True, "shape_name": "Sales", "has_title": False,
            "has_legend": True, "note": _NOTHING_ASKED,
        }
        # The same object, with the same XML and its animation still on it.
        assert deck.shapes == [before]
        assert deck.shapes[0].package == _fixture_bytes("chart")
        assert deck.order == ["snapshot", "copy", "claim", "restore"]

    def test_format_chart_axis_with_no_argument_applies_nothing(self):
        from ppt_mac.charts import _NOTHING_ASKED, _format_chart_axis_impl

        with _fake_deck(shapes=[("Sales", "chart", _fixture_bytes("chart"))]) as deck:
            result = _format_chart_axis_impl(1, "Sales", "value", *([None] * 14))

        assert {k: v for k, v in result.items() if k != "warnings"} == {
            "success": True, "shape_name": "Sales", "axis": "value", "applied": [],
            "note": _NOTHING_ASKED,
        }
        assert "paste" not in deck.order

    def test_set_chart_series_with_no_argument_changes_no_series(self):
        from ppt_mac.charts import _NOTHING_ASKED, _set_chart_series_impl

        with _fake_deck(shapes=[("Sales", "chart", _fixture_bytes("chart"))]) as deck:
            result = _set_chart_series_impl(1, "Sales", 2, None, None, None)

        assert {k: v for k, v in result.items() if k != "warnings"} == {
            "success": True, "shape_name": "Sales", "series_index": 2, "note": _NOTHING_ASKED,
        }
        assert "paste" not in deck.order

    def test_the_target_is_still_checked_so_windows_still_raises(self):
        """The copy is what finds out whether the axis or the series is
        there, so the shortcut keeps it rather than answering from the
        arguments alone."""
        from ppt_mac.charts import (
            _change_chart_type_impl, _format_chart_axis_impl, _set_chart_series_impl,
        )

        with _fake_deck(shapes=[("Sales", "chart", _fixture_bytes("chart"))]) as deck:
            with pytest.raises(ValueError, match="series_index 7 out of range"):
                _set_chart_series_impl(1, "Sales", 7, None, None, None)
            _change_chart_type_impl(1, "Sales", "pie")
            with pytest.raises(ValueError, match="Axis 'value' is not available"):
                _format_chart_axis_impl(1, "Sales", "value", *([None] * 14))

        assert deck.order.count("paste") == 1

    @pytest.mark.parametrize("call,field", [
        (lambda impl: impl["format"](1, "Sales", None, False, *([None] * 8)), "has_legend"),
        (lambda impl: impl["axis"](1, "Sales", "value", *([None] * 9), False, *([None] * 4)), "log_scale"),
        (lambda impl: impl["series"](1, "Sales", 1, None, False, None), "show_data_labels"),
    ])
    def test_an_argument_that_is_false_is_still_an_argument(self, call, field):
        """False is a request, not a missing one, so these do replace the chart."""
        from ppt_mac.charts import (
            _format_chart_axis_impl, _format_chart_impl, _set_chart_series_impl,
        )

        impl = {"format": _format_chart_impl, "axis": _format_chart_axis_impl,
                "series": _set_chart_series_impl}
        with _fake_deck(shapes=[("Sales", "chart", _fixture_bytes("chart"))]) as deck:
            result = call(impl)

        assert result["success"] is True and "note" not in result
        assert deck.order.count("paste") == 1, field


# ---------------------------------------------------------------------------
# The four freeform editors
# ---------------------------------------------------------------------------
@macos_only
class TestFreeformEditors:
    """JSON strings, in the Windows shape, with the path read back."""

    def test_set_node_position_reads_the_position_back(self):
        from ppt_mac.freeform import _set_node_position_impl

        deck_shapes = [("A", "auto"), ("Blob", "freeform", _fixture_bytes("freeform"))]
        with _fake_deck(shapes=deck_shapes) as deck:
            payload = json.loads(_set_node_position_impl(1, "Blob", None, 2, 300.0, 70.0))

        assert {k: v for k, v in payload.items() if k != "warnings"} == {
            "success": True, "shape_name": "Blob", "node_index": 2, "x": 300.0, "y": 70.0,
        }
        assert payload["warnings"][0].startswith("'Blob' was recreated")
        assert [s.name() for s in deck.shapes] == ["A", "Blob"]
        # The moved node is the leftmost and topmost point now, so the shape's
        # box starts there, and that position was written after the paste.
        assert (deck.shapes[1].left_position(), deck.shapes[1].top()) == (300.0, 70.0)

    def test_insert_node(self):
        from ppt_mac.freeform import _get_shape_nodes_impl, _insert_node_impl

        with _fake_deck(shapes=[("Blob", "freeform", _fixture_bytes("freeform"))]):
            payload = json.loads(_insert_node_impl(1, "Blob", None, 2, 0, 0, 10.0, 20.0, None, None, None, None))
            nodes = json.loads(_get_shape_nodes_impl(1, "Blob", None))

        assert {k: v for k, v in payload.items() if k != "warnings"} == {
            "success": True, "shape_name": "Blob", "new_node_count": 10,
        }
        assert nodes["node_count"] == 10
        assert (nodes["nodes"][2]["x"], nodes["nodes"][2]["y"]) == (10.0, 20.0)

    def test_delete_node(self):
        """Node 3 is a curve's first control point, so the whole curve goes,
        three nodes of the nine, and the pasted path says so."""
        from ppt_mac.freeform import _delete_node_impl, _get_shape_nodes_impl

        with _fake_deck(shapes=[("Blob", "freeform", _fixture_bytes("freeform"))]):
            payload = json.loads(_delete_node_impl(1, "Blob", None, 3))
            nodes = json.loads(_get_shape_nodes_impl(1, "Blob", None))

        assert {k: v for k, v in payload.items() if k != "warnings"} == {
            "success": True, "shape_name": "Blob", "remaining_node_count": 6,
        }
        assert nodes["node_count"] == 6
        assert [n["segment_type"] for n in nodes["nodes"]] == [
            "line", "curve", "inaccessible", "inaccessible", "line", "inaccessible",
        ]

    def test_set_segment_type_carries_the_windows_note_when_the_count_changes(self):
        from ppt_mac.freeform import _set_segment_type_impl

        with _fake_deck(shapes=[("Blob", "freeform", _fixture_bytes("freeform"))]):
            to_curve = json.loads(_set_segment_type_impl(1, "Blob", None, 1, 1))
            same = json.loads(_set_segment_type_impl(1, "Blob", None, 1, 1))

        assert {k: v for k, v in to_curve.items() if k != "warnings"} == {
            "success": True, "shape_name": "Blob", "node_index": 1, "segment_type": "curve",
            "old_node_count": 9, "new_node_count": 11,
            "note": "Node count changed — switching line↔curve adds or removes control-point nodes. "
                    "Re-call ppt_get_shape_nodes to see updated indices.",
        }
        assert same["old_node_count"] == same["new_node_count"] == 11
        assert "note" not in same

    def test_an_index_out_of_range_raises_before_anything_is_pasted(self):
        from ppt_mac.freeform import _delete_node_impl

        with _fake_deck(shapes=[("Blob", "freeform", _fixture_bytes("freeform"))]) as deck:
            with pytest.raises(ValueError, match="node_index 99 out of range \\(shape has 9 nodes\\)"):
                _delete_node_impl(1, "Blob", None, 99)

        assert "paste" not in deck.order and deck.order[-1] == "restore"
        assert [s.name() for s in deck.shapes] == ["Blob"]

    def test_a_dropped_paste_leaves_the_path_untouched(self):
        from ppt_mac.freeform import _delete_node_impl

        with _fake_deck(shapes=[("Blob", "freeform", _fixture_bytes("freeform"))], deaf=True) as deck:
            payload = json.loads(_delete_node_impl(1, "Blob", None, 3))

        assert "success" not in payload
        assert deck.shapes[0].package == _fixture_bytes("freeform")

    def test_a_path_that_reads_back_differently_is_removed_and_refused(self):
        from ppt_mac.freeform import _set_node_position_impl

        with _fake_deck(shapes=[("Blob", "freeform", _fixture_bytes("freeform"))],
                        stale_package=_fixture_bytes("freeform")) as deck:
            payload = json.loads(_set_node_position_impl(1, "Blob", None, 2, 1.0, 1.0))

        assert payload["error"] == "ppt_set_node_position pasted a freeform that did not carry the edit"
        assert "9 node(s) read back where 9 were written, or at other positions" in payload["reason"]
        assert [s.name() for s in deck.shapes] == ["Blob"]

    def test_lost_animations_are_counted_here_too(self):
        from ppt_mac.freeform import _delete_node_impl

        with _fake_deck(shapes=[("Blob", "freeform", _fixture_bytes("freeform"))], effects=["Blob"]):
            payload = json.loads(_delete_node_impl(1, "Blob", None, 2))

        assert "1 animation effect(s) on 'Blob' were lost" in payload["warnings"][1]

    def test_the_editing_type_tool_still_says_the_xml_has_no_word_for_it(self):
        from ppt_mac.freeform import _set_node_editing_type_impl

        with _fake_deck(shapes=[("Blob", "freeform", None)]) as deck:
            payload = json.loads(_set_node_editing_type_impl(1, "Blob", None, 2, 2))

        assert "no such attribute either" in payload["reason"]
        assert "ppt_set_node_position" in payload["reason"]
        assert "success" not in payload
        assert deck.order == []


# ---------------------------------------------------------------------------
# The fakes
# ---------------------------------------------------------------------------
_LINE_NODES = [
    {"seg_int": 0, "et_int": 0, "x1": 110.0, "y1": 20.0, "x2": None, "y2": None, "x3": None, "y3": None},
    {"seg_int": 0, "et_int": 0, "x1": 110.0, "y1": 120.0, "x2": None, "y2": None, "x3": None, "y3": None},
]
_CURVE_NODES = [
    {"seg_int": 0, "et_int": 0, "x1": 200, "y1": 100, "x2": None, "y2": None, "x3": None, "y3": None},
    {"seg_int": 1, "et_int": 0, "x1": 250, "y1": 200, "x2": None, "y2": None, "x3": None, "y3": None},
    {"seg_int": 1, "et_int": 1, "x1": 220, "y1": 260, "x2": 160, "y2": 260, "x3": 120, "y3": 200},
]


def _command_error(number):
    """An appscript CommandError carrying an Apple Event error number.

    ``errornumber`` is read only on the real class, so a stub subclasses it.
    """
    from appscript.reference import CommandError

    class _Stub(CommandError):
        def __init__(self, number):
            Exception.__init__(self, number)
            self._number = number

        @property
        def errornumber(self):
            return self._number

        def __str__(self):
            return f"Command failed: OSERROR: {self._number}"

    return _Stub(number)


class _FakeBoard:
    """A pasteboard: one dict of contents and a change count that moves."""

    def __init__(self, deck):
        self.deck = deck
        self.contents = {"public.utf8-plain-text": b"kept"}
        self.count = 100

    # The surface of backend.pasteboard, patched in for the length of a test.
    def change_count(self):
        return self.count

    def read(self, uti):
        return self.contents.get(uti)

    def write(self, uti, data):
        self.deck.order.append("write")
        self.contents = {uti: data}
        self.count += 1
        if self.deck.rewrite_after_write:
            self.contents = {"public.utf8-plain-text": b"someone else"}
            self.count += 1
        return self.count if not self.deck.rewrite_after_write else self.count - 1

    def snapshot(self):
        from backend.pasteboard import Snapshot

        self.deck.order.append("snapshot")
        return Snapshot(types=dict(self.contents), change_count=self.count,
                        skipped=[], kept=not self.deck.too_large)

    def restore(self, snap):
        self.deck.order.append("restore")
        self.contents = dict(snap.types)
        self.count += 1
        return True

    SNAPSHOT_LIMIT = 20 * 1024 * 1024


class _Property:
    def __init__(self, value, on_set=None):
        self.value = value
        self.on_set = on_set

    def __call__(self):
        return self.value

    def set(self, value):
        if self.on_set:
            self.on_set()
        self.value = value


class _FakeShape:
    def __init__(self, deck, name, kind, package=None, left=10.0, top=20.0, width=100.0, height=60.0):
        from appscript import k

        self.deck = deck
        self._name = name
        self._type = {
            "auto": k.shape_type_auto, "chart": k.shape_type_chart,
            "freeform": k.shape_type_free_form, "group": k.shape_type_group,
        }[kind]
        self.package = package
        # Writing the position is one of the steps the procedure is made of,
        # so the left write is what marks it in the order.
        self.left_position = _Property(left, on_set=lambda: deck.order.append("position"))
        self.top = _Property(top)
        self.width = _Property(width)
        self.height = _Property(height)

    def name(self):
        return self._name

    def shape_type(self):
        return self._type

    def z_order_position(self):
        return self.deck.shapes.index(self) + 1

    def copy_shape(self):
        self.deck.order.append("copy")
        if self.package is None:
            self.deck.board.contents = {"public.utf8-plain-text": b"a copy with no package"}
        else:
            self.deck.board.contents = {"com.microsoft.Art--GVML-ClipFormat": self._named(self.package)}
        self.deck.board.count += 1

    def _named(self, raw):
        """The package with this shape's name in it, as PowerPoint writes it.

        A fixture recorded from a shape called HandChart stands in for a
        shape called Sales; PowerPoint stamps the current name on every copy.
        """
        import xml.etree.ElementTree as ET

        from gvml import Package, canvas, shapes

        package = Package.from_bytes(raw)
        root = ET.fromstring(package.drawing())
        for element in canvas.children(canvas.canvas_of(root)):
            shapes.rename(element, self._name)
        package.parts[package.drawing_part()] = canvas.serialize_drawing(root).encode()
        return package.to_bytes()

    def delete(self):
        self.deck.order.append("delete")
        if self._name in self.deck.undeletable:
            return
        self.deck.shapes.remove(self)

    def z_order(self, z_order_position=None):
        """Only `send shape backward` is sent here: one step down."""
        from appscript import k

        assert z_order_position == k.send_shape_backward
        self.deck.order.append("backward")
        if self.deck.stuck_z:
            return
        shapes = self.deck.shapes
        i = shapes.index(self)
        if i > 0:
            shapes[i - 1], shapes[i] = shapes[i], shapes[i - 1]

    def _own_package(self):
        """What `copy shape` would put on the pasteboard for this shape."""
        from gvml import build, canvas, shapes

        e = canvas.EMU_PER_PT
        x, y = int(self.left_position() * e), int(self.top() * e)
        cx, cy = int(self.width() * e), int(self.height() * e)
        return build(canvas.wrap(shapes.shape_xml(2, self._name, x, y, cx, cy), x, y, cx, cy))


class _FakeShapes:
    def __init__(self, deck):
        self.deck = deck

    def get(self):
        if not self.deck.shapes:
            raise _command_error(-1728)
        return list(self.deck.shapes)

    def __getitem__(self, index):
        if index < 1 or index > len(self.deck.shapes):
            raise _command_error(-1728)
        return self.deck.shapes[index - 1]

    @property
    def name(self):
        deck = self.deck

        class _Names:
            def get(self):
                if not deck.shapes:
                    raise _command_error(-1728)
                return [s.name() for s in deck.shapes]

        return _Names()


class _FakeView:
    def __init__(self, deck):
        self.deck = deck

    def go_to_slide(self, number=None):
        if self.deck.goto_error is not None:
            raise _command_error(self.deck.goto_error)
        self.deck.order.append("goto")

    def paste_object(self):
        from gvml import Package, canvas, shapes

        deck = self.deck
        deck.order.append("paste")
        if deck.deaf:
            return
        raw = deck.board.contents.get("com.microsoft.Art--GVML-ClipFormat")
        if raw is None:
            return
        element = canvas.children(canvas.parse(Package.from_bytes(raw).drawing()))[0]
        info = shapes.describe(element)
        kind = deck.lands_as or {
            shapes.MSO_AUTO_SHAPE: "auto", shapes.MSO_CHART: "chart",
            shapes.MSO_FREEFORM: "freeform", shapes.MSO_GROUP: "group",
        }[info.type]
        # Where PowerPoint drops a paste: the middle of the view, not the
        # package's own offset.
        landed = _FakeShape(deck, info.name, kind, left=380.0, top=210.0,
                            width=canvas.pt(info.cx), height=canvas.pt(info.cy))
        # What `copy shape` on the new shape answers: the package pasted,
        # unless the test wants PowerPoint to have kept something else.
        landed.package = deck.stale_package or raw
        if deck.stuck_position:
            landed.left_position = _Property(deck.stuck_position[0], on_set=lambda: deck.order.append("position"))
            landed.left_position.set = lambda v: deck.order.append("position")
            landed.top = _Property(deck.stuck_position[1])
            landed.top.set = lambda v: None
        deck.shapes.append(landed)
        if deck.lands_twice:
            deck.shapes.append(_FakeShape(deck, info.name + " again", kind))
        if deck.rewrite_after_paste:
            deck.board.contents = {"public.utf8-plain-text": b"newer"}
            deck.board.count += 1


class _FakeWindow:
    def __init__(self, deck):
        self.view = _FakeView(deck)
        deck_ = deck

        class _Selection:
            def unselect(self):
                deck_.order.append("unselect")

        self.selection = _Selection()


class _FakeDeck:
    def __init__(self, shapes, deaf=False, goto_error=None, lands_as=None, lands_twice=False,
                 stuck_position=None, rewrite_after_write=False, rewrite_after_paste=False,
                 too_large=False, undeletable=(), effects=(), stale_package=None,
                 stuck_z=False, no_timeline=False):
        self.order = []
        # Shape names, one per effect of the slide's main sequence.
        self.effects = list(effects)
        self.stale_package = stale_package
        self.stuck_z = stuck_z
        self.no_timeline = no_timeline
        self.deaf = deaf
        self.goto_error = goto_error
        self.lands_as = lands_as
        self.lands_twice = lands_twice
        self.stuck_position = stuck_position
        self.rewrite_after_write = rewrite_after_write
        self.rewrite_after_paste = rewrite_after_paste
        self.too_large = too_large
        self.undeletable = set(undeletable)
        self.board = _FakeBoard(self)
        self.shapes = []
        for spec in shapes:
            name, kind = spec[0], spec[1]
            shape = _FakeShape(self, name, kind)
            shape.package = spec[2] if len(spec) > 2 else shape._own_package()
            self.shapes.append(shape)
        self.window = _FakeWindow(self)

    @property
    def slide(self):
        deck = self

        class _Effect:
            def __init__(self, name):
                self.shape = mock.Mock()
                self.shape.name = mock.Mock(return_value=name)

        class _Sequence:
            def count(self, each=None):
                deck.order.append("effects")
                return len(deck.effects)

            effects = _FakeShapesLike([_Effect(n) for n in deck.effects])

        class _Timeline:
            main_sequence = _Sequence()

        class _Slide:
            shapes = _FakeShapes(deck)
            if not deck.no_timeline:
                timeline = _Timeline()

        return _Slide()

    @property
    def presentation(self):
        deck = self

        class _Pres:
            slides = _FakeShapesLike([deck.slide])
            document_windows = _FakeShapesLike([deck.window])

            def count(self, each=None):
                return 1

            def name(self):
                return "Fake"

        return _Pres()


class _FakeShapesLike:
    def __init__(self, items):
        self._items = items

    def get(self):
        return list(self._items)

    def __getitem__(self, index):
        return self._items[index - 1]


@contextmanager
def _fake_deck(**kwargs):
    from backend.mac_ae import ppt
    from ppt_mac import gvml_paste

    deck = _FakeDeck(**kwargs)
    board = deck.board

    class _Pasteboard:
        SNAPSHOT_LIMIT = board.SNAPSHOT_LIMIT
        change_count = staticmethod(board.change_count)
        read = staticmethod(board.read)
        write = staticmethod(board.write)
        snapshot = staticmethod(board.snapshot)
        restore = staticmethod(board.restore)

    original_claim = gvml_paste.Clipboard.claim

    def claim(self):
        deck.order.append("claim")
        self.mine = board.count

    with mock.patch.object(ppt, "_get_pres_impl", return_value=deck.presentation), \
            mock.patch.object(ppt, "_get_app_impl", return_value=mock.Mock()), \
            mock.patch.object(gvml_paste, "pasteboard", _Pasteboard), \
            mock.patch.object(gvml_paste, "_recreation_explained", False), \
            mock.patch.object(gvml_paste.Clipboard, "claim", claim):
        yield deck
    gvml_paste.Clipboard.claim = original_claim


@contextmanager
def _no_powerpoint():
    from backend.mac_ae import ppt

    def explode(*args, **kwargs):
        raise AssertionError("PowerPoint must not be reached")

    with mock.patch.object(ppt, "_get_pres_impl", explode), \
            mock.patch.object(ppt, "_get_app_impl", explode):
        yield
