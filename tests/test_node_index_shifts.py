"""Node indexes move when control points come and go; the node tools say so.

A curve's two control points are nodes of their own, so making a segment a
curve, or smoothing a node, renumbers every node after it. Measured the same
through COM on Windows and through VBA on the Mac: after
ppt_set_segment_type(2, curve), node 2 is a control point, and smoothing it
straightens the curve again. A caller acting on an index it read before would
edit the wrong point, so the index arguments explain it and the two tools that
change the count as a side effect answer the nodes as they are afterwards.

The macOS side is covered in tests/test_mac_gvml.py. Nothing here reaches
PowerPoint.
"""

import json
import sys

import pytest

sys.path.insert(0, "src")

from ppt_com import freeform  # noqa: E402

windows_only = pytest.mark.skipif(
    sys.platform != "win32", reason="drives the COM implementation"
)


class TestTheArgumentsSaySo:
    @pytest.mark.parametrize("model, field", [
        (freeform.SetNodePositionInput, "node_index"),
        (freeform.DeleteNodeInput, "node_index"),
        (freeform.SetNodeEditingTypeInput, "node_index"),
        (freeform.SetSegmentTypeInput, "node_index"),
        (freeform.InsertNodeInput, "after_index"),
    ])
    def test_every_node_index_explains_the_shift(self, model, field):
        description = model.model_fields[field].description
        assert "control points" in description
        assert "ppt_get_shape_nodes" in description

    def test_the_segment_type_argument_still_says_which_segment(self):
        assert "ends at it" in freeform.SetSegmentTypeInput.model_fields["node_index"].description


# ---------------------------------------------------------------------------
# Windows
# ---------------------------------------------------------------------------
class _Node:
    def __init__(self, x, y, editing=1, segment=0, readable=True):
        self.x, self.y = x, y
        self._editing, self._segment, self.readable = editing, segment, readable

    @property
    def EditingType(self):  # noqa: N802 - the COM name
        if not self.readable:
            raise RuntimeError("Method 'EditingType' of object 'ShapeNode' failed")
        return self._editing

    @property
    def SegmentType(self):  # noqa: N802
        if not self.readable:
            raise RuntimeError("Method 'SegmentType' of object 'ShapeNode' failed")
        return self._segment


class _Nodes:
    """Four anchors in a row, behaving as PowerPoint did when measured."""

    def __init__(self):
        self.items = [_Node(240, 300, 0), _Node(320, 320), _Node(280, 380), _Node(360, 400, 0)]

    @property
    def Count(self):  # noqa: N802
        return len(self.items)

    def Item(self, i):  # noqa: N802
        return self.items[i - 1]

    def SetSegmentType(self, i, segment):  # noqa: N802
        if segment == 1 and self.items[i - 1]._segment == 0:
            # The segment after node i becomes a curve: its two control
            # points become nodes i+1 and i+2, and the rest move up by two.
            a, b = self.items[i - 1], self.items[i]
            self.items[i:i] = [
                _Node(a.x + (b.x - a.x) / 3, a.y + (b.y - a.y) / 3, readable=False),
                _Node(a.x + 2 * (b.x - a.x) / 3, a.y + 2 * (b.y - a.y) / 3, readable=False),
            ]
            a._segment = 1

    def SetEditingType(self, i, editing):  # noqa: N802
        if not self.items[i - 1].readable and editing == 2:
            # Smoothing a control point of a fresh curve straightened it.
            first = i - 1 if self.items[i - 2].readable else i - 2
            del self.items[first:first + 2]
            self.items[first - 1]._segment = 0
        else:
            self.items[i - 1]._editing = editing


class _Shape:
    Name = "Free"
    Type = 5  # msoFreeform

    def __init__(self):
        self.Nodes = _Nodes()

    @property
    def Vertices(self):  # noqa: N802
        return tuple((n.x, n.y) for n in self.Nodes.items)


@pytest.fixture
def shape(monkeypatch):
    s = _Shape()

    class _Deck:
        def Slides(self, index):  # noqa: N802
            return object()

    monkeypatch.setattr(freeform.ppt, "_get_app_impl", lambda *a, **k: object())
    monkeypatch.setattr(freeform.ppt, "_get_pres_impl", lambda *a, **k: _Deck())
    monkeypatch.setattr(freeform, "goto_slide", lambda app, index: None)
    monkeypatch.setattr(freeform, "_get_shape", lambda slide, _, shape_name=None, shape_index=None: s)
    return s


@windows_only
class TestOnWindows:
    def test_a_new_curve_answers_the_renumbered_nodes(self, shape):
        result = json.loads(freeform._set_segment_type_impl(1, "Free", None, 2, 1))

        assert (result["old_node_count"], result["new_node_count"]) == (4, 6)
        assert result["note"] == freeform.NODES_MOVED_NOTE
        assert [n["index"] for n in result["nodes"]] == [1, 2, 3, 4, 5, 6]
        # The anchor that was node 3 is node 5 now.
        assert (result["nodes"][4]["x"], result["nodes"][4]["y"]) == (280, 380)
        assert result["nodes"][2]["editing_type"] == "inaccessible"

    def test_smoothing_a_control_point_answers_the_straightened_path(self, shape):
        freeform._set_segment_type_impl(1, "Free", None, 2, 1)
        result = json.loads(freeform._set_node_editing_type_impl(1, "Free", None, 3, 2))

        assert (result["old_node_count"], result["new_node_count"]) == (6, 4)
        assert result["note"] == freeform.NODES_MOVED_NOTE
        assert [(n["x"], n["y"]) for n in result["nodes"]] == [(240, 300), (320, 320), (280, 380), (360, 400)]

    def test_an_edit_that_moves_nothing_says_nothing_about_it(self, shape):
        result = json.loads(freeform._set_node_editing_type_impl(1, "Free", None, 2, 1))

        assert result["old_node_count"] == result["new_node_count"] == 4
        assert "note" not in result
        assert len(result["nodes"]) == 4
