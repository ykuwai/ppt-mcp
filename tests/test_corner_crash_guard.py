"""ppt_set_node_editing_type refuses the one call that makes PowerPoint quit.

SetEditingType(2, corner) while node 2 is a control point, that is while the
first segment is a curve, ended PowerPoint every time it was tried: on Windows
(16.0.20326, "The remote procedure call failed", a fault in
mso20win32client.dll) and on the Mac through VBA (16.114). Corner on the other
control points, and the other editing types on node 2, work. Nothing here
reaches PowerPoint.
"""

import sys

import pytest

sys.path.insert(0, "src")

from ppt_com import freeform  # noqa: E402

LINE, CURVE = 0, 1
AUTO, CORNER, SMOOTH, SYMMETRIC = 0, 1, 2, 3


class _Node:
    def __init__(self, segment):
        self.SegmentType = segment
        self.EditingType = CORNER


class _Nodes:
    """A freeform whose first segment is ``first``; records what was set."""

    def __init__(self, first, count=6):
        self.items = [_Node(first)] + [_Node(LINE) for _ in range(count - 1)]
        self.calls = []

    @property
    def Count(self):  # noqa: N802 - the COM name
        return len(self.items)

    def Item(self, i):  # noqa: N802
        return self.items[i - 1]

    def SetEditingType(self, i, et):  # noqa: N802
        self.calls.append((i, et))


class _Shape:
    Name = "Free"
    Type = 5

    def __init__(self, nodes):
        self.Nodes = nodes


windows_only = pytest.mark.skipif(
    sys.platform != "win32", reason="on macOS the module's impl is the Mac one"
)


def test_corner_on_the_first_curves_control_point_is_refused():
    with pytest.raises(ValueError, match="makes PowerPoint quit") as raised:
        freeform._refuse_known_crash(_Nodes(CURVE), 2, CORNER)
    assert "ppt_set_segment_type" in str(raised.value)


@pytest.mark.parametrize("first, node_index, et, count", [
    (LINE, 2, CORNER, 6),        # node 2 is an anchor
    (CURVE, 3, CORNER, 6),       # the curve's other control point
    (CURVE, 1, CORNER, 6),       # the start
    (CURVE, 2, AUTO, 6),
    (CURVE, 2, SMOOTH, 6),
    (CURVE, 2, SYMMETRIC, 6),
    (CURVE, 2, CORNER, 2),       # two nodes: no control point
])
def test_every_other_call_is_let_through(first, node_index, et, count):
    freeform._refuse_known_crash(_Nodes(first, count), node_index, et)


@windows_only
def test_the_tool_never_sends_the_call(monkeypatch):
    nodes = _Nodes(CURVE)
    shape = _Shape(nodes)
    monkeypatch.setattr(freeform.ppt, "_get_app_impl", lambda *a, **k: object())
    monkeypatch.setattr(freeform.ppt, "_get_pres_impl", lambda *a, **k: type("P", (), {"Slides": lambda self, i: object()})())
    monkeypatch.setattr(freeform, "goto_slide", lambda *a: None)
    monkeypatch.setattr(freeform, "_get_shape", lambda *a, **k: shape)
    monkeypatch.setattr(freeform, "_check_freeform", lambda s: None)
    with pytest.raises(ValueError, match="makes PowerPoint quit"):
        freeform._set_node_editing_type_impl(1, "Free", None, 2, CORNER)
    assert nodes.calls == []
    freeform._set_node_editing_type_impl(1, "Free", None, 3, CORNER)
    assert nodes.calls == [(3, CORNER)]


def test_the_tool_description_warns():
    import inspect
    assert "'corner' on node 2 is refused while the first segment is a curve" in inspect.getsource(freeform)
