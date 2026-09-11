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
class TestGvmlPasteIsNotATool:
    def test_gvml_paste_defines_no_impl_of_its_own(self):
        """MACOS_PORT's tool count walks every `_impl` in ppt_mac."""
        from ppt_mac import gvml_paste

        assert not [n for n in dir(gvml_paste) if n.endswith("_impl")]


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

    def test_the_editing_type_tool_says_the_xml_has_no_word_for_it_either(self):
        from ppt_mac.freeform import _set_node_editing_type_impl

        with _fake_deck(shapes=[("Blob", "freeform", None)]):
            payload = json.loads(_set_node_editing_type_impl(1, "Blob", None, 2, 2))

        assert "no such attribute either" in payload["reason"]
        assert "success" not in payload


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
            self.deck.board.contents = {"com.microsoft.Art--GVML-ClipFormat": self.package}
        self.deck.board.count += 1

    def delete(self):
        self.deck.order.append("delete")
        if self._name in self.deck.undeletable:
            return
        self.deck.shapes.remove(self)

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
                 too_large=False, undeletable=()):
        self.order = []
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

        class _Slide:
            shapes = _FakeShapes(deck)

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
