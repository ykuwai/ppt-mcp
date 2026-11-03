"""SmartArt nodes on macOS, through the graphic's clipboard package (#189).

``tests/fixtures/gvml/smartart.gvml.zip`` is the package PowerPoint for Mac
16.114 wrote for a "Basic Block List" graphic with five nodes, the first
three reading Plan, Build and Ship. Its connections are not in node order
(``srcOrd`` 2 comes before 1), which is what the ordering test leans on.
"""

import os
import sys
import xml.etree.ElementTree as ET

import pytest

sys.path.insert(0, "src")

from gvml import Package  # noqa: E402
from gvml import smartart as sa  # noqa: E402
from gvml.text import body_text  # noqa: E402

macos_only = pytest.mark.skipif(
    sys.platform != "darwin", reason="the Apple Event backend needs appscript"
)

FIXTURE = os.path.join(os.path.dirname(__file__), "fixtures", "gvml", "smartart.gvml.zip")


def _package():
    with open(FIXTURE, "rb") as handle:
        return Package.from_bytes(handle.read())


def _parts(package):
    data = next(p for p in package.parts if "/diagrams/data" in p)
    drawing = next(p for p in package.parts if "/diagrams/drawing" in p)
    return ET.fromstring(package.parts[data]), ET.fromstring(package.parts[drawing])


def _drawn_texts(drawing):
    return [body_text(sp.find(f"{sa._S}txBody")) for sp in drawing.iter(f"{sa._S}sp")
            if sp.find(f"{sa._S}txBody") is not None]


class TestTheDataModel:
    def test_nodes_in_all_nodes_order(self):
        data, _ = _parts(_package())
        assert sa.node_texts(data) == ["Plan", "Build", "Ship", "", ""]

    def test_the_order_follows_src_ord_not_the_file(self):
        data, _ = _parts(_package())
        cxns = [c for c in data.iter(f"{sa._D}cxn") if c.get("type", "parOf") == "parOf"]
        assert [c.get("srcOrd") for c in cxns] != sorted(c.get("srcOrd") for c in cxns)
        assert len(sa.node_ids(data)) == 5

    def test_a_node_s_text_is_written_in_the_data_and_the_drawing(self):
        data, drawing = _parts(_package())
        sa.set_node_text(data, drawing, 2, "Make")
        assert sa.node_texts(data)[1] == "Make"
        assert _drawn_texts(drawing)[1] == "Make"

    def test_paragraphs_and_breaks(self):
        data, drawing = _parts(_package())
        sa.set_node_text(data, drawing, 1, "Design\nphase\vone")
        assert sa.node_texts(data)[0] == "Design\nphase\vone"
        assert _drawn_texts(drawing)[0] == "Design\nphase\vone"

    def test_an_empty_node_stops_being_a_placeholder(self):
        data, drawing = _parts(_package())
        node = data.find(f"{sa._D}ptLst/{sa._D}pt[@modelId='{sa.node_ids(data)[3]}']")
        assert node.find(f"{sa._D}prSet").get("phldr") == "1"
        sa.set_node_text(data, drawing, 4, "Learn")
        assert "phldr" not in node.find(f"{sa._D}prSet").attrib
        assert sa.node_texts(data)[3] == "Learn" and _drawn_texts(drawing)[3] == "Learn"

    def test_the_drawn_text_keeps_its_size(self):
        data, drawing = _parts(_package())
        sa.set_node_text(data, drawing, 4, "Learn")
        sp = [s for s in drawing.iter(f"{sa._S}sp") if s.find(f"{sa._S}txBody") is not None][3]
        rpr = sp.find(f"{sa._S}txBody").find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}rPr")
        assert rpr.get("sz") == "6500"

    @pytest.mark.parametrize("index", [0, 6])
    def test_an_index_out_of_range(self, index):
        data, drawing = _parts(_package())
        with pytest.raises(ValueError, match=r"out of range \(1-5\)"):
            sa.set_node_text(data, drawing, index, "x")

    def test_without_a_drawing_the_data_is_still_written(self):
        data, _ = _parts(_package())
        sa.set_node_text(data, None, 3, "Launch")
        assert sa.node_texts(data)[2] == "Launch"


@pytest.fixture
def wired(monkeypatch):
    from ppt_mac import gvml_paste
    from ppt_mac import smartart as mac_smartart

    state = {"raw": None}

    class FakeClipboard:
        @classmethod
        def take(cls):
            clip = cls()
            clip.warnings = []
            return clip

        def restore(self):
            state["restored"] = True

    def replace(pres, slide, slide_index, index, raw, clip, tool, expected, left, top,
                what, verify, alternatives):
        state["raw"], state["index"] = raw, index
        state["problem"] = verify(Package.from_bytes(raw))
        return gvml_paste.Replaced(None, "Diagram 1", Package.from_bytes(raw), ["recreated"])

    class Shape:
        def name(self):
            return "Diagram 1"

        def left_position(self):
            return 60.0

        def top(self):
            return 100.0

    monkeypatch.setattr(gvml_paste, "Clipboard", FakeClipboard)
    monkeypatch.setattr(gvml_paste, "copy_shape_package", lambda shape, clip, tool: _package())
    monkeypatch.setattr(gvml_paste, "replace_shape", replace)
    monkeypatch.setattr(mac_smartart, "_shape_index", lambda slide, name: 1)
    state["call"] = lambda index, text: mac_smartart._edit_nodes(
        object(), object(), 1, Shape(), "set_text",
        lambda data, drawing: sa.set_node_text(data, drawing, index, text))
    state["edit"] = lambda action, edit: mac_smartart._edit_nodes(
        object(), object(), 1, Shape(), action, edit)
    return state


@macos_only
class TestSmartArtNodeText:
    def test_the_text_is_written_pasted_and_checked(self, wired):
        result = wired["call"](1, "Design")
        assert result == {"success": True, "action": "set_text", "total_nodes": 5,
                          "shape_name": "Diagram 1", "warnings": ["recreated"]}
        assert wired["problem"] is None and wired["restored"]
        data, drawing = _parts(Package.from_bytes(wired["raw"]))
        assert sa.node_texts(data)[0] == "Design" and _drawn_texts(drawing)[0] == "Design"

    def test_the_prefixes_powerpoint_wrote_are_kept(self, wired):
        wired["call"](2, "Make")
        package = Package.from_bytes(wired["raw"])
        raw = next(package.parts[p] for p in package.parts if "/diagrams/data" in p)
        assert b"<dgm:dataModel" in raw and b"ns0:" not in raw

    def test_an_index_out_of_range_changes_nothing(self, wired):
        with pytest.raises(ValueError, match=r"out of range \(1-5\)"):
            wired["call"](8, "x")
        assert wired["raw"] is None

    def test_the_tool_needs_a_node_and_text(self, monkeypatch):
        from ppt_mac import smartart as mac_smartart

        monkeypatch.setattr(mac_smartart.ppt, "_get_pres_impl", lambda *a, **kw: object())
        monkeypatch.setattr(mac_smartart, "_slide", lambda pres, index: object())
        monkeypatch.setattr(mac_smartart, "_get_shape", lambda slide, name: object())
        monkeypatch.setattr(mac_smartart, "_require_smartart", lambda shape: None)
        args = (None, None, None, None, None, None, None, None, None, None, None)
        with pytest.raises(ValueError, match="node_index is required"):
            mac_smartart._modify_smartart_impl(1, "Diagram 1", "set_text", None, "x", *args)
        with pytest.raises(ValueError, match="text is required"):
            mac_smartart._modify_smartart_impl(1, "Diagram 1", "set_text", 1, None, *args)


class TestAddingAndDeletingNodes:
    def test_a_node_after_another_takes_its_place_and_moves_the_rest_on(self):
        data, _ = _parts(_package())
        assert sa.add_node(data, 1, "After plan") == 2
        assert sa.node_texts(data) == ["Plan", "After plan", "Build", "Ship", "", ""]

    def test_without_an_index_a_node_is_added_last(self):
        data, _ = _parts(_package())
        assert sa.add_node(data, None, "Ops") == 6
        assert sa.node_texts(data)[-1] == "Ops"

    def test_a_new_node_has_its_own_ids_and_one_link(self):
        data, _ = _parts(_package())
        sa.add_node(data, 2, None)
        ids = [p.get("modelId") for p in data.iter(f"{sa._D}pt")]
        assert len(ids) == len(set(ids))
        new = sa.node_ids(data)[2]
        links = [c for c in data.iter(f"{sa._D}cxn")
                 if c.get("type", "parOf") == "parOf" and c.get("destId") == new]
        assert len(links) == 1
        for kind in ("parTransId", "sibTransId"):
            assert any(p.get("modelId") == links[0].get(kind) for p in data.iter(f"{sa._D}pt"))

    def test_an_empty_new_node_is_a_placeholder(self):
        data, _ = _parts(_package())
        sa.add_node(data, 1, None)
        node = data.find(f"{sa._D}ptLst/{sa._D}pt[@modelId='{sa.node_ids(data)[1]}']")
        assert node.find(f"{sa._D}prSet").get("phldr") == "1"
        assert sa.node_texts(data)[1] == ""

    def test_deleting_takes_the_node_its_points_and_its_links(self):
        data, _ = _parts(_package())
        gone = sa.node_ids(data)[1]
        pres = [c.get("destId") for c in data.iter(f"{sa._D}cxn")
                if c.get("type") == "presOf" and c.get("srcId") == gone]
        assert sa.delete_node(data, 2) == 1
        assert sa.node_texts(data) == ["Plan", "Ship", "", ""]
        left = {p.get("modelId") for p in data.iter(f"{sa._D}pt")}
        assert gone not in left and not set(pres) & left
        assert not any(gone in (c.get("srcId"), c.get("destId")) for c in data.iter(f"{sa._D}cxn"))

    def test_the_siblings_after_a_deleted_node_move_back(self):
        data, _ = _parts(_package())
        sa.delete_node(data, 1)
        orders = sorted(int(c.get("srcOrd")) for c in data.iter(f"{sa._D}cxn")
                        if c.get("type", "parOf") == "parOf")
        assert orders == [0, 1, 2, 3]

    def test_deleting_a_node_takes_the_nodes_under_it(self):
        data, _ = _parts(_package())
        # Put a node under Plan, the way a hierarchy has one.
        sa.add_node(data, None, "Child")
        child = sa.node_ids(data)[-1]
        link = next(c for c in data.iter(f"{sa._D}cxn") if c.get("destId") == child)
        link.set("srcId", sa.node_ids(data)[0])
        link.set("srcOrd", "0")
        assert sa.node_texts(data)[:2] == ["Plan", "Child"]
        assert sa.delete_node(data, 1) == 2
        assert "Child" not in sa.node_texts(data) and "Plan" not in sa.node_texts(data)

    def test_deleting_every_node_is_refused(self):
        data, _ = _parts(_package())
        for _ in range(4):
            sa.delete_node(data, 1)
        with pytest.raises(ValueError, match="needs one"):
            sa.delete_node(data, 1)
        assert len(sa.node_ids(data)) == 1

    @pytest.mark.parametrize("call", [lambda d: sa.add_node(d, 9, "x"),
                                      lambda d: sa.delete_node(d, 0)])
    def test_an_index_out_of_range(self, call):
        data, _ = _parts(_package())
        with pytest.raises(ValueError, match=r"out of range \(1-5\)"):
            call(data)


@macos_only
class TestTheNodeActionsOnTheMac:
    def test_add_node_is_pasted_and_checked(self, wired):
        result = wired["edit"]("add_node", lambda d, w: sa.add_node(d, None, "Ops"))
        assert (result["action"], result["total_nodes"]) == ("add_node", 6)
        assert wired["problem"] is None
        data, _ = _parts(Package.from_bytes(wired["raw"]))
        assert sa.node_texts(data)[-1] == "Ops"

    def test_delete_node_is_pasted_and_checked(self, wired):
        result = wired["edit"]("delete_node", lambda d, w: sa.delete_node(d, 2))
        assert (result["action"], result["total_nodes"]) == ("delete_node", 4)
        assert wired["problem"] is None
