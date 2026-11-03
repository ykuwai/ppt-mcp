"""SmartArt layouts, colour schemes and styles on macOS.

PowerPoint for Mac installs every definition it offers as an XML file
(``ppt_mac/smartart_defs.py``). These tests build a small install of the same
shape under ``tmp_path``, so they run where PowerPoint is not installed, and
check the listing, the order, the name lookup, the definition swap in a
graphic's package and the package a new graphic is built as.
"""

import os
import sys
import xml.etree.ElementTree as ET

import pytest

sys.path.insert(0, "src")

from gvml import Package  # noqa: E402
from gvml import smartart as sa  # noqa: E402

macos_only = pytest.mark.skipif(
    sys.platform != "darwin", reason="the Apple Event backend needs appscript"
)

FIXTURE = os.path.join(os.path.dirname(__file__), "fixtures", "gvml", "smartart.gvml.zip")
URN = "urn:microsoft.com/office/officeart"
HEAD = '<?xml version="1.0" encoding="utf-8"?>\n<!-- Copyright (c) Microsoft Corporation. -->\n'
NS = 'xmlns="http://schemas.openxmlformats.org/drawingml/2006/diagram"'


def _layout(uid, name, cat, pri, samp=""):
    title = '<title val=""/>' + (f'<!-- "{name}" -->' if name else "")
    return (f'{HEAD}<layoutDef {NS} uniqueId="{uid}">\n\t{title}\n\t<desc val="" />'
            f'<catLst><cat type="{cat}" pri="{pri}" /></catLst>{samp}<layoutNode/></layoutDef>')


def _flat(n):
    pts = '<pt modelId="0" type="doc" />' + "".join(
        f'<pt modelId="{i}" type="node"><prSet phldr="1" /></pt>' for i in range(1, n + 1))
    cxns = "".join(f'<cxn modelId="{10 + i}" type="parOf" srcId="0" destId="{i}" '
                   f'srcOrd="{i - 1}" destOrd="0" />' for i in range(1, n + 1))
    return f"<sampData><dataModel><ptLst>{pts}</ptLst><cxnLst>{cxns}</cxnLst></dataModel></sampData>"


# A root with two children, the first with one child of its own, written in
# a file order that is not the tree's.
TREE = (
    '<sampData><dataModel><ptLst><pt modelId="0" type="doc" />'
    '<pt modelId="1"><prSet phldr="1" /></pt><pt modelId="2"><prSet phldr="1" /></pt>'
    '<pt modelId="3"><prSet phldr="1" /></pt><pt modelId="4"><prSet phldr="1" /></pt>'
    '<pt modelId="9" type="sibTrans"><prSet phldrT="1" /><t/></pt></ptLst><cxnLst>'
    '<cxn modelId="5" srcId="1" destId="3" srcOrd="1" destOrd="0" />'
    '<cxn modelId="6" srcId="0" destId="1" srcOrd="0" destOrd="0" />'
    '<cxn modelId="7" srcId="1" destId="2" srcOrd="0" destOrd="0" />'
    '<cxn modelId="8" srcId="2" destId="4" srcOrd="0" destOrd="0" />'
    '</cxnLst></dataModel></sampData>'
)


def _colors(uid, name, cat, pri):
    return (f'{HEAD}<colorsDef {NS} uniqueId="{uid}"><title val=""/><!-- "{name}" -->'
            f'<desc val=""/><catLst><cat type="{cat}" pri="{pri}"/></catLst></colorsDef>')


def _style(uid, name, cat, pri):
    return (f'{HEAD}<styleDef {NS} uniqueId="{uid}"><title val=""/><!-- "{name}" -->'
            f'<desc val=""/><!-- "{name}, described" --><catLst><cat type="{cat}" pri="{pri}"/>'
            '</catLst></styleDef>')


FILES = {
    "lo/default.glo": _layout(f"{URN}/2005/8/layout/default", "Basic Block List", "list", 400, _flat(5)),
    "lo/process1.glo": _layout(f"{URN}/2005/8/layout/process1", "Basic Process", "process", 100),
    "lo/hierarchy1.glo": _layout(f"{URN}/2005/8/layout/hierarchy1", "Hierarchy", "hierarchy", 100, TREE),
    "lo/BendingPictureCaption.glo": _layout(
        f"{URN}/2008/layout/BendingPictureCaption", "", "picture", 100, _flat(2)),
    "lo/vList2.glo": _layout(f"{URN}/2005/8/layout/vList2", "Vertical Bullet List", "list", 100, _flat(2)),
    "cs/accent1_2.gcs": _colors(f"{URN}/2005/8/colors/accent1_2", "Colored Fill - Accent 1", "accent1", 11200),
    "cs/colorful1.gcs": _colors(f"{URN}/2005/8/colors/colorful1", "Colorful - Accent Colors", "colorful", 10100),
    "cs/accent0_1.gcs": _colors(f"{URN}/2005/8/colors/accent0_1", "Dark 1 Outline", "mainScheme", 1000),
    "qs/3d1.gqs": _style(f"{URN}/2005/8/quickstyle/3d1", "Inset", "3D", 11100),
    "qs/simple1.gqs": _style(f"{URN}/2005/8/quickstyle/simple1", "Simple Fill", "simple", 10100),
    "qs/simple2.gqs": _style(f"{URN}/2005/8/quickstyle/simple2", "White Outline", "simple", 10200),
}


@pytest.fixture
def install(monkeypatch, tmp_path):
    app = tmp_path / "PowerPoint.app"
    root = app / "Contents/Frameworks/SmartArt.framework/Versions/A/Resources"
    for rel, text in FILES.items():
        (root / rel).parent.mkdir(parents=True, exist_ok=True)
        (root / rel).write_text(text, encoding="utf-8")
    monkeypatch.setenv("PPT_MCP_POWERPOINT_APP", str(app))
    return root


def _fixture_package():
    with open(FIXTURE, "rb") as handle:
        return Package.from_bytes(handle.read())


def _data(package):
    part = next(p for p in package.parts if "/diagrams/data" in p)
    return ET.fromstring(package.parts[part])


class TestTheInstalledDefinitions:
    def test_each_kind_in_windows_order(self, install):
        from ppt_mac import smartart_defs as defs

        assert [d.name for d in defs.definitions("layouts")] == [
            "Vertical Bullet List", "Basic Block List", "Basic Process", "Hierarchy",
            "Bending Picture Caption"]
        assert [d.name for d in defs.definitions("colors")] == [
            "Dark 1 Outline", "Colorful - Accent Colors", "Colored Fill - Accent 1"]
        assert [d.name for d in defs.definitions("styles")] == [
            "Simple Fill", "White Outline", "Inset"]
        assert [d.index for d in defs.definitions("styles")] == [1, 2, 3]

    def test_a_layout_without_a_name_is_named_from_its_id(self, install):
        from ppt_mac import smartart_defs as defs

        assert defs.by_id("layouts", f"{URN}/2008/layout/BendingPictureCaption").name == \
            "Bending Picture Caption"

    def test_the_description_comes_from_its_comment(self, install):
        from ppt_mac import smartart_defs as defs

        assert defs.by_index("styles", 1).description == "Simple Fill, described"
        assert defs.by_index("colors", 1).description == ""

    def test_a_layout_is_found_by_part_of_its_name(self, install):
        from ppt_mac import smartart_defs as defs

        assert defs.resolve_layout("basic proc").id.endswith("/process1")
        with pytest.raises(ValueError, match="SmartArt layout 'Nope' not found"):
            defs.resolve_layout("Nope")

    def test_an_index_out_of_range_says_how_many_there_are(self, install):
        from ppt_mac import smartart_defs as defs

        with pytest.raises(ValueError, match=r"color index 4 is out of range.*has 3 \(1-3\)"):
            defs.by_index("colors", 4)

    def test_no_install_says_where_it_looked(self, monkeypatch, tmp_path):
        from ppt_mac import smartart_defs as defs

        monkeypatch.setenv("PPT_MCP_POWERPOINT_APP", str(tmp_path / "Gone.app"))
        with pytest.raises(defs.DefinitionsMissing, match="Gone.app"):
            defs.definitions("layouts")


class TestSwappingADefinition:
    def test_ids_as_the_document_point_names_them(self):
        assert sa.definition_ids(_data(_fixture_package())) == {
            "layouts": f"{URN}/2005/8/layout/default",
            "colors": f"{URN}/2005/8/colors/accent1_2",
            "styles": f"{URN}/2005/8/quickstyle/simple1",
        }

    def test_a_colour_scheme_and_its_suffix(self):
        data = _data(_fixture_package())
        sa.set_definition(data, "colors", f"{URN}/2005/8/colors/colorful1", "colorful")
        pr = next(p for p in data.iter(f"{sa._D}pt") if p.get("type") == "doc").find(f"{sa._D}prSet")
        assert pr.get("csTypeId") == f"{URN}/2005/8/colors/colorful1#1"
        assert pr.get("csCatId") == "colorful"
        part = sa.definition_part("colors", FILES["cs/colorful1.gcs"].encode(),
                                  f"{URN}/2005/8/colors/colorful1")
        assert f'uniqueId="{URN}/2005/8/colors/colorful1#1"'.encode() in part

    def test_a_layout_and_the_presentation_connections(self):
        data = _data(_fixture_package())
        sa.set_definition(data, "layouts", f"{URN}/2005/8/layout/process1", "process")
        assert sa.definition_ids(data)["layouts"] == f"{URN}/2005/8/layout/process1"
        pres = {c.get("presId") for c in data.iter(f"{sa._D}cxn") if c.get("presId")}
        assert pres == {f"{URN}/2005/8/layout/process1"}
        part = sa.definition_part("layouts", FILES["lo/process1.glo"].encode(),
                                  f"{URN}/2005/8/layout/process1")
        assert f'uniqueId="{URN}/2005/8/layout/process1"'.encode() in part

    def test_a_file_that_is_not_the_definition_named(self):
        from gvml import PackageError

        with pytest.raises(PackageError, match="does not carry the id"):
            sa.definition_part("colors", FILES["cs/colorful1.gcs"].encode(), "urn:other")


IDS = {"layouts": (f"{URN}/2005/8/layout/default", "list"),
       "colors": (f"{URN}/2005/8/colors/accent1_2", "accent1"),
       "styles": (f"{URN}/2005/8/quickstyle/simple1", "simple")}


def _ids(layout):
    return dict(IDS, layouts=(f"{URN}/2005/8/layout/{layout}", "x"))


class TestANewGraphic:
    def test_the_layout_s_sample_nodes_are_where_it_starts(self):
        data = sa.new_data_model(FILES["lo/default.glo"].encode(), IDS, None)
        assert sa.node_texts(data) == [""] * 5
        assert sa.definition_ids(data)["layouts"] == f"{URN}/2005/8/layout/default"

    def test_texts_fill_the_nodes_and_add_more_past_them(self):
        layout = FILES["lo/default.glo"].encode()
        assert sa.node_texts(sa.new_data_model(layout, IDS, ["a", "b"])) == ["a", "b", "", "", ""]
        texts = ["1", "2", "3", "4", "5", "6", "7"]
        assert sa.node_texts(sa.new_data_model(layout, IDS, texts)) == texts

    def test_a_layout_with_no_sample_starts_with_three(self):
        data = sa.new_data_model(FILES["lo/process1.glo"].encode(), _ids("process1"), ["x"])
        assert sa.node_texts(data) == ["x", "", ""]

    def test_a_sample_tree_keeps_its_shape(self):
        data = sa.new_data_model(FILES["lo/hierarchy1.glo"].encode(), _ids("hierarchy1"),
                                 ["Root", "A", "A1", "B"])
        assert sa.node_texts(data) == ["Root", "A", "A1", "B"]
        points = {p.get("modelId"): p for p in data.iter(f"{sa._D}pt")}
        ids = sa.node_ids(data)
        parent = {c.get("destId"): c.get("srcId") for c in data.iter(f"{sa._D}cxn")}
        assert points[parent[ids[0]]].get("type") == "doc"
        assert parent[ids[1]] == ids[0] and parent[ids[2]] == ids[1] and parent[ids[3]] == ids[0]

    def test_every_connection_has_its_two_transition_points(self):
        data = sa.new_data_model(FILES["lo/default.glo"].encode(), IDS, ["a"])
        points = {p.get("modelId"): p for p in data.iter(f"{sa._D}pt")}
        for cxn in data.iter(f"{sa._D}cxn"):
            assert points[cxn.get("parTransId")].get("type") == "parTrans"
            assert points[cxn.get("sibTransId")].get("type") == "sibTrans"
            assert points[cxn.get("parTransId")].get("cxnId") == cxn.get("modelId")

    def test_the_package_passes_its_own_check(self):
        from gvml.package import validate

        raw = sa.new_graphic(FILES["lo/default.glo"].encode(), FILES["cs/accent1_2.gcs"].encode(),
                             FILES["qs/simple1.gqs"].encode(), IDS, ["a"], 1, 2, 300, 400,
                             "Diagram 3", fonts={"font_size": 20, "bold": True})
        validate(raw)
        package = Package.from_bytes(raw)
        assert {p.rsplit("/", 1)[-1] for p in package.parts if "/diagrams/" in p} == {
            "data1.xml", "layout1.xml", "quickStyle1.xml", "colors1.xml"}
        drawing = package.drawing().decode()
        assert 'name="Diagram 3"' in drawing and 'r:dm="rId1"' in drawing
        assert '<a:off x="1" y="2"/><a:ext cx="300" cy="400"/>' in drawing
        data = _data(package)
        assert (sa.node_formats(data)[0]["sz"], sa.node_formats(data)[0]["b"]) == ("2000", "1")
        # An empty node has no run; every node's paragraph end carries the
        # font too, and text typed into it later takes it from there.
        ends = [e.get("sz") for e in data.iter(f"{{{sa.NS_A}}}endParaRPr")]
        assert ends.count("2000") == 5


# ---------------------------------------------------------------------------
# The three tools on the Mac, with the paste faked
# ---------------------------------------------------------------------------
@pytest.fixture
def pasted(monkeypatch, install):
    from ppt_mac import gvml_paste
    from ppt_mac import smartart as mac_smartart

    state = {}

    class FakeClipboard:
        @classmethod
        def take(cls):
            clip = cls()
            clip.warnings = []
            return clip

        def restore(self):
            state["restored"] = True

    class Shape:
        def name(self):
            return "Diagram 1"

        def left_position(self):
            return 60.0

        def top(self):
            return 100.0

        def delete(self):
            state["deleted"] = True

    def replace(pres, slide, slide_index, index, raw, clip, tool, expected, left, top,
                what, verify, alternatives):
        state["raw"] = raw
        state["problem"] = verify(Package.from_bytes(raw))
        return gvml_paste.Replaced(None, "Diagram 1", Package.from_bytes(raw), [])

    def paste(pres, slide, slide_index, raw, clip, tool, expected, left, top, alternatives):
        state["raw"], state["at"] = raw, (left, top)
        return gvml_paste.Pasted(Shape(), "Diagram 1", [])

    def copy(shape, clip, tool):
        return Package.from_bytes(state["raw"]) if "raw" in state else _fixture_package()

    monkeypatch.setattr(gvml_paste, "Clipboard", FakeClipboard)
    monkeypatch.setattr(gvml_paste, "copy_shape_package", copy)
    monkeypatch.setattr(gvml_paste, "replace_shape", replace)
    monkeypatch.setattr(gvml_paste, "paste_package", paste)
    monkeypatch.setattr(mac_smartart, "_shape_index", lambda slide, name: 1)
    monkeypatch.setattr(mac_smartart.ppt, "_get_pres_impl", lambda *a, **kw: object())
    monkeypatch.setattr(mac_smartart, "_slide", lambda pres, index: object())
    monkeypatch.setattr(mac_smartart, "_get_shape", lambda slide, name: Shape())
    monkeypatch.setattr(mac_smartart, "_require_smartart", lambda shape: None)
    import ppt_mac.shapes as mac_shapes
    monkeypatch.setattr(mac_shapes, "_shape_names", lambda slide: ["Title 1", "Diagram 1"])
    return state


def _modify(action, layout_name=None, layout_index=None, color_index=None, style_index=None):
    from ppt_mac.smartart import _modify_smartart_impl

    return _modify_smartart_impl(1, "Diagram 1", action, None, None, layout_name, layout_index,
                                 color_index, style_index, None, None, None, None,
                                 None, None, None)


@macos_only
class TestRestylingOnTheMac:
    def test_change_layout_swaps_the_part_and_names_it(self, pasted):
        result = _modify("change_layout", layout_name="Basic Process")
        assert result == {"success": True, "action": "change_layout", "total_nodes": 5,
                          "shape_name": "Diagram 1", "layout_name": "Basic Process"}
        assert pasted["problem"] is None and pasted["restored"]
        package = Package.from_bytes(pasted["raw"])
        assert sa.definition_ids(_data(package))["layouts"] == f"{URN}/2005/8/layout/process1"
        layout = next(p for p in package.parts if "/diagrams/layout" in p)
        assert b'"Basic Process"' in package.parts[layout]
        assert sa.node_texts(_data(package)) == ["Plan", "Build", "Ship", "", ""]

    def test_change_layout_by_index(self, pasted):
        assert _modify("change_layout", layout_index=4)["layout_name"] == "Hierarchy"

    def test_change_color(self, pasted):
        result = _modify("change_color", color_index=2)
        assert result["color_name"] == "Colorful - Accent Colors"
        ids = sa.definition_ids(_data(Package.from_bytes(pasted["raw"])))
        assert ids["colors"] == f"{URN}/2005/8/colors/colorful1"

    def test_change_style_with_a_colour(self, pasted):
        result = _modify("change_style", style_index=3, color_index=1)
        assert (result["style_name"], result["color_name"]) == ("Inset", "Dark 1 Outline")
        ids = sa.definition_ids(_data(Package.from_bytes(pasted["raw"])))
        assert ids["styles"] == f"{URN}/2005/8/quickstyle/3d1"
        assert ids["colors"] == f"{URN}/2005/8/colors/accent0_1"

    def test_an_index_out_of_range_copies_nothing(self, pasted):
        with pytest.raises(ValueError, match="out of range"):
            _modify("change_color", color_index=99)
        assert "raw" not in pasted and "restored" not in pasted

    def test_a_paste_that_kept_the_old_layout_is_caught(self, pasted, monkeypatch):
        from ppt_mac import gvml_paste

        def stale(pres, slide, slide_index, index, raw, clip, tool, expected, left, top,
                  what, verify, alternatives):
            pasted["problem"] = verify(_fixture_package())
            return gvml_paste.Replaced(None, "Diagram 1", _fixture_package(), [])

        monkeypatch.setattr(gvml_paste, "replace_shape", stale)
        _modify("change_layout", layout_name="Basic Process")
        assert "layout reads back as" in pasted["problem"]


def _add(**kw):
    from ppt_mac.smartart import _add_smartart_impl

    args = dict(slide_index=1, layout_name=None, layout_index=None, left=50.0, top=60.0,
                width=400.0, height=300.0, node_texts=None, color_index=None,
                style_index=None, font_name=None, font_size=None, bold=None, font_color=None)
    args.update(kw)
    return _add_smartart_impl(**args)


@macos_only
class TestAddingOnTheMac:
    def test_a_graphic_is_built_pasted_and_read_back(self, pasted):
        result = _add(layout_name="process", node_texts=["Plan", "Build"])
        assert result == {"success": True, "shape_name": "Diagram 1", "node_count": 3,
                          "layout_name": "Basic Process"}
        assert pasted["at"] == (50.0, 60.0) and pasted["restored"]
        package = Package.from_bytes(pasted["raw"])
        assert sa.node_texts(_data(package)) == ["Plan", "Build", ""]
        assert 'name="Diagram 2"' in package.drawing().decode()
        assert '<a:ext cx="5080000" cy="3810000"/>' in package.drawing().decode()

    def test_the_default_layout_colours_and_style(self, pasted):
        result = _add()
        assert result["layout_name"] == "Vertical Bullet List"
        ids = sa.definition_ids(_data(Package.from_bytes(pasted["raw"])))
        assert ids["colors"].endswith("/accent1_2") and ids["styles"].endswith("/simple1")

    def test_colour_style_and_fonts(self, pasted):
        _add(layout_index=2, color_index=2, style_index=3, font_size=18, font_color="#FF0000",
             node_texts=["a", "b", "c", "d", "e"])
        data = _data(Package.from_bytes(pasted["raw"]))
        ids = sa.definition_ids(data)
        assert ids["colors"].endswith("/colorful1") and ids["styles"].endswith("/3d1")
        assert {(f["sz"], f["font_color"]) for f in sa.node_formats(data)} == {("1800", "FF0000")}

    def test_a_paste_whose_nodes_did_not_land_is_removed(self, pasted, monkeypatch):
        from ppt_mac import gvml_paste

        monkeypatch.setattr(gvml_paste, "copy_shape_package",
                            lambda shape, clip, tool: _fixture_package())
        result = _add(node_texts=["a"])
        assert pasted["deleted"] and "success" not in result
        assert result["error"] == "ppt_add_smartart could not confirm the graphic's nodes"

    def test_a_bad_colour_or_index_pastes_nothing(self, pasted):
        with pytest.raises(ValueError):
            _add(font_color="accent1")
        with pytest.raises(ValueError, match="out of range"):
            _add(style_index=9)
        assert "raw" not in pasted


@macos_only
class TestListingOnTheMac:
    def _list(self, list_type, category=None, keyword=None, desc=False):
        from ppt_mac.smartart import _list_smartart_options_impl

        return _list_smartart_options_impl(list_type, category, keyword, desc)

    def test_layouts_in_the_windows_shape(self, install):
        result = self._list("layouts", keyword="process")
        assert result == {
            "success": True, "list_type": "layouts", "total_count": 5, "filtered_count": 1,
            "layouts": [{"index": 3, "name": "Basic Process",
                         "english_name": "Basic Process", "category": "process"}],
        }

    def test_a_category_filter(self, install):
        result = self._list("layouts", category="list")
        assert [i["name"] for i in result["layouts"]] == ["Vertical Bullet List", "Basic Block List"]

    def test_styles_with_descriptions(self, install):
        result = self._list("styles", desc=True)
        assert result["styles"][0] == {"index": 1, "name": "Simple Fill",
                                       "description": "Simple Fill, described"}

    def test_categories(self, install):
        assert self._list("categories") == {
            "success": True, "list_type": "categories", "total_layouts": 5,
            "categories": [{"category": "list", "count": 2}, {"category": "process", "count": 1},
                           {"category": "hierarchy", "count": 1},
                           {"category": "picture", "count": 1}],
        }
