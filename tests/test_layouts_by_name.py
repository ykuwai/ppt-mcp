"""Finding a template's layout by name on the Mac, through the saved file.

PowerPoint for Mac names no layouts through scripting and will not hand a
master's layout to a slide, so ``utils.layouts`` reads the names from the
deck's file and writes a one-slide copy to paste from. The package here is
built by hand: two masters that share a layout name, a custom layout, an
``mc:Ignorable`` declaration that must survive the rewrite, and one saved
slide.
"""

import os
import sys
import zipfile
from xml.etree import ElementTree

import pytest

sys.path.insert(0, "src")

from utils import layouts  # noqa: E402

macos_only = pytest.mark.skipif(
    sys.platform != "darwin", reason="the Apple Event backend needs appscript"
)

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS = {"p": P, "a": A}


def _rels(*items):
    body = "".join(f'<Relationship Id="{i}" Type="{R}/{t}" Target="{target}"/>'
                   for i, t, target in items)
    return f'<?xml version="1.0"?><Relationships xmlns="{REL}">{body}</Relationships>'


def _ph(id_, name, type_=None, idx=None):
    attrs = (f' type="{type_}"' if type_ else "") + (f' idx="{idx}"' if idx else "")
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{id_}" name="{name}"/><p:cNvSpPr/>'
            f'<p:nvPr><p:ph{attrs}/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>')


def _layout(name, type_, shapes):
    t = f' type="{type_}"' if type_ else ""
    return (f'<p:sldLayout xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}"{t}>'
            f'<p:cSld name="{name}"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/>'
            f'<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{shapes}</p:spTree></p:cSld>'
            '</p:sldLayout>')


def slide_layouts(path):
    """Each slide's id to the layout part it uses, read from the package."""
    import posixpath
    z = zipfile.ZipFile(path)
    pres = ElementTree.fromstring(z.read("ppt/presentation.xml"))
    rels = ElementTree.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
    target = {r.get("Id"): "ppt/" + r.get("Target") for r in rels}
    out = {}
    for sid in pres.findall("p:sldIdLst/p:sldId", NS):
        part = target[sid.get(f"{{{R}}}id")]
        folder, name = posixpath.split(part)
        srels = ElementTree.fromstring(z.read(f"{folder}/_rels/{name}.rels"))
        layout = next(r.get("Target") for r in srels if r.get("Type").endswith("/slideLayout"))
        out[int(sid.get("id"))] = posixpath.normpath(posixpath.join(folder, layout))
    return out


def build(path, with_slide=True):
    """Masters "Acme" (Title Only, Big Number) and "Plain" (Title Only)."""
    files = {
        "[Content_Types].xml": (
            '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '</Types>'),
        "ppt/presentation.xml": (
            f'<?xml version="1.0"?><p:presentation xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}" '
            f'xmlns:mc="{MC}" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" '
            'mc:Ignorable="p14">'
            '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/>'
            '<p:sldMasterId id="2147483660" r:id="rId2"/></p:sldMasterIdLst>'
            + ('<p:sldIdLst><p:sldId id="300" r:id="rId3"/></p:sldIdLst>' if with_slide else "")
            + '<p:sldSz cx="12192000" cy="6858000"/></p:presentation>'),
        "ppt/_rels/presentation.xml.rels": _rels(
            ("rId1", "slideMaster", "slideMasters/slideMaster1.xml"),
            ("rId2", "slideMaster", "slideMasters/slideMaster2.xml"),
            *([("rId3", "slide", "slides/slide1.xml")] if with_slide else [])),
        "ppt/slideMasters/slideMaster1.xml": (
            f'<p:sldMaster xmlns:r="{R}" xmlns:p="{P}"><p:cSld name="Acme"/>'
            '<p:sldLayoutIdLst><p:sldLayoutId id="1" r:id="rId1"/>'
            '<p:sldLayoutId id="2" r:id="rId2"/></p:sldLayoutIdLst></p:sldMaster>'),
        "ppt/slideMasters/_rels/slideMaster1.xml.rels": _rels(
            ("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml"),
            ("rId2", "slideLayout", "../slideLayouts/slideLayout2.xml"),
            ("rId3", "theme", "../theme/theme1.xml")),
        "ppt/slideMasters/slideMaster2.xml": (
            f'<p:sldMaster xmlns:r="{R}" xmlns:p="{P}"><p:cSld/>'
            '<p:sldLayoutIdLst><p:sldLayoutId id="3" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>'),
        "ppt/slideMasters/_rels/slideMaster2.xml.rels": _rels(
            ("rId1", "slideLayout", "../slideLayouts/slideLayout3.xml"),
            ("rId2", "theme", "../theme/theme2.xml")),
        "ppt/theme/theme1.xml": f'<a:theme xmlns:a="{A}" name="Acme Theme"/>',
        "ppt/theme/theme2.xml": f'<a:theme xmlns:a="{A}" name="Plain Theme"/>',
        "ppt/slideLayouts/slideLayout1.xml": _layout(
            "Title Only", "titleOnly",
            _ph(2, "Title 1", "title") + _ph(3, "Date Placeholder 2", "dt", 10)
            + _ph(4, "Slide Number Placeholder 3", "sldNum", 12)),
        "ppt/slideLayouts/slideLayout2.xml": _layout(
            "Big Number", "cust",
            _ph(2, "Title 1", "title") + _ph(3, "Number &amp; Note", None, 13)),
        "ppt/slideLayouts/slideLayout3.xml": _layout("Title Only", "titleOnly", _ph(2, "Title 1", "title")),
    }
    if with_slide:
        files["ppt/slides/slide1.xml"] = f'<p:sld xmlns:p="{P}"><p:cSld/></p:sld>'
        files["ppt/slides/_rels/slide1.xml.rels"] = _rels(
            ("rId1", "slideLayout", "../slideLayouts/slideLayout2.xml"))
    with zipfile.ZipFile(path, "w") as z:
        for name, text in files.items():
            z.writestr(name, text)
    return str(path)


class TestReadingTheFile:
    def test_masters_and_their_layouts_in_order(self, tmp_path):
        found = layouts.masters(build(tmp_path / "d.pptx"))
        assert [(m.index, m.name) for m in found] == [(1, "Acme"), (2, "Plain Theme")]
        assert [(lay.name, lay.type) for lay in found[0].layouts] == [
            ("Title Only", "titleOnly"), ("Big Number", "cust")]

    def test_a_name_matches_exactly_as_windows_matches_it(self, tmp_path):
        path = build(tmp_path / "d.pptx")
        assert [m.name for m, _ in layouts.find(path, "Title Only")] == ["Acme", "Plain Theme"]
        assert layouts.find(path, "title only") == []
        assert layouts.find(path, "Title") == []
        assert [m.index for m, _ in layouts.find(path, "Title Only", 2)] == [2]
        assert layouts.find(path, "Big Number", 2) == []



def _copy_of(tmp_path, with_slide, part="ppt/slideLayouts/slideLayout2.xml"):
    out = tmp_path / "copy.pptx"
    layouts.one_slide_copy(build(tmp_path / "d.pptx", with_slide), part, str(out))
    return zipfile.ZipFile(out)


class TestTheOneSlideCopy:
    @pytest.mark.parametrize("with_slide", [True, False])
    def test_it_holds_one_new_slide_on_the_layout(self, tmp_path, with_slide):
        z = _copy_of(tmp_path, with_slide)
        ids = ElementTree.fromstring(z.read("ppt/presentation.xml")).findall("p:sldIdLst/p:sldId", NS)
        assert len(ids) == 1
        assert slide_layouts(z.filename) == {256: "ppt/slideLayouts/slideLayout2.xml"}

    def test_the_slide_list_sits_after_the_masters_when_there_was_none(self, tmp_path):
        text = _copy_of(tmp_path, False).read("ppt/presentation.xml").decode()
        assert text.index("</p:sldMasterIdLst>") < text.index("<p:sldIdLst>") < text.index("<p:sldSz")

    def test_the_namespaces_mc_ignorable_names_are_kept(self, tmp_path):
        text = _copy_of(tmp_path, True).read("ppt/presentation.xml").decode()
        assert 'mc:Ignorable="p14"' in text and "xmlns:p14=" in text

    def test_the_layout_s_placeholders_come_without_date_and_number(self, tmp_path):
        z = _copy_of(tmp_path, True, "ppt/slideLayouts/slideLayout1.xml")
        name = next(n for n in z.namelist() if n.startswith("ppt/slides/slide") and n != "ppt/slides/slide1.xml")
        slide = ElementTree.fromstring(z.read(name))
        assert [s.find("p:nvSpPr/p:cNvPr", NS).get("name") for s in slide.iter(f"{{{P}}}sp")] == ["Title 1"]

    def test_a_name_with_markup_characters_is_escaped(self, tmp_path):
        z = _copy_of(tmp_path, True)
        name = next(n for n in z.namelist() if n.startswith("ppt/slides/slide") and n != "ppt/slides/slide1.xml")
        slide = ElementTree.fromstring(z.read(name))
        names = [s.find("p:nvSpPr/p:cNvPr", NS).get("name") for s in slide.iter(f"{{{P}}}sp")]
        assert names == ["Title 1", "Number & Note"]
        ph = list(slide.iter(f"{{{P}}}ph"))[1]
        assert ph.get("idx") == "13"

    def test_the_new_part_has_a_content_type_and_a_relationship(self, tmp_path):
        z = _copy_of(tmp_path, True)
        assert "/ppt/slides/slide2.xml" in z.read("[Content_Types].xml").decode()
        rels = z.read("ppt/_rels/presentation.xml.rels").decode()
        assert rels.count("/relationships/slide\"") == 1 and "slides/slide2.xml" in rels

    def test_a_layout_not_in_the_deck_is_refused(self, tmp_path):
        with pytest.raises(ValueError, match="not in"):
            layouts.one_slide_copy(build(tmp_path / "d.pptx"), "ppt/slideLayouts/nope.xml",
                                   str(tmp_path / "c.pptx"))


# ---------------------------------------------------------------------------
# The Mac side, with PowerPoint faked
# ---------------------------------------------------------------------------
class _Ref:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value

    def __call__(self):
        return self.value


class _Slide:
    def __init__(self, layout):
        self.custom_layout = _Ref(layout)


class _Slides(dict):
    def __getitem__(self, i):
        return dict.__getitem__(self, i)


class _Pres:
    def __init__(self, path, slide_ids, slide_layouts):
        self.full_name = _Ref(path)
        self.ids = slide_ids
        self.slides = _Slides({i + 1: _Slide(lay) for i, lay in enumerate(slide_layouts)})


@pytest.fixture
def mac(monkeypatch):
    from ppt_mac import slides as ms

    monkeypatch.setattr(ms, "_slide_ids", lambda pres: list(pres.ids))
    return ms


@macos_only
class TestFindingItOnTheMac:
    def test_an_unsaved_deck_says_to_save_it(self, mac, tmp_path):
        with pytest.raises(ValueError, match="has not been saved"):
            mac._layout_by_name(None, _Pres(str(tmp_path / "gone.pptx"), [], []), "Title Only", None, 1)

    def test_an_unknown_name_lists_the_layouts(self, mac, tmp_path):
        pres = _Pres(build(tmp_path / "d.pptx"), [], [])
        with pytest.raises(ValueError, match=r"not found.*'Acme': \['Title Only', 'Big Number'\].*saved file"):
            mac._layout_by_name(None, pres, "Agenda", None, 1)
        with pytest.raises(ValueError, match=r"Design index 3 out of range \(1-2\)"):
            mac._layout_by_name(None, pres, "Title Only", 3, 1)

    @pytest.fixture
    def pasting(self, mac, tmp_path, monkeypatch):
        from backend import mac_ae
        from ppt_mac import presentation

        staged = tmp_path / "staging"
        monkeypatch.setattr(mac_ae, "EXPORT_STAGING_DIR", str(staged))
        state = {"events": [], "staged": staged, "fail_at": None, "deleted": []}

        class Helper:
            def close(self, saving):
                state["events"].append("closed")

        def open_copy(app, path):
            state["events"].append(("opened", sorted(slide_layouts(path).values())))
            return Helper()

        def paste(app, src, index, dst):
            n = sum(1 for e in state["events"] if e == "pasted") + 1
            if n == state["fail_at"]:
                raise RuntimeError("paste failed")
            state["events"].append("pasted")
            dst.ids.append(900 + n)
            return 900 + n

        monkeypatch.setattr(presentation, "_open_and_find", open_copy)
        monkeypatch.setattr(mac, "_paste_copy", paste)
        monkeypatch.setattr(mac, "_remove_slides",
                            lambda pres, ids: state["deleted"].extend(reversed(list(ids))))
        return state

    def test_each_slide_is_pasted_from_a_one_slide_copy(self, mac, tmp_path, pasting):
        pres = _Pres(build(tmp_path / "d.pptx"), [300], ["x"])
        found = mac._layout_by_name(None, pres, "Title Only", None, 2)
        assert found["slide_ids"] == [901, 902] and found["design_name"] == "Acme"
        assert found["ambiguous"] == ["Acme", "Plain Theme"]
        assert pasting["events"] == [("opened", ["ppt/slideLayouts/slideLayout1.xml"]),
                                     "pasted", "pasted", "closed"]
        assert os.listdir(pasting["staged"]) == []

    def test_a_failed_paste_removes_the_slides_already_pasted(self, mac, tmp_path, pasting):
        pasting["fail_at"] = 3
        pres = _Pres(build(tmp_path / "d.pptx"), [300], ["x"])
        with pytest.raises(RuntimeError, match="paste failed"):
            mac._layout_by_name(None, pres, "Big Number", None, 3)
        assert pasting["deleted"] == [902, 901]
        assert pasting["events"][-1] == "closed" and os.listdir(pasting["staged"]) == []


def test_the_saved_file_says_which_layout_each_slide_uses(tmp_path):
    path = build(tmp_path / "d.pptx")
    assert layouts.slide_layouts(path) == {300: "ppt/slideLayouts/slideLayout2.xml"}
    assert layouts.slide_layouts(build(tmp_path / "e.pptx", with_slide=False)) == {}


@macos_only
class TestLikeASlideWhoseLayoutWillNotPass:
    def test_its_layout_is_read_from_the_file_and_pasted(self, mac, tmp_path, monkeypatch):
        calls = []
        monkeypatch.setattr(mac, "_paste_on_layout",
                            lambda app, pres, path, part, count: calls.append((part, count)) or [950])
        pres = _Pres(build(tmp_path / "d.pptx"), [256, 300], ["a", "b"])
        found = mac._like_saved_slide(None, pres, 2, 1)
        assert found == {"layout_name": "Big Number", "design_name": "Acme", "slide_ids": [950]}
        assert calls == [("ppt/slideLayouts/slideLayout2.xml", 1)]

    def test_a_slide_not_saved_yet_says_to_save(self, mac, tmp_path):
        pres = _Pres(build(tmp_path / "d.pptx"), [256, 300, 301], ["a", "b", "c"])
        with pytest.raises(ValueError, match="slide 3 is not in it yet. Save the deck"):
            mac._like_saved_slide(None, pres, 3, 1)
