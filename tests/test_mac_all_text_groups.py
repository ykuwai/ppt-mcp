"""ppt_get_all_text reads the text inside groups on macOS.

A group answers no members over Apple Events (``shapes`` counts 0), so a
slide whose text was all inside groups came back as "(no text)". The members
are read from the package ``copy shape`` writes for the group instead, in
the shape PowerPoint for Mac 16.114 writes it, faked here.
"""

import sys
import xml.etree.ElementTree as ET

import pytest

sys.path.insert(0, "src")

pytestmark = pytest.mark.skipif(
    sys.platform != "darwin", reason="the Apple Event backend needs appscript"
)

NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _sp(name, x, y, paragraphs):
    body = "".join(paragraphs)
    return (f'<a:sp><a:nvSpPr><a:cNvPr id="2" name="{name}"/><a:cNvSpPr/></a:nvSpPr>'
            f'<a:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="1270000" cy="635000"/></a:xfrm></a:spPr>'
            f'<a:txSp><a:txBody><a:bodyPr/><a:lstStyle/>{body}</a:txBody></a:txSp></a:sp>')


def _p(*runs, ppr=""):
    out = []
    for text, attrs in runs:
        if text == "\n":
            out.append("<a:br/>")
        else:
            out.append(f'<a:r><a:rPr lang="en-US" {attrs}/><a:t>{text}</a:t></a:r>')
    return f"<a:p>{ppr}{''.join(out)}</a:p>"


GROUP = (
    '<a:grpSp><a:nvGrpSpPr><a:cNvPr id="1" name="Group 4"/><a:cNvGrpSpPr/></a:nvGrpSpPr>'
    '<a:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5080000" cy="5080000"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="5080000" cy="5080000"/></a:xfrm></a:grpSpPr>'
    + _sp("Heading", 0, 0, [_p(("Card heading", 'b="1"'))])
    + _sp("Body", 0, 1270000, [
        _p(("Point ", ""), ("one", 'i="1"')),
        _p(("two", ""), ppr='<a:pPr lvl="1"><a:buChar char="•"/></a:pPr>'),
        _p(("a", ""), ("\n", ""), ("b", "")),
    ])
    + _sp("Empty", 0, 2540000, [_p()])
    + '<a:grpSp><a:nvGrpSpPr><a:cNvPr id="9" name="Inner"/><a:cNvGrpSpPr/></a:nvGrpSpPr>'
    '<a:grpSpPr><a:xfrm><a:off x="2540000" y="3810000"/><a:ext cx="1270000" cy="635000"/>'
    '<a:chOff x="2540000" y="3810000"/><a:chExt cx="1270000" cy="635000"/></a:xfrm></a:grpSpPr>'
    + _sp("Note", 2540000, 3810000, [_p(("Nested", ""))])
    + '</a:grpSp></a:grpSp>'
)


def _drawing():
    from gvml import canvas

    return canvas.wrap(GROUP, 0, 0, 5080000, 5080000)


class _Package:
    def drawing(self):
        return _drawing().encode("utf-8")


@pytest.fixture
def copied(monkeypatch):
    from ppt_mac import gvml_paste

    events = []

    class FakeClipboard:
        warnings = ()

        @classmethod
        def take(cls):
            events.append("take")
            return cls()

        def restore(self):
            events.append("restore")

    monkeypatch.setattr(gvml_paste, "Clipboard", FakeClipboard)
    monkeypatch.setattr(gvml_paste, "copy_shape_package", lambda shape, clip, tool: _Package())
    return events


def test_each_member_with_text_is_an_entry_at_its_box(copied):
    from ppt_mac.text import _group_text_infos

    infos = _group_text_infos(object())
    assert [(i["top"], i["left"]) for i in infos] == [(0.0, 0.0), (100.0, 0.0), (300.0, 200.0)]
    assert all(i["shape"] is None and not i["is_title"] for i in infos)
    assert copied == ["take", "restore"]


def test_the_markdown_keeps_emphasis_bullets_and_breaks(copied):
    from ppt_mac.text import _group_text_infos, _shape_info_to_markdown

    heading, body, nested = _group_text_infos(object())
    assert _shape_info_to_markdown(heading) == "## Card heading"
    assert _shape_info_to_markdown(heading, subheading_level="###") == "### Card heading"
    assert _shape_info_to_markdown(body) == "Point *one*\n  - two\na\nb"
    assert _shape_info_to_markdown(nested) == "Nested"


def test_a_group_whose_package_cannot_be_read_is_skipped(monkeypatch):
    from ppt_mac import gvml_paste
    from ppt_mac.text import _group_text_infos

    class FakeClipboard:
        warnings = ()

        @classmethod
        def take(cls):
            return cls()

        def restore(self):
            pass

    def refuse(shape, clip, tool):
        raise gvml_paste.Refused({"error": "no package"})

    monkeypatch.setattr(gvml_paste, "Clipboard", FakeClipboard)
    monkeypatch.setattr(gvml_paste, "copy_shape_package", refuse)
    assert _group_text_infos(object()) == []


def test_paragraphs_as_read_from_a_member():
    from ppt_mac.text import _package_paragraphs

    root = ET.fromstring(_drawing())
    body = next(e for e in root.iter(f"{{{NS}}}sp")
                if e.find(f"{{{NS}}}nvSpPr/{{{NS}}}cNvPr").get("name") == "Body")
    paragraphs = _package_paragraphs(body)
    assert [p["bullet"] for p in paragraphs] == [None, "bullet", None]
    assert [p["level"] for p in paragraphs] == [1, 2, 1]
    assert paragraphs[0]["runs"][1] == {"text": "one", "bold": False, "italic": True}
