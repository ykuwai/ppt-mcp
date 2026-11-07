"""ppt_get_all_text on macOS reads a text shape's paragraphs from the package
`copy shape` writes for it.

Asked for a range, PowerPoint for Mac answers its bold or italic as true when
any of it is (a body with one bold word read as bold throughout, and came out
as one "##" heading), and a per-paragraph walk cost about 250 Apple Events a
slide. The package carries each run's emphasis and each paragraph's bullet and
level, with what a placeholder inherits resolved into the text body's list
style, in one event a shape. A deck that is not in front gets no package and
is read shape by shape as before.
"""

import sys
import xml.etree.ElementTree as ET

import pytest

sys.path.insert(0, "src")

pytestmark = pytest.mark.skipif(
    sys.platform != "darwin", reason="the Apple Event backend needs appscript"
)

NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _shape(paragraphs, lst_style=""):
    xml = (f'<a:sp xmlns:a="{NS}"><a:txSp><a:txBody><a:bodyPr/>'
           f'<a:lstStyle>{lst_style}</a:lstStyle>{"".join(paragraphs)}</a:txBody></a:txSp></a:sp>')
    return ET.fromstring(xml)


def _p(*runs, ppr=""):
    out = []
    for text, attrs in runs:
        if text == "\n":
            out.append("<a:br/>")
        else:
            out.append(f'<a:r><a:rPr lang="en-US" {attrs}/><a:t>{text}</a:t></a:r>')
    return f"<a:p>{ppr}{''.join(out)}</a:p>"


# A body placeholder as `copy shape` writes it: bullets come from the levels
# of the list style, not from the paragraphs.
BODY_STYLE = ('<a:lvl1pPr marL="342900"><a:buChar char="•"/><a:defRPr sz="3200"/></a:lvl1pPr>'
              '<a:lvl2pPr marL="742950"><a:buChar char="–"/><a:defRPr sz="2800"/></a:lvl2pPr>')


def test_bullets_and_emphasis_come_from_the_list_style_unless_set():
    from ppt_mac.text import _package_paragraphs

    shape = _shape([
        _p(("Revenue grew", "")),
        _p(("Costs ", ""), ("fell", 'b="1"'), (" sharply", 'i="1"')),
        _p(("Sub point", ""), ppr='<a:pPr lvl="1"/>'),
        _p(("No bullet here", ""), ppr="<a:pPr><a:buNone/></a:pPr>"),
    ], BODY_STYLE)
    paragraphs = _package_paragraphs(shape)
    assert [(p["bullet"], p["level"]) for p in paragraphs] == [
        ("bullet", 1), ("bullet", 1), ("bullet", 2), (None, 1)]
    assert [(r["text"], r["bold"], r["italic"]) for r in paragraphs[1]["runs"]] == [
        ("Costs ", False, False), ("fell", True, False), (" sharply", False, True)]


def test_a_level_default_bold_holds_unless_a_run_turns_it_off():
    from ppt_mac.text import _package_paragraphs

    shape = _shape([_p(("Bold by style", ""), ("plain", 'b="0"'))],
                   '<a:lvl1pPr><a:defRPr b="1"/></a:lvl1pPr>')
    runs = _package_paragraphs(shape)[0]["runs"]
    assert [(r["text"], r["bold"]) for r in runs] == [("Bold by style", True), ("plain", False)]


def _info(paragraphs, **flags):
    info = {"shape": object(), "paragraphs": paragraphs, "is_title": False, "is_subtitle": False,
            "has_table": False, "top": 0, "left": 0, "width": 100, "height": 20}
    info.update(flags)
    return info


def test_a_body_with_one_bold_word_is_a_list_not_a_heading():
    from ppt_mac.text import _package_paragraphs, _shape_info_to_markdown

    paragraphs = _package_paragraphs(_shape([
        _p(("Revenue grew", "")),
        _p(("Costs ", ""), ("fell", 'b="1"')),
        _p(("Sub point", ""), ppr='<a:pPr lvl="1"/>'),
        _p(),
        _p(("Line one", ""), ("\n", ""), ("line two", "")),
    ], BODY_STYLE))
    assert _shape_info_to_markdown(_info(paragraphs)) == (
        "- Revenue grew\n- Costs **fell**\n  - Sub point\n\n- Line one\nline two")


def test_titles_subtitles_and_all_bold_shapes():
    from ppt_mac.text import _package_paragraphs, _shape_info_to_markdown

    title = _package_paragraphs(_shape([_p(("Annual", "")), _p(("Report", ""))]))
    assert _shape_info_to_markdown(_info(title, is_title=True)) == "# Annual Report"
    subtitle = _package_paragraphs(_shape([_p(("For the ", ""), ("board", 'b="1"'))]))
    assert _shape_info_to_markdown(_info(subtitle, is_subtitle=True)) == "For the **board**"
    bold = _package_paragraphs(_shape([_p(("All bold", 'b="1"'), ("\n", ""), ("still", 'b="1"'))]))
    assert _shape_info_to_markdown(_info(bold), subheading_level="###") == "### All bold still"
    numbered = _package_paragraphs(_shape([_p(("Plan", ""), ppr='<a:pPr><a:buAutoNum type="arabicPeriod"/></a:pPr>')]))
    assert _shape_info_to_markdown(_info(numbered)) == "1. Plan"


class _Shape:
    """A slide shape for the collector; text frame reads are counted."""

    def __init__(self, name, text=True):
        self.name, self.text, self.slow_reads = name, text, 0

    @property
    def text_frame(self):
        shape = self

        class Frame:
            def has_text(self):
                shape.slow_reads += 1
                return shape.text

        return Frame()


def _record(top, kind="text"):
    from appscript import k

    from backend.mac_enums import MsoShapeType, to_keyword
    from ppt_com.constants import msoTextBox

    return {k.shape_type: to_keyword(MsoShapeType, msoTextBox, "shape type"), k.placeholder_type: k.missing_value,
            k.top: top, k.left_position: 10.0, k.width: 100.0, k.height: 20.0,
            k.has_table: False, k.has_text_frame: True}


def test_each_text_shape_is_read_from_its_package_and_one_that_will_not_copy_the_slow_way(monkeypatch):
    from ppt_mac import text

    copied, refused = _Shape("Copied"), _Shape("Refused")
    monkeypatch.setattr(text, "_shape_records", lambda slide: [(copied, _record(10.0)), (refused, _record(50.0))])
    paragraphs = [{"runs": [{"text": "Hello", "bold": False, "italic": False}], "bullet": None, "level": 1}]
    monkeypatch.setattr(text, "_shape_package_paragraphs",
                        lambda shape, clip: paragraphs if shape is copied else None)
    infos = text._collect_text_shapes(object(), clip=object())
    assert [i["shape"] for i in infos] == [copied, refused]
    assert infos[0]["paragraphs"] == paragraphs and "paragraphs" not in infos[1]
    assert (infos[0]["top"], infos[0]["left"], infos[0]["width"]) == (10.0, 10.0, 100.0)
    assert copied.slow_reads == 0 and refused.slow_reads == 1


def test_a_copied_shape_with_no_text_is_left_out(monkeypatch):
    from ppt_mac import text

    empty = _Shape("Empty")
    monkeypatch.setattr(text, "_shape_records", lambda slide: [(empty, _record(10.0))])
    monkeypatch.setattr(text, "_shape_package_paragraphs",
                        lambda shape, clip: [{"runs": [], "bullet": None, "level": 1}])
    assert text._collect_text_shapes(object(), clip=object()) == []


def test_the_clipboard_is_taken_only_for_the_deck_in_front_and_put_back(monkeypatch):
    from types import SimpleNamespace

    from ppt_mac import gvml_paste, text

    taken, seen = [], []

    class Clip:
        warnings = []

        @classmethod
        def take(cls):
            clip = cls()
            taken.append(clip)
            clip.restored = False
            return clip

        def restore(self):
            self.restored = True

    deck = SimpleNamespace(slides={1: "slide 1"})
    monkeypatch.setattr(text.ppt, "_get_pres_impl", lambda: deck, raising=False)
    monkeypatch.setattr(text, "count", lambda ref: 1)
    monkeypatch.setattr(text, "_slide_to_markdown", lambda slide, idx, clip=None: seen.append(clip) or "== Slide 1 ==")
    monkeypatch.setattr(gvml_paste, "Clipboard", Clip)

    monkeypatch.setattr(text, "_in_front", lambda pres: True)
    text._get_all_text_impl(None)
    assert len(taken) == 1 and seen[-1] is taken[0] and taken[0].restored

    monkeypatch.setattr(text, "_in_front", lambda pres: False)
    text._get_all_text_impl(None)
    assert len(taken) == 1 and seen[-1] is None
