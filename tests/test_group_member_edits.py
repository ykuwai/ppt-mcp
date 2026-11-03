"""Part of a group member's text, and its formatting, on macOS (after #294).

#294 writes a member's whole text through the group's clipboard package.
This adds a span (``search_text`` or ``start``/``length``) inside one line,
which keeps its run's formatting, and ``ppt_format_text`` on a member. The
XML half is tested everywhere; the Mac half with the copy and paste faked,
since that route (``replace_shape``) is #294's and is unchanged.
"""

import sys
import xml.etree.ElementTree as ET

import pytest

sys.path.insert(0, "src")

from gvml import text as gt  # noqa: E402
from gvml.package import NS_A  # noqa: E402

macos_only = pytest.mark.skipif(
    sys.platform != "darwin", reason="the Apple Event backend needs appscript"
)

A = f"{{{NS_A}}}"
WHITE = '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'


def body(paragraphs):
    return ET.fromstring(
        f'<a:sp xmlns:a="{NS_A}"><a:nvSpPr><a:cNvPr id="2" name="Body"/><a:cNvSpPr/></a:nvSpPr>'
        f'<a:spPr/><a:txSp><a:txBody><a:bodyPr/><a:lstStyle/>{paragraphs}</a:txBody></a:txSp></a:sp>')


CARD = (f'<a:p><a:r><a:rPr lang="en-US">{WHITE}</a:rPr><a:t>First line</a:t></a:r></a:p>'
        f'<a:p><a:r><a:rPr lang="en-US" b="1">{WHITE}</a:rPr><a:t>Second</a:t></a:r>'
        f'<a:r><a:rPr lang="en-US">{WHITE}</a:rPr><a:t> line</a:t></a:r>'
        '<a:endParaRPr lang="en-US"/></a:p>')


@pytest.fixture
def card():
    return body(CARD)


class TestReplacingASpan:
    def test_inside_one_run_keeps_its_formatting(self, card):
        gt.replace_span(card, 12, 6, "2nd")
        assert gt.text_of(card) == "First line\n2nd line"
        assert gt.run_formats(card)[1]["b"] == "1"

    def test_across_runs_takes_the_first_runs_formatting(self, card):
        gt.replace_span(card, 15, 6, "X")
        assert gt.text_of(card) == "First line\nSecXne"

    def test_an_insertion_at_the_end(self, card):
        gt.replace_span(card, 23, 0, "!")
        assert gt.text_of(card) == "First line\nSecond line!"

    def test_typing_at_a_run_boundary_goes_on_in_the_run_before(self, card):
        gt.replace_span(card, 18, 0, "s")
        assert gt.text_of(card) == "First line\nSeconds line"
        assert gt.run_formats(card)[1]["b"] == "1"

    def test_an_empty_paragraph_takes_its_end_formatting(self):
        shape = body('<a:p><a:endParaRPr lang="en-US" sz="2400"/></a:p>')
        gt.replace_span(shape, 1, 0, "New")
        assert gt.text_of(shape) == "New" and gt.run_formats(shape)[0]["sz"] == "2400"

    @pytest.mark.parametrize("start,length,text", [
        (9, 4, "x"),       # across the paragraph break
        (1, 2, "a\nb"),    # a new paragraph in the text
        (1, 1, "a\vb"),    # a line break in the text
        (40, 1, "x"),      # past the end
    ])
    def test_what_crosses_a_line_is_refused_and_nothing_changes(self, card, start, length, text):
        before = ET.tostring(card)
        with pytest.raises(ValueError, match="Pass the whole text"):
            gt.replace_span(card, start, length, text)
        assert ET.tostring(card) == before


class TestFormatting:
    def test_every_run_and_the_paragraph_ends(self, card):
        n = gt.format_runs(card, font_size=20, bold=False, italic=True, underline=True,
                           color="accent2", font_name="Georgia")
        assert n == 3
        assert gt.run_formats(card) == [{"sz": "2000", "b": "0", "i": "1", "u": "sng",
                                         "color": "accent2", "latin": "Georgia"}] * 3
        end = card.find(f"{A}txSp/{A}txBody/{A}p[2]/{A}endParaRPr")
        assert end.get("sz") == "2000"

    def test_a_hex_colour_replaces_the_old_fill(self, card):
        gt.format_runs(card, color="#1f2937")
        rpr = card.find(f"{A}txSp/{A}txBody/{A}p/{A}r/{A}rPr")
        assert [c.tag[len(A):] for c in rpr] == ["solidFill"]
        assert rpr.find(f"{A}solidFill/{A}srgbClr").get("val") == "1F2937"

    def test_children_go_where_the_schema_has_them(self, card):
        gt.format_runs(card, font_name="Georgia", color="dark1")
        rpr = card.find(f"{A}txSp/{A}txBody/{A}p/{A}r/{A}rPr")
        assert [c.tag[len(A):] for c in rpr] == ["solidFill", "latin", "ea"]
        assert rpr.find(f"{A}solidFill/{A}schemeClr").get("val") == "dk1"

    def test_a_run_with_no_properties_gets_some(self):
        shape = body('<a:p><a:r><a:t>x</a:t></a:r></a:p>')
        gt.format_runs(shape, bold=True)
        assert gt.run_formats(shape)[0]["b"] == "1"


# ---------------------------------------------------------------------------
# The Mac tools
# ---------------------------------------------------------------------------
@pytest.fixture
def member(monkeypatch):
    """Stand in for _edit_member: run the edit on the card, then the check on
    the card as edited, as a paste would read it back."""
    pytest.importorskip("appscript")
    import ppt_mac.groups as groups

    state = {"shape": body(CARD), "names": []}

    def edit_member(slide_index, name, tool_name, edit):
        state["names"].append((tool_name, name))
        if name.split("/")[-1] != "Body":
            return None
        check, fields = edit(state["shape"], "Body")
        state["problem"] = check(state["shape"])
        return dict(fields, status="success", shape_name="Body")

    monkeypatch.setattr(groups, "_edit_member", edit_member)
    return state


@macos_only
class TestTheMacTools:
    def test_a_search_replaces_inside_the_member(self, member):
        from ppt_mac.groups import set_member_text

        result = set_member_text(1, "Group 4/Body", "2nd", search_text="Second")
        assert (result["start"], result["replaced_length"], result["text"]) == (
            12, 6, "First line\n2nd line")
        assert member["problem"] is None

    def test_a_start_and_length(self, member):
        from ppt_mac.groups import set_member_text

        result = set_member_text(1, "Body", "Top", start=1, length=5)
        assert result["text"] == "Top line\nSecond line"

    def test_the_whole_text_as_before(self, member):
        from ppt_mac.groups import set_member_text

        result = set_member_text(1, "Body", "New body")
        assert (result["text_length"], result["paragraph_count"]) == (8, 1)
        assert "text" not in result

    def test_a_search_that_finds_nothing_says_so(self, member):
        from ppt_mac.groups import set_member_text

        with pytest.raises(ValueError, match="'Third' not found"):
            set_member_text(1, "Body", "x", search_text="Third")

    def test_formatting_is_written_and_checked(self, member):
        from ppt_mac.groups import format_member_text

        result = format_member_text(1, "Body", font_name="Georgia", font_size=20,
                                    bold=True, color="Accent 2")
        assert result["runs_formatted"] == 3 and member["problem"] is None
        assert {r["color"] for r in gt.run_formats(member["shape"])} == {"accent2"}

    def test_a_name_in_no_group_is_none(self, member):
        from ppt_mac.groups import format_member_text

        assert format_member_text(1, "Nope", bold=True) is None


@macos_only
class TestRoutingFromTheTools:
    @pytest.fixture
    def off_the_slide(self, monkeypatch):
        pytest.importorskip("appscript")
        import ppt_mac.groups as groups
        import ppt_mac.text as mac_text

        calls = []

        def missing(*args, **kwargs):
            raise ValueError("Shape 'Body' not found on slide")

        monkeypatch.setattr(mac_text, "_text_frame_of", missing)
        monkeypatch.setattr(mac_text, "_get_shape", missing)
        monkeypatch.setattr(mac_text, "_top_level_names", lambda slide_index: ["Group 4"])
        monkeypatch.setattr(mac_text, "goto_slide", lambda app, index: None)
        monkeypatch.setattr(mac_text.ppt, "_get_app_impl", lambda *a, **kw: object())

        class Pres:
            slides = {1: object()}

        monkeypatch.setattr(mac_text.ppt, "_get_pres_impl", lambda *a, **kw: Pres())
        monkeypatch.setattr(groups, "set_member_text",
                            lambda *args: calls.append(("set", args)) or {"status": "success"})
        monkeypatch.setattr(groups, "format_member_text",
                            lambda *args: calls.append(("format", args)) or {"status": "success"})
        return calls

    def test_set_text_hands_a_search_to_the_group(self, off_the_slide):
        from ppt_mac.text import _set_text_impl

        _set_text_impl(1, "Body", "2nd", search_text="Second")
        assert off_the_slide == [("set", (1, "Body", "2nd", None, None, "Second", 1))]

    def test_runs_are_refused_before_any_group_is_read(self, off_the_slide):
        from ppt_mac.text import _set_text_impl

        with pytest.raises(ValueError, match="runs cannot be written"):
            _set_text_impl(1, "Body", None, runs=[{"text": "a"}])
        assert off_the_slide == []

    def test_format_text_hands_a_member_to_the_group(self, off_the_slide):
        from ppt_mac.text import _format_text_impl

        _format_text_impl(1, "Body", "Georgia", None, 20, True, None, None, "#FF0000",
                          "accent2", None)
        # The theme colour, when given, is the one written, as on a shape.
        assert off_the_slide == [("format", (1, "Body", "Georgia", None, 20, True, None,
                                             None, "accent2"))]

    def test_a_highlight_is_refused_before_any_group_is_read(self, off_the_slide):
        from ppt_mac.text import _format_text_impl

        with pytest.raises(ValueError, match="highlight cannot be written"):
            _format_text_impl(1, "Body", None, None, None, None, None, None, None, None,
                              "#FFFF00")
        assert off_the_slide == []
