"""Text and formatting on a group's member on macOS (#220).

The XML half (``gvml.text``) is tested on every platform, against a group in
the shape PowerPoint for Mac 16.114 wrote it to the clipboard. The Mac half
(``ppt_mac.text`` handing a member to ``ppt_mac.group_text``) is tested with
the paste faked; the paste itself is ``replace_shape``, tested with charts
and freeforms, and the whole route was run live.
"""

import copy
import sys
import xml.etree.ElementTree as ET

import pytest

sys.path.insert(0, "src")

from gvml import text as gt  # noqa: E402
from gvml.package import NS_A, PackageError  # noqa: E402

macos_only = pytest.mark.skipif(
    sys.platform != "darwin", reason="the Apple Event backend needs appscript"
)

A = f"{{{NS_A}}}"
WHITE = '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'


def _sp(name, body):
    return (f'<a:sp><a:nvSpPr><a:cNvPr id="2" name="{name}"/><a:cNvSpPr txBox="1"/></a:nvSpPr>'
            '<a:spPr/>' + (f'<a:txSp><a:txBody><a:bodyPr/><a:lstStyle/>{body}</a:txBody></a:txSp>'
                           if body is not None else '') + '</a:sp>')


GROUP = (
    f'<a:grpSp xmlns:a="{NS_A}"><a:nvGrpSpPr><a:cNvPr id="1" name="Group 4"/><a:cNvGrpSpPr/></a:nvGrpSpPr>'
    '<a:grpSpPr/>'
    + _sp("Card", None)
    + _sp("Title", f'<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="2800">{WHITE}</a:rPr>'
                   f'<a:t>Card title</a:t></a:r><a:endParaRPr lang="en-US" sz="2800"/></a:p>')
    + _sp("Body", f'<a:p><a:r><a:rPr lang="en-US">{WHITE}</a:rPr><a:t>First line</a:t></a:r></a:p>'
                  f'<a:p><a:r><a:rPr lang="en-US" b="1">{WHITE}</a:rPr><a:t>Second</a:t></a:r>'
                  f'<a:r><a:rPr lang="en-US">{WHITE}</a:rPr><a:t> line</a:t></a:r>'
                  '<a:endParaRPr lang="en-US"/></a:p>')
    + '<a:grpSp><a:nvGrpSpPr><a:cNvPr id="9" name="Inner"/><a:cNvGrpSpPr/></a:nvGrpSpPr><a:grpSpPr/>'
    + _sp("Note", '<a:p><a:r><a:rPr lang="en-US"/><a:t>a</a:t></a:r><a:br><a:rPr lang="en-US"/></a:br>'
                  '<a:r><a:rPr lang="en-US"/><a:t>b</a:t></a:r></a:p>')
    + _sp("Title", '<a:p><a:r><a:t>Inner title</a:t></a:r></a:p>')
    + '</a:grpSp></a:grpSp>'
)


@pytest.fixture
def group():
    return ET.fromstring(GROUP)


def member(group, identifier):
    return gt.find_member(group, "Group 4", identifier)[1]


class TestFinding:
    def test_every_member_at_any_depth_with_its_path(self, group):
        assert [p for p, _ in gt.members(group, "Group 4")] == [
            "Group 4/Card", "Group 4/Title", "Group 4/Body", "Group 4/Inner",
            "Group 4/Inner/Note", "Group 4/Inner/Title",
        ]

    def test_a_unique_bare_name(self, group):
        assert gt.find_member(group, "Group 4", "Note")[0] == "Group 4/Inner/Note"

    def test_a_path_names_one_of_two_of_the_same_name(self, group):
        path, found = gt.find_member(group, "Group 4", "Group 4/Inner/Title")
        assert path == "Group 4/Inner/Title" and gt.read_text(found) == "Inner title"

    def test_a_bare_name_two_members_share_is_refused_with_both_paths(self, group):
        with pytest.raises(gt.AmbiguousMember, match="Group 4/Title, Group 4/Inner/Title"):
            gt.find_member(group, "Group 4", "Title")

    def test_a_name_that_is_not_there(self, group):
        with pytest.raises(gt.MemberNotFound):
            gt.find_member(group, "Group 4", "Nope")


class TestReading:
    def test_paragraphs_and_line_breaks_as_the_tools_see_them(self, group):
        assert gt.read_text(member(group, "Body")) == "First line\nSecond line"
        assert gt.read_text(member(group, "Note")) == "a\vb"

    def test_a_member_with_no_text(self, group):
        assert gt.text_body(member(group, "Card")) is None
        with pytest.raises(PackageError, match="takes no text"):
            gt.read_text(member(group, "Card"))


class TestWriting:
    def test_the_whole_text_keeps_each_paragraphs_formatting(self, group):
        body = member(group, "Body")
        gt.write_text(body, "One\nTwo\vbreak\nThree")
        assert gt.read_text(body) == "One\nTwo\vbreak\nThree"
        # The second paragraph and the third (more than there were) take the
        # old second paragraph's first run: bold.
        assert [r["b"] for r in gt.run_formats(body)] == [None, "1", "1", "1"]
        assert all(r["color"] == "FFFFFF" for r in gt.run_formats(body))

    def test_paragraph_settings_are_kept(self, group):
        title = member(group, "Group 4/Title")
        gt.write_text(title, "New")
        assert title.find(f"{A}txSp/{A}txBody/{A}p/{A}pPr").get("algn") == "ctr"
        assert gt.run_formats(title)[0]["sz"] == "2800"

    def test_empty_text_leaves_one_empty_paragraph(self, group):
        title = member(group, "Group 4/Title")
        gt.write_text(title, "")
        assert gt.read_text(title) == ""


class TestReplacingASpan:
    def test_inside_one_run_keeps_its_formatting(self, group):
        body = member(group, "Body")
        gt.replace_span(body, 12, 6, "2nd")
        assert gt.read_text(body) == "First line\n2nd line"
        assert gt.run_formats(body)[1]["b"] == "1"

    def test_across_runs_takes_the_first_runs_formatting(self, group):
        body = member(group, "Body")
        gt.replace_span(body, 15, 6, "X")  # "ond li" across the bold run and the next
        assert gt.read_text(body) == "First line\nSecXne"

    def test_an_insertion_at_the_end(self, group):
        body = member(group, "Body")
        gt.replace_span(body, 23, 0, "!")
        assert gt.read_text(body) == "First line\nSecond line!"

    def test_an_insertion_at_a_run_boundary_goes_on_in_the_run_before(self, group):
        body = member(group, "Body")
        gt.replace_span(body, 18, 0, "s")
        assert gt.read_text(body) == "First line\nSeconds line"
        assert gt.run_formats(body)[1]["b"] == "1"

    @pytest.mark.parametrize("start,length,text", [(9, 4, "x"), (1, 2, "a\nb"), (1, 1, "a\vb")])
    def test_across_a_paragraph_or_with_a_break_is_refused(self, group, start, length, text):
        body = member(group, "Body")
        before = ET.tostring(body)
        with pytest.raises(ValueError, match="Pass the whole text"):
            gt.replace_span(body, start, length, text)
        assert ET.tostring(body) == before


class TestFormatting:
    def test_every_run_and_the_paragraph_ends(self, group):
        body = member(group, "Body")
        n = gt.format_runs(body, font_size=20, bold=False, italic=True, underline=True,
                           color="accent2", font_name="Georgia")
        assert n == 3
        assert gt.run_formats(body) == [{"sz": "2000", "b": "0", "i": "1", "u": "sng",
                                         "color": "accent2", "latin": "Georgia"}] * 3
        end = body.find(f"{A}txSp/{A}txBody/{A}p/{A}endParaRPr")
        assert end is None or end.get("sz") == "2000"

    def test_a_hex_colour_replaces_the_old_fill(self, group):
        body = member(group, "Body")
        gt.format_runs(body, color="#1f2937")
        rpr = body.find(f"{A}txSp/{A}txBody/{A}p/{A}r/{A}rPr")
        assert [c.tag[len(A):] for c in rpr] == ["solidFill"]
        assert rpr.find(f"{A}solidFill/{A}srgbClr").get("val") == "1F2937"

    def test_children_go_where_the_schema_has_them(self, group):
        body = member(group, "Body")
        gt.format_runs(body, font_name="Georgia", color="dark1")
        rpr = body.find(f"{A}txSp/{A}txBody/{A}p/{A}r/{A}rPr")
        assert [c.tag[len(A):] for c in rpr] == ["solidFill", "latin", "ea"]
        assert rpr.find(f"{A}solidFill/{A}schemeClr").get("val") == "dk1"

    def test_a_run_with_no_properties_gets_some(self, group):
        title = member(group, "Group 4/Inner/Title")
        gt.format_runs(title, bold=True)
        assert gt.run_formats(title)[0]["b"] == "1"


# ---------------------------------------------------------------------------
# The Mac tools handing a member over
# ---------------------------------------------------------------------------
@pytest.fixture
def fake_edit(monkeypatch):
    """Run ppt_mac.text's edit and check on a member of GROUP, as edit_member
    would around a real copy and paste."""
    pytest.importorskip("appscript")
    import ppt_mac.group_text as group_text
    import ppt_mac.text as mac_text

    state = {"group": ET.fromstring(GROUP), "calls": []}

    def edit_member(slide_index, identifier, tool_name, edit, check):
        state["calls"].append((tool_name, identifier))
        try:
            path, found = gt.find_member(state["group"], "Group 4", identifier)
        except gt.MemberNotFound:
            return None
        fields = edit(found)
        back = copy.deepcopy(found)  # what the pasted group reads back as
        state["problem"] = check(back)
        return dict(fields, path=path)

    monkeypatch.setattr(group_text, "edit_member", edit_member)
    monkeypatch.setattr(group_text, "not_found",
                        lambda slide_index, identifier: ValueError(f"no '{identifier}'"))
    monkeypatch.setattr(mac_text, "_in_a_group", lambda slide_index, name: True)
    state["member"] = lambda name: gt.find_member(state["group"], "Group 4", name)[1]
    return state


@macos_only
class TestTheMacTools:
    def test_set_text_writes_the_whole_text(self, fake_edit):
        from ppt_mac.text import _set_text_impl

        result = _set_text_impl(1, "Body", "New body")
        assert result["path"] == "Group 4/Body" and result["paragraph_count"] == 1
        assert gt.read_text(fake_edit["member"]("Body")) == "New body"
        assert fake_edit["problem"] is None

    def test_set_text_replaces_what_search_text_finds(self, fake_edit):
        from ppt_mac.text import _set_text_impl

        result = _set_text_impl(1, "Body", "2nd", search_text="Second")
        assert (result["start"], result["replaced_length"], result["text"]) == (
            12, 6, "First line\n2nd line")

    def test_runs_are_refused_before_the_group_is_touched(self, fake_edit):
        from ppt_mac.text import _set_text_impl

        with pytest.raises(ValueError, match="runs cannot be written"):
            _set_text_impl(1, "Body", None, runs=[{"text": "a"}])
        assert fake_edit["calls"] == []

    def test_a_name_in_no_group_falls_to_not_found(self, fake_edit):
        from ppt_mac.text import _set_text_impl

        with pytest.raises(ValueError, match="no 'Nope'"):
            _set_text_impl(1, "Nope", "x")

    def test_format_text_sets_every_run_and_is_checked(self, fake_edit):
        from ppt_mac.text import _format_text_impl

        result = _format_text_impl(1, "Body", "Georgia", None, 20, True, None, None,
                                   "Accent 2", None, None)
        assert result["runs_formatted"] == 3 and fake_edit["problem"] is None
        assert {r["color"] for r in gt.run_formats(fake_edit["member"]("Body"))} == {"accent2"}

    def test_a_theme_colour_by_its_own_argument(self, fake_edit):
        from ppt_mac.text import _format_text_impl

        _format_text_impl(1, "Note", None, None, None, None, None, None, None, "dark2", None)
        assert {r["color"] for r in gt.run_formats(fake_edit["member"]("Note"))} == {"dk2"}

    def test_a_highlight_is_refused_before_the_group_is_touched(self, fake_edit):
        from ppt_mac.text import _format_text_impl

        with pytest.raises(ValueError, match="highlight cannot be written"):
            _format_text_impl(1, "Body", None, None, None, None, None, None, None, None,
                              "#FFFF00")
        assert fake_edit["calls"] == []
