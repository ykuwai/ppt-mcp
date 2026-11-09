"""ppt_get_all_text from a copy of the deck: utils.deck_text reads a saved
.pptx, looking up what a slide leaves to its layout and master the way
PowerPoint does, and anything that stops the copy only means the deck is read
the old way. The files here are written by hand; no PowerPoint is involved.
"""

import json
import sys
import zipfile
from pathlib import Path
from unittest.mock import patch

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from ppt_com import text as text_mod  # noqa: E402
from utils import deck_text  # noqa: E402

XMLNS = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
         'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
         'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
         'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"')
RELS = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"'
EMU = 12700


def xfrm(left, top, width, height):
    return (f'<a:xfrm><a:off x="{left * EMU}" y="{top * EMU}"/>'
            f'<a:ext cx="{width * EMU}" cy="{height * EMU}"/></a:xfrm>')


def sp(body="", ph=None, box=None, name="S", lst=""):
    nv = f"<p:ph {ph}/>" if ph is not None else ""
    props = f"<p:spPr>{xfrm(*box) if box else ''}</p:spPr>"
    text = f"<p:txBody><a:bodyPr/><a:lstStyle>{lst}</a:lstStyle>{body}</p:txBody>" if body else ""
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="2" name="{name}"/><p:cNvSpPr/><p:nvPr>{nv}</p:nvPr></p:nvSpPr>'
            f"{props}{text}</p:sp>")


def para(*runs, ppr=""):
    return f"<a:p>{ppr}{''.join(runs)}</a:p>"


def run(text, attrs=""):
    return f"<a:r><a:rPr {attrs}/><a:t>{text}</a:t></a:r>"


def tree(*shapes):
    return f"<p:cSld><p:spTree>{''.join(shapes)}</p:spTree></p:cSld>"


MASTER_STYLES = (
    "<p:txStyles><p:titleStyle><a:lvl1pPr><a:buNone/><a:defRPr/></a:lvl1pPr></p:titleStyle>"
    '<p:bodyStyle><a:lvl1pPr><a:buChar char="x"/><a:defRPr/></a:lvl1pPr>'
    '<a:lvl2pPr><a:buChar char="x"/><a:defRPr/></a:lvl2pPr></p:bodyStyle>'
    "<p:otherStyle/></p:txStyles>")


def deck(tmp_path, *slides, layout=None, default_style=""):
    """A .pptx with these slide bodies, all on one layout and master."""
    layout = layout if layout is not None else (
        sp(ph='type="title"', box=(40, 20, 800, 60))
        + sp(ph='idx="1"', box=(40, 100, 400, 200)))
    master = sp(ph='type="title"', box=(1, 2, 3, 4)) + sp(ph='type="body" idx="1"', box=(5, 6, 7, 8))
    path = tmp_path / "copy.pptx"
    ids = "".join(f'<p:sldId id="{255 + n}" r:id="rId{n}"/>' for n in range(1, len(slides) + 1))
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("ppt/presentation.xml",
                   f"<p:presentation {XMLNS}><p:sldIdLst>{ids}</p:sldIdLst>"
                   f"<p:defaultTextStyle>{default_style}</p:defaultTextStyle></p:presentation>")
        z.writestr("ppt/_rels/presentation.xml.rels", f"<Relationships {RELS}>" + "".join(
            f'<Relationship Id="rId{n}" Target="slides/slide{n}.xml"/>' for n in range(1, len(slides) + 1))
            + "</Relationships>")
        for n, slide in enumerate(slides, start=1):
            attrs, body = slide if isinstance(slide, tuple) else ("", slide)
            z.writestr(f"ppt/slides/slide{n}.xml", f"<p:sld {XMLNS} {attrs}>{tree(body)}</p:sld>")
            z.writestr(f"ppt/slides/_rels/slide{n}.xml.rels",
                       f'<Relationships {RELS}><Relationship Id="rId1" '
                       'Target="../slideLayouts/slideLayout1.xml"/></Relationships>')
        z.writestr("ppt/slideLayouts/slideLayout1.xml", f"<p:sldLayout {XMLNS}>{tree(layout)}</p:sldLayout>")
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                   f'<Relationships {RELS}><Relationship Id="rId1" '
                   'Target="../slideMasters/slideMaster1.xml"/></Relationships>')
        z.writestr("ppt/slideMasters/slideMaster1.xml",
                   f"<p:sldMaster {XMLNS}>{tree(master)}{MASTER_STYLES}</p:sldMaster>")
    return str(path)


def shapes_of(path, index=1):
    return deck_text.read_slides(path)[1][index]["shapes"]


def render(tmp_path, *body, **kwargs):
    return [deck_text.shape_markdown(s) for s in shapes_of(deck(tmp_path, "".join(body), **kwargs))]


class TestWhatASlideInherits:
    def test_a_placeholder_with_no_box_takes_its_layouts(self, tmp_path):
        [title] = shapes_of(deck(tmp_path, sp(para(run("T")), ph='type="title"')))
        assert (title["left"], title["top"], title["width"], title["height"]) == (40, 20, 800, 60)
        assert title["is_title"]

    def test_a_placeholder_is_matched_by_idx(self, tmp_path):
        [body] = shapes_of(deck(tmp_path, sp(para(run("B")), ph='idx="1"')))
        assert (body["left"], body["top"]) == (40, 100)

    def test_its_own_box_wins(self, tmp_path):
        [body] = shapes_of(deck(tmp_path, sp(para(run("B")), ph='idx="1"', box=(7, 8, 9, 10))))
        assert (body["left"], body["top"]) == (7, 8)

    def test_a_bullet_comes_from_the_masters_body_style(self, tmp_path):
        assert render(tmp_path, sp(para(run("One")), ph='idx="1"')) == ["- One"]

    def test_a_level_is_indented(self, tmp_path):
        assert render(tmp_path, sp(para(run("One"), ppr='<a:pPr lvl="1"/>'), ph='idx="1"')) == ["  - One"]

    def test_a_paragraph_can_turn_its_bullet_off(self, tmp_path):
        assert render(tmp_path, sp(para(run("One"), ppr="<a:pPr><a:buNone/></a:pPr>"), ph='idx="1"')) == ["One"]

    def test_numbering_is_written_as_a_number(self, tmp_path):
        ppr = '<a:pPr><a:buAutoNum type="arabicPeriod"/></a:pPr>'
        assert render(tmp_path, sp(para(run("One"), ppr=ppr), ph='idx="1"')) == ["1. One"]

    def test_a_text_box_has_no_bullet_from_the_masters_body(self, tmp_path):
        assert render(tmp_path, sp(para(run("One")))) == ["One"]

    def test_bold_can_come_from_the_layout(self, tmp_path):
        layout = sp(ph='idx="1"', box=(0, 0, 9, 9), body=para(run("x")),
                    lst='<a:lvl1pPr><a:defRPr b="1"/></a:lvl1pPr>')
        assert render(tmp_path, sp(para(run("All bold")), ph='idx="1"'), layout=layout) == ["## All bold"]

    def test_a_run_can_turn_inherited_bold_off(self, tmp_path):
        layout = sp(ph='idx="1"', box=(0, 0, 9, 9), body=para(run("x")),
                    lst='<a:lvl1pPr><a:buNone/><a:defRPr b="1"/></a:lvl1pPr>')
        body = sp(para(run("Loud"), run(" quiet", 'b="0"')), ph='idx="1"')
        assert render(tmp_path, body, layout=layout) == ["**Loud** quiet"]


class TestTheText:
    def test_emphasis_inside_a_paragraph(self, tmp_path):
        body = sp(para(run("Costs "), run("fell", 'b="1"'), run(" "), run("sharply", 'i="1"')))
        assert render(tmp_path, body) == ["Costs **fell** *sharply*"]

    def test_a_shape_all_in_bold_is_a_subheading(self, tmp_path):
        assert render(tmp_path, sp(para(run("Head", 'b="1"')))) == ["## Head"]

    def test_a_title_is_plain_text_on_one_line(self, tmp_path):
        out = render(tmp_path, sp(para(run("One", 'b="1"'), "<a:br/>", run("two")), ph='type="title"'))
        assert out == ["# One two"]

    def test_a_subtitle_is_not_a_heading_even_in_bold(self, tmp_path):
        body = sp(para(run("Sub", 'b="1"'), ppr="<a:pPr><a:buNone/></a:pPr>"), ph='type="subTitle" idx="1"')
        assert render(tmp_path, body) == ["**Sub**"]

    def test_a_line_break_stays_inside_the_paragraph(self, tmp_path):
        assert render(tmp_path, sp(para(run("a"), "<a:br/>", run("b")))) == ["a\nb"]

    def test_an_empty_paragraph_is_a_blank_line_but_not_the_last(self, tmp_path):
        assert render(tmp_path, sp(para(run("a")) + para() + para(run("b")) + para())) == ["a\n\nb"]

    def test_a_field_keeps_its_text(self, tmp_path):
        field = '<a:fld id="{1}" type="datetime1"><a:rPr/><a:t>Today</a:t></a:fld>'
        assert render(tmp_path, sp(para(run("On "), field))) == ["On Today"]

    def test_a_shape_with_no_text_is_left_out(self, tmp_path):
        assert render(tmp_path, sp(para()), sp()) == []

    @pytest.mark.parametrize("kind", ["dt", "ftr", "sldNum"])
    def test_dates_footers_and_numbers_are_left_out(self, tmp_path, kind):
        assert render(tmp_path, sp(para(run("x")), ph=f'type="{kind}"')) == []

    def test_alternate_content_is_read_as_its_first_choice(self, tmp_path):
        chosen = (f'<mc:AlternateContent><mc:Choice Requires="a14">{sp(para(run("Live")))}</mc:Choice>'
                  f"<mc:Fallback>{sp(para(run('Picture')))}</mc:Fallback></mc:AlternateContent>")
        assert render(tmp_path, chosen) == ["Live"]


class TestEquations:
    # As PowerPoint saves one: the whole paragraph is inside AlternateContent,
    # and the picture in the fallback carries placeholder text.
    MATH = ('<a14:m xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">'
            '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            "<m:r><m:t>E=m</m:t></m:r><m:r><m:t>c2</m:t></m:r></m:oMath></a14:m>")

    def wrapped(self, inner):
        return (f'<mc:AlternateContent><mc:Choice Requires="a14">{inner}</mc:Choice>'
                f"<mc:Fallback>{para(run('picture of it'))}</mc:Fallback></mc:AlternateContent>")

    def test_a_paragraph_wrapped_whole_keeps_its_place(self, tmp_path):
        body = para(run("Before")) + self.wrapped(para(self.MATH)) + para(run("After"))
        assert render(tmp_path, sp(body)) == ["Before\nE=mc2\nAfter"]

    def test_an_equation_beside_words(self, tmp_path):
        assert render(tmp_path, sp(self.wrapped(para(run("So "), self.MATH)))) == ["So E=mc2"]

    def test_a_shape_holding_only_an_equation_is_read(self, tmp_path):
        assert render(tmp_path, sp(self.wrapped(para(self.MATH)))) == ["E=mc2"]

    def test_an_equation_in_a_table_cell(self, tmp_path):
        cell = f"<a:tc><a:txBody><a:bodyPr/>{self.wrapped(para(self.MATH))}</a:txBody></a:tc>"
        assert render(tmp_path, TestTables().table([cell])) == ["| E=mc2 |\n| --- |"]


class TestGroups:
    def group(self, *members, off=(100, 50), ext=(200, 100), ch_off=(0, 0), ch_ext=(100, 50)):
        x = (f'<a:xfrm><a:off x="{off[0] * EMU}" y="{off[1] * EMU}"/><a:ext cx="{ext[0] * EMU}" cy="{ext[1] * EMU}"/>'
             f'<a:chOff x="{ch_off[0] * EMU}" y="{ch_off[1] * EMU}"/>'
             f'<a:chExt cx="{ch_ext[0] * EMU}" cy="{ch_ext[1] * EMU}"/></a:xfrm>')
        return (f'<p:grpSp><p:nvGrpSpPr><p:cNvPr id="4" name="G"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
                f"<p:grpSpPr>{x}</p:grpSpPr>{''.join(members)}</p:grpSp>")

    def test_a_member_is_placed_through_the_groups_scale(self, tmp_path):
        [member] = shapes_of(deck(tmp_path, self.group(sp(para(run("M")), box=(10, 10, 20, 20)))))
        assert (member["left"], member["top"], member["width"], member["height"]) == (120, 70, 40, 40)

    def test_a_group_inside_a_group(self, tmp_path):
        inner = self.group(sp(para(run("M")), box=(0, 0, 10, 10)), off=(10, 10), ext=(10, 10), ch_ext=(10, 10))
        [member] = shapes_of(deck(tmp_path, self.group(inner)))
        assert (member["left"], member["top"]) == (120, 70)


class TestTables:
    def table(self, *rows):
        body = "".join("<a:tr>" + "".join(rows_cell) + "</a:tr>" for rows_cell in rows)
        return ('<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="T"/><p:cNvGraphicFramePr/><p:nvPr/>'
                f'</p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="{EMU}" cy="{EMU}"/></p:xfrm>'
                f"<a:graphic><a:graphicData><a:tbl>{body}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>")

    @staticmethod
    def cell(text="", attrs=""):
        return f'<a:tc {attrs}><a:txBody><a:bodyPr/>{para(run(text)) if text else "<a:p/>"}</a:txBody></a:tc>'

    def test_a_table_is_a_markdown_table(self, tmp_path):
        out = render(tmp_path, self.table([self.cell("a|b"), self.cell("c")], [self.cell("1"), self.cell("2")]))
        assert out == ["| a\\|b | c |\n| --- | --- |\n| 1 | 2 |"]

    def test_a_merged_cell_answers_at_every_place_it_covers(self, tmp_path):
        out = render(tmp_path, self.table(
            [self.cell("wide", 'gridSpan="2"'), self.cell("", 'hMerge="1"')],
            [self.cell("tall", 'rowSpan="2"'), self.cell("x")],
            [self.cell("", 'vMerge="1"'), self.cell("y")]))
        assert out == ["| wide | wide |\n| --- | --- |\n| tall | x |\n| tall | y |"]


class TestTheDeck:
    def test_slides_come_in_the_decks_order_and_say_when_hidden(self, tmp_path):
        path = deck(tmp_path, sp(para(run("one"))), ('show="0"', sp(para(run("two")))))
        total, slides = deck_text.read_slides(path, [2])
        assert total == 2 and list(slides) == [2] and slides[2]["hidden"]

    def test_markdown_of_the_whole_copy(self, tmp_path):
        path = deck(tmp_path, sp(para(run("T")), ph='type="title"'), ('show="0"', ""))
        assert text_mod._markdown_of_copy(path, None, 2) == "== Slide 1 ==\n# T\n\n== Slide 2 (hidden) ==\n(no text)"

    def test_a_slide_the_deck_does_not_have_says_so(self, tmp_path):
        path = deck(tmp_path, sp(para(run("x"))))
        assert text_mod._markdown_of_copy(path, [3], 1) == "== Slide 3 ==\n(invalid slide index, presentation has 1 slides)"

    def test_a_copy_with_another_number_of_slides_is_not_the_deck(self, tmp_path):
        assert text_mod._markdown_of_copy(deck(tmp_path, sp(para(run("x")))), None, 2) is None


class TestAnythingThatStopsTheCopyIsNotAFailure:
    """The deck is then read the old way, and the tool still answers."""

    def run_tool(self, execute):
        with patch.object(text_mod.ppt, "execute", side_effect=execute):
            return text_mod.get_all_text(text_mod.GetAllTextInput())

    @staticmethod
    def old_way(func, *args):
        if func is text_mod._slide_count_impl:
            return 1
        if func is text_mod._get_all_text_impl:
            return "the old way"
        raise AssertionError(func)

    def test_when_there_is_no_copy(self):
        def execute(func, *args):
            return None if func is text_mod._copy_deck_impl else self.old_way(func, *args)
        assert self.run_tool(execute) == "the old way"

    def test_when_the_copy_fails(self):
        def execute(func, *args):
            if func is text_mod._copy_deck_impl:
                raise RuntimeError("protected")
            return self.old_way(func, *args)
        assert self.run_tool(execute) == "the old way"

    def test_when_the_copy_cannot_be_read_and_its_folder_still_goes(self, tmp_path):
        folder = tmp_path / "held"
        folder.mkdir()
        (folder / "copy.pptx").write_bytes(b"not a zip")

        def execute(func, *args):
            if func is text_mod._copy_deck_impl:
                return {"path": str(folder / "copy.pptx"), "folder": str(folder), "slides": 1}
            return self.old_way(func, *args)
        assert self.run_tool(execute) == "the old way"
        assert not folder.exists()

    def test_a_copy_that_reads_is_used_and_deleted(self, tmp_path):
        folder = tmp_path / "held"
        folder.mkdir()
        path = deck(folder, sp(para(run("from the copy"))))

        def execute(func, *args):
            assert func is text_mod._copy_deck_impl
            return {"path": path, "folder": str(folder), "slides": 1}
        assert self.run_tool(execute) == "== Slide 1 ==\nfrom the copy"
        assert not folder.exists()

    def test_the_output_file_says_how_many_slides(self, tmp_path):
        folder = tmp_path / "held"
        folder.mkdir()
        path = deck(folder, sp(para(run("a"))), sp(para(run("b"))))
        out = tmp_path / "out.md"

        def execute(func, *args):
            return {"path": path, "folder": str(folder), "slides": 2}
        with patch.object(text_mod.ppt, "execute", side_effect=execute):
            answer = json.loads(text_mod.get_all_text(text_mod.GetAllTextInput(output_path=str(out))))
        assert answer["slide_count"] == 2 and "== Slide 2 ==" in out.read_text(encoding="utf-8")


windows_only = pytest.mark.skipif(
    sys.platform != "win32", reason="on macOS the Apple Event impl takes the place of the COM one")


@windows_only
class TestGroupMembersReadOneByOne:
    class Items:
        """Shapes and GroupItems both count, and are called with an index."""

        def __init__(self, *members):
            self.members, self.Count = members, len(members)

        def __call__(self, index):
            return self.members[index - 1]

    def test_a_members_position_is_not_counted_twice(self):
        frame = type("F", (), {"HasText": True})()
        member = type("M", (), {"Type": 17, "Top": 380, "Left": 30, "Width": 9, "Height": 9,
                                "HasTable": False, "HasTextFrame": True, "TextFrame": frame})()
        group = type("G", (), {"Type": text_mod.msoGroup, "Top": 380, "Left": 30,
                               "GroupItems": self.Items(member)})()
        slide = type("Sl", (), {"Shapes": self.Items(group)})()
        [info] = text_mod._collect_text_shapes(slide)
        assert (info["top"], info["left"]) == (380, 30)


@windows_only
class TestWhichDecksAreCopied:
    class Pres:
        def __init__(self, full_name, raises=None):
            self.FullName, self.raises, self.saved = full_name, raises, None
            self.Slides = type("S", (), {"Count": 3})()

        def SaveCopyAs(self, path):
            if self.raises:
                raise self.raises
            self.saved = path

    def copy(self, pres):
        with patch.object(text_mod.ppt, "_get_pres_impl", return_value=pres):
            return text_mod._copy_deck_impl()

    @pytest.mark.parametrize("name", [r"C:\decks\macros.pptm", r"C:\decks\old.ppt", r"C:\decks\show.ppsx"])
    def test_a_deck_in_another_format_is_not(self, name):
        pres = self.Pres(name)
        assert self.copy(pres) is None and pres.saved is None

    def test_a_very_large_deck_is_not(self, monkeypatch):
        monkeypatch.setattr(text_mod.os.path, "getsize", lambda p: text_mod._COPY_SIZE_LIMIT + 1)
        assert self.copy(self.Pres(r"C:\decks\film.pptx")) is None

    def test_a_new_deck_is(self):
        pres = self.Pres("Presentation1")
        answer = self.copy(pres)
        assert answer["slides"] == 3 and answer["path"] == pres.saved
        text_mod.shutil.rmtree(answer["folder"])

    def test_a_failed_copy_leaves_no_folder(self, monkeypatch, tmp_path):
        monkeypatch.setattr(text_mod.tempfile, "tempdir", str(tmp_path))
        with pytest.raises(RuntimeError):
            self.copy(self.Pres("Presentation1", raises=RuntimeError("protected")))
        assert list(tmp_path.iterdir()) == []


@pytest.mark.skipif(sys.platform != "darwin", reason="the Apple Event backend")
def test_on_the_mac_without_the_add_in_there_is_no_copy():
    from ppt_mac import text as mac_text

    assert mac_text._copy_deck_impl() is None


def test_the_copy_goes_through_the_add_in_on_the_mac():
    from backend.vba import routes

    assert ("ppt_com.text", "_copy_deck_impl") in routes.ROUTES
