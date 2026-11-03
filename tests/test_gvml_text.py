"""gvml.text: a canvas shape's text read and rewritten, for group members (#220)."""

import sys
import xml.etree.ElementTree as ET

sys.path.insert(0, "src")

from gvml import text  # noqa: E402
from gvml.package import NS_A  # noqa: E402

A = f"{{{NS_A}}}"

GROUP = f"""<a:grpSp xmlns:a="{NS_A}">
<a:nvGrpSpPr><a:cNvPr id="1" name="Card"/><a:cNvGrpSpPr/></a:nvGrpSpPr>
<a:grpSpPr/>
<a:sp><a:nvSpPr><a:cNvPr id="2" name="Title"/><a:cNvSpPr/></a:nvSpPr><a:spPr/>
<a:txSp><a:txBody><a:bodyPr/><a:lstStyle/>
<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="ja-JP" sz="1800" b="1"/><a:t>Old</a:t></a:r></a:p>
<a:p><a:r><a:rPr lang="ja-JP" sz="1400"/><a:t>second</a:t></a:r></a:p>
</a:txBody></a:txSp></a:sp>
<a:grpSp><a:nvGrpSpPr><a:cNvPr id="3" name="Inner"/><a:cNvGrpSpPr/></a:nvGrpSpPr><a:grpSpPr/>
<a:sp><a:nvSpPr><a:cNvPr id="4" name="Note"/><a:cNvSpPr/></a:nvSpPr><a:spPr/>
<a:txSp><a:txBody><a:bodyPr/><a:p><a:endParaRPr lang="en-US" sz="1200"/></a:p></a:txBody></a:txSp></a:sp>
<a:pic><a:nvPicPr><a:cNvPr id="5" name="Logo"/><a:cNvPicPr/></a:nvPicPr></a:pic>
</a:grpSp>
</a:grpSp>"""


def group():
    return ET.fromstring(GROUP)


def test_a_member_is_found_by_name_at_any_depth_or_by_its_path():
    g = group()
    assert text.find_member(g, ["Title"]) is not None
    assert text.find_member(g, ["Note"]) is not None
    assert text.find_member(g, ["Inner", "Note"]) is not None
    assert text.find_member(g, ["Title", "Note"]) is None
    assert text.find_member(g, ["Missing"]) is None


def test_text_reads_paragraphs_and_breaks():
    assert text.text_of(text.find_member(group(), ["Title"])) == "Old\nsecond"


def test_new_text_keeps_the_first_paragraph_and_run_look():
    member = text.find_member(group(), ["Title"])
    text.set_text(member, "New\nline\vbreak")
    assert text.text_of(member) == "New\nline\vbreak"
    paragraphs = member.findall(f"{A}txSp/{A}txBody/{A}p")
    assert len(paragraphs) == 2
    for p in paragraphs:
        assert p.find(f"{A}pPr").get("algn") == "ctr"
        for r in p.findall(f"{A}r"):
            assert r.find(f"{A}rPr").get("sz") == "1800" and r.find(f"{A}rPr").get("b") == "1"
    assert paragraphs[1].find(f"{A}br") is not None


def test_an_empty_member_takes_its_look_from_the_end_of_paragraph():
    member = text.find_member(group(), ["Note"])
    text.set_text(member, "Hello")
    run = member.find(f"{A}txSp/{A}txBody/{A}p/{A}r")
    assert run.find(f"{A}rPr").get("sz") == "1200"
    assert text.text_of(member) == "Hello"


def test_a_picture_has_no_text_body():
    assert not text.has_text_body(text.find_member(group(), ["Logo"]))
