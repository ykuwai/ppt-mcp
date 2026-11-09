"""utils.deck_text.format_runs: bold and italic as Markdown that renders.

Emphasis that starts or ends with a space is not emphasis in CommonMark, so a
run's spaces at either end go outside its markers. Both platforms' shape
by shape reads and the read from a copy use it.
"""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from utils.deck_text import format_runs  # noqa: E402


def runs(*parts):
    return [{"text": t, "bold": "b" in f, "italic": "i" in f} for t, f in parts]


@pytest.mark.parametrize("parts, markdown", [
    ((("Costs ", ""), ("fell", "b"), (" sharply", "i")), "Costs **fell** *sharply*"),
    ((("Costs ", ""), ("fell ", "b"), ("sharply", "i")), "Costs **fell** *sharply*"),
    ((("A ", ""), (" both ", "bi"), ("end", "")), "A  ***both*** end"),
    ((("Plain", ""), ("   ", "b"), ("text", "")), "Plain   text"),
    ((("one ", "b"), ("two", "b")), "**one two**"),
    ((("Title", "b"),), "**Title**"),
    ((("no emphasis at all", ""),), "no emphasis at all"),
])
def test_spaces_stay_outside_the_markers(parts, markdown):
    assert format_runs(runs(*parts)) == markdown


def test_it_renders_as_meant():
    markdown_it = pytest.importorskip("markdown_it")
    html = markdown_it.MarkdownIt("commonmark").renderInline(
        format_runs(runs(("Costs ", ""), ("fell", "b"), (" sharply", "i"))))
    assert html == "Costs <strong>fell</strong> <em>sharply</em>"


def test_the_read_from_a_copy_writes_the_same(tmp_path):
    # A body read from the deck's file, as ppt_get_all_text reads a copy.
    import zipfile

    from utils import deck_text

    ns = ('xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
          'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
    slide = (f'<p:sld {ns}><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr txBox="1"/>'
             '<p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>'
             '<p:txBody><a:bodyPr/><a:p><a:r><a:t>Costs </a:t></a:r><a:r><a:rPr b="1"/><a:t>fell</a:t></a:r>'
             '<a:r><a:rPr i="1"/><a:t> sharply</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>')
    path = tmp_path / "deck.pptx"
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("ppt/presentation.xml", f'<p:presentation {ns}><p:sldIdLst><p:sldId id="256" r:id="rId2"/>'
                   '</p:sldIdLst></p:presentation>')
        z.writestr("ppt/_rels/presentation.xml.rels",
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId2" Target="slides/slide1.xml" Type="slide"/></Relationships>')
        z.writestr("ppt/slides/slide1.xml", slide)
        z.writestr("ppt/slides/_rels/slide1.xml.rels",
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
    _, slides = deck_text.read_slides(str(path))
    [shape] = slides[1]["shapes"]
    assert "Costs **fell** *sharply*" in deck_text.shape_markdown(shape)


def test_the_mac_reads_through_the_same_one():
    source = (Path(__file__).resolve().parent.parent / "src/ppt_mac/text.py").read_text(encoding="utf-8")
    body = source[source.index("def _format_runs("):]
    assert "format_runs(raw)" in body[:400]
