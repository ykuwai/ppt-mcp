"""Speaker notes keep their paragraphs both ways (#194).

Notes sent with \\n came back from ppt_get_slide_notes as one line. PowerPoint
keeps a paragraph break as \\r and a line break inside one as \\v, so the
notes tools now convert as ppt_set_text does: \\n to \\r on the way in and
back on the way out, \\v left as it is. The conversion is in the shared tool
functions, so these run against a stand-in for either platform's impl.
"""

import json
import sys
from unittest.mock import patch

sys.path.insert(0, "src")

from ppt_com import slides  # noqa: E402


class _Notes:
    """One slide's notes as PowerPoint holds them: \\r between paragraphs."""

    def __init__(self, stored=""):
        self.stored = stored
        self.written = []

    def execute(self, fn, *args):
        name = fn.__name__
        if name == "_set_slide_notes_impl":
            if args[1] is not None:
                self.written.append(args[1])
                self.stored = args[1]
            return {"success": True}
        if name == "_get_slide_notes_impl":
            return {"slide_index": args[0], "notes_text": self.stored}
        if name == "_get_slide_info_impl":
            return {"slide_index": args[0], "notes_text": self.stored}
        raise AssertionError(name)


def _set(notes, **kwargs):
    with patch.object(slides.ppt, "execute", notes.execute):
        return json.loads(slides.set_slide_notes(
            slides.SetSlideNotesInput(slide_index=1, **kwargs)))


def _get(notes, tool=slides.get_slide_notes, model=slides.GetSlideNotesInput):
    with patch.object(slides.ppt, "execute", notes.execute):
        return json.loads(tool(model(slide_index=1)))["notes_text"]


def test_a_paragraph_break_goes_to_powerpoint_as_its_own():
    notes = _Notes()
    _set(notes, notes_text="Open with the bill\nThen the chart")
    assert notes.written == ["Open with the bill\rThen the chart"]


def test_a_line_break_inside_a_paragraph_goes_as_it_is():
    notes = _Notes()
    _set(notes, notes_text="Two lines\vone paragraph")
    assert notes.written == ["Two lines\vone paragraph"]


def test_notes_round_trip():
    notes = _Notes()
    sent = "Open with the bill\nThen the chart\vand its source\nClose"
    _set(notes, notes_text=sent)
    assert _get(notes) == sent


def test_notes_written_in_powerpoint_read_with_new_lines():
    assert _get(_Notes("First\rSecond\r\nThird")) == "First\nSecond\nThird"


def test_slide_info_reads_the_notes_the_same_way():
    notes = _Notes("First\rSecond")
    assert _get(notes, slides.get_slide_info, slides.GetSlideInfoInput) == "First\nSecond"


def test_formatting_alone_writes_no_text():
    notes = _Notes("Kept")
    assert _set(notes, bold=True) == {"success": True}
    assert notes.written == [] and notes.stored == "Kept"
