"""ppt_find_replace_text on Windows with whole_words and a phrase.

PowerPoint's WholeWords finds nothing when the search text has a space in it
("acme corp": found 0, replaced 0, no error), measured on PowerPoint 16.0.
A phrase is found without it and kept only on word boundaries, as the Mac's
regex does. The TextRange here is a stand-in that searches like Find.
"""

import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from ppt_com import text as text_mod  # noqa: E402


def _units(text):
    return text.encode("utf-16-le")


def _from_units(data):
    return data.decode("utf-16-le")


class Found:
    """A range, counted in UTF-16 code units as PowerPoint counts."""

    def __init__(self, owner, start, length):
        self._owner, self.Start, self.Length = owner, start, length

    @property
    def Text(self):
        data = _units(self._owner.Text)
        return _from_units(data[2 * (self.Start - 1):2 * (self.Start - 1 + self.Length)])

    @Text.setter
    def Text(self, value):
        o, data = self._owner, _units(self._owner.Text)
        o.Text = _from_units(data[:2 * (self.Start - 1)]) + value + _from_units(
            data[2 * (self.Start - 1 + self.Length):])


class TextRange:
    """Find(what, after, match_case, whole_words) like PowerPoint's, for
    plain substring search, with Start and Length in UTF-16 code units; a
    whole-word search of a phrase finds nothing."""

    def __init__(self, text):
        self.Text, self.finds, self.replaces = text, [], 0

    def Find(self, what, after, match_case, whole_words):
        self.finds.append((after, whole_words))
        if whole_words == text_mod.msoTrue and " " in what:
            return None
        hay, needle = (self.Text, what) if match_case == text_mod.msoTrue else (self.Text.lower(), what.lower())
        hay_units, needle_units = _units(hay), _units(needle)
        at = 2 * after
        while True:
            at = hay_units.find(needle_units, at)
            if at < 0:
                return None
            if at % 2 == 0:
                return Found(self, at // 2 + 1, len(needle_units) // 2)
            at += 1

    def Characters(self, start, length):
        return Found(self, start, length)

    def Replace(self, *args):
        self.replaces += 1
        raise AssertionError("a phrase is not replaced through Replace")


# On macOS ppt_com.text's impl is swapped for the Mac one, which matches in
# Python already; the Windows path is the one under test.
windows_only = pytest.mark.skipif(sys.platform != "win32", reason="the COM impl")


def _pres(tr):
    shape = MagicMock(HasTextFrame=True, Name="TB")
    shape.TextFrame.TextRange = tr
    slide = MagicMock(SlideIndex=1)
    slide.Shapes.Count = 1
    slide.Shapes.return_value = shape
    pres = MagicMock()
    pres.Slides.Count = 1
    pres.Slides.return_value = slide
    return pres


def _run(tr, find, replace, match_case=False, dry_run=False):
    with patch.object(text_mod.ppt, "_get_pres_impl", return_value=_pres(tr)):
        return text_mod._find_replace_text_impl(
            find_text=find, replace_text=replace, dry_run=dry_run, match_case=match_case,
            whole_words=True, slide_indices=None, shape_name=None, context_chars=0,
            include_groups=False)


@windows_only
def test_a_phrase_is_replaced_where_it_stands_as_whole_words():
    tr = TextRange("ACME Corp and acme corp. Not acme corporation or xacme corp.")
    out = _run(tr, "acme corp", "Contoso Ltd")
    assert tr.Text == "Contoso Ltd and Contoso Ltd. Not acme corporation or xacme corp."
    assert out["match_count"] == 2
    assert [(m["start"], m["length"]) for m in out["matches"]] == [(1, 11), (17, 11)]
    assert {ww for _, ww in tr.finds} == {text_mod.msoFalse} and tr.replaces == 0


@windows_only
def test_a_phrase_is_found_without_being_changed():
    tr = TextRange("New York and new york, not new yorker.")
    out = _run(tr, "new york", None)
    assert out["mode"] == "find" and out["match_count"] == 2
    assert [m["start"] for m in out["matches"]] == [1, 14]
    assert tr.Text == "New York and new york, not new yorker."


@windows_only
def test_match_case_still_applies_to_a_phrase():
    tr = TextRange("ACME Corp and acme corp.")
    assert _run(tr, "acme corp", "Contoso", match_case=True)["match_count"] == 1
    assert tr.Text == "ACME Corp and Contoso."


@windows_only
def test_a_replacement_containing_the_phrase_does_not_loop():
    tr = TextRange("new york and new york")
    assert _run(tr, "new york", "new york city")["match_count"] == 2
    assert tr.Text == "new york city and new york city"


@windows_only
def test_a_single_word_still_goes_through_powerpoints_whole_words():
    tr = TextRange("acme")
    tr.Replace = MagicMock(return_value=None)
    _run(tr, "acme", "x")
    assert tr.Replace.call_args.args[4] == text_mod.msoTrue


@windows_only
def test_an_emoji_before_a_phrase_does_not_shift_it():
    # PowerPoint counts the rocket as two; Python sees one character.
    tr = TextRange("\U0001F680 acme corp and \U0001F600xacme corp")
    out = _run(tr, "acme corp", "Contoso")
    assert out["match_count"] == 1
    assert tr.Text == "\U0001F680 Contoso and \U0001F600xacme corp"


def test_a_replacement_powerpoint_did_not_take_is_an_error():
    tr = TextRange("acme corp")
    tr.Characters = lambda start, length: MagicMock(Text="acme corp")
    with pytest.raises(RuntimeError, match="did not take the replacement"):
        text_mod._replace_whole_phrase(tr, "acme corp", "Contoso", 0, text_mod.msoFalse)


def test_a_paragraph_break_in_the_replacement_is_read_back_as_powerpoint_holds_it():
    tr = TextRange("acme corp")
    replaced = text_mod._replace_whole_phrase(tr, "acme corp", "Contoso\nLtd", 0, text_mod.msoFalse)
    assert tr.Text == "Contoso\rLtd" and replaced.Length == 11


def test_word_boundaries_are_counted_as_powerpoint_counts():
    at = text_mod._whole_phrase_at
    # "\U0001F680 acme corp": the phrase starts at unit 4 for PowerPoint.
    assert at("\U0001F680 acme corp", 4, 9)
    # "\U0001F600xa b": "a b" at unit 4 runs into the x before it.
    assert not at("\U0001F600xa b", 4, 3)
    assert text_mod._utf16_length("\U0001F680ab") == 4


def test_word_boundaries():
    at = text_mod._whole_phrase_at
    assert at("acme corp", 1, 9)
    assert not at("xacme corp", 2, 9)
    assert not at("acme corps", 1, 9)
    assert at("(acme corp)", 2, 9)
    assert at("Q3-plan and q3 plan", 1, 2) and at("e-mail", 1, 6)
    assert at("x acme corp.", 3, 10)  # ends on a full stop, nothing after counts
