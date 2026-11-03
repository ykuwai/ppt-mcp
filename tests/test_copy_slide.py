"""Tests for ppt_copy_slide (_copy_slide_impl) on Windows.

Pure Python tests -- fake COM presentations stand in for PowerPoint.

The slides come from a snapshot of the source (SaveCopyAs) via
Slides.InsertFromFile. Only a source whose snapshot is not a zip
(password-protected) goes through the clipboard, which right after Copy()
can still hold the previous slide; there every pasted slide is checked
against its source.
"""

import sys
import types
import zipfile
from unittest.mock import patch

import pytest

sys.path.insert(0, "src")

from ppt_com import slides  # noqa: E402

# On macOS the impl is replaced by its Apple Event counterpart, which refuses
# copies between presentations (see ppt_mac/slides.py).
pytestmark = pytest.mark.skipif(
    sys.platform != "win32",
    reason="drives the COM Slides collection directly",
)


# --- Fake COM object graph -------------------------------------------------

class _TextRange:
    def __init__(self, text):
        self.Text = text


class _TextFrame:
    def __init__(self, text):
        self.HasText = bool(text)
        self.TextRange = _TextRange(text)


class _Shape:
    Type = 17  # msoTextBox

    def __init__(self, text):
        self.HasTextFrame = True
        self.TextFrame = _TextFrame(text)


class _SlideNumber:
    """Slide number placeholder: its text is the slide's current position."""

    Type = 14  # msoPlaceholder
    HasTextFrame = True

    def __init__(self, slide):
        self._slide = slide
        self.PlaceholderFormat = types.SimpleNamespace(Type=13)

    @property
    def TextFrame(self):
        return _TextFrame(str(self._slide.SlideIndex))


# A slide's content is a tuple of texts; NUMBER stands for a slide number
# placeholder.
NUMBER = object()


class _Shapes:
    def __init__(self, slide, texts):
        self._shapes = [_SlideNumber(slide) if t is NUMBER else _Shape(t) for t in texts]
        self.HasTitle = bool(texts)

    @property
    def Count(self):
        return len(self._shapes)

    @property
    def Title(self):
        return self._shapes[0]

    def __iter__(self):
        return iter(self._shapes)


class _Slide:
    _next_id = 256

    def __init__(self, owner, texts):
        _Slide._next_id += 1
        self.SlideID = _Slide._next_id
        self.texts = tuple(texts)
        self.Shapes = _Shapes(self, texts)
        self._owner = owner

    @property
    def SlideIndex(self):
        return self._owner._slides.index(self) + 1

    def Copy(self):
        self._owner.app.copy(self.texts)

    def Delete(self):
        self._owner._slides.remove(self)


class _SlideRange:
    def __init__(self, items):
        self._items = items
        self.Count = len(items)

    def __call__(self, i):
        return self._items[i - 1]

    def Delete(self):
        for s in self._items:
            s.Delete()


class _Slides:
    def __init__(self, pres):
        self._pres = pres

    @property
    def Count(self):
        return len(self._pres._slides)

    def __call__(self, i):
        return self._pres._slides[i - 1]

    def __iter__(self):
        return iter(list(self._pres._slides))

    def FindBySlideID(self, sid):
        return next(s for s in self._pres._slides if s.SlideID == sid)

    def InsertFromFile(self, path, index, start, end):
        self._pres.insert_calls.append((index, start, end))
        saved = self._pres.app.files[path]
        new = [_Slide(self._pres, t) for t in saved[start - 1:end]]
        self._pres._slides[index:index] = new
        return len(new)

    def Paste(self, index=None):
        texts = self._pres.app.paste()
        slide = _Slide(self._pres, texts)
        pos = len(self._pres._slides) if index is None else index - 1
        self._pres._slides.insert(pos, slide)
        return _SlideRange([slide])


class _Pres:
    def __init__(self, app, name, titles, password="", numbered=False):
        self.app = app
        self.Name = name
        self.Password = password
        extra = (NUMBER,) if numbered else ()
        self._slides = [_Slide(self, (t, "body") + extra) for t in titles]
        self.Slides = _Slides(self)
        self.insert_calls = []
        self.save_fails = False

    def SaveCopyAs(self, path):
        if self.save_fails:
            raise Exception("Presentation.SaveCopyAs : rights-managed")
        self.app.files[path] = [s.texts for s in self._slides]
        if self.Password:
            # what PowerPoint writes for a password-protected deck
            with open(path, "wb") as f:
                f.write(bytes.fromhex("d0cf11e0a1b11ae1") + bytes(504))
        else:
            with zipfile.ZipFile(path, "w") as z:
                z.writestr("[Content_Types].xml", "")

    def titles(self):
        return [s.texts[0] for s in self._slides]


class _App:
    """Holds the files written by SaveCopyAs and a clipboard.

    With lag=n the clipboard only shows a copy n copies later, the way a
    stale clipboard looks from the paste side.
    """

    def __init__(self, lag=0):
        self.files = {}
        self.lag = lag
        self._copies = []
        self.pastes = 0

    def copy(self, texts):
        self._copies.append(texts)

    def paste(self):
        self.pastes += 1
        i = max(0, len(self._copies) - 1 - self.lag)
        return self._copies[i]


def _run(app, decks, sources, src="src.pptx", dst="dst.pptx", insert_at=None):
    def resolve(_app, presentation_index=None, presentation_name=None):
        return decks[presentation_name]

    single = sources[0] if len(sources) == 1 else None
    many = None if single is not None else sources
    with patch.object(slides.ppt, "_get_app_impl", return_value=app), \
            patch.object(slides, "_resolve_presentation", side_effect=resolve), \
            patch.object(slides, "nav_goto_slide"), \
            patch.object(slides.time, "sleep"):
        return slides._copy_slide_impl(single, many, None, src, None, dst, insert_at)


def _decks(app, password="", name="src.pptx", numbered=False):
    src = _Pres(app, name, [f"S{i}" for i in range(1, 7)], password, numbered)
    dst = _Pres(app, "dst.pptx", ["D1", "D2", "D3"])
    return src, dst, {name: src, "dst.pptx": dst}


# --- InsertFromFile ----------------------------------------------------------

def test_consecutive_slides_in_one_insert_from_file_call():
    app = _App(lag=1)  # the clipboard would get this wrong
    src, dst, decks = _decks(app)
    result = _run(app, decks, [2, 3, 4, 5, 6], insert_at=2)

    assert dst.titles() == ["D1", "S2", "S3", "S4", "S5", "S6", "D2", "D3"]
    assert dst.insert_calls == [(1, 2, 6)]
    assert app.pastes == 0
    assert result["method"] == "insert_from_file"
    assert result["new_slide_indices"] == [2, 3, 4, 5, 6]
    assert [(c["source_index"], c["new_index"], c["title"]) for c in result["copied"]] == [
        (2, 2, "S2"), (3, 3, "S3"), (4, 4, "S4"), (5, 5, "S5"), (6, 6, "S6")]


def test_given_order_and_repeats_are_kept():
    app = _App()
    src, dst, decks = _decks(app)
    result = _run(app, decks, [5, 2, 3, 3])

    assert dst.titles() == ["D1", "D2", "D3", "S5", "S2", "S3", "S3"]
    assert dst.insert_calls == [(3, 5, 5), (4, 2, 3), (6, 3, 3)]
    assert result["new_slide_indices"] == [4, 5, 6, 7]


def test_snapshot_includes_unsaved_changes():
    app = _App()
    src, dst, decks = _decks(app)
    src._slides[1] = _Slide(src, ("S2 edited", "body"))
    _run(app, decks, [2], insert_at=1)
    assert dst.titles()[0] == "S2 edited"


def test_copy_within_one_presentation():
    app = _App()
    src, _dst, decks = _decks(app)
    result = _run(app, decks, [2, 3], dst="src.pptx", insert_at=1)
    assert src.titles() == ["S2", "S3", "S1", "S2", "S3", "S4", "S5", "S6"]
    assert result["new_slide_indices"] == [1, 2]


def test_mismatching_insert_is_an_error_and_is_undone():
    app = _App()
    src, dst, decks = _decks(app)

    def wrong_slides(path, index, start, end):
        return _Slides.InsertFromFile(dst.Slides, path, index, start - 1, end - 1)

    with patch.object(dst.Slides, "InsertFromFile", side_effect=wrong_slides), \
            pytest.raises(RuntimeError, match="source slide 2 does not match"):
        _run(app, decks, [2, 3])
    assert dst.titles() == ["D1", "D2", "D3"]


def test_slide_numbers_do_not_count_as_content():
    app = _App()
    src, dst, decks = _decks(app, numbered=True)
    result = _run(app, decks, [2, 3], insert_at=1)
    assert dst.titles()[:2] == ["S2", "S3"]
    assert result["method"] == "insert_from_file"


def test_insert_that_raises_after_inserting_is_undone():
    app = _App()
    src, dst, decks = _decks(app)
    calls = []

    def insert_then_fail(path, index, start, end):
        calls.append(start)
        n = _Slides.InsertFromFile(dst.Slides, path, index, start, end)
        if len(calls) == 2:
            raise Exception("RPC_E_CALL_REJECTED")
        return n

    with patch.object(dst.Slides, "InsertFromFile", side_effect=insert_then_fail),             pytest.raises(Exception, match="RPC_E_CALL_REJECTED"):
        _run(app, decks, [2, 5])
    assert dst.titles() == ["D1", "D2", "D3"]


def test_failed_snapshot_falls_back_to_the_clipboard():
    app = _App()
    src, dst, decks = _decks(app)
    src.save_fails = True
    result = _run(app, decks, [4])
    assert result["method"] == "clipboard"
    assert dst.titles() == ["D1", "D2", "D3", "S4"]


# --- Clipboard (password-protected source) -----------------------------------

def test_password_protected_source_uses_the_clipboard():
    app = _App()
    src, dst, decks = _decks(app, password="secret")
    result = _run(app, decks, [2, 3, 4], insert_at=2)
    assert result["method"] == "clipboard"
    assert dst.insert_calls == []
    assert dst.titles() == ["D1", "S2", "S3", "S4", "D2", "D3"]
    assert result["new_slide_indices"] == [2, 3, 4]


def test_stale_clipboard_is_detected_and_copied_again():
    # lag=1: every first paste brings the previous slide
    app = _App(lag=1)
    src, dst, decks = _decks(app, password="secret")
    result = _run(app, decks, [2, 3, 4, 5, 6], insert_at=2)
    assert dst.titles() == ["D1", "S2", "S3", "S4", "S5", "S6", "D2", "D3"]
    assert [c["title"] for c in result["copied"]] == ["S2", "S3", "S4", "S5", "S6"]


def test_clipboard_that_never_catches_up_is_an_error_and_is_undone():
    app = _App(lag=10)
    src, dst, decks = _decks(app, password="secret")
    app._copies = [("Other", "body")] * 10
    with pytest.raises(RuntimeError, match="Source slide 2 could not be copied"):
        _run(app, decks, [2, 3])
    assert dst.titles() == ["D1", "D2", "D3"]


def test_legacy_source_goes_via_the_clipboard_without_a_snapshot():
    app = _App()
    src, dst, decks = _decks(app, name="src.ppt")
    result = _run(app, decks, [2], src="src.ppt")
    assert result["method"] == "clipboard"
    assert app.files == {}
    assert dst.titles() == ["D1", "D2", "D3", "S2"]


# --- helpers -------------------------------------------------------------------

def test_runs():
    assert slides._runs([2, 3, 4, 7, 5]) == [(2, 4), (7, 7), (5, 5)]
    assert slides._runs([3, 3]) == [(3, 3), (3, 3)]
    assert slides._runs([1]) == [(1, 1)]
