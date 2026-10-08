"""Tests for opening a presentation by its SharePoint or OneDrive URL.

PowerPoint opens a document library file straight from its URL
(Presentations.Open("https://contoso.sharepoint.com/.../Deck.pptx")), but
ppt_open_presentation checked the path with os.path.exists first, which is
False for every URL, and answered "File not found" without asking PowerPoint.

PowerPoint then reports the deck's FullName percent-decoded (spaces, not %20),
so the URL copied from the browser must still name that deck in the
`presentation` argument and in ppt_activate_presentation.

Covered here without COM and without PowerPoint.
"""

from __future__ import annotations

import sys
from unittest.mock import MagicMock

import pytest

sys.path.insert(0, "src")

from utils.com_wrapper import (  # noqa: E402
    PowerPointCOMWrapper,
    full_name_key,
    pick_presentation,
)

_ENCODED = "https://contoso.sharepoint.com/sites/Team/Shared%20Documents/Q3%20Review.pptx"
_DECODED = "https://contoso.sharepoint.com/sites/Team/Shared Documents/Q3 Review.pptx"


# ---------------------------------------------------------------------------
# full_name_key / pick_presentation
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("name, key", [
    (_ENCODED, _DECODED.lower()),
    ("  " + _DECODED + "  ", _DECODED.lower()),
    ("C:\\Decks\\100%20Done.pptx", "c:\\decks\\100%20done.pptx"),
])
def test_full_name_key_decodes_urls_only(name, key):
    assert full_name_key(name) == key


def test_pick_matches_an_encoded_url_to_the_decoded_full_name():
    candidates = [(_DECODED, "Q3 Review.pptx", "q3"), ("C:\\Decks\\Other.pptx", "Other.pptx", "o")]
    assert pick_presentation(candidates, _ENCODED) == "q3"
    assert pick_presentation(candidates, _DECODED) == "q3"


# ---------------------------------------------------------------------------
# ppt_activate_presentation
# ---------------------------------------------------------------------------
def _make_pres(full_name, name):
    pres = MagicMock()
    pres.FullName = full_name
    pres.Name = name
    pres.Windows.Count = 1
    return pres


def test_activate_by_encoded_url():
    deck = _make_pres(_DECODED, "Q3 Review.pptx")
    other = _make_pres("C:\\Decks\\Other.pptx", "Other.pptx")
    app = MagicMock()
    app.Presentations.Count = 2
    app.Presentations.side_effect = lambda i: [other, deck][i - 1]
    w = PowerPointCOMWrapper()
    w._app = app
    w._get_app_impl = lambda allow_launch=False: app

    result = w._set_target_pres_impl(_ENCODED)

    assert w._target_pres_full_name == _DECODED
    assert result["full_name"] == _DECODED


# ---------------------------------------------------------------------------
# ppt_open_presentation (Windows impl)
# ---------------------------------------------------------------------------
windows_only = pytest.mark.skipif(
    sys.platform != "win32",
    reason="on macOS the Apple Event impl takes the place of the COM one",
)


@pytest.fixture
def open_impl(monkeypatch):
    from ppt_com import presentation

    deck = _make_pres(_DECODED, "Q3 Review.pptx")
    deck.Slides.Count = 12
    deck.ReadOnly = -1
    app = MagicMock()
    app.Visible = True
    app.Presentations.Open.return_value = deck
    app.Presentations.Count = 1
    app.Presentations.side_effect = lambda i: deck
    fake_ppt = MagicMock()
    fake_ppt._get_app_impl.return_value = app
    monkeypatch.setattr(presentation, "ppt", fake_ppt)
    return presentation._open_presentation_impl, app


@windows_only
def test_open_passes_a_url_to_powerpoint(open_impl):
    impl, app = open_impl

    result = impl(_ENCODED, True, False, False)

    assert app.Presentations.Open.call_args.kwargs["FileName"] == _ENCODED
    assert result["full_name"] == _DECODED
    assert result["slides_count"] == 12


@windows_only
def test_open_still_refuses_a_missing_local_file(open_impl, tmp_path):
    impl, app = open_impl
    missing = str(tmp_path / "missing.pptx")

    with pytest.raises(FileNotFoundError, match="File not found"):
        impl(missing, False, True, True)
    app.Presentations.Open.assert_not_called()


def _open_tool_with_failing_open(monkeypatch, error):
    import json
    from ppt_com import presentation

    fake_ppt = MagicMock()
    fake_ppt.execute.side_effect = error
    monkeypatch.setattr(presentation, "ppt", fake_ppt)
    return lambda path: json.loads(presentation.open_presentation(
        presentation.OpenPresentationInput(file_path=path)
    ))["error"]


@windows_only
def test_open_explains_a_url_powerpoint_cannot_open(monkeypatch):
    import pywintypes

    e_fail = pywintypes.com_error(-2147352567, "Exception occurred.", None, None)
    open_tool = _open_tool_with_failing_open(monkeypatch, e_fail)

    error = open_tool(_ENCODED)

    assert error.startswith(f"PowerPoint could not open {_ENCODED}.")
    assert "signed in" in error
    assert "-2147352567" in error


def test_open_reports_a_non_com_error_on_a_url_unchanged(monkeypatch):
    open_tool = _open_tool_with_failing_open(monkeypatch, TimeoutError("timed out"))

    assert open_tool(_ENCODED) == "timed out"


def test_open_reports_a_local_error_unchanged(monkeypatch):
    open_tool = _open_tool_with_failing_open(monkeypatch, RuntimeError("E_FAIL"))

    assert open_tool(r"C:\Decks\Broken.pptx") == "E_FAIL"


@windows_only
def test_open_strips_whitespace_around_a_url(open_impl):
    impl, app = open_impl

    impl("  " + _ENCODED + "\n", True, False, False)

    assert app.Presentations.Open.call_args.kwargs["FileName"] == _ENCODED
