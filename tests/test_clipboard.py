"""Keeping the user's clipboard across a copy and paste (utils/clipboard.py)."""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from utils import clipboard  # noqa: E402
from utils.clipboard import (  # noqa: E402
    CF_BITMAP,
    CF_DIB,
    CF_DIBV5,
    CF_ENHMETAFILE,
    CF_METAFILEPICT,
    CF_OWNERDISPLAY,
    CF_PALETTE,
    SavedClipboard,
    plan,
)

CF_UNICODETEXT, CF_LOCALE, CF_TEXT, CF_OEMTEXT = 13, 16, 1, 7
HTML = 0xC0A0  # a registered format, as "HTML Format" is


def test_memory_formats_are_kept_in_their_order():
    formats = [CF_UNICODETEXT, CF_LOCALE, CF_TEXT, CF_OEMTEXT, HTML]
    assert plan(formats) == (formats, [])


def test_a_bitmap_is_left_to_windows_to_make_from_the_dib():
    assert plan([CF_BITMAP, CF_DIB, CF_DIBV5]) == ([CF_DIB, CF_DIBV5], [])


def test_a_palette_and_a_metafile_picture_are_made_again_too():
    keep, lost = plan([CF_ENHMETAFILE, CF_METAFILEPICT, CF_DIB, CF_PALETTE])
    assert keep == [CF_ENHMETAFILE, CF_DIB]
    assert lost == []


def test_handles_with_nothing_to_make_them_from_are_reported():
    keep, lost = plan([CF_BITMAP, CF_PALETTE, CF_METAFILEPICT, CF_OWNERDISPLAY, 0x0300])
    assert keep == []
    assert lost == [CF_BITMAP, CF_PALETTE, CF_METAFILEPICT, CF_OWNERDISPLAY, 0x0300]


def test_standard_formats_have_names():
    assert clipboard.format_name(CF_DIB) == "CF_DIB"
    assert clipboard.format_name(CF_OWNERDISPLAY) == "CF_OWNERDISPLAY"


def test_kept_clipboard_reports_what_could_not_be_put_back(monkeypatch):
    saved = SavedClipboard(entries=[(CF_UNICODETEXT, "global", b"a\0\0")], lost=["CF_OWNERDISPLAY"])
    put_back = []
    monkeypatch.setattr(clipboard, "save", lambda: saved)
    monkeypatch.setattr(clipboard, "restore", lambda s: put_back.append(s) or list(s.lost))
    with clipboard.kept_clipboard() as lost:
        assert lost == [], "nothing is known until the block is done"
    assert put_back == [saved]
    assert lost == ["CF_OWNERDISPLAY"]


def test_kept_clipboard_puts_back_even_when_the_block_fails(monkeypatch):
    put_back = []
    monkeypatch.setattr(clipboard, "save", lambda: SavedClipboard())
    monkeypatch.setattr(clipboard, "restore", lambda s: put_back.append(s) or [])
    with pytest.raises(RuntimeError):
        with clipboard.kept_clipboard():
            raise RuntimeError("paste failed")
    assert len(put_back) == 1


def test_a_failure_to_put_back_is_reported_not_raised(monkeypatch):
    def restore(saved):
        raise RuntimeError("The clipboard is in use by another program.")

    monkeypatch.setattr(clipboard, "save", lambda: SavedClipboard())
    monkeypatch.setattr(clipboard, "restore", restore)
    with clipboard.kept_clipboard() as lost:
        pass
    assert lost and "in use" in lost[0]


@pytest.mark.skipif(sys.platform != "win32", reason="uses the Windows clipboard")
def test_global_memory_round_trips():
    data = b"\x00\x01chart\xff" * 100
    handle = clipboard.global_block(data)
    try:
        assert clipboard.global_bytes(handle)[: len(data)] == data
    finally:
        clipboard.GlobalFree(handle)


@pytest.mark.skipif(sys.platform != "win32", reason="uses the Windows clipboard")
def test_the_clipboard_is_opened_with_a_window_of_its_own(monkeypatch):
    hwnd = clipboard.owner_window()
    assert hwnd and clipboard.IsWindow(hwnd)
    assert clipboard.owner_window() == hwnd, "made once"
    opened_with = []
    monkeypatch.setattr(clipboard, "OpenClipboard", lambda h: opened_with.append(h) or True)
    monkeypatch.setattr(clipboard, "CloseClipboard", lambda: True)
    with clipboard.opened():
        pass
    assert opened_with == [hwnd]


@pytest.mark.skipif(sys.platform != "win32", reason="makes a Windows window")
class TestTheOwnerWindow:
    @pytest.fixture(autouse=True)
    def fresh(self, monkeypatch):
        monkeypatch.setattr(clipboard, "_owner", None)
        monkeypatch.setattr(clipboard, "_pending", None)

    def test_a_window_that_could_not_be_made_is_tried_again(self, monkeypatch, caplog):
        real = clipboard.CreateWindowEx
        monkeypatch.setattr(clipboard, "CreateWindowEx", lambda *a: None)
        with caplog.at_level("WARNING"):
            assert clipboard.owner_window() is None
        assert "could not be made" in caplog.text
        monkeypatch.setattr(clipboard, "CreateWindowEx", real)
        hwnd = clipboard.owner_window()
        assert hwnd and clipboard.IsWindow(hwnd)

    def test_without_a_window_the_clipboard_is_opened_without_one(self, monkeypatch):
        monkeypatch.setattr(clipboard, "owner_window", lambda: None)
        opened_with = []
        monkeypatch.setattr(clipboard, "OpenClipboard", lambda h: opened_with.append(h) or True)
        monkeypatch.setattr(clipboard, "CloseClipboard", lambda: True)
        with clipboard.opened():
            pass
        assert opened_with == [None]

    def test_a_window_not_made_in_time_is_waited_for_again_not_made_twice(self, monkeypatch, caplog):
        import threading

        go = threading.Event()
        made = []
        real = clipboard._run_owner

        def slow(ready, box):
            go.wait(5)
            made.append(1)
            real(ready, box)

        monkeypatch.setattr(clipboard, "_run_owner", slow)
        monkeypatch.setattr(clipboard, "OWNER_WAIT", 0.05)
        with caplog.at_level("WARNING"):
            assert clipboard.owner_window() is None
        assert "not made within" in caplog.text
        go.set()
        monkeypatch.setattr(clipboard, "OWNER_WAIT", 5.0)
        hwnd = clipboard.owner_window()
        assert hwnd and clipboard.IsWindow(hwnd)
        assert made == [1], "the window being made was waited for, not made again"
