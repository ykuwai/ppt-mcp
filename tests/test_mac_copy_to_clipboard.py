"""ppt_copy_to_clipboard on macOS (#247).

It always refused, as needing an NSPasteboard implementation written from
scratch; backend/pasteboard.py has been that since #210. One slide goes on the
pasteboard as PNG data, several as file URLs, as Windows puts a DIB or an
HDROP there. Rendering (the deck's PDF through Quartz) and the pasteboard are
faked here; both were run live.
"""

import os
import sys

import pytest

sys.path.insert(0, "src")

pytestmark = pytest.mark.skipif(
    sys.platform != "darwin", reason="the Apple Event backend needs appscript"
)

PNG = b"\x89PNG\r\n\x1a\nfake"


@pytest.fixture
def deck(monkeypatch, tmp_path):
    from backend import pasteboard
    from ppt_mac import export

    state = {"written": None, "files": None, "rendered": [], "shown": 2, "went_to": None}

    class Pres:
        slides = object()

    class App:
        presentations = object()

    monkeypatch.setattr(export.ppt, "_get_app_impl", lambda *a, **kw: App())
    monkeypatch.setattr(export.ppt, "_get_pres_impl", lambda *a, **kw: Pres())
    monkeypatch.setattr(export, "count", lambda ref: 4 if ref is Pres.slides else 1)
    monkeypatch.setattr(export, "_staging_path", lambda suffix: str(tmp_path / f"deck{suffix}"))
    monkeypatch.setattr(export, "_save_pdf_of", lambda pres, path, numbers: list(numbers))

    def render(pdf, pages, width, height, uti, ext):
        state["rendered"].append((list(pages), width, height, uti))
        out = []
        for page in pages:
            path = tmp_path / f"page{page}{ext}"
            path.write_bytes(PNG)
            out.append((page, str(path), 1280, 720))
        return out

    monkeypatch.setattr(export, "_render_pdf_pages", render)
    monkeypatch.setattr(pasteboard, "write", lambda uti, data: state.update(written=(uti, data)))
    monkeypatch.setattr(pasteboard, "write_files", lambda paths: state.update(files=list(paths)))

    import backend.mac_ae as mac_ae
    import utils.navigation as navigation

    monkeypatch.setattr(mac_ae, "target_window", lambda pres: object())
    monkeypatch.setattr(navigation, "shown_slide", lambda window: state["shown"])
    monkeypatch.setattr(navigation, "goto_slide",
                        lambda app, index: state.update(went_to=index))
    state["tmp"] = tmp_path
    return state


def test_one_slide_goes_on_as_png_data(deck):
    from ppt_mac.export import _copy_to_clipboard_impl

    result = _copy_to_clipboard_impl([3], 1600, None)
    assert result == {"success": True, "slide_indices": [3], "count": 1,
                      "method": "png", "temp_dir": None}
    assert deck["written"] == ("public.png", PNG)
    assert deck["rendered"] == [([3], 1600, None, "public.png")]
    assert not (deck["tmp"] / "page3.png").exists()


def test_the_slide_shown_is_the_default(deck):
    from ppt_mac.export import _copy_to_clipboard_impl

    result = _copy_to_clipboard_impl(None, None, None)
    assert result["slide_indices"] == [2] and deck["went_to"] == 2


def test_several_slides_go_on_as_files_that_stay(deck):
    from ppt_mac.export import _copy_to_clipboard_impl

    result = _copy_to_clipboard_impl([1, 4], None, None)
    assert (result["method"], result["count"]) == ("file_urls", 2)
    assert [os.path.basename(p) for p in deck["files"]] == ["Slide1.png", "Slide4.png"]
    assert all(os.path.dirname(p) == result["temp_dir"] for p in deck["files"])
    assert all(open(p, "rb").read() == PNG for p in deck["files"])
    assert deck["went_to"] == 4


def test_an_index_out_of_range_is_refused_before_anything_is_rendered(deck):
    from ppt_mac.export import _copy_to_clipboard_impl

    with pytest.raises(ValueError, match=r"Slide index 9 out of range \(1-4\)"):
        _copy_to_clipboard_impl([1, 9], None, None)
    assert deck["rendered"] == [] and deck["written"] is None


def test_file_urls_are_written_as_file_urls(monkeypatch):
    from backend import pasteboard

    class Board:
        def clearContents(self):
            self.cleared = True

        def writeObjects_(self, objects):
            self.objects = objects
            return True

        def changeCount(self):
            return 7

    board = Board()
    monkeypatch.setattr(pasteboard, "_board", lambda: board)
    assert pasteboard.write_files(["/tmp/a b.png"]) == 7
    assert board.cleared and [str(u.path()) for u in board.objects] == ["/tmp/a b.png"]
    assert board.objects[0].isFileURL()
