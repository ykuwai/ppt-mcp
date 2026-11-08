"""ppt_set_window_state on the Mac: maximised and normal through the front
deck window's `zoomed`, minimised refused by name.

PowerPoint for Mac 16.114: `zoomed` of a document window fills the screen,
gives the old bounds back, reads back what it holds, and does not bring
PowerPoint to the front. `miniaturized` and `collapsed` read empty and setting
them does nothing; VBA's WindowState is taken and ignored. Nothing here reaches
PowerPoint.
"""

import sys

import pytest

sys.path.insert(0, "src")

pytestmark = pytest.mark.skipif(sys.platform != "darwin", reason="the Apple Event backend needs appscript")


class _Prop:
    def __init__(self, value, sticks=True):
        self.value, self.sticks, self.writes = value, sticks, []

    def get(self):
        return self.value

    def set(self, value):
        self.writes.append(value)
        if self.sticks:
            self.value = value


class _Window:
    def __init__(self, zoomed=False, sticks=True):
        self.zoomed = _Prop(zoomed, sticks)

    def caption(self):
        return "Deck.pptx"


class _Windows:
    def __init__(self, windows):
        self.windows = windows

    def get(self):
        return list(self.windows)

    @property
    def first(self):
        return self.windows[0]


@pytest.fixture
def app(monkeypatch):
    from ppt_mac import app as mac_app

    def make(windows):
        fake = type("App", (), {"document_windows": _Windows(windows)})()
        monkeypatch.setattr(mac_app.ppt, "_get_app_impl", lambda *a, **k: fake)
        return mac_app._set_window_state_impl
    return make


@pytest.mark.parametrize("state, zoomed", [("maximized", True), ("normal", False), ("MAXIMIZED", True)])
def test_maximised_and_normal_set_zoomed_and_read_it_back(app, state, zoomed):
    window = _Window(zoomed=not zoomed)
    result = app([window])(state)
    assert result == {"success": True, "window_state": state.lower(), "window": "Deck.pptx"}
    assert window.zoomed.writes == [zoomed]


def test_minimised_is_refused_by_name_and_nothing_is_touched(app):
    window = _Window()
    result = app([window])("minimized")
    assert result["error"] == "ppt_set_window_state cannot set window_state='minimized' on macOS"
    assert "zoomed" in result["reason"] and "miniaturized" in result["reason"]
    assert window.zoomed.writes == []


def test_a_value_that_is_no_state_is_the_windows_error(app):
    result = app([_Window()])("full")
    assert result == {"error": "Invalid window_state: full. Use 'normal', 'minimized', or 'maximized'."}


def test_with_no_deck_open_there_is_no_window(app):
    result = app([])("maximized")
    assert "No presentation window is open" in result["error"]


def test_a_zoom_that_did_not_take_is_an_error(app):
    window = _Window(zoomed=False, sticks=False)
    assert app([window])("maximized") == {"error": "PowerPoint kept the window normal, not maximized."}


@pytest.mark.parametrize("zoomed, state", [(True, "maximized"), (False, "normal")])
def test_app_info_reads_the_window_state(monkeypatch, zoomed, state):
    from ppt_mac import app as mac_app

    fake = type("App", (), {"document_windows": _Windows([_Window(zoomed)]),
                            "presentations": _Windows([]),
                            "name": lambda self: "Microsoft PowerPoint",
                            "version": lambda self: "16.114"})()
    monkeypatch.setattr(mac_app.ppt, "_get_app_impl", lambda *a, **k: fake)
    monkeypatch.setattr(mac_app, "count", lambda x: len(x.get()))
    assert mac_app._get_app_info_impl()["window_state"] == state
