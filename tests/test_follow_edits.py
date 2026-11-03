"""PPT_FOLLOW_EDITS: whether a write moves the deck's window to its slide."""

import logging
import sys

import pytest

sys.path.insert(0, "src")

from utils import navigation  # noqa: E402


@pytest.mark.parametrize("value, follows", [
    ("on", True), ("", True), (" ON ", True), ("1", True),
    ("off", False), ("OFF", False), ("0", False), ("no", False),
])
def test_the_setting(value, follows):
    assert navigation._follow_edits(value) is follows


def test_a_mistyped_setting_follows_with_a_warning(caplog):
    with caplog.at_level(logging.WARNING):
        assert navigation._follow_edits("of") is True
    assert "PPT_FOLLOW_EDITS='of' is not a setting" in caplog.text


def test_off_leaves_the_window_alone(monkeypatch):
    def reached(*args, **kwargs):
        raise AssertionError("the window was looked up")

    monkeypatch.setattr(navigation, "FOLLOW_EDITS", False)
    monkeypatch.setattr(navigation.ppt, "_get_pres_impl", reached, raising=False)
    monkeypatch.setattr(navigation.ppt, "_get_target_window_impl", reached, raising=False)
    navigation.goto_slide(None, 3)


def test_on_moves_the_window(monkeypatch):
    moved = []

    class View:
        def GotoSlide(self, index):
            moved.append(index)

        def go_to_slide(self, number):
            moved.append(number)

    view = View()

    class Window:
        View = view

    class Pres:
        pass

    monkeypatch.setattr(navigation, "FOLLOW_EDITS", True)
    monkeypatch.setattr(navigation.ppt, "_get_target_window_impl", lambda: Window(), raising=False)
    monkeypatch.setattr(navigation.ppt, "_get_pres_impl", lambda: Pres(), raising=False)
    if navigation.IS_MACOS:
        from backend import mac_ae
        monkeypatch.setattr(mac_ae, "target_window",
                            lambda pres: type("MacWindow", (), {"view": view})())
    navigation.goto_slide(None, 3)
    assert moved == [3]


def test_a_tool_that_needs_the_slide_on_screen_moves_it_anyway(monkeypatch):
    moved = []

    class View:
        def GotoSlide(self, index):
            moved.append(index)

        def go_to_slide(self, number):
            moved.append(number)

    view = View()
    monkeypatch.setattr(navigation, "FOLLOW_EDITS", False)
    monkeypatch.setattr(navigation.ppt, "_get_target_window_impl",
                        lambda: type("W", (), {"View": view})(), raising=False)
    monkeypatch.setattr(navigation.ppt, "_get_pres_impl", lambda: object(), raising=False)
    if navigation.IS_MACOS:
        from backend import mac_ae
        monkeypatch.setattr(mac_ae, "target_window",
                            lambda pres: type("MacWindow", (), {"view": view})())
    navigation.goto_slide(None, 4, force=True)
    assert moved == [4]


class _Window:
    """A deck's window on either platform, showing one slide."""

    def __init__(self, showing):
        self.showing, self.moves = showing, []
        window = self

        class Slide:
            SlideIndex = showing

            def slide_index(self):
                return window.showing

        shown = Slide()

        class View:
            slide = shown
            Slide = shown

            def GotoSlide(self, index):
                window.moves.append(index)

            def go_to_slide(self, number):
                window.moves.append(number)

        self.view = self.View = View()


def test_a_paste_puts_the_users_slide_back_while_off(monkeypatch):
    window = _Window(showing=2)
    monkeypatch.setattr(navigation, "FOLLOW_EDITS", False)
    was = navigation.shown_slide(window)
    assert was == 2
    navigation.return_to(window, was)
    assert window.moves == [2]


def test_a_paste_leaves_the_window_on_its_slide_while_on(monkeypatch):
    window = _Window(showing=2)
    monkeypatch.setattr(navigation, "FOLLOW_EDITS", True)
    navigation.return_to(window, navigation.shown_slide(window))
    assert window.moves == []
