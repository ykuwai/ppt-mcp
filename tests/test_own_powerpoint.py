"""A PowerPoint this server starts stays out of the way, and goes when done.

Two things a user saw every time a tool call started PowerPoint:

* on macOS, PowerPoint came to the front, taking the keyboard from whatever
  the user was typing in, though a new deck's window arrives on screen by
  itself;
* once the last deck was closed, PowerPoint stayed running, with an empty
  window on Windows or its start screen on macOS, until someone quit it.

Now the server remembers whether it started PowerPoint. It launches it in the
background on macOS, activates a deck's window only when another deck's is in
front, and quits PowerPoint when a close leaves no deck open, unless the user
started it or the caller passed keep_powerpoint_open=true.

No COM and no PowerPoint required.
"""

from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import MagicMock

import pytest

_src_dir = str(Path(__file__).resolve().parents[1] / "src")
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)

import backend  # noqa: E402

macos_only = pytest.mark.skipif(sys.platform != "darwin", reason="Apple Event impl")
# The COM close is replaced by the Apple Event one on macOS.
windows_only = pytest.mark.skipif(
    sys.platform != "win32",
    reason="on macOS the Apple Event impl takes the place of the COM one",
)


@pytest.fixture
def server_started(monkeypatch):
    monkeypatch.setattr(backend.ppt, "_launched_by_us", True)
    monkeypatch.setattr(backend.ppt, "_still_ours", lambda: True)
    monkeypatch.setattr(backend.ppt, "_app", object())
    monkeypatch.setattr(backend.ppt, "_target_pres_full_name", "/decks/a.pptx")


# ---------------------------------------------------------------------------
# The rule both platforms share
# ---------------------------------------------------------------------------
class TestQuitIfOurs:
    def test_a_powerpoint_this_server_started_is_quit(self, server_started):
        quit = MagicMock()
        result = backend.quit_if_ours(quit)
        quit.assert_called_once_with()
        assert result["powerpoint_quit"] is True
        assert "starts it again" in result["note"]

    def test_and_forgotten_so_the_next_call_starts_it_again(self, server_started):
        backend.quit_if_ours(MagicMock())
        assert backend.ppt._app is None
        assert backend.ppt._launched_by_us is False
        assert backend.ppt._target_pres_full_name is None

    def test_a_powerpoint_the_user_started_is_left_alone(self, monkeypatch):
        monkeypatch.setattr(backend.ppt, "_launched_by_us", False)
        quit = MagicMock()
        assert backend.quit_if_ours(quit) == {}
        quit.assert_not_called()

    def test_one_that_took_the_place_of_ours_is_left_alone(self, server_started, monkeypatch):
        # The user quit the PowerPoint this server started and started another.
        monkeypatch.setattr(backend.ppt, "_still_ours", lambda: False)
        quit = MagicMock()
        assert backend.quit_if_ours(quit) == {}
        quit.assert_not_called()
        assert backend.ppt._launched_by_us is False

    def test_a_quit_that_fails_does_not_fail_the_close(self, server_started):
        quit = MagicMock(side_effect=RuntimeError("busy"))
        assert backend.quit_if_ours(quit) == {}
        assert backend.ppt._launched_by_us is True


@macos_only
class TestMacKnowsWhichPowerPointItStarted:
    def _wrapper(self, monkeypatch, started_pid, now_pid):
        from backend import mac_ae

        wrapper = mac_ae.PowerPointAppleEventWrapper()
        wrapper._launched_by_us = True
        wrapper._launched_pid = started_pid
        monkeypatch.setattr(mac_ae, "_powerpoint_pid", lambda: now_pid)
        return wrapper

    def test_the_same_process_is_ours(self, monkeypatch):
        assert self._wrapper(monkeypatch, 100, 100)._still_ours() is True

    def test_another_process_is_not(self, monkeypatch):
        assert self._wrapper(monkeypatch, 100, 200)._still_ours() is False

    def test_a_process_that_is_gone_is_not(self, monkeypatch):
        assert self._wrapper(monkeypatch, 100, None)._still_ours() is False

    def test_a_process_we_could_not_read_at_launch_is_not_quit(self, monkeypatch):
        assert self._wrapper(monkeypatch, None, 200)._still_ours() is False

    def test_what_we_did_not_start_is_never_ours(self, monkeypatch):
        wrapper = self._wrapper(monkeypatch, 100, 100)
        wrapper._launched_by_us = False
        assert wrapper._still_ours() is False


class TestWindowsReadsTheProcessIdFromAWindow:
    def _wrapper(self, monkeypatch, app, pid=4242, frames=()):
        import sys
        from utils.com_wrapper import PowerPointCOMWrapper

        process = MagicMock()
        process.GetWindowThreadProcessId.side_effect = lambda hwnd: (0, pid + hwnd)
        monkeypatch.setitem(sys.modules, "win32process", process)
        gui = MagicMock()
        gui.GetClassName.return_value = "PPTFrameClass"
        gui.EnumWindows.side_effect = lambda cb, extra: [cb(h, extra) for h in frames]
        monkeypatch.setitem(sys.modules, "win32gui", gui)
        wrapper = PowerPointCOMWrapper()
        wrapper._app = app
        return wrapper

    def test_the_active_window_is_used_first(self, monkeypatch):
        app = MagicMock(ActiveWindow=MagicMock(HWND=1))
        assert self._wrapper(monkeypatch, app)._app_pid() == 4243

    def test_the_application_handle_is_not_used(self, monkeypatch):
        # Late binding answers Application.HWND with a bound method.
        app = MagicMock(HWND=1, ActiveWindow=MagicMock(HWND=2))
        assert self._wrapper(monkeypatch, app)._app_pid() == 4244

    def test_the_main_frame_answers_when_there_is_no_document_window(self, monkeypatch):
        app = MagicMock()
        app.ActiveWindow.HWND = 0
        app.Windows.side_effect = RuntimeError("no window")
        assert self._wrapper(monkeypatch, app, frames=[5])._app_pid() == 4247

    def test_two_frames_are_ambiguous(self, monkeypatch):
        app = MagicMock()
        app.ActiveWindow.HWND = 0
        app.Windows.side_effect = RuntimeError("no window")
        assert self._wrapper(monkeypatch, app, frames=[5, 6])._app_pid() is None

    def test_a_zero_handle_is_not_an_answer(self, monkeypatch):
        app = MagicMock(ActiveWindow=MagicMock(HWND=0))
        app.Windows.side_effect = RuntimeError("no window")
        assert self._wrapper(monkeypatch, app)._app_pid() is None


class TestWindowsKnowsWhichPowerPointItStarted:
    def _wrapper(self, started_pid, now_pid, ours=True):
        from utils.com_wrapper import PowerPointCOMWrapper

        wrapper = PowerPointCOMWrapper()
        wrapper._launched_by_us = ours
        wrapper._launched_pid = started_pid
        wrapper._app_pid = lambda: now_pid
        return wrapper

    def test_the_same_process_is_ours(self):
        assert self._wrapper(100, 100)._still_ours() is True

    def test_another_process_is_not(self):
        assert self._wrapper(100, 200)._still_ours() is False

    def test_without_a_process_id_the_flag_decides(self):
        assert self._wrapper(None, 200)._still_ours() is True
        assert self._wrapper(None, 200, ours=False)._still_ours() is False


# ---------------------------------------------------------------------------
# Windows: who started it
# ---------------------------------------------------------------------------
class TestWindowsRemembersWhoStartedIt:
    def _connect(self, monkeypatch, running):
        from utils import com_wrapper

        client = MagicMock()
        if running:
            client.GetActiveObject.return_value = MagicMock(Visible=True)
        else:
            client.GetActiveObject.side_effect = com_wrapper.pywintypes.com_error()
            client.Dispatch.return_value = MagicMock(Visible=False)
        monkeypatch.setattr(com_wrapper, "win32com", MagicMock(client=client))
        wrapper = com_wrapper.PowerPointCOMWrapper()
        wrapper._connect_impl(allow_launch=True)
        return wrapper, client

    def test_launched_here(self, monkeypatch):
        wrapper, client = self._connect(monkeypatch, running=False)
        client.Dispatch.assert_called_once_with("PowerPoint.Application")
        assert wrapper._launched_by_us is True

    def test_found_running(self, monkeypatch):
        wrapper, client = self._connect(monkeypatch, running=True)
        client.Dispatch.assert_not_called()
        assert wrapper._launched_by_us is False

    def test_not_known_until_connected(self):
        from utils.com_wrapper import PowerPointCOMWrapper

        assert PowerPointCOMWrapper()._launched_by_us is False


@windows_only
class TestWindowsClose:
    def _close(self, monkeypatch, left_open, keep=False, protected=0):
        import ppt_com.presentation as presentation

        app = MagicMock()
        pres = MagicMock(Name="a.pptx")
        app.Presentations.Count = left_open
        app.ProtectedViewWindows.Count = protected
        monkeypatch.setattr(presentation.ppt, "_get_app_impl", lambda allow_launch=False: app)
        monkeypatch.setattr(presentation, "_resolve_presentation", lambda app, **kw: pres)
        return presentation._close_presentation_impl(False, None, None, keep), app

    def test_the_last_deck_quits_it(self, monkeypatch, server_started):
        result, app = self._close(monkeypatch, left_open=0)
        app.Quit.assert_called_once_with()
        assert result["powerpoint_quit"] is True

    def test_not_while_a_deck_is_open(self, monkeypatch, server_started):
        result, app = self._close(monkeypatch, left_open=1)
        app.Quit.assert_not_called()
        assert "powerpoint_quit" not in result

    def test_not_when_asked_to_keep_it(self, monkeypatch, server_started):
        _, app = self._close(monkeypatch, left_open=0, keep=True)
        app.Quit.assert_not_called()

    def test_not_while_a_file_is_open_in_protected_view(self, monkeypatch, server_started):
        result, app = self._close(monkeypatch, left_open=0, protected=1)
        app.Quit.assert_not_called()
        assert "powerpoint_quit" not in result

    def test_not_one_the_user_started(self, monkeypatch):
        monkeypatch.setattr(backend.ppt, "_launched_by_us", False)
        result, app = self._close(monkeypatch, left_open=0)
        app.Quit.assert_not_called()
        assert result == {"success": True, "closed": "a.pptx"}

    def _protected(self, monkeypatch, save):
        # #264: closing a file open in Protected View raised at Saved = True.
        from ppt_com import presentation

        app = MagicMock()
        pres = MagicMock(Name="a.pptx", FullName=r"C:\a.pptx")
        type(pres).Saved = property(lambda self: True, _refuse)
        window = MagicMock()
        window.Presentation.FullName = r"C:\a.pptx"
        app.ProtectedViewWindows.Count = 1
        app.ProtectedViewWindows.return_value = window
        app.Presentations.Count = 0
        monkeypatch.setattr(presentation.ppt, "_get_app_impl", lambda allow_launch=False: app)
        monkeypatch.setattr(presentation, "_resolve_presentation", lambda app, **kw: pres)
        return presentation._close_presentation_impl(save, None, None, True), pres, window

    def test_a_file_in_protected_view_is_closed_through_its_window(self, monkeypatch):
        result, pres, window = self._protected(monkeypatch, save=False)
        window.Close.assert_called_once_with()
        pres.Close.assert_not_called()
        assert result["closed"] == "a.pptx"

    def test_and_is_not_saved(self, monkeypatch):
        with pytest.raises(ValueError, match="Protected View, which is read-only"):
            self._protected(monkeypatch, save=True)


def _refuse(self, value):
    raise RuntimeError("Presentation.Saved : Invalid request.")


# ---------------------------------------------------------------------------
# macOS: launched in the background, and closed like Windows
# ---------------------------------------------------------------------------
@macos_only
class TestMacStartsItInTheBackground:
    def _connect(self, monkeypatch, running, visible=None):
        from backend import mac_ae

        candidate = MagicMock()
        candidate.isrunning.return_value = running
        monkeypatch.setattr(mac_ae, "app", lambda **kw: candidate)
        wrapper = mac_ae.PowerPointAppleEventWrapper()
        wrapper._connect_impl(visible=visible, allow_launch=True)
        return wrapper, candidate

    def test_launched_without_taking_the_front(self, monkeypatch):
        wrapper, candidate = self._connect(monkeypatch, running=False)
        candidate.launch.assert_called_once_with()
        candidate.activate.assert_not_called()
        assert wrapper._launched_by_us is True

    def test_unless_the_caller_asks_to_see_it(self, monkeypatch):
        _, candidate = self._connect(monkeypatch, running=False, visible=True)
        candidate.activate.assert_called_once_with()

    def test_found_running(self, monkeypatch):
        wrapper, candidate = self._connect(monkeypatch, running=True)
        candidate.launch.assert_not_called()
        assert wrapper._launched_by_us is False


@macos_only
class TestMacBringsAWindowForwardOnlyWhenItIsNot:
    def _bring(self, active_full_name):
        from ppt_mac.presentation import _bring_window_forward

        app, pres = MagicMock(), MagicMock()
        app.active_window.presentation.full_name.return_value = active_full_name
        _bring_window_forward(app, pres, "/decks/new.pptx")
        return pres.document_windows[1].activate

    def test_a_new_deck_is_already_in_front_in_powerpoint(self):
        self._bring("/decks/new.pptx").assert_not_called()

    def test_another_decks_window_is_in_the_way(self):
        self._bring("/decks/other.pptx").assert_called_once_with()


@macos_only
class TestMacClose:
    def _close(self, monkeypatch, left_open, keep=False):
        import ppt_mac.presentation as presentation

        app = MagicMock()
        pres = MagicMock()
        pres.name.return_value = "Presentation1"
        pres.full_name.return_value = "Presentation1"
        monkeypatch.setattr(presentation.ppt, "_get_app_impl", lambda allow_launch=False: app)
        monkeypatch.setattr(presentation, "_resolve_presentation", lambda app, **kw: pres)
        monkeypatch.setattr(presentation, "_full_names", lambda app: list(left_open))
        return presentation._close_presentation_impl(False, None, None, keep), app

    def test_the_last_deck_quits_it(self, monkeypatch, server_started):
        result, app = self._close(monkeypatch, left_open=[])
        app.quit.assert_called_once_with()
        assert result["powerpoint_quit"] is True

    def test_not_while_a_deck_is_open(self, monkeypatch, server_started):
        result, app = self._close(monkeypatch, left_open=["Other"])
        app.quit.assert_not_called()
        assert "powerpoint_quit" not in result

    def test_not_when_asked_to_keep_it(self, monkeypatch, server_started):
        _, app = self._close(monkeypatch, left_open=[], keep=True)
        app.quit.assert_not_called()

    def test_not_one_the_user_started(self, monkeypatch):
        monkeypatch.setattr(backend.ppt, "_launched_by_us", False)
        result, app = self._close(monkeypatch, left_open=[])
        app.quit.assert_not_called()
        assert result == {"success": True, "closed": "Presentation1"}
