"""backend.mac_ae.quit_and_wait: closing the last deck quits PowerPoint and
waits until it has gone, so the next call does not reach a PowerPoint that is
shutting down (Apple Event error -600)."""

import sys
from pathlib import Path
from unittest.mock import MagicMock

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

pytestmark = pytest.mark.skipif(sys.platform != "darwin", reason="Apple Event impl")


class Clock:
    def __init__(self):
        self.now, self.paused = 0.0, []

    def __call__(self):
        return self.now

    def pause(self, seconds):
        self.paused.append(seconds)
        self.now += seconds


def test_it_waits_until_that_powerpoint_has_gone():
    from backend.mac_ae import quit_and_wait

    app, clock = MagicMock(), Clock()
    answers = iter([4242, 4242, 4242, None])
    assert quit_and_wait(app, 4242, running=lambda: next(answers), pause=clock.pause, clock=clock) is True
    app.quit.assert_called_once_with()
    assert len(clock.paused) == 3


def test_another_powerpoint_already_running_is_not_waited_for():
    from backend.mac_ae import quit_and_wait

    clock = Clock()
    assert quit_and_wait(MagicMock(), 4242, running=lambda: 5555, pause=clock.pause, clock=clock) is True
    assert clock.paused == []


def test_past_the_limit_it_says_so_and_goes_on(caplog):
    from backend.mac_ae import quit_and_wait

    clock = Clock()
    assert quit_and_wait(MagicMock(), 4242, limit=2, running=lambda: 4242, pause=clock.pause, clock=clock) is False
    assert clock.now >= 2 and "still running" in caplog.text


def test_with_no_process_id_it_only_quits():
    from backend.mac_ae import quit_and_wait

    app = MagicMock()
    assert quit_and_wait(app, None, running=lambda: pytest.fail("not asked")) is True
    app.quit.assert_called_once_with()


def test_closing_the_last_deck_waits_for_the_powerpoint_this_server_started(monkeypatch):
    import backend
    import ppt_mac.presentation as presentation
    from backend import mac_ae

    app, pres = MagicMock(), MagicMock()
    pres.name.return_value = pres.full_name.return_value = "Presentation1"
    monkeypatch.setattr(backend.ppt, "_launched_by_us", True)
    monkeypatch.setattr(backend.ppt, "_launched_pid", 4242)
    monkeypatch.setattr(backend.ppt, "_still_ours", lambda: True)
    monkeypatch.setattr(backend.ppt, "_app", app)
    monkeypatch.setattr(presentation.ppt, "_get_app_impl", lambda allow_launch=False: app)
    monkeypatch.setattr(presentation, "_resolve_presentation", lambda app, **kw: pres)
    monkeypatch.setattr(presentation, "_full_names", lambda app: [])
    asked = []
    monkeypatch.setattr(mac_ae, "_powerpoint_pid", lambda: asked.append(1) or (4242 if len(asked) < 3 else None))
    monkeypatch.setattr(mac_ae.time, "sleep", lambda s: None)
    result = presentation._close_presentation_impl(False, None, None, False)
    assert result["powerpoint_quit"] is True
    app.quit.assert_called_once_with()
    assert len(asked) == 3
