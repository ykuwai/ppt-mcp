"""backend.mac_excel: the Excel a chart call brings up on the Mac is hidden
when it took the front, and quit when the call started it and closed the
chart's workbook."""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

pytestmark = pytest.mark.skipif(sys.platform != "darwin", reason="AppKit")


class Excel:
    def __init__(self, pid, active=False):
        self.pid, self.active, self.calls = pid, active, []

    def processIdentifier(self):
        return self.pid

    def isActive(self):
        return self.active

    def hide(self):
        self.calls.append("hide")
        return False  # as the real one answers

    def terminate(self):
        self.calls.append("terminate")
        return True


class World:
    """Which Excel is running, call by call: a list played in order, the last
    one kept."""

    def __init__(self, *states):
        self.states, self.now = list(states), 0.0

    def running(self):
        return self.states.pop(0) if len(self.states) > 1 else self.states[0]

    def clock(self):
        return self.now

    def pause(self, s):
        self.now += s


def _call(world, closed=True, raises=False):
    from backend.mac_excel import out_of_the_way

    with out_of_the_way(running=world.running, pause=world.pause, clock=world.clock, limit=2) as call:
        if raises:
            raise RuntimeError("Excel stopped")
        call.workbook_closed = closed


def test_an_excel_the_call_started_is_quit_and_waited_out():
    excel = Excel(700, active=True)
    world = World(None, excel, excel, excel, None)
    _call(world)
    assert excel.calls == ["terminate"]
    assert world.now == pytest.approx(0.1)


def test_one_the_call_started_but_left_a_workbook_open_in_is_only_hidden():
    excel = Excel(700, active=True)
    _call(World(None, excel), closed=False)
    assert excel.calls == ["hide"]


def test_one_the_call_failed_in_is_only_hidden():
    excel = Excel(700, active=True)
    with pytest.raises(RuntimeError):
        _call(World(None, excel), raises=True)
    assert excel.calls == ["hide"]


def test_the_users_excel_is_never_quit_and_is_hidden_only_if_it_took_the_front():
    from backend.mac_excel import out_of_the_way

    excel = Excel(500, active=False)
    with out_of_the_way(running=lambda: excel) as call:
        excel.active = True  # Activate brought it to the front
        call.workbook_closed = True
    assert excel.calls == ["hide"]


def test_an_excel_that_was_in_front_already_stays_there():
    excel = Excel(500, active=True)
    _call(World(excel))
    assert excel.calls == []


def test_a_quit_that_does_not_end_is_logged_and_left(caplog):
    excel = Excel(700, active=True)
    world = World(None, excel)
    _call(world)
    assert excel.calls == ["terminate"] and world.now >= 2
    assert "still running" in caplog.text


def test_no_excel_at_all_is_left_alone():
    _call(World(None))


def test_the_left_open_words_are_in_the_charts_warning():
    from backend.vba import LEFT_OPEN_WORDS
    from ppt_com.charts import WORKBOOK_LEFT_OPEN

    assert LEFT_OPEN_WORDS in WORKBOOK_LEFT_OPEN
    assert LEFT_OPEN_WORDS in str({"warnings": [WORKBOOK_LEFT_OPEN]})


def test_a_chart_call_says_whether_it_closed_the_workbook(monkeypatch):
    from backend import mac_excel, vba

    used = []

    class Spy:
        def __enter__(self):
            used.append("enter")
            self.call = mac_excel.Call()
            return self.call

        def __exit__(self, *exc):
            used.append(self.call.workbook_closed)

    monkeypatch.setattr(mac_excel, "out_of_the_way", lambda: Spy())

    def chart_tool():
        return {"success": True, "warnings": [vba.LEFT_OPEN_WORDS]}
    chart_tool.__module__ = "ppt_com.charts"
    assert vba._with_excel_put_away(chart_tool, (), {})["success"]
    assert used == ["enter", False]
    used.clear()

    def clean():
        return {"success": True}
    vba._with_excel_put_away(clean, (), {})
    assert used == ["enter", True]
