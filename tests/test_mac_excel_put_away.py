"""backend.mac_excel: after a chart call that opened a chart's workbook, the
Excel it brought up is hidden when it took the front, and quit when the call
started it, closed the workbook, and it shows no window of the user's. After a
call that did not open one, Excel is left alone."""

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


def _call(world, opened=True, left_open=False, raises=False, windows=False):
    from backend.mac_excel import out_of_the_way

    with out_of_the_way(running=world.running, shows_a_window=lambda pid: windows,
                        pause=world.pause, clock=world.clock, limit=2) as call:
        call.opened = opened
        call.left_open = left_open
        if raises:
            raise RuntimeError("Excel stopped")


def test_an_excel_the_call_started_is_quit_and_waited_out():
    excel = Excel(700, active=True)
    world = World(None, excel, excel, excel, None)
    _call(world)
    assert excel.calls == ["terminate"]
    assert world.now == pytest.approx(0.1)


def test_one_the_call_started_but_left_a_workbook_open_in_is_only_hidden():
    excel = Excel(700, active=True)
    _call(World(None, excel), left_open=True)
    assert excel.calls == ["hide"]


def test_one_the_call_failed_in_is_only_hidden():
    excel = Excel(700, active=True)
    with pytest.raises(RuntimeError):
        _call(World(None, excel), raises=True)
    assert excel.calls == ["hide"]


def test_one_showing_a_window_the_user_opened_meanwhile_is_only_hidden():
    excel = Excel(700, active=True)
    world = World(None, excel)
    _call(world, windows=True)
    assert excel.calls == ["hide"]
    assert world.now >= 3  # the closed workbook's window was given time to go


def test_the_closed_workbooks_window_is_waited_out_before_the_quit():
    from backend.mac_excel import out_of_the_way

    excel = Excel(700, active=True)
    world = World(None, excel, excel, None)
    left = [True] * 10  # gone after half a second
    with out_of_the_way(running=world.running, shows_a_window=lambda pid: bool(left and left.pop()),
                        pause=world.pause, clock=world.clock, limit=2) as call:
        call.opened = True
    assert excel.calls == ["terminate"]


def test_a_call_that_opened_no_workbook_leaves_excel_alone():
    # The user started Excel while ppt_format_chart ran: not the call's.
    excel = Excel(700, active=True)
    _call(World(None, excel), opened=False)
    assert excel.calls == []


def test_the_users_excel_is_never_quit_and_is_hidden_only_if_it_took_the_front():
    from backend.mac_excel import out_of_the_way

    excel = Excel(500, active=False)
    with out_of_the_way(running=lambda: excel, shows_a_window=lambda pid: False) as call:
        excel.active = True  # Activate brought it to the front
        call.opened = True
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


def test_a_failed_look_before_the_call_does_not_stop_it():
    from backend.mac_excel import out_of_the_way

    def broken():
        raise OSError("no AppKit")
    ran = []
    with out_of_the_way(running=broken) as call:
        ran.append(call)
    assert ran


class _Workbook:
    def __init__(self, closes=True):
        self.closes = closes

    def Close(self, save):
        if not self.closes:
            raise RuntimeError("busy")


def test_closing_a_charts_workbook_notes_it_on_the_call():
    from backend.mac_excel import Call
    from ppt_com import charts

    call = Call()
    token = charts.excel_call.set(call)
    try:
        assert charts._close_chart_workbook(None, _Workbook())
        assert (call.opened, call.left_open) == (True, False)
        stuck = type("Chart", (), {"ChartData": type("D", (), {"Workbook": _Workbook(False)})()})()
        assert not charts._close_chart_workbook(stuck, _Workbook(False), attempts=2, wait=0)
        assert call.left_open
    finally:
        charts.excel_call.reset(token)


def test_a_chart_call_on_the_bridge_is_put_away_after_its_handles(monkeypatch):
    from backend import mac_excel, vba
    from ppt_com import charts

    order = []

    class Spy:
        def __enter__(self):
            self.call = mac_excel.Call()
            return self.call

        def __exit__(self, *exc):
            order.append(("put away", self.call.opened))

    class Bridge:
        stale = False

        def application(self):
            return object()

        def release(self):
            order.append("release")

    monkeypatch.setattr(mac_excel, "out_of_the_way", lambda: Spy())

    def chart_tool():
        charts.excel_call.get().opened = True
        order.append("tool")
    chart_tool.__module__ = "ppt_com.charts"
    vba._run(Bridge(), chart_tool, (), {})
    assert order == ["tool", "release", ("put away", True)]
    assert charts.excel_call.get() is None

    order.clear()

    def other_tool():
        assert charts.excel_call.get() is None
        order.append("tool")
    other_tool.__module__ = "ppt_com.shapes"
    vba._run(Bridge(), other_tool, (), {})
    assert order == ["tool", "release"]
