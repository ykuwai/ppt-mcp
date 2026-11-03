"""ppt_com.charts._close_chart_workbook: a chart's data workbook is closed, or
the failure is reported, never left open in silence.

A workbook left open keeps the chart in edit mode, and the next chart
PowerPoint is asked to insert or edit answers with a dialog, after which every
call times out. Runs anywhere, against stand-ins for the chart and workbook.
"""

import sys

sys.path.insert(0, "src")

import pytest  # noqa: E402

from ppt_com.charts import _close_chart_workbook  # noqa: E402


class Workbook:
    def __init__(self, failures, error=AttributeError("<unknown>.Close")):
        self.failures, self.error, self.closed = failures, error, False

    def Close(self, save):
        assert save is False
        if self.failures:
            self.failures -= 1
            raise self.error
        self.closed = True


class ChartData:
    def __init__(self, workbook, busy=0):
        self.workbook, self.busy, self.asked, self.activated = workbook, busy, 0, 0

    def Activate(self):
        self.activated += 1

    @property
    def Workbook(self):
        self.asked += 1
        if self.busy:
            self.busy -= 1
            raise RuntimeError("Call was rejected by callee")
        return self.workbook


class Chart:
    def __init__(self, data):
        self.ChartData = data


def test_a_workbook_that_closes_is_closed_once():
    workbook = Workbook(failures=0)
    chart = Chart(ChartData(workbook))
    assert _close_chart_workbook(chart, workbook, wait=0) is True
    assert workbook.closed and chart.ChartData.asked == 0


def test_close_not_found_while_excel_starts_is_retried_on_the_workbook_asked_again():
    # Excel still starting up hands back a workbook without Close.
    workbook = Workbook(failures=2)
    chart = Chart(ChartData(workbook))
    assert _close_chart_workbook(chart, workbook, wait=0) is True
    assert workbook.closed
    assert chart.ChartData.asked == 2


def test_a_busy_refusal_to_hand_out_the_workbook_is_retried_not_taken_as_closed():
    workbook = Workbook(failures=1)
    chart = Chart(ChartData(workbook, busy=2))
    assert _close_chart_workbook(chart, workbook, wait=0) is True
    assert workbook.closed
    assert chart.ChartData.asked == 3


def test_a_new_charts_workbook_is_opened_and_closed():
    workbook = Workbook(failures=0)
    chart = Chart(ChartData(workbook))
    assert _close_chart_workbook(chart, wait=0) is True
    assert workbook.closed and chart.ChartData.activated == 1


def test_a_workbook_that_cannot_be_had_at_all_is_reported():
    chart = Chart(ChartData(Workbook(failures=0), busy=99))
    assert _close_chart_workbook(chart, attempts=3, wait=0) is False


@pytest.mark.skipif(sys.platform != "win32",
                    reason="drives the COM impls, which macOS replaces with its own")
class TestTheToolsSayWhenItWasLeftOpen:
    """add_chart and set_chart_data carry a warning when the workbook stayed open."""

    def deck(self, monkeypatch, workbook):
        from types import SimpleNamespace

        from ppt_com import charts

        data = ChartData(workbook)
        chart = SimpleNamespace(ChartData=data, _oleobj_=SimpleNamespace(InvokeTypes=lambda *a: None))
        shape = SimpleNamespace(Name="Chart 1", ZOrderPosition=1, Chart=chart, HasChart=True)
        cells = SimpleNamespace(Clear=lambda: None)
        sheet = SimpleNamespace(Cells=lambda *a: SimpleNamespace(Value=None) if a else cells,
                                Range=lambda *a: SimpleNamespace(Address="$A$1:$B$2"))
        sheet.Cells.Clear = lambda: None
        workbook.Worksheets = lambda i: sheet
        shapes = SimpleNamespace(AddChart2=lambda *a: shape)
        slide = SimpleNamespace(Shapes=shapes)
        pres = SimpleNamespace(Slides=lambda i: slide)
        monkeypatch.setattr(charts.ppt, "_get_app_impl", lambda *a, **k: None, raising=False)
        monkeypatch.setattr(charts.ppt, "_get_pres_impl", lambda *a, **k: pres, raising=False)
        monkeypatch.setattr(charts, "goto_slide", lambda *a: None)
        monkeypatch.setattr(charts, "_get_chart_shape", lambda slide, ref: shape, raising=False)
        monkeypatch.setattr(charts, "_wait_pumping", lambda s: None)
        return charts

    def test_add_chart(self, monkeypatch):
        charts = self.deck(monkeypatch, Workbook(failures=99))
        assert "warning" in charts._add_chart_impl(1, "column", 0, 0, 100, 100)
        charts = self.deck(monkeypatch, Workbook(failures=0))
        assert "warning" not in charts._add_chart_impl(1, "column", 0, 0, 100, 100)

    def test_set_chart_data(self, monkeypatch):
        charts = self.deck(monkeypatch, Workbook(failures=99))
        result = charts._set_chart_data_impl(1, "Chart 1", ["Q1"], [{"name": "S", "values": [1]}])
        assert result["warning"] == charts.WORKBOOK_LEFT_OPEN
        charts = self.deck(monkeypatch, Workbook(failures=0))
        result = charts._set_chart_data_impl(1, "Chart 1", ["Q1"], [{"name": "S", "values": [1]}])
        assert "warning" not in result


def test_a_workbook_that_never_closes_is_reported(caplog):
    workbook = Workbook(failures=99, error=OSError("Excel is busy"))
    chart = Chart(ChartData(workbook))
    with caplog.at_level("WARNING"):
        assert _close_chart_workbook(chart, workbook, attempts=3, wait=0) is False
    assert not workbook.closed
    assert "Excel is busy" in caplog.text


@pytest.mark.skipif(sys.platform != "win32", reason="marshals a real COM object to a thread")
class TestOnAThreadOfItsOwn:
    """The workbook work runs on a thread of its own, handed a COM object
    through marshaling, so letting go of Excel there cannot hold up the worker."""

    @pytest.fixture(autouse=True)
    def host(self):
        """A Dictionary living in a multithreaded apartment of its own.

        Like PowerPoint's objects, it then answers calls from any thread
        without the test's thread having to pump messages for it.
        """
        import threading

        import pythoncom
        import win32com.client

        ready, finished, held = threading.Event(), threading.Event(), {}

        def serve():
            pythoncom.CoInitializeEx(pythoncom.COINIT_MULTITHREADED)
            try:
                d = win32com.client.Dispatch("Scripting.Dictionary")
                held["stream"] = pythoncom.CoMarshalInterThreadInterfaceInStream(
                    pythoncom.IID_IDispatch, d._oleobj_)
                ready.set()
                finished.wait(30)
                d = None
            finally:
                pythoncom.CoUninitialize()

        thread = threading.Thread(target=serve, daemon=True)
        thread.start()
        ready.wait(10)
        self._stream = held["stream"]
        yield
        finished.set()
        thread.join(10)

    def dictionary(self):
        import pythoncom
        import win32com.client

        return win32com.client.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(
            self._stream, pythoncom.IID_IDispatch))

    def test_the_work_runs_elsewhere_on_the_same_object_and_answers_here(self):
        import threading

        from ppt_com.charts import _on_a_thread_of_its_own

        here = threading.get_ident()
        d = self.dictionary()

        def work(there):
            there.Add("from", "the other thread")
            return threading.get_ident()

        assert _on_a_thread_of_its_own(d, work) != here
        assert d.Item("from") == "the other thread"

    def test_an_error_there_is_raised_here(self):
        from ppt_com.charts import _on_a_thread_of_its_own

        def work(there):
            raise ValueError("Excel said no")

        with pytest.raises(ValueError, match="Excel said no"):
            _on_a_thread_of_its_own(self.dictionary(), work)

    def test_work_that_does_not_finish_in_time_is_left_to_finish_there(self):
        import time

        from ppt_com.charts import TIMED_OUT, _on_a_thread_of_its_own

        started = time.monotonic()
        assert _on_a_thread_of_its_own(self.dictionary(), lambda there: time.sleep(3),
                                       timeout=0.5) is TIMED_OUT
        assert time.monotonic() - started < 2

    def test_a_stand_in_runs_right_here(self):
        from types import SimpleNamespace

        from ppt_com.charts import _on_a_thread_of_its_own

        assert _on_a_thread_of_its_own(SimpleNamespace(), lambda c: "inline") == "inline"
