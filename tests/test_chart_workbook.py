"""ppt_com.charts._close_chart_workbook: a chart's data workbook is closed, or
the failure is reported, never left open in silence.

A workbook left open keeps the chart in edit mode, and the next chart
PowerPoint is asked to insert or edit answers with a dialog, after which every
call times out. Runs anywhere, against stand-ins for the chart and workbook.
"""

import sys
from pathlib import Path

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
        shown = SimpleNamespace(Name="S", Values=(1.0,), XValues=("Q1",))
        chart.SeriesCollection = lambda *i: shown if i else SimpleNamespace(Count=1)
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
        assert "warning" in charts._add_chart_impl(1, "3d_column", 0, 0, 100, 100)
        charts = self.deck(monkeypatch, Workbook(failures=0))
        assert "warning" not in charts._add_chart_impl(1, "3d_column", 0, 0, 100, 100)

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


class TestTheDataIsReadBack:
    """set_chart_data reads the chart back and fails when the data did not land (#315)."""

    def compare(self, categories, series, read_categories, read_series):
        from ppt_com.charts import _compare_chart_data

        return _compare_chart_data(categories, series, read_categories, read_series)

    def test_what_was_written_is_fine(self):
        assert self.compare(["Q1", "Q2"], [{"name": "S", "values": [1, 2.5]}],
                            ["Q1", "Q2"], [{"name": "S", "values": [1.0, 2.5]}]) == (None, None)

    def test_an_empty_chart_is_a_problem(self):
        problem, _ = self.compare(["Q1", "Q2"], [{"name": "S", "values": [1, 2]}],
                                  [], [{"name": "系列1", "values": []}])
        assert problem == "it has 0 categories, not 2"

    def test_a_missing_series_is_a_problem(self):
        problem, _ = self.compare(["Q1"], [{"name": "A", "values": [1]}, {"name": "B", "values": [2]}],
                                  ["Q1"], [{"name": "A", "values": [1.0]}])
        assert problem == "it has 1 series, not 2"

    def test_a_different_value_is_a_problem(self):
        problem, _ = self.compare(["Q1", "Q2"], [{"name": "S", "values": [1, 2]}],
                                  ["Q1", "Q2"], [{"name": "S", "values": [1.0, 3.0]}])
        assert problem == "series 1 value 2 reads 3.0, not 2"

    def test_a_tiny_value_read_as_zero_is_a_problem(self):
        problem, _ = self.compare(["Q1"], [{"name": "S", "values": [1e-10]}],
                                  ["Q1"], [{"name": "S", "values": [0.0]}])
        assert problem == "series 1 value 1 reads 0.0, not 1e-10"

    def test_rounding_in_the_last_digits_is_fine(self):
        assert self.compare(["Q1"], [{"name": "S", "values": [0.1 + 0.2]}],
                            ["Q1"], [{"name": "S", "values": [0.3]}]) == (None, None)

    def test_short_values_show_as_empty_points(self):
        assert self.compare(["Q1", "Q2"], [{"name": "S", "values": [1]}],
                            ["Q1", "Q2"], [{"name": "S", "values": [1.0, None]}]) == (None, None)

    def test_a_number_label_read_as_a_number_is_fine(self):
        assert self.compare(["2024", "2025"], [{"name": "2", "values": [1, 2]}],
                            ["2024.0", "2025.0"], [{"name": 2.0, "values": [1.0, 2.0]}]) == (None, None)

    def test_a_label_excel_turned_into_a_date_is_a_warning(self):
        problem, renamed = self.compare(["1/2"], [{"name": "S", "values": [1]}],
                                        ["1月2日"], [{"name": "S", "values": [1.0]}])
        assert problem is None
        assert "category 1 reads '1月2日'" in renamed


@pytest.mark.skipif(sys.platform != "win32",
                    reason="drives the COM impls, which macOS replaces with its own")
class TestSetChartDataFailsWhenTheChartStaysEmpty:

    def deck(self, monkeypatch, workbook):
        return TestTheToolsSayWhenItWasLeftOpen().deck(monkeypatch, workbook)

    def test_an_empty_chart_fails(self, monkeypatch):
        from types import SimpleNamespace

        charts = self.deck(monkeypatch, Workbook(failures=0))
        empty = SimpleNamespace(Name="系列1", Values=(), XValues=())
        chart = charts._get_chart_shape(None, "Chart 1").Chart
        chart.SeriesCollection = lambda *i: empty if i else SimpleNamespace(Count=1)
        with pytest.raises(RuntimeError, match="did not take the data: it has 0 categories, not 1"):
            charts._set_chart_data_impl(1, "Chart 1", ["Q1"], [{"name": "S", "values": [1]}])

    def test_data_that_lands_late_is_waited_for(self, monkeypatch):
        from types import SimpleNamespace

        charts = self.deck(monkeypatch, Workbook(failures=0))
        reads = iter([((), ()), (("Q1",), (1.0,))])

        def series(*i):
            if not i:
                return SimpleNamespace(Count=1)
            categories, values = next(reads)
            return SimpleNamespace(Name="S", Values=values, XValues=categories)

        charts._get_chart_shape(None, "Chart 1").Chart.SeriesCollection = series
        result = charts._set_chart_data_impl(1, "Chart 1", ["Q1"], [{"name": "S", "values": [1]}])
        assert result["success"] and "warning" not in result


@pytest.mark.skipif(sys.platform != "win32",
                    reason="drives the COM impls, which macOS replaces with its own")
class TestExcelIsLetGoOfOnceItHasEnded:
    """The tools hand the Excel objects they got to ``held``, with the Excel
    process behind them, rather than letting go of them on return (#299)."""

    def deck(self, monkeypatch):
        from types import SimpleNamespace

        workbook = Workbook(failures=0)
        workbook.Application = SimpleNamespace(Hwnd=555)
        charts = TestTheToolsSayWhenItWasLeftOpen().deck(monkeypatch, workbook)
        handed = []
        monkeypatch.setattr(charts, "window_process", lambda hwnd: {555: 4321}.get(hwnd))
        monkeypatch.setattr(charts, "held", SimpleNamespace(
            hold=lambda objects, pid: handed.append((list(objects), pid)),
            wait_until_ended=lambda: handed.append("waited")))
        return charts, workbook, handed

    def test_add_chart_hands_over_the_closed_workbook(self, monkeypatch):
        charts, workbook, handed = self.deck(monkeypatch)
        charts._add_chart_impl(1, "3d_column", 0, 0, 100, 100)
        assert workbook.closed
        assert handed == ["waited", ([workbook], 4321)]

    def test_set_chart_data_hands_over_workbook_sheet_and_range(self, monkeypatch):
        charts, workbook, handed = self.deck(monkeypatch)
        charts._set_chart_data_impl(1, "Chart 1", ["Q1"], [{"name": "S", "values": [1]}])
        assert workbook.closed
        [(objects, pid)] = handed
        assert pid == 4321
        assert objects[0] is workbook and len(objects) == 3

    def test_with_no_excel_window_there_is_no_process(self, monkeypatch):
        charts, workbook, handed = self.deck(monkeypatch)
        del workbook.Application
        charts._add_chart_impl(1, "3d_column", 0, 0, 100, 100)
        assert handed == ["waited", ([workbook], None)]


@pytest.mark.skipif(sys.platform != "win32",
                    reason="drives the COM impls, which macOS replaces with its own")
class TestANewChartStartsNoExcel:
    """A type python-pptx writes is brought in from a file, and AddChart2,
    which starts an Excel, is left for the rest (#299)."""

    def deck(self, monkeypatch):
        from types import SimpleNamespace

        from ppt_com import charts

        calls = []

        def add_chart2(*a):
            calls.append(("AddChart2",) + a)
            raise AssertionError("AddChart2 starts an Excel")

        slide = SimpleNamespace(Shapes=SimpleNamespace(AddChart2=add_chart2))
        pres = SimpleNamespace(Slides=lambda i: slide)
        monkeypatch.setattr(charts.ppt, "_get_app_impl", lambda *a, **k: "app", raising=False)
        monkeypatch.setattr(charts.ppt, "_get_pres_impl", lambda *a, **k: pres, raising=False)
        monkeypatch.setattr(charts, "goto_slide", lambda *a: None)
        shape = SimpleNamespace(Name="Chart 1", ZOrderPosition=3)

        def from_file(app, p, s, type_int, left, top, width, height):
            calls.append(("from file", type_int, left, top, width, height))
            return shape, []

        monkeypatch.setattr(charts, "_add_chart_from_file", from_file)
        return charts, calls

    @pytest.mark.parametrize("chart_type, type_int", [
        ("column", 51), ("bar", 57), ("line", 4), ("pie", 5), ("scatter", -4169),
        ("area", 1), ("doughnut", -4120), ("radar", -4151), ("bubble", 15), (74, 74),
    ])
    def test_these_come_from_a_file(self, monkeypatch, chart_type, type_int):
        charts, calls = self.deck(monkeypatch)
        result = charts._add_chart_impl(1, chart_type, 10, 20, 300, 200)
        assert calls == [("from file", type_int, 10, 20, 300, 200)]
        assert result["shape_name"] == "Chart 1" and result["shape_index"] == 3
        assert result["chart_type_int"] == type_int
        assert "warning" not in result

    def test_without_python_pptx_addchart2_is_used(self, monkeypatch):
        charts, calls = self.deck(monkeypatch)
        monkeypatch.setattr(charts.chart_file, "available", lambda: False)
        monkeypatch.setattr(charts.held, "wait_until_ended", lambda: calls.append("waited"))
        with pytest.raises(AssertionError, match="starts an Excel"):
            charts._add_chart_impl(1, "column", 0, 0, 100, 100)
        assert calls[0] == "waited" and calls[1][0] == "AddChart2"

    def test_clipboard_formats_that_were_lost_are_said(self, monkeypatch):
        from types import SimpleNamespace

        charts, _ = self.deck(monkeypatch)
        monkeypatch.setattr(charts, "_add_chart_from_file", lambda *a: (
            SimpleNamespace(Name="Chart 1", ZOrderPosition=1), ["CF_OWNERDISPLAY"]))
        result = charts._add_chart_impl(1, "column", 0, 0, 100, 100)
        assert "CF_OWNERDISPLAY" in result["warning"]

    def test_three_d_waits_for_the_last_excel_and_uses_addchart2(self, monkeypatch):
        from types import SimpleNamespace

        charts, calls = self.deck(monkeypatch)
        monkeypatch.setattr(charts, "held", SimpleNamespace(
            wait_until_ended=lambda: calls.append("waited"),
            hold=lambda objects, pid: None))
        with pytest.raises(AssertionError, match="starts an Excel"):
            charts._add_chart_impl(1, "3d_column", 0, 0, 100, 100)
        assert calls[0] == "waited" and calls[1][0] == "AddChart2"


@pytest.mark.skipif(sys.platform != "win32",
                    reason="drives the COM impls, which macOS replaces with its own")
class TestTheChartIsBroughtInFromItsFile:
    def run(self, monkeypatch, paste_fails=False, names=("Title 1",), seen=None):
        import contextlib
        import os
        from types import SimpleNamespace

        from ppt_com import charts

        seen = {} if seen is None else seen
        seen.update(closed=False, written=None, kept=False)
        chart = SimpleNamespace(ChartStyle=None, cleared=False)
        chart.ClearToMatchStyle = lambda: setattr(chart, "cleared", True)
        pasted = SimpleNamespace(Name="Chart 1", Chart=chart, Left=0, Top=0, Width=0, Height=0)
        shapes_on_slide = [SimpleNamespace(Name=n) for n in names] + [pasted]

        def paste():
            if paste_fails:
                raise RuntimeError("nothing to paste")
            return lambda i: pasted

        slide = SimpleNamespace()
        slide.Shapes = type("Shapes", (), {
            "Paste": staticmethod(paste),
            "Count": len(shapes_on_slide),
            "__call__": lambda self, i: shapes_on_slide[i - 1],
        })()
        source_shape = SimpleNamespace(Copy=lambda: None)
        source = SimpleNamespace(
            Slides=lambda i: SimpleNamespace(Shapes=lambda j: source_shape),
            Close=lambda: seen.__setitem__("closed", True))

        def open_(path, *a):
            seen["opened"] = (path, a)
            assert os.path.exists(path)
            return source

        app = SimpleNamespace(Presentations=SimpleNamespace(Open=open_))
        pres = SimpleNamespace(PageSetup=SimpleNamespace(SlideWidth=960, SlideHeight=540))

        def write(path, *args):
            seen["written"] = args

        @contextlib.contextmanager
        def kept():
            seen["kept"] = True
            yield ["CF_PALETTE"]

        monkeypatch.setattr(charts.chart_file, "write", write)
        monkeypatch.setattr(charts, "kept_clipboard", kept)
        monkeypatch.setattr(charts, "_wait_pumping", lambda s: None)
        monkeypatch.setattr(charts, "_office_language", lambda app: "ja")
        result = charts._add_chart_from_file(app, pres, slide, 51, 10, 20, 300, 200)
        return result, seen, pasted, chart

    def test_it_lands_where_asked_in_its_style(self, monkeypatch):
        (shape, lost), seen, pasted, chart = self.run(monkeypatch)
        assert shape is pasted and lost == ["CF_PALETTE"]
        assert (pasted.Left, pasted.Top, pasted.Width, pasted.Height) == (10, 20, 300, 200)
        assert chart.ChartStyle == 201 and chart.cleared
        assert seen["written"] == (51, 960, 540, 10, 20, 300, 200, "ja")
        assert seen["opened"][1] == (-1, 0, 0), "read only, without a window"
        assert seen["closed"] and seen["kept"]
        assert not Path(seen["opened"][0]).exists(), "the file is removed"

    def test_a_name_already_on_the_slide_is_not_repeated(self, monkeypatch):
        (shape, _), *_ = self.run(monkeypatch, names=("Chart 1", "Chart 2"))
        assert shape.Name == "Chart 3"

    def test_a_failed_paste_still_closes_and_removes_the_file(self, monkeypatch):
        seen = {}
        with pytest.raises(RuntimeError, match="Could not paste"):
            self.run(monkeypatch, paste_fails=True, seen=seen)
        assert seen["closed"] and seen["kept"]
        assert not Path(seen["opened"][0]).exists()
