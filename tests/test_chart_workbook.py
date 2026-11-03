"""ppt_com.charts._close_chart_workbook: a chart's data workbook is closed, or
the failure is reported, never left open in silence.

A workbook left open keeps the chart in edit mode, and the next chart
PowerPoint is asked to insert or edit answers with a dialog, after which every
call times out. Runs anywhere, against stand-ins for the chart and workbook.
"""

import sys

sys.path.insert(0, "src")

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
    def __init__(self, workbook, gone=False):
        self.workbook, self.gone, self.asked = workbook, gone, 0

    @property
    def Workbook(self):
        self.asked += 1
        if self.gone:
            raise RuntimeError("The chart's data is not open")
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


def test_a_workbook_that_cannot_be_had_any_more_counts_as_closed():
    workbook = Workbook(failures=1)
    chart = Chart(ChartData(workbook, gone=True))
    assert _close_chart_workbook(chart, workbook, wait=0) is True


def test_a_workbook_that_never_closes_is_reported(caplog):
    workbook = Workbook(failures=99, error=OSError("Excel is busy"))
    chart = Chart(ChartData(workbook))
    with caplog.at_level("WARNING"):
        assert _close_chart_workbook(chart, workbook, attempts=3, wait=0) is False
    assert not workbook.closed
    assert "Excel is busy" in caplog.text
