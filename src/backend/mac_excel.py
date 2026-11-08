"""The Excel a chart's data brings up on the Mac, put out of the way again.

Through the VBA bridge, ``ppt_set_chart_data`` and the other chart tools reach
the chart's data through ``ChartData.Activate``, which on the Mac starts
Excel, or wakes the one running, and brings it to the front: the user's own
window lost the front on every call. Closing the chart's workbook does not
quit an Excel PowerPoint started either, so one was left running with no
workbook at all (measured 2026-11-08, PowerPoint and Excel 16.114).

So around each such call, ``out_of_the_way`` gives the front back: an Excel
that took it is hidden, unless it was in front already. And an Excel that was
not running before the call is quit once the call has finished with the
chart's workbook closed, and waited out, so the next chart call does not reach
an Excel that is shutting down. An Excel the user started is never quit, and
neither is one the call failed in or left a workbook open in: its window is
what the tool's message tells the user to close.

Neither takes an Apple Event, so neither asks the user for permission to
control Excel: ``NSRunningApplication`` hides and quits it.
"""

import contextlib
import logging
import time
from typing import Optional

logger = logging.getLogger(__name__)

EXCEL = "com.microsoft.Excel"
# Excel quits in about a tenth of a second with no workbook open.
QUIT_WAIT = 10


def _running():
    """The running Excel, as an NSRunningApplication, or None."""
    from AppKit import NSRunningApplication

    found = NSRunningApplication.runningApplicationsWithBundleIdentifier_(EXCEL)
    return found[0] if found else None


class Call:
    """What the call inside ``out_of_the_way`` says about how it ended."""

    def __init__(self):
        # Set once the chart's workbook is known to be closed.
        self.workbook_closed = False


@contextlib.contextmanager
def out_of_the_way(running=None, pause=None, clock=None, limit: float = QUIT_WAIT):
    """Around a call that may bring Excel up: put it away again after."""
    running = running or _running
    pause = pause or time.sleep
    clock = clock or time.monotonic
    before = running()
    pid_before: Optional[int] = before.processIdentifier() if before is not None else None
    was_front = bool(before is not None and before.isActive())
    call = Call()
    try:
        yield call
    finally:
        try:
            _put_away(running, pid_before, was_front, call.workbook_closed, pause, clock, limit)
        except Exception:  # noqa: BLE001 - never in the way of the call's own answer
            logger.warning("Could not put Excel out of the way after a chart call", exc_info=True)


def _put_away(running, pid_before, was_front, workbook_closed, pause, clock, limit) -> None:
    excel = running()
    if excel is None:
        return
    if pid_before is None and workbook_closed:
        pid = excel.processIdentifier()
        excel.terminate()
        deadline = clock() + limit
        while True:
            now = running()
            if now is None or now.processIdentifier() != pid:
                return
            if clock() >= deadline:
                logger.warning("Excel (pid %s) was asked to quit after a chart call and is still running "
                               "after %.0f s", pid, limit)
                return
            pause(0.05)
    if not was_front and excel.isActive():
        # hide() answers False here even when it works; whether it did is
        # seen in the front app.
        excel.hide()
