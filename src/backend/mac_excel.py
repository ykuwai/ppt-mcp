"""The Excel a chart's data brings up on the Mac, put out of the way again.

Through the VBA bridge, ``ppt_set_chart_data`` and the other chart tools reach
the chart's data through ``ChartData.Activate``, which on the Mac starts
Excel, or wakes the one running, and brings it to the front: the user's own
window lost the front on every call. Closing the chart's workbook does not
quit an Excel PowerPoint started either, so one was left running with no
workbook at all (measured 2026-11-08, PowerPoint and Excel 16.114).

So around each chart call, ``out_of_the_way`` gives the front back, but only
when the call opened a chart's workbook (``ppt_com.charts`` notes that on the
``Call``): a call that never touched Excel leaves it alone, whatever the user
did with Excel meanwhile. After such a call, an Excel that took the front is
hidden, unless it was in front already. And an Excel that was not running
before the call is quit once the call has closed the chart's workbook, and
waited out, so the next chart call does not reach an Excel that is shutting
down. It is not quit when it shows a window of its own: the chart's workbook
is closed by then, so that window is one the user opened in it while the call
ran. An Excel the user started is never quit, and neither is one the call
failed in or left a workbook open in: its window is what the tool's message
tells the user to close.

None of this takes an Apple Event, so none of it asks the user for permission
to control Excel: ``NSRunningApplication`` hides and quits it, and the window
list says whether it shows a window.
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


def _shows_a_window(pid: int) -> bool:
    """Whether process ``pid`` has a document window on screen."""
    from Quartz import (CGWindowListCopyWindowInfo, kCGNullWindowID,
                        kCGWindowListOptionOnScreenOnly)

    windows = CGWindowListCopyWindowInfo(kCGWindowListOptionOnScreenOnly, kCGNullWindowID) or []
    return any(w.get("kCGWindowOwnerPID") == pid and w.get("kCGWindowLayer") == 0 for w in windows)


class Call:
    """What the chart call inside ``out_of_the_way`` did with Excel."""

    def __init__(self):
        # A chart's workbook was opened in Excel.
        self.opened = False
        # One could not be closed again.
        self.left_open = False


@contextlib.contextmanager
def out_of_the_way(running=None, shows_a_window=None, pause=None, clock=None,
                   limit: float = QUIT_WAIT):
    """Around a chart call: put the Excel it brought up away again after."""
    running = running or _running
    shows_a_window = shows_a_window or _shows_a_window
    pause = pause or time.sleep
    clock = clock or time.monotonic
    call = Call()
    try:
        before = running()
        pid_before: Optional[int] = before.processIdentifier() if before is not None else None
        was_front = bool(before is not None and before.isActive())
    except Exception:  # noqa: BLE001 - never in the way of the call itself
        logger.warning("Could not see whether Excel is running before a chart call", exc_info=True)
        yield call
        return
    failed = True
    try:
        yield call
        failed = False
    finally:
        try:
            if call.opened:
                quit_it = pid_before is None and not call.left_open and not failed
                _put_away(running, shows_a_window, quit_it, was_front, pause, clock, limit)
        except Exception:  # noqa: BLE001 - never in the way of the call's own answer
            logger.warning("Could not put Excel out of the way after a chart call", exc_info=True)


def _put_away(running, shows_a_window, quit_it, was_front, pause, clock, limit) -> None:
    excel = running()
    if excel is None:
        return
    pid = excel.processIdentifier()
    if quit_it and not shows_a_window(pid):
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
