"""Recovering from a COM worker stuck inside a call (issue #199).

A COM call that never returns used to wedge the server for good.  Every later
request queued behind the stuck worker and timed out, and only restarting the
client made PowerPoint usable again.

The worker cannot be reclaimed, because an outgoing COM call is not
interruptible from this side.  So it is abandoned instead: the generation is
bumped, a fresh STA worker takes over a fresh queue, and everything the stuck
worker had not started is cancelled rather than applied minutes later.  While
the replacement is still proving that PowerPoint answers, calls fail fast with
an actionable message instead of each costing another full timeout.

No COM and no PowerPoint required.
"""

from __future__ import annotations

import sys
import threading
from concurrent.futures import Future, TimeoutError as FuturesTimeoutError
from pathlib import Path
from queue import Queue
from unittest.mock import MagicMock, patch

import pytest
import pywintypes

_src_dir = str(Path(__file__).resolve().parents[1] / "src")
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)

from utils.com_wrapper import PowerPointCOMWrapper  # noqa: E402

# RPC_E_CALL_REJECTED — PowerPoint answering "busy", most likely a dialog.
BUSY = -2147418111
# MK_E_UNAVAILABLE — PowerPoint is not there at all.
NOT_RUNNING = -2147221021


def _com_error(hresult):
    return pywintypes.com_error(hresult, "test", None, None)


def _wrapper(running=True):
    w = PowerPointCOMWrapper()
    w._running = running
    return w


def _queued(future):
    return (MagicMock(), (), {}, future, False)


# --- deciding that the worker is wedged -----------------------------------

def test_execute_replaces_the_worker_when_the_call_can_neither_be_cancelled_nor_read():
    """The discriminator.  A future that will not cancel and is not done means
    the worker is inside the call and has been for the whole timeout."""
    w = _wrapper()

    with patch.object(Future, "result", side_effect=FuturesTimeoutError), \
            patch.object(Future, "cancel", return_value=False), \
            patch.object(Future, "done", return_value=False), \
            patch.object(PowerPointCOMWrapper, "_recover_from_wedge",
                         return_value=True) as recover:
        with pytest.raises(RuntimeError, match="fresh connection"):
            w.execute(MagicMock())

    recover.assert_called_once_with(0)


def test_a_call_that_cancels_cleanly_does_not_replace_the_worker():
    """Everything queued behind the stuck call cancels normally.  Those
    callers must not each start another replacement."""
    w = _wrapper()

    with patch.object(Future, "result", side_effect=FuturesTimeoutError), \
            patch.object(Future, "cancel", return_value=True), \
            patch.object(PowerPointCOMWrapper, "_recover_from_wedge") as recover:
        with pytest.raises(RuntimeError, match="was cancelled"):
            w.execute(MagicMock())

    recover.assert_not_called()


def test_a_call_that_lands_during_the_timeout_race_is_still_returned():
    w = _wrapper()

    with patch.object(Future, "cancel", return_value=False), \
            patch.object(Future, "done", return_value=True), \
            patch.object(Future, "result",
                         side_effect=[FuturesTimeoutError, {"ok": True}]), \
            patch.object(PowerPointCOMWrapper, "_recover_from_wedge") as recover:
        assert w.execute(MagicMock()) == {"ok": True}

    recover.assert_not_called()


# --- the replacement itself -----------------------------------------------

def test_recovery_starts_a_new_worker_on_a_new_queue():
    w = _wrapper()
    old_queue = w._queue

    with patch.object(PowerPointCOMWrapper, "_start_worker") as start:
        assert w._recover_from_wedge(0) is True

    assert w._generation == 1
    assert w._queue is not old_queue, (
        "the abandoned worker keeps its own queue, so the replacement needs a "
        "fresh one or the stuck thread would drain it on unwedging"
    )
    start.assert_called_once_with(w._queue, 1)
    assert not w._healthy.is_set()


def test_only_the_first_caller_of_a_generation_recovers_it():
    """Several callers can time out together on the same wedge."""
    w = _wrapper()

    with patch.object(PowerPointCOMWrapper, "_start_worker"):
        assert w._recover_from_wedge(0) is True
        assert w._recover_from_wedge(0) is False

    assert w._generation == 1


def test_recovery_cancels_the_work_the_stuck_worker_never_started():
    w = _wrapper()
    old_queue = w._queue
    first, second = Future(), Future()
    old_queue.put(_queued(first))
    old_queue.put(_queued(second))

    with patch.object(PowerPointCOMWrapper, "_start_worker"):
        w._recover_from_wedge(0)

    assert first.cancelled() and second.cancelled()
    assert old_queue.get_nowait() is None, (
        "a sentinel must be left behind, or a worker that unwedges blocks "
        "forever in get() on a queue nobody feeds"
    )


def test_no_recovery_once_the_wrapper_is_stopping():
    w = _wrapper(running=False)

    with patch.object(PowerPointCOMWrapper, "_start_worker") as start:
        assert w._recover_from_wedge(0) is False

    start.assert_not_called()


# --- the health gate ------------------------------------------------------

def test_execute_fails_fast_while_the_replacement_is_still_probing():
    w = _wrapper()
    w._healthy.clear()

    with patch("utils.com_wrapper._HEALTH_WAIT", 0.01):
        with pytest.raises(RuntimeError, match="being rebuilt"):
            w.execute(MagicMock())

    assert w._queue.empty(), "nothing should be queued behind a wedge"


def test_the_probe_reattaches_and_reopens_the_gate():
    w = _wrapper()
    w._generation = 1
    w._healthy.clear()
    app = MagicMock()

    with patch("utils.com_wrapper.win32com.client.GetActiveObject",
               return_value=app):
        w._reattach_after_wedge(1)

    assert w._app is app
    assert w._healthy.is_set()


def test_the_probe_keeps_the_gate_shut_while_powerpoint_is_still_busy():
    """A busy rejection most likely means the dialog that caused the wedge is
    still up.  Reporting health then would turn the next call into a
    misleading "partially applied" error for work that never started."""
    w = _wrapper()
    w._generation = 1
    w._healthy.clear()
    app = MagicMock()

    with patch("utils.com_wrapper.win32com.client.GetActiveObject",
               side_effect=[_com_error(BUSY), _com_error(BUSY), app]), \
            patch("utils.com_wrapper.time.sleep") as sleep:
        w._reattach_after_wedge(1)

    assert sleep.call_count == 2
    assert w._app is app
    assert w._healthy.is_set()


def test_the_probe_gives_up_on_anything_that_is_not_busy():
    """PowerPoint being absent is not a reason to stay shut.  _connect_impl
    relaunches it the way it always has."""
    w = _wrapper()
    w._generation = 1
    w._healthy.clear()

    with patch("utils.com_wrapper.win32com.client.GetActiveObject",
               side_effect=_com_error(NOT_RUNNING)):
        w._reattach_after_wedge(1)

    assert w._app is None
    assert w._healthy.is_set()


def test_the_probe_drops_the_proxy_the_wedged_worker_held():
    """It belongs to an apartment this thread cannot call into."""
    w = _wrapper()
    w._generation = 1
    w._app = MagicMock(name="proxy from the dead apartment")

    with patch("utils.com_wrapper.win32com.client.GetActiveObject",
               side_effect=_com_error(NOT_RUNNING)):
        w._reattach_after_wedge(1)

    assert w._app is None


# --- what the abandoned worker may and may not do -------------------------

def test_an_abandoned_worker_leaves_the_live_connection_alone():
    """Its finally clause must not clear the app its replacement is using."""
    w = _wrapper()
    live_app = MagicMock()
    w._app = live_app
    w._generation = 1  # this worker is generation 0, so it has been replaced

    with patch("utils.com_wrapper.pythoncom"):
        w._com_worker(Queue(), 0)

    assert w._app is live_app


def test_an_abandoned_worker_drops_an_item_that_arrives_late():
    """Recovery drains the old queue, but the worker may pick something up in
    the gap.  It must be cancelled, not run."""
    w = _wrapper()
    future: Future = Future()

    class BumpOnGet(Queue):
        def get(self, *args, **kwargs):
            item = super().get(*args, **kwargs)
            w._generation += 1  # abandoned while this item sat in the queue
            return item

    queue = BumpOnGet()
    queue.put(_queued(future))

    with patch("utils.com_wrapper.pythoncom"), \
            patch.object(PowerPointCOMWrapper, "_run_item") as run_item:
        w._com_worker(queue, 0)

    run_item.assert_not_called()
    assert future.cancelled()


# --- end to end, with real threads ----------------------------------------

def test_the_next_call_is_served_after_a_wedge():
    """The whole point of the exercise.  One stuck call used to end the
    session; now the operation after it goes through."""
    w = _wrapper(running=False)
    entered = threading.Event()
    release = threading.Event()

    def stuck():
        entered.set()
        release.wait(10)

    app = MagicMock()

    with patch("utils.com_wrapper.pythoncom"), \
            patch("utils.com_wrapper.win32com.client.GetActiveObject",
                  return_value=app), \
            patch("utils.com_wrapper._BUSY_WAIT_BUDGET", 0.05), \
            patch("utils.com_wrapper._CALL_TIMEOUT", 0.25):
        w.start()
        try:
            with pytest.raises(RuntimeError, match="fresh connection"):
                w.execute(stuck)

            assert entered.is_set(), "the worker really was inside the call"
            assert w._generation == 1
            assert w.execute(lambda: "served") == "served"
        finally:
            release.set()
            w.stop()
