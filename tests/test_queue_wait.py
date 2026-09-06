"""Waiting in the queue is not the operation's own time (issue #192).

execute() used to give one combined budget to the whole wait, starting when
the item was queued. Two things went wrong.

A caller behind other operations was charged for time it spent in line, so a
perfectly healthy call could be reported as failed while the worker had not
even reached it, and the worker then ran it anyway against a deck that had
moved on. On macOS the same shape produced duplicate pictures.

It also made a wedge impossible to identify. A call that used up its budget
queued looks exactly like one the worker is stuck inside, so the recovery
added for #199 would abandon a healthy worker and let a replacement issue
overlapping edits.

So the wait is in two parts. The queue wait belongs to the queue and takes the
operation back if it never starts; the call budget starts only once the worker
picks the item up.

No COM and no PowerPoint required.
"""

from __future__ import annotations

import sys
import threading
import time
from concurrent.futures import Future
from pathlib import Path
from queue import Queue
from unittest.mock import MagicMock, patch

import pytest

_src_dir = str(Path(__file__).resolve().parents[1] / "src")
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)

from utils.com_wrapper import (  # noqa: E402
    _BUSY_WAIT_BUDGET,
    _CALL_TIMEOUT,
    _QUEUE_WAIT,
    PowerPointCOMWrapper,
)


def _wrapper(running=True):
    w = PowerPointCOMWrapper()
    w._running = running
    return w


# --- the budgets themselves -----------------------------------------------

def test_the_queue_wait_outlasts_one_whole_call():
    """Otherwise a caller behind a single slow but healthy operation is taken
    back for no reason."""
    assert _QUEUE_WAIT > _BUSY_WAIT_BUDGET + _CALL_TIMEOUT


def test_the_two_waits_are_separate_budgets():
    w = _wrapper()

    with patch.object(threading.Event, "wait", return_value=True) as waited, \
            patch.object(Future, "result", return_value=None) as result:
        w.execute(MagicMock())

    assert _QUEUE_WAIT in [call.args[0] for call in waited.call_args_list], (
        "the dequeue is waited for with the queue budget"
    )
    assert result.call_args.kwargs["timeout"] == _BUSY_WAIT_BUDGET + _CALL_TIMEOUT, (
        "and the operation still gets its whole budget afterwards"
    )


# --- an operation that never starts ---------------------------------------

def test_an_operation_that_never_starts_is_taken_back():
    """Reported as a failure and cancelled, so the worker skips it rather than
    running it later against a deck that has moved on."""
    w = _wrapper()

    with patch("utils.com_wrapper._QUEUE_WAIT", 0.05), \
            patch.object(PowerPointCOMWrapper, "_recover_from_wedge") as recover:
        with pytest.raises(RuntimeError, match="did not start the operation"):
            w.execute(MagicMock())

    item = w._queue.get_nowait()
    assert item[3].cancelled()
    # A queue that is merely busy is not a wedge.
    recover.assert_not_called()


def test_the_queue_timeout_says_something():
    """concurrent.futures.TimeoutError stringifies to the empty string, which
    tools rendered as "Failed to add picture: " and nothing more."""
    w = _wrapper()

    with patch("utils.com_wrapper._QUEUE_WAIT", 0.05):
        with pytest.raises(RuntimeError) as exc:
            w.execute(MagicMock())

    assert len(str(exc.value)) > 40


def test_a_dequeue_that_wins_the_race_still_gets_its_budget():
    """The wait expired, but the worker had just picked the item up, so the
    cancel fails and the operation keeps its own budget."""
    w = _wrapper()

    with patch("utils.com_wrapper._QUEUE_WAIT", 0.01), \
            patch.object(Future, "cancel", return_value=False), \
            patch.object(Future, "result", return_value={"ok": True}) as result:
        assert w.execute(MagicMock()) == {"ok": True}

    assert result.call_args.kwargs["timeout"] == _BUSY_WAIT_BUDGET + _CALL_TIMEOUT


# --- interaction with the wedge recovery ----------------------------------

def test_time_spent_in_the_queue_is_not_read_as_a_wedge():
    """The regression this guards. With one combined wait, an operation that
    queued behind another for longer than a call budget was reported as a
    wedge, and a healthy worker was abandoned while it kept mutating the
    deck alongside its replacement."""
    w = _wrapper(running=False)

    class SlowQueue(Queue):
        """Stands in for a worker busy with earlier operations."""

        def get(self, *args, **kwargs):
            item = super().get(*args, **kwargs)
            time.sleep(0.4)  # longer than a whole call budget
            return item

    w._queue = SlowQueue()

    with patch("utils.com_wrapper.pythoncom"), \
            patch("utils.com_wrapper._BUSY_WAIT_BUDGET", 0.05), \
            patch("utils.com_wrapper._CALL_TIMEOUT", 0.2), \
            patch("utils.com_wrapper._QUEUE_WAIT", 5.0), \
            patch.object(PowerPointCOMWrapper, "_recover_from_wedge") as recover:
        w.start()
        try:
            assert w.execute(lambda: "served") == "served"
        finally:
            w.stop()

    recover.assert_not_called()


def test_recovery_releases_a_caller_still_waiting_in_the_queue():
    w = _wrapper()
    future: Future = Future()
    dequeued = threading.Event()
    w._queue.put((MagicMock(), (), {}, future, False, dequeued))

    with patch.object(PowerPointCOMWrapper, "_start_worker"):
        w._recover_from_wedge(0)

    assert future.cancelled()
    assert dequeued.is_set(), (
        "otherwise the caller waits out the whole queue budget for an answer "
        "that already exists"
    )


# --- picking the queue and the generation together ------------------------

class _RecordingLock:
    """A lock that says whether it is currently held."""

    def __init__(self):
        self._lock = threading.Lock()
        self.held = False

    def __enter__(self):
        self._lock.acquire()
        self.held = True
        return self

    def __exit__(self, *exc):
        self.held = False
        self._lock.release()


def test_the_queue_is_chosen_under_the_recovery_lock():
    """Recovery swaps the queue and bumps the generation together. Reading one
    and putting on the other would leave the item behind the sentinel on a
    queue no worker serves."""
    w = _wrapper()
    lock = _RecordingLock()
    w._recover_lock = lock

    class CheckingQueue(Queue):
        def put(self, item, *args, **kwargs):
            assert lock.held, (
                "a recovery could swap the queue between the read and the put"
            )
            super().put(item, *args, **kwargs)

    w._queue = CheckingQueue()

    with patch.object(threading.Event, "wait", return_value=True), \
            patch.object(Future, "result", return_value=None):
        w.execute(MagicMock())


def test_a_recovery_starting_while_the_gate_is_read_is_reported():
    """The health gate is checked again under the lock, so a call cannot slip
    onto a queue whose worker has just been abandoned."""
    w = _wrapper()
    w._healthy.clear()

    with patch.object(threading.Event, "wait", return_value=True):
        with pytest.raises(RuntimeError, match="being rebuilt"):
            w.execute(MagicMock())

    assert w._queue.empty()
