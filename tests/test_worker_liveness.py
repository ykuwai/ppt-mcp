"""Tests for COM worker liveness and abandoned work (issue #199).

Two guarantees are pinned here:

1. An operation the caller has stopped waiting for is dropped, not applied
   later.  execute() cancels the future when it times out; the worker skips
   any item whose future was cancelled before it started.  Without this, a
   backlog that built up behind a slow call was replayed once the worker
   caught up — applying edits the caller had already been told had failed.
2. The worker thread survives whatever a queued operation throws, and a caller
   fails fast if it ever does die instead of waiting out the full timeout on a
   queue nothing is draining.

No COM and no PowerPoint required.
"""

from __future__ import annotations

import sys
from concurrent.futures import Future, TimeoutError as FuturesTimeoutError
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

_src_dir = str(Path(__file__).resolve().parents[1] / "src")
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)

from utils.com_wrapper import PowerPointCOMWrapper  # noqa: E402


# ---------------------------------------------------------------------------
# Abandoned work is dropped rather than applied late
# ---------------------------------------------------------------------------
def test_cancelled_operation_is_not_run():
    """The core of issue #199: work the caller gave up on must not be applied."""
    w = PowerPointCOMWrapper()
    func = MagicMock()
    future: Future = Future()
    assert future.cancel()

    w._run_item(func, (), {}, future, False)

    func.assert_not_called()


def test_running_operation_can_no_longer_be_cancelled():
    """Once the worker has started an item, cancelling must not skip it —
    otherwise the future would be left unsettled forever."""
    w = PowerPointCOMWrapper()
    func = MagicMock(return_value={"ok": True})
    future: Future = Future()

    w._run_item(func, (), {}, future, False)

    assert func.call_count == 1
    assert not future.cancel(), "a finished future cannot be cancelled"
    assert future.result() == {"ok": True}


def test_execute_cancels_the_future_when_it_times_out():
    w = PowerPointCOMWrapper()

    with patch.object(Future, "result", side_effect=FuturesTimeoutError), \
         patch.object(Future, "cancel", return_value=True) as cancel_mock:
        with pytest.raises(RuntimeError, match="cancelled"):
            w.execute(MagicMock())

    cancel_mock.assert_called_once()


def test_execute_says_so_when_the_operation_could_not_be_cancelled():
    """An operation already in flight cannot be cancelled — COM outgoing calls
    are not interruptible — so the caller is told it may still be applied."""
    w = PowerPointCOMWrapper()

    with patch.object(Future, "result", side_effect=FuturesTimeoutError), \
         patch.object(Future, "cancel", return_value=False):
        with pytest.raises(RuntimeError, match="may still be applied"):
            w.execute(MagicMock())


def test_timeout_does_not_leak_a_bare_futures_timeout():
    """Tools render errors with str(e); a bare TimeoutError says nothing."""
    w = PowerPointCOMWrapper()

    with patch.object(Future, "result", side_effect=FuturesTimeoutError):
        with pytest.raises(RuntimeError) as exc:
            w.execute(MagicMock())

    assert "PowerPoint did not finish" in str(exc.value)


# ---------------------------------------------------------------------------
# The worker thread stays alive, and a dead one is reported immediately
# ---------------------------------------------------------------------------
def test_base_exception_settles_the_future_instead_of_killing_the_worker():
    """A SystemExit escaping the worker loop used to end the thread while
    _running stayed True, wedging every later request behind it."""
    w = PowerPointCOMWrapper()
    w._app = MagicMock()
    func = MagicMock(side_effect=SystemExit("boom"))
    future: Future = Future()

    w._run_item(func, (), {}, future, False)  # must not propagate

    assert future.done()
    with pytest.raises(SystemExit):
        future.result()


def test_execute_fails_fast_when_the_worker_thread_is_dead():
    w = PowerPointCOMWrapper()
    dead = MagicMock()
    dead.is_alive.return_value = False
    w._com_thread = dead

    with pytest.raises(RuntimeError, match="no longer running"):
        w.execute(MagicMock())

    assert w._queue.empty(), "nothing should be queued for a dead worker"


def test_execute_still_queues_while_the_worker_is_alive():
    w = PowerPointCOMWrapper()
    alive = MagicMock()
    alive.is_alive.return_value = True
    w._com_thread = alive
    func = MagicMock()

    with patch.object(Future, "result", return_value=None):
        w.execute(func)

    queued_func, *_ = w._queue.get_nowait()
    assert queued_func is func
