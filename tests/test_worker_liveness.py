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
import pywintypes

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


def test_operation_stays_cancellable_through_the_preflight_wait():
    """A caller that times out while the worker is still waiting for a busy
    PowerPoint must be able to cancel: nothing has been applied yet."""
    w = PowerPointCOMWrapper()
    app = MagicMock()
    busy = pywintypes.com_error(-2147418111, "busy", None, None)   # RPC_E_CALL_REJECTED
    type(app).Name = property(lambda _self: (_ for _ in ()).throw(busy))
    w._app = app
    func = MagicMock()
    future: Future = Future()

    cancelled_during_wait = []
    clock = {"now": 0.0}

    def _sleep(seconds):
        # Stand in for the caller giving up mid-wait; advance a fake clock so
        # the wait budget still bounds the loop without real waiting.
        clock["now"] += seconds
        if not cancelled_during_wait:
            cancelled_during_wait.append(future.cancel())

    with patch("utils.com_wrapper.time.sleep", _sleep), \
         patch("utils.com_wrapper.time.monotonic", lambda: clock["now"]):
        w._run_item(func, (), {}, future, False)

    assert cancelled_during_wait == [True], "must still be cancellable while waiting"
    # A cancelled operation must never be applied.
    func.assert_not_called()


def test_timeout_error_raised_by_the_operation_is_not_mistaken_for_ours():
    """concurrent.futures.TimeoutError *is* the builtin TimeoutError since
    3.11, so an operation raising it must not be reported as our 45s wait."""
    w = PowerPointCOMWrapper()
    w._app = MagicMock()
    future: Future = Future()
    w._run_item(MagicMock(side_effect=TimeoutError("COM call timed out")),
                (), {}, future, False)

    with patch.object(Future, "result", side_effect=future.exception()), \
         patch.object(Future, "done", return_value=True), \
         patch.object(Future, "cancel") as cancel_mock:
        with pytest.raises(TimeoutError, match="COM call timed out"):
            w.execute(MagicMock())

    # A finished future must not be cancelled.
    cancel_mock.assert_not_called()


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

    with patch.object(Future, "result", side_effect=FuturesTimeoutError), \
         patch.object(Future, "done", return_value=False):
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
    with pytest.raises(RuntimeError, match="SystemExit"):
        future.result()


def test_base_exception_does_not_reach_the_calling_thread_unchanged():
    """Keeping the worker alive must not cost the caller its process.

    Tools wrap execute() in `except Exception`, which does not catch
    SystemExit or KeyboardInterrupt — those would sail past the tool, past the
    server's dispatch, and take the event loop down. They are reported as a
    RuntimeError instead.
    """
    w = PowerPointCOMWrapper()
    w._app = MagicMock()

    for raised in (SystemExit("bye"), KeyboardInterrupt()):
        future: Future = Future()
        w._run_item(MagicMock(side_effect=raised), (), {}, future, False)
        with pytest.raises(Exception) as exc:   # must be catchable as Exception
            future.result()
        assert isinstance(exc.value, RuntimeError)
        assert type(raised).__name__ in str(exc.value)


def test_ordinary_exception_is_passed_through_unwrapped():
    """Only non-Exception BaseExceptions get rewritten."""
    w = PowerPointCOMWrapper()
    w._app = MagicMock()
    err = ValueError("bad slide index")
    future: Future = Future()

    w._run_item(MagicMock(side_effect=err), (), {}, future, False)

    with pytest.raises(ValueError, match="bad slide index"):
        future.result()


def test_worker_loop_survives_an_item_that_escapes_run_item():
    """The backstop in _com_worker: if _run_item itself blows up, the future
    is still settled and the loop keeps going."""
    w = PowerPointCOMWrapper()
    w._running = True
    future: Future = Future()
    w._queue.put((MagicMock(), (), {}, future, False))
    w._queue.put(None)  # sentinel so the loop exits after the bad item

    with patch("utils.com_wrapper.pythoncom"):
        with patch.object(PowerPointCOMWrapper, "_run_item",
                          side_effect=RuntimeError("escaped")):
            w._com_worker()

    assert future.done(), "the caller must not be left waiting forever"
    with pytest.raises(RuntimeError, match="COM worker failed"):
        future.result()


def test_execute_returns_the_result_if_it_lands_during_the_timeout_race():
    """cancel() failing because the operation just finished is not a failure."""
    w = PowerPointCOMWrapper()

    with patch.object(Future, "cancel", return_value=False):
        with patch.object(Future, "done", return_value=True):
            with patch.object(
                Future, "result",
                side_effect=[FuturesTimeoutError, {"ok": True}],
            ):
                assert w.execute(MagicMock()) == {"ok": True}


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
