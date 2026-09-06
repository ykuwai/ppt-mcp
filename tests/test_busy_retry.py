"""Tests for busy-handling policy in the COM worker (issue #200).

The rule under test: PowerPoint being busy is waited out *before* an operation
starts, and an operation rejected part-way through is NOT re-run — re-running
it would repeat whatever it already applied.  Only callers that opt in with
idempotent=True are retried wholesale, and all waiting for one operation shares
a single budget.

No COM and no PowerPoint required: the app object is a mock, and a fake clock
replaces time.sleep/time.monotonic so the backoff costs no wall-clock time
while the budget still bounds the loops exactly as it would in production.
"""

from __future__ import annotations

import sys
import threading
from concurrent.futures import Future
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
import pywintypes

_src_dir = str(Path(__file__).resolve().parents[1] / "src")
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)

from utils.com_wrapper import (  # noqa: E402
    _BUSY_WAIT_BUDGET,
    _CALL_TIMEOUT,
    PowerPointCOMWrapper,
)

# RPC_E_CALL_REJECTED — PowerPoint answering "busy", e.g. a modal dialog is up.
BUSY = -2147418111
# MK_E_UNAVAILABLE — deliberately NOT a busy HRESULT.
NOT_BUSY = -2147221021


def _com_error(hresult):
    return pywintypes.com_error(hresult, "test", None, None)


class FakeClock:
    """Sleeping advances time, so budgets bound loops without real waiting."""

    def __init__(self):
        self.now = 0.0
        self.sleeps = []

    def monotonic(self):
        return self.now

    def sleep(self, seconds):
        self.sleeps.append(seconds)
        self.now += seconds

    @property
    def slept(self):
        return sum(self.sleeps)


def _wrapper(probe_script=None):
    """Wrapper with a mock app whose `.Name` probe can be scripted.

    probe_script entries are returned in order; an exception entry is raised.
    Once exhausted the probe succeeds.  Pass None for an unscripted app.
    """
    w = PowerPointCOMWrapper()
    app = MagicMock()
    if probe_script is not None:
        script = list(probe_script)

        def _name(_self):
            value = script.pop(0) if script else "PowerPoint"
            if isinstance(value, BaseException):
                raise value
            return value

        type(app).Name = property(_name)
    w._app = app
    return w


def _always_busy_wrapper():
    w = PowerPointCOMWrapper()
    app = MagicMock()
    type(app).Name = property(
        lambda _self: (_ for _ in ()).throw(_com_error(BUSY))
    )
    w._app = app
    return w


def _run(w, func, idempotent=False):
    clock = FakeClock()
    future: Future = Future()
    with patch("utils.com_wrapper.time.sleep", clock.sleep), \
         patch("utils.com_wrapper.time.monotonic", clock.monotonic):
        w._run_item(func, (), {}, future, idempotent)
    return future, clock


# ---------------------------------------------------------------------------
# The core of issue #200: a mid-operation rejection must not re-run the op
# ---------------------------------------------------------------------------
def test_mid_operation_busy_does_not_rerun_the_operation():
    w = _wrapper()
    func = MagicMock(side_effect=_com_error(BUSY))

    future, _ = _run(w, func)

    assert func.call_count == 1, "the operation must not be repeated"
    with pytest.raises(RuntimeError, match="partially applied"):
        future.result()

def _execute(w, *args, **kwargs):
    """Call execute() as if the worker had picked the item up at once.

    execute() waits for the dequeue before the call budget starts, so that
    time spent queued behind other operations is not charged to the operation
    (issue #192).  These tests run no worker and are about what happens once
    the operation has started.
    """
    with patch.object(threading.Event, "wait", return_value=True):
        return w.execute(*args, **kwargs)


def test_idempotent_operation_is_retried_wholesale():
    w = _wrapper()
    func = MagicMock(side_effect=[_com_error(BUSY), {"ok": True}])

    future, clock = _run(w, func, idempotent=True)

    assert func.call_count == 2
    assert clock.sleeps, "must back off before re-running"
    assert future.result() == {"ok": True}


def test_idempotent_retry_backs_off_even_when_nothing_is_connected():
    """With _app None there is no object to probe, so the backoff between
    attempts is the only thing giving a modal dialog time to clear."""
    w = PowerPointCOMWrapper()  # _app is None — the probe is skipped entirely
    func = MagicMock(side_effect=[_com_error(BUSY), {"ok": True}])

    future, clock = _run(w, func, idempotent=True)

    assert func.call_count == 2
    assert clock.slept > 0, "back-to-back retries with no delay are useless"
    assert future.result() == {"ok": True}


def test_idempotent_operation_gives_up_within_the_shared_budget():
    w = _wrapper()
    func = MagicMock(side_effect=_com_error(BUSY))

    future, clock = _run(w, func, idempotent=True)

    assert func.call_count > 1, "an idempotent operation should be retried"
    assert clock.slept <= _BUSY_WAIT_BUDGET, "retries must share one budget"
    with pytest.raises(RuntimeError, match="did not respond"):
        future.result()


def test_waiting_and_retrying_share_one_budget():
    """A retry must not start a second full budget on top of the first."""
    w = _always_busy_wrapper()
    func = MagicMock(side_effect=_com_error(BUSY))

    _future, clock = _run(w, func, idempotent=True)

    assert clock.slept <= _BUSY_WAIT_BUDGET


# ---------------------------------------------------------------------------
# Waiting happens before the operation starts
# ---------------------------------------------------------------------------
def test_busy_is_waited_out_before_the_operation_runs():
    """Two busy probes, then responsive — func runs exactly once."""
    w = _wrapper([_com_error(BUSY), _com_error(BUSY), "PowerPoint"])
    func = MagicMock(return_value={"ok": True})

    future, clock = _run(w, func)

    assert func.call_count == 1
    assert len(clock.sleeps) == 2, "should back off once per busy probe"
    assert future.result() == {"ok": True}


def test_operation_is_not_started_when_the_wait_budget_runs_out():
    w = _always_busy_wrapper()
    func = MagicMock()

    future, clock = _run(w, func)

    func.assert_not_called()
    assert clock.slept <= _BUSY_WAIT_BUDGET
    with pytest.raises(RuntimeError, match="did not respond"):
        future.result()


def test_non_busy_probe_error_does_not_block_the_operation():
    """A stale reference must reach _get_app_impl, not be swallowed here."""
    w = _wrapper([_com_error(NOT_BUSY)])
    func = MagicMock(return_value={"ok": True})

    future, clock = _run(w, func)

    assert func.call_count == 1
    assert not clock.sleeps
    assert future.result() == {"ok": True}


def test_probe_is_skipped_when_not_connected_yet():
    w = PowerPointCOMWrapper()  # _app is None
    func = MagicMock(return_value={"ok": True})

    future, clock = _run(w, func)

    assert func.call_count == 1
    assert not clock.sleeps
    assert future.result() == {"ok": True}


# ---------------------------------------------------------------------------
# Non-busy failures are passed through untouched
# ---------------------------------------------------------------------------
def test_non_busy_com_error_propagates_unchanged():
    w = _wrapper()
    func = MagicMock(side_effect=_com_error(NOT_BUSY))

    future, _ = _run(w, func)

    assert func.call_count == 1
    with pytest.raises(pywintypes.com_error) as exc:
        future.result()
    assert exc.value.hresult == NOT_BUSY


def test_ordinary_exception_propagates_unchanged():
    w = _wrapper()
    func = MagicMock(side_effect=ValueError("bad slide index"))

    future, _ = _run(w, func)

    assert func.call_count == 1
    with pytest.raises(ValueError, match="bad slide index"):
        future.result()


# ---------------------------------------------------------------------------
# execute(): flag plumbing and a deadline that covers the wait
# ---------------------------------------------------------------------------
def test_execute_queues_the_idempotent_flag():
    w = PowerPointCOMWrapper()
    func = MagicMock()

    # Don't start a COM thread — inspect what execute() put on the queue.
    with patch.object(Future, "result", return_value=None):
        _execute(w, func, 1, 2, idempotent=True, keyword="x")

    queued_func, args, kwargs, _future, idempotent, _dequeued = \
        w._queue.get_nowait()
    assert queued_func is func
    assert args == (1, 2)
    assert kwargs == {"keyword": "x"}, "idempotent must not leak into func kwargs"
    assert idempotent is True


def test_execute_defaults_to_non_idempotent():
    w = PowerPointCOMWrapper()
    with patch.object(Future, "result", return_value=None):
        _execute(w, MagicMock())
    *_, idempotent, _dequeued = w._queue.get_nowait()
    assert idempotent is False


def test_execute_timeout_covers_the_busy_wait_as_well_as_the_call():
    """Otherwise the caller could abandon the future while the worker, having
    spent its wait budget first, is still applying the edit."""
    w = PowerPointCOMWrapper()
    with patch.object(Future, "result", return_value=None) as result_mock:
        _execute(w, MagicMock())

    timeout = result_mock.call_args.kwargs["timeout"]
    assert timeout == _BUSY_WAIT_BUDGET + _CALL_TIMEOUT
    assert timeout > _BUSY_WAIT_BUDGET
