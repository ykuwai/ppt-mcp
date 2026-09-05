"""Tests for busy-handling policy in the COM worker (issue #200).

The rule under test: PowerPoint being busy is waited out *before* an operation
starts, and an operation rejected part-way through is NOT re-run — re-running
it would repeat whatever it already applied.  Only callers that opt in with
idempotent=True are retried wholesale.

No COM and no PowerPoint required: the app object is a mock and time.sleep is
patched out so the backoff costs nothing.
"""

from __future__ import annotations

import sys
from concurrent.futures import Future
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
import pywintypes

_src_dir = str(Path(__file__).resolve().parents[1] / "src")
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)

from utils.com_wrapper import PowerPointCOMWrapper  # noqa: E402

# RPC_E_CALL_REJECTED — PowerPoint answering "busy", e.g. a modal dialog is up.
BUSY = -2147418111
# MK_E_UNAVAILABLE — deliberately NOT a busy HRESULT.
NOT_BUSY = -2147221021


def _com_error(hresult):
    return pywintypes.com_error(hresult, "test", None, None)


def _wrapper(name_side_effect=None):
    """Wrapper with a mock app whose `.Name` probe can be scripted."""
    w = PowerPointCOMWrapper()
    app = MagicMock()
    if name_side_effect is not None:
        type(app).Name = property(lambda self: _next(name_side_effect))
    w._app = app
    return w


def _next(script):
    """Pop the next scripted probe result; raise it if it is an exception."""
    value = script.pop(0) if script else "PowerPoint"
    if isinstance(value, BaseException):
        raise value
    return value


def _run(w, func, idempotent=False):
    future: Future = Future()
    with patch("utils.com_wrapper.time.sleep") as sleep_mock:
        w._run_item(func, (), {}, future, idempotent)
    return future, sleep_mock


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


def test_idempotent_operation_is_retried_wholesale():
    w = _wrapper()
    func = MagicMock(side_effect=[_com_error(BUSY), {"ok": True}])

    future, _ = _run(w, func, idempotent=True)

    assert func.call_count == 2
    assert future.result() == {"ok": True}


def test_idempotent_operation_still_busy_on_retry_reports_clearly():
    w = _wrapper()
    func = MagicMock(side_effect=_com_error(BUSY))

    future, _ = _run(w, func, idempotent=True)

    assert func.call_count == 2
    with pytest.raises(RuntimeError, match="not responding"):
        future.result()


# ---------------------------------------------------------------------------
# Waiting happens before the operation starts
# ---------------------------------------------------------------------------
def test_busy_is_waited_out_before_the_operation_runs():
    """Two busy probes, then responsive — func runs exactly once."""
    w = _wrapper([_com_error(BUSY), _com_error(BUSY), "PowerPoint"])
    func = MagicMock(return_value={"ok": True})

    future, sleep_mock = _run(w, func)

    assert func.call_count == 1
    assert sleep_mock.call_count == 2, "should back off once per busy probe"
    assert future.result() == {"ok": True}


def test_operation_is_not_started_when_the_wait_budget_runs_out():
    w = _wrapper()
    type(w._app).Name = property(lambda self: (_ for _ in ()).throw(_com_error(BUSY)))
    func = MagicMock()

    # Budget is consumed by advancing the clock instead of really sleeping.
    clock = iter([0.0] + [100.0] * 20)
    with patch("utils.com_wrapper.time.sleep"), \
         patch("utils.com_wrapper.time.monotonic", side_effect=lambda: next(clock)):
        future: Future = Future()
        w._run_item(func, (), {}, future, False)

    func.assert_not_called()
    with pytest.raises(RuntimeError, match="did not respond"):
        future.result()


def test_non_busy_probe_error_does_not_block_the_operation():
    """A stale reference must reach _get_app_impl, not be swallowed here."""
    w = _wrapper([_com_error(NOT_BUSY)])
    func = MagicMock(return_value={"ok": True})

    future, sleep_mock = _run(w, func)

    assert func.call_count == 1
    sleep_mock.assert_not_called()
    assert future.result() == {"ok": True}


def test_probe_is_skipped_when_not_connected_yet():
    w = PowerPointCOMWrapper()  # _app is None
    func = MagicMock(return_value={"ok": True})

    future, sleep_mock = _run(w, func)

    assert func.call_count == 1
    sleep_mock.assert_not_called()
    assert future.result() == {"ok": True}


# ---------------------------------------------------------------------------
# Non-busy failures are passed through untouched
# ---------------------------------------------------------------------------
def test_non_busy_com_error_propagates_unchanged():
    w = _wrapper()
    err = _com_error(NOT_BUSY)
    func = MagicMock(side_effect=err)

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
# execute() queues the idempotent flag and keeps it away from func
# ---------------------------------------------------------------------------
def test_execute_queues_the_idempotent_flag():
    w = PowerPointCOMWrapper()
    func = MagicMock()

    # Don't start a COM thread — inspect what execute() put on the queue.
    with patch.object(Future, "result", return_value=None):
        w.execute(func, 1, 2, idempotent=True, keyword="x")

    queued_func, args, kwargs, _future, idempotent = w._queue.get_nowait()
    assert queued_func is func
    assert args == (1, 2)
    assert kwargs == {"keyword": "x"}, "idempotent must not leak into func kwargs"
    assert idempotent is True


def test_execute_defaults_to_non_idempotent():
    w = PowerPointCOMWrapper()
    with patch.object(Future, "result", return_value=None):
        w.execute(MagicMock())
    *_, idempotent = w._queue.get_nowait()
    assert idempotent is False
