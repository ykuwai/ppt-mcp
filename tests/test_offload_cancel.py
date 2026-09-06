"""Cancelling a request must drop the COM work it queued (issues #198, #199).

anyio waits for a worker thread by default, which would leave a cancelled
request neither returning promptly nor stopping the edit it asked for. The edit
would land minutes later with nobody to report it to, the exact failure #199
removed for timeouts.

run_offloaded abandons the await and cancels the futures the call had queued,
so anything the COM worker has not started yet is skipped. Work already in
flight cannot be recalled, which these tests also pin.

No COM and no PowerPoint required.
"""

from __future__ import annotations

import sys
import threading
from concurrent.futures import Future
from pathlib import Path

import anyio
import pytest

_src_dir = str(Path(__file__).resolve().parents[1] / "src")
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)

from utils.com_wrapper import pending_com_futures  # noqa: E402
from utils.offload import run_offloaded  # noqa: E402


def test_returns_the_value_from_the_worker_thread():
    async def main():
        return await run_offloaded(lambda a, b: a + b, 2, 3)

    assert anyio.run(main) == 5


def test_exceptions_propagate_from_the_worker_thread():
    def boom():
        raise ValueError("from the thread")

    async def main():
        return await run_offloaded(boom)

    with pytest.raises(ValueError, match="from the thread"):
        anyio.run(main)


def test_cancellation_cancels_the_queued_com_work():
    """The point of the exercise. A future queued but not started is dropped."""
    queued: Future = Future()
    entered = threading.Event()
    release = threading.Event()

    def blocking():
        # Stand in for ppt.execute: register the future, then block on it.
        pending_com_futures.get().append(queued)
        entered.set()
        release.wait(5)
        return "finished"

    async def main():
        async with anyio.create_task_group() as tg:
            tg.start_soon(run_offloaded, blocking)
            await anyio.to_thread.run_sync(entered.wait, 5)
            tg.cancel_scope.cancel()

    anyio.run(main)

    assert queued.cancelled(), "queued COM work must be cancelled with the request"
    release.set()


def test_cancellation_leaves_work_already_started_alone():
    """A COM call in flight cannot be recalled, and pretending otherwise would
    be worse than saying so. cancel() simply fails on a running future."""
    started: Future = Future()
    assert started.set_running_or_notify_cancel()
    entered = threading.Event()
    release = threading.Event()

    def blocking():
        pending_com_futures.get().append(started)
        entered.set()
        release.wait(5)

    async def main():
        async with anyio.create_task_group() as tg:
            tg.start_soon(run_offloaded, blocking)
            await anyio.to_thread.run_sync(entered.wait, 5)
            tg.cancel_scope.cancel()

    anyio.run(main)

    assert not started.cancelled()
    release.set()


def test_the_registry_is_cleared_afterwards():
    """A leaked context var would let one request cancel another's work."""
    async def main():
        await run_offloaded(lambda: None)
        return pending_com_futures.get()

    assert anyio.run(main) is None
