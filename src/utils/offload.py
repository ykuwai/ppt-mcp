"""Run a blocking tool call in a worker thread without blocking the loop.

Every MCP tool handler is `async def`, and everything underneath it is
synchronous COM work that blocks until the STA worker answers.  Called inline
it stops the event loop outright, so the server cannot answer a ping, a
cancellation, or even tools/list while one COM call is in flight (issue #198).

run_offloaded() hands the call to a worker thread and, crucially, keeps it
cancellable.  anyio's default is to wait for the thread, which would leave a
cancelled request neither returning promptly nor dropping the edit it asked
for.  Instead the await is abandoned on cancellation and every COM operation
the call had queued is cancelled, so anything the worker has not started yet is
skipped rather than applied after the client gave up (issue #199).

A COM call already in flight cannot be stopped.  Outgoing COM calls are not
interruptible from the client side, so an operation that has begun will finish;
what cancellation buys is that the queue behind it does not run.
"""

from typing import Any, Callable, TypeVar

import anyio

from utils.com_wrapper import QueuedCalls, pending_com_futures

T = TypeVar("T")


async def run_offloaded(func: Callable[..., T], *args: Any) -> T:
    """Await `func(*args)` on a worker thread, cancelling its COM work if the
    caller goes away.

    Args:
        func: A blocking callable, normally a tool's synchronous wrapper or
            ppt.execute itself.
        *args: Positional arguments for func.  anyio's run_sync takes no
            keyword arguments, and no handler needs any.

    Returns:
        Whatever func returns.
    """
    queued = QueuedCalls()
    token = pending_com_futures.set(queued)
    try:
        return await anyio.to_thread.run_sync(func, *args, abandon_on_cancel=True)
    except anyio.get_cancelled_exc_class():
        # The thread keeps running, so drop whatever it queued but has not
        # started, and anything it queues from here on.  Cancelling a future
        # the worker already picked up is a no-op, which is the honest outcome
        # for work COM cannot recall.
        queued.cancel_all()
        raise
    finally:
        pending_com_futures.reset(token)
