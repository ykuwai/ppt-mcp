"""COM connection lifecycle management for PowerPoint automation.

Handles CoInitialize, Dispatch, GetActiveObject, error recovery, and cleanup.
PowerPoint supports only a single running instance, so this provides
singleton-like access to the Application COM object.
"""

import gc
import logging
import os
import threading
import time
from concurrent.futures import (
    CancelledError as FuturesCancelledError,
    Future,
    TimeoutError as FuturesTimeoutError,
)
from contextvars import ContextVar
from queue import Empty, Queue
from typing import Any, Callable, Optional

import pythoncom
import pywintypes
import win32com.client

logger = logging.getLogger(__name__)

class QueuedCalls:
    """The COM futures one tool call has queued, so they can be cancelled.

    utils.offload puts one of these on pending_com_futures for the duration of
    a request and cancels it if the caller goes away.  A tool wrapper is free
    to call execute() several times, and the worker thread keeps running for a
    moment after cancellation, so registration and cancellation can genuinely
    overlap.  The lock makes that deterministic, and anything queued after the
    cancellation is cancelled on arrival rather than quietly running.
    """

    def __init__(self):
        self._lock = threading.Lock()
        self._futures = []
        self._cancelled = False

    def add(self, future) -> None:
        """Register a queued future, or cancel it if we are already too late."""
        with self._lock:
            if not self._cancelled:
                self._futures.append(future)
                return
        future.cancel()

    def cancel_all(self) -> None:
        """Cancel everything queued so far, and everything queued from now on."""
        with self._lock:
            self._cancelled = True
            futures = list(self._futures)
        for future in futures:
            future.cancel()


# The QueuedCalls for the request running on this context, so that
# utils.offload can cancel its COM work if the caller goes away (#198, #199).
# None means nobody is watching, which is the case for internal callers.
pending_com_futures: ContextVar = ContextVar("pending_com_futures", default=None)

# HRESULTs that indicate PowerPoint is temporarily busy (e.g. modal dialog open).
# RPC_E_CALL_REJECTED (0x80010001): server rejected the call outright.
# RPC_E_SERVERCALL_RETRYLATER (0x8001010A): server explicitly says retry later.
#
# Both mean that *individual COM call* never started, so retrying THAT CALL is
# safe.  It does NOT follow that re-running the whole impl function is safe:
# an impl typically makes many COM calls, and if the rejection lands on the
# fifth one, re-running from the top repeats the first four (issue #200 -- e.g.
# ppt_add_slide(count=5) rejected on slide 3 would create slides 1-2 twice).
#
# So the worker waits for PowerPoint to become responsive BEFORE starting the
# operation, and does not re-run an operation that was rejected part-way
# through.  Only callers that pass idempotent=True are retried wholesale.
_BUSY_HRESULTS = frozenset({-2147418111, -2147417846})
# Total budget for waiting out a busy PowerPoint, in seconds.  Deliberately
# below the 30 s timeout in execute(): with the old fixed 5 x 3 s retries the
# worker could still be working after the caller had already given up, and
# then apply the operation the caller was told had failed.
_BUSY_WAIT_BUDGET = 15.0
# Backoff between busy probes.  The last value repeats until the budget runs out.
_BUSY_BACKOFF = (0.5, 1.0, 2.0, 3.0)
# How long the operation itself may take, once PowerPoint is responsive.
_CALL_TIMEOUT = 30.0
# Shown whenever the wait budget runs out, wherever in the wait it happens.
_BUSY_TIMEOUT_MESSAGE = (
    "PowerPoint did not respond within %.0fs. A dialog or menu is probably "
    "open. Close it and retry." % _BUSY_WAIT_BUDGET
)
# How long execute() gives a replacement worker to prove that PowerPoint
# answers again before it reports the connection as still being rebuilt.
# Short on purpose: a recovery that goes well should be invisible, and one
# that does not should not cost every later call another full timeout.
_HEALTH_WAIT = 2.0
_REBUILDING_MESSAGE = (
    "PowerPoint stopped responding and the connection is being rebuilt. "
    "Close any dialog open in PowerPoint, then retry."
)
# How long a caller waits for the worker to pick its operation up.  Time spent
# in the queue is not the operation's own time, and charging it to the call
# budget would misreport a healthy call that simply queued behind another one
# (issue #192).  It also matters for #199: a call that spends its whole budget
# queued cannot be told apart from a call the worker is stuck inside, and
# treating it as a wedge would abandon a perfectly good worker and let a
# replacement issue overlapping edits.
#
# Deliberately longer than one whole call budget, so a caller behind one slow
# but healthy operation is served rather than taken back.
_QUEUE_WAIT = 60.0
_QUEUE_TIMEOUT_MESSAGE = (
    "PowerPoint did not start the operation within %.0fs because earlier "
    "operations were still running. It was cancelled rather than left to run "
    "later." % _QUEUE_WAIT
)
# When True, the server sends ESC to PowerPoint on the first busy rejection to
# dismiss any blocking modal dialog automatically.
# Opt-in: set PPT_AUTO_DISMISS_DIALOG=true in mcp.json env to enable:
#   "env": {"PPT_AUTO_DISMISS_DIALOG": "true"}
AUTO_DISMISS_DIALOG: bool = os.getenv("PPT_AUTO_DISMISS_DIALOG", "false").lower() in ("true", "1", "yes")


def _try_dismiss_ppt_dialog() -> None:
    """Send ESC to the PowerPoint window to dismiss any open modal dialog.

    Called once on the first RPC_E_CALL_REJECTED so the next retry can
    succeed without waiting for the user to notice.  ESC is safe: it cancels
    without committing, so no destructive side-effects occur.

    Implementation notes:
    - Uses win32gui (part of pywin32) to find the PowerPoint main window by
      class name "PPTFrameClass", then SetForegroundWindow + win32api.keybd_event
      to deliver the keystroke reliably without side-effects on keyboard state.
    - All errors are swallowed — this is best-effort only.
    """
    try:
        import win32api   # part of pywin32, already a project dependency
        import win32con
        import win32gui
        hwnd = win32gui.FindWindow("PPTFrameClass", None)
        if not hwnd:
            logger.debug("_try_dismiss_ppt_dialog: PPTFrameClass window not found")
            return
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.15)  # brief pause for focus to settle
        # Use keybd_event instead of WScript.Shell.SendKeys — SendKeys resets
        # Num Lock / Caps Lock state before sending, causing spurious Windows
        # accessibility notifications ("Num Lock Off").  keybd_event sends
        # only the ESC key with no side-effects on keyboard toggle state.
        win32api.keybd_event(win32con.VK_ESCAPE, 0, 0, 0)
        win32api.keybd_event(win32con.VK_ESCAPE, 0, win32con.KEYEVENTF_KEYUP, 0)
        logger.info("Sent ESC to PowerPoint to dismiss open dialog")
    except Exception as exc:
        logger.debug("_try_dismiss_ppt_dialog failed (ignored): %s", exc)


def _caller_safe(exc: BaseException) -> BaseException:
    """Return an exception the calling thread can safely receive.

    The worker settles a future with whatever the operation raised, and the
    caller gets it back from Future.result().  Every tool wraps that call in
    `except Exception`, which does not catch SystemExit or KeyboardInterrupt:
    those would sail past the tool, past the server's dispatch, and take the
    event loop down with them.  Keeping the worker alive must not cost the
    caller its process, so a non-Exception BaseException is reported as a
    RuntimeError instead.
    """
    if isinstance(exc, Exception):
        return exc
    logger.error("Operation raised %s: %s", type(exc).__name__, exc)
    return RuntimeError(
        "The operation raised %s: %s" % (type(exc).__name__, exc)
    )


class PowerPointCOMWrapper:
    """Manages the lifecycle of a PowerPoint COM Application object.

    All COM operations are routed through a dedicated STA thread to ensure
    thread safety. The MCP server (which runs async) calls methods on this
    wrapper, which internally queues operations to the COM thread.
    """

    def __init__(self):
        self._app = None
        self._com_thread: Optional[threading.Thread] = None
        self._queue: Queue = Queue()
        self._running = False
        self._target_pres_full_name: Optional[str] = None  # session-level target (FullName for uniqueness)
        # Which worker generation is the live one.  A worker abandoned after a
        # wedge keeps its own generation, sees that it is no longer current,
        # and stops touching shared state (issue #199).
        self._generation = 0
        # Clear while a replacement worker is still proving that PowerPoint
        # answers.  execute() then fails fast instead of queueing behind a
        # connection that may not be there.
        self._healthy = threading.Event()
        self._healthy.set()
        self._recover_lock = threading.Lock()

    def start(self) -> None:
        """Start the COM worker thread."""
        if self._running:
            return
        self._running = True
        self._healthy.set()
        self._start_worker(self._queue, self._generation)
        logger.info("COM worker thread started")

    def _start_worker(self, queue: Queue, generation: int) -> None:
        """Spawn a worker bound to one queue and one generation."""
        self._com_thread = threading.Thread(
            target=self._com_worker,
            args=(queue, generation),
            daemon=True,
            name="COM-Worker-%d" % generation,
        )
        self._com_thread.start()

    def stop(self) -> None:
        """Stop the COM worker thread and clean up."""
        if not self._running:
            return
        self._running = False
        # Send a sentinel to unblock the worker
        self._queue.put(None)
        if self._com_thread and self._com_thread.is_alive():
            self._com_thread.join(timeout=5.0)
        logger.info("COM worker thread stopped")

    def _wait_until_responsive(self, deadline: float) -> None:
        """Block until PowerPoint stops rejecting calls, or the deadline passes.

        A cheap property read is used as the probe.  Doing this *before* the
        operation is what makes the busy handling safe: a dialog or open menu
        is almost always already up when the request arrives, so waiting here
        costs nothing in the common case and avoids having to re-run a
        partially-applied operation afterwards (issue #200).

        Only busy HRESULTs are handled.  Any other com_error is left alone so
        that _get_app_impl keeps its own stale-reference detection, and a
        missing app object simply skips the probe -- connecting is the job of
        _connect_impl.

        Args:
            deadline: time.monotonic() value after which to give up.  Shared
                across every wait for one queued operation, so a retry cannot
                start a second full budget on top of the first.

        Raises:
            pywintypes.com_error: the last busy rejection, if PowerPoint is
                still busy at the deadline.
        """
        if self._app is None:
            return
        attempt = 0
        while True:
            try:
                _ = self._app.Name
                return
            except pywintypes.com_error as e:
                if e.hresult not in _BUSY_HRESULTS:
                    # Stale reference or anything else: not this method's job.
                    return
                if attempt == 0 and AUTO_DISMISS_DIALOG:
                    # Optionally dismiss the blocking dialog via ESC so the
                    # next probe likely succeeds immediately.
                    _try_dismiss_ppt_dialog()
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    raise
                delay = _BUSY_BACKOFF[min(attempt, len(_BUSY_BACKOFF) - 1)]
                logger.warning(
                    "PowerPoint is busy (modal dialog open?). "
                    "Waiting %.1fs before probing again (%.1fs of budget left)...",
                    delay, remaining,
                )
                time.sleep(min(delay, remaining))
                attempt += 1
            except AttributeError:
                return  # app went away; _get_app_impl will reconnect

    def _run_item(self, func, args, kwargs, future, idempotent: bool) -> None:
        """Run one queued operation and settle its future.

        Returns immediately if the caller already gave up on this operation:
        execute() cancels the future on timeout, so anything still queued
        behind a slow operation is dropped instead of being applied minutes
        later, after the caller was told it had failed (issue #199).

        All waiting for one operation shares a single _BUSY_WAIT_BUDGET
        deadline, so however many times we probe or retry, the worker starts
        the operation within the window execute() allows for waiting.

        Split out of _com_worker so the busy-handling policy is unit-testable
        without a COM thread.
        """
        if future.cancelled():
            logger.warning(
                "Dropping %s: the caller stopped waiting for it",
                getattr(func, "__name__", func),
            )
            return

        deadline = time.monotonic() + _BUSY_WAIT_BUDGET
        attempt = 0
        started = False
        while True:
            try:
                self._wait_until_responsive(deadline)
            except pywintypes.com_error as e:
                logger.warning("Gave up waiting for PowerPoint: %s", e)
                if not started and not future.set_running_or_notify_cancel():
                    # The caller gave up while we waited.  Nothing was
                    # applied, so there is nobody to report the failure to.
                    return
                future.set_exception(RuntimeError(_BUSY_TIMEOUT_MESSAGE))
                return

            if not started:
                # Only now, with PowerPoint responsive and the call about to
                # go out, does this stop being cancellable.  Marking it
                # earlier meant a caller timing out during the wait could not
                # cancel, and the edit landed anyway.
                if not future.set_running_or_notify_cancel():
                    logger.warning(
                        "Dropping %s: the caller stopped waiting for it",
                        getattr(func, "__name__", func),
                    )
                    return
                started = True

            try:
                future.set_result(func(*args, **kwargs))
                return
            except pywintypes.com_error as e:
                if e.hresult not in _BUSY_HRESULTS:
                    future.set_exception(e)
                    return
                if not idempotent:
                    # The operation was rejected part-way through.  Re-running
                    # it would repeat whatever it already applied, so surface
                    # the failure instead of duplicating work (issue #200).
                    future.set_exception(RuntimeError(
                        "PowerPoint became busy in the middle of the operation "
                        "(a dialog or menu is open?). The operation may have "
                        "been partially applied -- check the slide and retry."
                    ))
                    return
                busy_error = e
            except BaseException as e:
                # BaseException, not Exception: a SystemExit or
                # KeyboardInterrupt escaping here used to end the worker
                # thread while _running stayed True, leaving every later
                # request to queue up behind a worker that no longer existed
                # (issue #199).
                future.set_exception(_caller_safe(e))
                return

            # Idempotent (connect / attach): safe to re-run from the top.  Back
            # off before doing so -- when nothing is connected yet there is no
            # object to probe, so this sleep is the only thing giving a modal
            # dialog time to clear.
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                logger.warning("Idempotent operation still busy: %s", busy_error)
                future.set_exception(RuntimeError(_BUSY_TIMEOUT_MESSAGE))
                return
            delay = _BUSY_BACKOFF[min(attempt, len(_BUSY_BACKOFF) - 1)]
            logger.warning(
                "Busy during an idempotent operation, retrying in %.1fs: %s",
                min(delay, remaining), busy_error,
            )
            time.sleep(min(delay, remaining))
            attempt += 1

    def _recover_from_wedge(self, generation: int) -> bool:
        """Abandon a worker stuck inside a COM call and start a fresh one.

        An outgoing COM call cannot be cancelled from this side, so the stuck
        thread cannot be reclaimed.  It is a daemon thread and is left to
        finish or not.  What recovery buys is that the next operation meets a
        working worker instead of queueing behind the stuck one forever, which
        previously needed a client restart (issue #199).

        Args:
            generation: the worker generation the timed-out call belonged to.
                Recovery is skipped if that generation has already been
                replaced, so several callers timing out together still
                produce one replacement.

        Returns:
            True if this call performed the recovery.
        """
        with self._recover_lock:
            if generation != self._generation or not self._running:
                return False
            old_queue = self._queue
            self._generation = generation + 1
            self._healthy.clear()
            self._queue = Queue()
            logger.error(
                "The COM worker is stuck inside a PowerPoint call. Abandoning "
                "it and starting worker generation %d.", self._generation,
            )
            self._start_worker(self._queue, self._generation)

        # Outside the lock.  Everything the abandoned worker had not started
        # is dropped rather than applied once it unwedges, long after the
        # callers were told their operations had failed.
        while True:
            try:
                item = old_queue.get_nowait()
            except Empty:
                break
            if item is not None:
                item[3].cancel()
                item[5].set()  # stop the caller waiting out the queue budget
        # If the abandoned worker ever does return from its call it goes back
        # to get().  The sentinel stops it blocking there for good.
        old_queue.put(None)
        return True

    def _reattach_after_wedge(self, generation: int) -> None:
        """Prove PowerPoint answers again before letting traffic resume.

        A replacement worker runs in a fresh apartment, so the proxy the
        wedged worker held is useless here and is dropped.  Connecting proper
        is still _connect_impl's job; all this does is find out whether
        PowerPoint is answering at all.

        Staying unhealthy is the point.  A busy rejection most likely means
        the modal dialog that caused the wedge is still up, and reporting
        health then would turn the next call into a misleading "partially
        applied" error for an operation that never started.  Any other failure
        is reported as healthy so that _connect_impl can relaunch PowerPoint
        the way it always has.
        """
        self._app = None
        attempt = 0
        while self._running and generation == self._generation:
            try:
                app = win32com.client.GetActiveObject("PowerPoint.Application")
                _ = app.Name
                self._app = app
                logger.info("Reattached to PowerPoint after a wedged call")
                break
            except pywintypes.com_error as e:
                if e.hresult not in _BUSY_HRESULTS:
                    logger.warning(
                        "Could not reattach to PowerPoint after the wedge "
                        "(%s). The next operation will connect normally.", e,
                    )
                    break
                delay = _BUSY_BACKOFF[min(attempt, len(_BUSY_BACKOFF) - 1)]
                logger.warning(
                    "PowerPoint is still busy after the wedge, probing again "
                    "in %.1fs", delay,
                )
                time.sleep(delay)
                attempt += 1
            except Exception as e:
                logger.warning(
                    "Could not reattach to PowerPoint after the wedge (%s). "
                    "The next operation will connect normally.", e,
                )
                break

        if generation == self._generation:
            self._healthy.set()

    def _com_worker(self, queue: Queue, generation: int) -> None:
        """Worker thread that processes COM operations in an STA apartment.

        Each worker owns the queue it was started with rather than reading
        self._queue, so a worker abandoned after a wedge cannot drain the
        queue its replacement is serving (issue #199).
        """
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        try:
            if generation > 0:
                self._reattach_after_wedge(generation)
            while self._running and generation == self._generation:
                item = queue.get()
                if item is None:
                    break
                func, args, kwargs, future, idempotent, dequeued = item
                # The caller's own budget for the operation starts here, not
                # when it was queued (issue #192).  Setting this on every path
                # out of the queue, the drop below included, is what stops a
                # caller waiting out the whole queue budget for an answer that
                # already exists.
                dequeued.set()
                if generation != self._generation:
                    # Abandoned while this item sat in the queue, so it
                    # arrived after _recover_from_wedge had drained it.
                    future.cancel()
                    break
                try:
                    self._run_item(func, args, kwargs, future, idempotent)
                except BaseException:
                    # _run_item settles the future itself; reaching here means
                    # something escaped it entirely.  Log and keep the worker
                    # alive rather than silently abandoning the queue.
                    logger.exception("COM worker item failed unexpectedly")
                    if not future.done():
                        future.set_exception(
                            RuntimeError("The COM worker failed unexpectedly.")
                        )
        finally:
            # An abandoned worker shares self._app with the live one and must
            # not clear it on the way out.
            if generation == self._generation:
                self._cleanup_com()
            pythoncom.CoUninitialize()

    def execute(self, func: Callable, *args: Any,
                idempotent: bool = False, **kwargs: Any) -> Any:
        """Execute a function on the COM thread and return its result.

        This is the main entry point for all COM operations from async code.
        It queues the operation and blocks until completion.

        Args:
            func: The function to execute on the COM thread
            idempotent: Keyword-only, and not forwarded to func.  Pass True
                only when re-running func from the top has no cumulative
                effect (connecting, attaching).  Such operations are retried
                wholesale if PowerPoint turns busy mid-call; everything else
                fails instead, so a half-applied operation is never silently
                repeated (issue #200).
            *args, **kwargs: Arguments to pass to func

        Returns:
            The return value of func

        The wait is in two parts.  Time spent queued behind other operations
        belongs to the queue, not to this operation, so the call budget only
        starts once the worker picks the item up (issue #192).  Keeping them
        apart is also what makes a wedge identifiable at all: a call that used
        up its budget waiting in line looks exactly like one the worker is
        stuck inside, and abandoning the worker for it would be a bug.

        Raises:
            RuntimeError: if the COM worker thread has died, if a previous
                call wedged it and the replacement has not reconnected yet,
                or if either wait times out.
            Any exception raised by func
        """
        if not self._healthy.wait(_HEALTH_WAIT):
            # A replacement worker is still probing.  Queueing here would
            # cost another full timeout to learn what is already known
            # (issue #199).
            raise RuntimeError(_REBUILDING_MESSAGE)

        if self._com_thread is not None and not self._com_thread.is_alive():
            # Nothing is draining the queue any more, so queueing would just
            # burn the full timeout before failing (issue #199).
            raise RuntimeError(
                "The COM worker thread is no longer running. "
                "Restart the MCP server to reconnect to PowerPoint."
            )

        future: Future = Future()
        dequeued = threading.Event()
        # Under the recovery lock, so that the generation this call is
        # attributed to and the queue it lands on are the same pair.  Without
        # it a recovery could swap the queue in between, and the item would
        # sit behind the sentinel on a queue no worker serves.
        with self._recover_lock:
            if not self._healthy.is_set():
                # A recovery started while the health gate was being read.
                raise RuntimeError(_REBUILDING_MESSAGE)
            generation = self._generation
            self._queue.put((func, args, kwargs, future, idempotent, dequeued))
        watching = pending_com_futures.get()
        if watching is not None:
            watching.add(future)

        if not dequeued.wait(_QUEUE_WAIT):
            # Still in the queue.  Take it back rather than let it run against
            # a deck that has moved on, long after this call reported failure.
            if future.cancel():
                raise RuntimeError(_QUEUE_TIMEOUT_MESSAGE)
            # It started in the moment the wait expired, so it gets its budget
            # after all.
        # The worker may spend up to _BUSY_WAIT_BUDGET waiting for PowerPoint
        # before it even starts, so the caller has to allow for that on top of
        # the time the operation itself gets.  Otherwise the caller could
        # abandon the future while the worker is still applying the edit.
        try:
            return future.result(timeout=_BUSY_WAIT_BUDGET + _CALL_TIMEOUT)
        except FuturesCancelledError:
            # The caller went away and utils.offload cancelled this before the
            # worker started it.  Report it as an ordinary error so the tool
            # wrapper can render it; nobody is waiting for the answer anyway.
            raise RuntimeError(
                "The operation was cancelled before PowerPoint started it."
            ) from None
        except FuturesTimeoutError:
            if future.done():
                # Not our timeout: either func itself raised TimeoutError
                # (concurrent.futures.TimeoutError *is* the builtin since
                # 3.11, so the two are indistinguishable by type), or the
                # operation finished in the moment the wait expired. Either
                # way there is a real outcome to hand back.
                return future.result()
            # Cancel so the worker skips this item if it has not started it.
            # An operation already in flight cannot be cancelled -- COM
            # outgoing calls are not interruptible -- but everything queued
            # behind it is dropped rather than applied after the fact.
            cancelled = future.cancel()
            if cancelled:
                raise RuntimeError(
                    "PowerPoint did not finish the operation within %.0fs "
                    "and it was cancelled."
                    % (_BUSY_WAIT_BUDGET + _CALL_TIMEOUT)
                ) from None
            if future.done():
                # It finished in the moment between the timeout and the
                # cancel; the caller may as well have the outcome.
                return future.result()

            # The future could be neither cancelled nor read, so the worker
            # is inside the call and has been for the whole budget.  The
            # budget started when the worker picked the item up, so this is
            # time spent in the call and not in the queue, which is what makes
            # it evidence of a wedge rather than of a busy worker.  Callers
            # queued behind it cancel cleanly and must not each trigger
            # another replacement.
            recovered = self._recover_from_wedge(generation)
            raise RuntimeError(
                "PowerPoint did not finish the operation within %.0fs and it "
                "was left running, so it may still be applied.%s"
                % (
                    _BUSY_WAIT_BUDGET + _CALL_TIMEOUT,
                    " A fresh connection is being established."
                    if recovered else "",
                )
            ) from None

    def _connect_impl(self, visible: Optional[bool] = None, allow_launch: bool = True) -> Any:
        """Internal: connect to PowerPoint on the COM thread.

        When allow_launch is False, this only attaches to an already-running
        PowerPoint instance; if none is running, ConnectionError is raised.
        """
        if self._app is not None:
            try:
                _ = self._app.Name
                if visible is not None:
                    self._app.Visible = visible
                return self._app
            except pywintypes.com_error as e:
                if e.hresult in _BUSY_HRESULTS:
                    raise  # PowerPoint busy — _run_item decides whether to retry
                logger.warning("Stale COM reference, reconnecting...")
                self._app = None
            except AttributeError:
                logger.warning("Stale COM reference, reconnecting...")
                self._app = None

        # Try existing instance first
        launched_new = False
        try:
            self._app = win32com.client.GetActiveObject("PowerPoint.Application")
            logger.info("Attached to existing PowerPoint instance")
        except pywintypes.com_error as e:
            if e.hresult in _BUSY_HRESULTS:
                # PowerPoint is running but busy (modal dialog). Re-raise as
                # pywintypes.com_error so _com_worker's retry loop handles it.
                raise
            if not allow_launch:
                raise ConnectionError(
                    "PowerPoint is not running. Call ppt_connect, "
                    "ppt_create_presentation, or ppt_open_presentation first."
                ) from e
            try:
                self._app = win32com.client.Dispatch("PowerPoint.Application")
                logger.info("Launched new PowerPoint instance via Dispatch")
                launched_new = True
            except pywintypes.com_error as e2:
                if e2.hresult in _BUSY_HRESULTS:
                    raise  # PowerPoint busy — _run_item decides whether to retry
                raise ConnectionError(
                    f"Failed to connect to PowerPoint. Is it installed? Error: {e2.strerror}"
                ) from e2

        if visible is not None:
            self._app.Visible = visible
        elif launched_new and not self._app.Visible:
            # Only force visibility when we ourselves started PowerPoint.
            # Don't yank a user-hidden running instance to the foreground.
            self._app.Visible = True

        return self._app

    def _get_app_impl(self, allow_launch: bool = False) -> Any:
        """Internal: get app on COM thread.

        By default refuses to launch PowerPoint — the vast majority of tools
        operate on an already-open presentation and should fail fast if
        PowerPoint is not running, instead of silently spawning it.
        """
        if self._app is None:
            return self._connect_impl(allow_launch=allow_launch)
        try:
            _ = self._app.Name
            return self._app
        except pywintypes.com_error as e:
            if e.hresult in _BUSY_HRESULTS:
                raise  # PowerPoint busy — _run_item decides whether to retry
            logger.warning("COM connection lost, reconnecting...")
            self._app = None
            return self._connect_impl(allow_launch=allow_launch)
        except AttributeError:
            logger.warning("COM connection lost, reconnecting...")
            self._app = None
            return self._connect_impl(allow_launch=allow_launch)

    def _find_target_pres_impl(self, app) -> Optional[Any]:
        """Internal: locate the session target presentation, or None.

        Matches on FullName.  When the target is no longer open, the stored
        target is cleared so callers fall back to ActivePresentation.
        """
        if not self._target_pres_full_name:
            return None
        for i in range(1, app.Presentations.Count + 1):
            try:
                p = app.Presentations(i)
                if p.FullName == self._target_pres_full_name:
                    return p
            except Exception:
                pass
        # Target was closed since last activation — clear and fall back
        logger.warning(
            "Target presentation '%s' is no longer open; "
            "falling back to ActivePresentation",
            self._target_pres_full_name,
        )
        self._target_pres_full_name = None
        return None

    def _get_pres_impl(self) -> Any:
        """Internal: get target presentation on COM thread.

        Returns the session-level target presentation if one has been set via
        _set_target_pres_impl and the file is still open.  Falls back to
        ActivePresentation when no target is set or when the target was closed.

        This deliberately does NOT activate the presentation's window: doing so
        on every tool call yanked PowerPoint to the foreground and stole focus
        from whatever the user was doing (issue #183).  Operations that need to
        follow the edit use _get_target_window_impl(), which drives the target
        deck's own window without activating it.
        """
        app = self._get_app_impl()
        pres = self._find_target_pres_impl(app)
        if pres is not None:
            return pres
        return app.ActivePresentation

    def _get_target_window_impl(self) -> Optional[Any]:
        """Internal: get the DocumentWindow to drive, without activating it.

        Returns the target presentation's first window when a session target is
        set, otherwise app.ActiveWindow.  Returns None when the target deck has
        no window (opened with with_window=False) — callers must not silently
        fall back to ActiveWindow there, since that belongs to another deck.
        """
        app = self._get_app_impl()
        pres = self._find_target_pres_impl(app)
        if pres is not None:
            try:
                if pres.Windows.Count == 0:
                    return None
                return pres.Windows(1)
            except Exception:
                return None
        try:
            if app.Windows.Count == 0:
                return None
            return app.ActiveWindow
        except Exception:
            return None

    def _activate_target_window_impl(self) -> Any:
        """Internal: bring the target presentation's window to the front.

        Reserved for the handful of operations COM will only perform on an
        active view — Shape.Select() and TextRange.Select() + ExecuteMso().
        Everything else must use _get_target_window_impl(), which drives the
        window without stealing focus (issue #183).

        Raises RuntimeError when the target deck has no window, rather than
        letting the caller fail later with an opaque COM error.
        """
        window = self._get_target_window_impl()
        if window is None:
            raise RuntimeError(
                "This operation needs an active PowerPoint window, but the "
                "target presentation has none (opened with with_window=False?)."
            )
        try:
            window.Activate()
        except Exception as e:
            logger.warning("Could not activate presentation window: %s", e)
        return window

    def _set_target_pres_impl(self, name_or_index) -> dict:
        """Internal: set session-level target presentation on COM thread."""
        app = self._get_app_impl()
        if app.Presentations.Count == 0:
            raise RuntimeError("No presentation is open in PowerPoint.")

        pres = None
        if isinstance(name_or_index, int):
            if name_or_index < 1 or name_or_index > app.Presentations.Count:
                raise ValueError(
                    f"Presentation index {name_or_index} out of range "
                    f"(1-{app.Presentations.Count})"
                )
            pres = app.Presentations(name_or_index)
        else:
            name_lower = name_or_index.lower()
            matches = []
            for i in range(1, app.Presentations.Count + 1):
                p = app.Presentations(i)
                if p.Name.lower() == name_lower or p.FullName.lower() == name_lower:
                    matches.append(p)
            if len(matches) == 0:
                open_names = [
                    app.Presentations(i).Name
                    for i in range(1, app.Presentations.Count + 1)
                ]
                raise ValueError(
                    f"Presentation '{name_or_index}' not found. "
                    f"Open presentations: {open_names}"
                )
            if len(matches) > 1:
                raise ValueError(
                    f"Multiple presentations match '{name_or_index}': "
                    f"{[p.Name for p in matches]}. Use a more specific name."
                )
            pres = matches[0]

        # Bring the presentation's window to the front
        try:
            pres.Windows(1).Activate()
        except Exception as e:
            logger.warning("Could not activate presentation window: %s", e)

        # Store FullName (includes path) to uniquely identify the presentation
        # even if another file with the same basename is later opened.
        self._target_pres_full_name = pres.FullName
        index = None
        for i in range(1, app.Presentations.Count + 1):
            if app.Presentations(i).FullName == pres.FullName:
                index = i
                break
        return {
            "success": True,
            "name": pres.Name,
            "full_name": pres.FullName,
            "index": index,
        }

    def _ensure_presentation_impl(self) -> Any:
        """Internal: ensure presentation on COM thread."""
        app = self._get_app_impl()
        if app.Presentations.Count == 0:
            raise RuntimeError(
                "No presentation is open in PowerPoint. "
                "Use ppt_create_presentation or ppt_open_presentation first."
            )
        return app.ActivePresentation

    def _cleanup_com(self) -> None:
        """Release COM references."""
        if self._app is not None:
            try:
                # Don't quit PowerPoint - the user may be using it
                self._app = None
            except Exception:
                pass
        gc.collect()


def handle_com_error(e: pywintypes.com_error) -> dict:
    """Parse a COM error into a structured dict for error responses."""
    result = {
        "hresult": e.hresult,
        "message": str(e.strerror) if e.strerror else "Unknown COM error",
        "source": None,
        "description": None,
    }
    if e.excepinfo:
        result["source"] = e.excepinfo[1] if len(e.excepinfo) > 1 else None
        result["description"] = e.excepinfo[2] if len(e.excepinfo) > 2 else None
    return result


# Global singleton instance
ppt = PowerPointCOMWrapper()
