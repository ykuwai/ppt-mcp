"""COM references kept until the process behind them has ended (#299).

When a chart's data workbook is closed, the Excel that held it closes down,
and letting go of the last reference to anything in it waits until that Excel
has ended: two to twelve seconds as a rule, and sometimes exactly a minute.
Done on the COM worker, that minute outlasted the worker's 45 s budget, the
worker was abandoned in the middle of letting go, and the tool reported a
timeout although the chart was made. Excel ends just as soon with the
references still held, and once it has, letting go of them takes no time.

So the references are handed to ``held`` instead, with the id of the process
they lead to, and the tool returns at once. Before each item the worker runs,
``let_go_of_ended`` lets go of those whose process has ended, on the thread
that holds them, since a COM reference belongs to the thread it was made on.

Not waiting at all went too far. A chart's workbook is opened in an Excel of
its own, and an Excel started while the one before it is still ending crashes,
measured straight through pywin32 too. Waiting for each Excel as its last
reference went had kept them one at a time. A new chart is now written without
Excel where gvml can write it, and pasted (ppt_com/chart_defaults.py); for the
types it cannot, ``wait_until_ended`` waits, before ``AddChart2``, for the Excel behind
every held set to end, for ``WAIT_LIMIT`` seconds at most.
"""

import logging
import threading
import time

logger = logging.getLogger(__name__)

# OpenProcess access to wait on a process, WaitForSingleObject's answer when
# it has ended, and OpenProcess's error for a process that is not there.
_SYNCHRONIZE = 0x00100000
_WAIT_OBJECT_0 = 0
_ERROR_INVALID_PARAMETER = 87

# Past this many sets held by one worker, something keeps an Excel from ending.
_MANY = 16

# How long AddChart2 waits for the Excel before it to end, in seconds. An
# Excel ends from two seconds to about a minute after its workbook is closed,
# so this waits out the short ones and lets a slow one overlap.
WAIT_LIMIT = 15.0


def _pause(seconds: float) -> None:
    """Wait on the STA thread while still answering COM calls into it."""
    try:
        import pythoncom

        pythoncom.PumpWaitingMessages()
    except Exception:  # noqa: BLE001 - then only sleep
        pass
    time.sleep(seconds)


def process_ended(pid: int) -> bool:
    """Whether the process ``pid`` has ended, or never was.

    Any other failure to open it, such as an Excel running elevated that this
    process may not wait on, raises, and the question is asked again before
    the next item: taking it for ended would let go while Excel is still on its
    way out, which is the wait this module is here to avoid.
    """
    import pywintypes
    import win32api
    import win32event

    try:
        handle = win32api.OpenProcess(_SYNCHRONIZE, False, pid)
    except pywintypes.error as e:
        if e.winerror == _ERROR_INVALID_PARAMETER:
            return True
        raise
    try:
        return win32event.WaitForSingleObject(handle, 0) == _WAIT_OBJECT_0
    finally:
        win32api.CloseHandle(handle)


def window_process(hwnd) -> int | None:
    """The id of the process that owns window ``hwnd``, or None."""
    try:
        import win32process

        pid = win32process.GetWindowThreadProcessId(int(hwnd))[1]
    except Exception:  # noqa: BLE001 - no window, no process to wait for
        return None
    return pid or None


class HeldUntilEnded:
    """COM references, each kept until its process has ended."""

    def __init__(self, ended=process_ended, pause=_pause, clock=time.monotonic):
        self._ended = ended
        self._pause = pause
        self._clock = clock
        self._lock = threading.Lock()
        # (thread that holds them, process id, the references). The thread
        # object itself, not its ident: the system hands an ended thread's
        # ident to a new one, and that worker must not let go of references
        # made in another apartment.
        self._held = []
        self._warned = False

    def hold(self, references, pid) -> None:
        """Keep ``references`` until process ``pid`` has ended.

        With no process to wait for, they are let go of at once, as before.
        """
        references = [r for r in references if r is not None]
        if not references:
            return
        if not pid:
            references.clear()
            return
        here = threading.current_thread()
        with self._lock:
            self._held.append((here, pid, references))
            many = sum(1 for entry in self._held if entry[0] is here)
        if many > _MANY and not self._warned:
            self._warned = True
            logger.warning(
                "%d sets of Excel references are waiting for their Excel to end; "
                "an Excel is not ending", many,
            )

    def let_go_of_ended(self) -> int:
        """Let go of what this thread holds for processes that have ended.

        They are let go of before the worker's next item, so after a quiet
        spell they stay held until the next call; Excel ends without waiting
        for that.

        Answers how many sets of references were let go of. Those another
        thread holds are left alone: a worker abandoned after a wedge keeps
        them, and they are let go of with its apartment.
        """
        here = threading.current_thread()
        with self._lock:
            mine = [entry for entry in self._held if entry[0] is here]
        if not mine:
            return 0
        done = []
        for entry in mine:
            try:
                if self._ended(entry[1]):
                    done.append(entry)
            except Exception:  # noqa: BLE001 - asked again before the next item
                logger.debug("Could not tell whether process %s has ended", entry[1], exc_info=True)
        if not done:
            return 0
        with self._lock:
            self._held = [entry for entry in self._held if not any(entry is d for d in done)]
        for _thread, _pid, references in done:
            references.clear()
        return len(done)

    def wait_until_ended(self, limit: float = WAIT_LIMIT) -> float:
        """Wait for the Excel behind every held set to end, then let go.

        Called before ``AddChart2``, so that its Excel does not start while
        another is still ending. The sets of every thread are waited for, an
        abandoned worker's too, since their Excel overlaps all the same, but
        only this thread's are let go of. After ``limit`` seconds the chart
        goes ahead, and that is logged. A process that cannot be asked
        about is not waited for. Answers the seconds waited.
        """
        start = self._clock()
        while True:
            with self._lock:
                pids = {entry[1] for entry in self._held}
            waiting = []
            for pid in sorted(pids):
                try:
                    if not self._ended(pid):
                        waiting.append(pid)
                except Exception:  # noqa: BLE001 - not waited for
                    logger.debug("Could not tell whether process %s has ended", pid, exc_info=True)
            waited = self._clock() - start
            if not waiting:
                break
            if waited >= limit:
                logger.warning(
                    "Excel %s has not ended %.0f s after its chart was done; "
                    "opening the next chart's workbook anyway", waiting, waited,
                )
                break
            self._pause(0.1)
        self.let_go_of_ended()
        if waited >= 0.05:
            logger.debug("Waited %.1f s for the last chart's Excel to end", waited)
        return waited

    def still_held(self) -> int:
        """How many sets of references are held, by any thread."""
        with self._lock:
            return len(self._held)

    def leave_behind(self) -> int:
        """At a worker's end, let go of what has ended and log what has not.

        What has not ended stays held for the rest of this process's life, so
        that the worker does not wait up to a minute for an Excel on its way
        out.
        """
        self.let_go_of_ended()
        here = threading.current_thread()
        with self._lock:
            left = sum(1 for entry in self._held if entry[0] is here)
        if left:
            logger.info("Leaving %d Excel reference set(s) to the end of the COM apartment", left)
        return left


held = HeldUntilEnded()
