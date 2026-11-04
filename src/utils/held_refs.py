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
"""

import logging
import threading

logger = logging.getLogger(__name__)

# OpenProcess access to wait on a process, and WaitForSingleObject's answer
# when it has ended.
_SYNCHRONIZE = 0x00100000
_WAIT_OBJECT_0 = 0


def process_ended(pid: int) -> bool:
    """Whether the process ``pid`` has ended, or never was."""
    import pywintypes
    import win32api
    import win32event

    try:
        handle = win32api.OpenProcess(_SYNCHRONIZE, False, pid)
    except pywintypes.error:
        return True
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

    def __init__(self, ended=process_ended):
        self._ended = ended
        self._lock = threading.Lock()
        # (thread that holds them, process id, the references)
        self._held = []

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
        with self._lock:
            self._held.append((threading.get_ident(), pid, references))

    def let_go_of_ended(self) -> int:
        """Let go of what this thread holds for processes that have ended.

        Answers how many sets of references were let go of. Those another
        thread holds are left alone: a worker abandoned after a wedge keeps
        them, and they are let go of with its apartment.
        """
        here = threading.get_ident()
        with self._lock:
            mine = [entry for entry in self._held if entry[0] == here]
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
        here = threading.get_ident()
        with self._lock:
            left = sum(1 for entry in self._held if entry[0] == here)
        if left:
            logger.info("Leaving %d Excel reference set(s) to the end of the COM apartment", left)
        return left


held = HeldUntilEnded()
