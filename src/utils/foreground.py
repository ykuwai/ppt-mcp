"""Hand the foreground back when a PowerPoint call took it (Windows).

Some of what the tools ask of PowerPoint brings its window to the front:
``Presentations.Add`` and ``Presentations.Open`` with a window, and
``Window.Activate``, which selecting shapes, Undo, Redo and ribbon commands
need, since COM only does those in the active view. Windows lets that happen
whenever its foreground lock is off, which by default is after 200 seconds
without input, so PowerPoint jumped in front of whatever the user was doing
while an agent worked.

The activation itself cannot be avoided, but it lands by the time the call
returns. So the window that was in front before a call is noted, and when the
call leaves a PowerPoint or Excel window in front instead, the noted window is
put back. PowerPoint keeps its own active window, so a following Undo or
Select still reaches the deck it was meant for.

Nothing is handed back when PowerPoint or Excel was already in front (the user
is working there), when the user touched the keyboard or mouse during the call
(they may have clicked PowerPoint themselves), or when something other than
PowerPoint took the front. A call whose point is to show PowerPoint, such as
starting a slide show, is marked with ``shows_powerpoint`` and left alone.
"""

import logging
import os

logger = logging.getLogger(__name__)

# The processes whose windows a PowerPoint call can bring forward: PowerPoint,
# and Excel for a chart's data.
OFFICE_PROCESSES = frozenset({"powerpnt.exe", "excel.exe"})


def shows_powerpoint(func):
    """Mark an impl whose purpose is to bring PowerPoint to the front."""
    func.shows_powerpoint = True
    return func


class Foreground:
    """Reads and sets the foreground through the Win32 API."""

    def __init__(self):
        import ctypes
        from ctypes import wintypes

        import win32api
        import win32gui
        import win32process

        self._ctypes, self._wintypes = ctypes, wintypes
        self._win32api, self._win32gui, self._win32process = win32api, win32gui, win32process
        self._kernel32 = ctypes.windll.kernel32
        self._user32 = ctypes.windll.user32

        class LastInput(ctypes.Structure):
            _fields_ = [("cbSize", wintypes.UINT), ("dwTime", wintypes.DWORD)]

        self._last_input = LastInput()
        self._last_input.cbSize = ctypes.sizeof(LastInput)

    def window(self):
        return self._win32gui.GetForegroundWindow()

    def process_name(self, hwnd):
        """The executable of the process that owns ``hwnd``, lower case."""
        if not hwnd:
            return ""
        _, pid = self._win32process.GetWindowThreadProcessId(hwnd)
        handle = self._kernel32.OpenProcess(0x1000, False, pid)  # QUERY_LIMITED_INFORMATION
        if not handle:
            return ""
        try:
            buffer = self._ctypes.create_unicode_buffer(1024)
            size = self._wintypes.DWORD(1024)
            if not self._kernel32.QueryFullProcessImageNameW(handle, 0, buffer, self._ctypes.byref(size)):
                return ""
            return os.path.basename(buffer.value).lower()
        finally:
            self._kernel32.CloseHandle(handle)

    def last_input(self):
        """The tick of the user's last keyboard or mouse input."""
        self._user32.GetLastInputInfo(self._ctypes.byref(self._last_input))
        return self._last_input.dwTime

    def is_window(self, hwnd):
        return bool(hwnd) and bool(self._win32gui.IsWindow(hwnd))

    def put_back(self, hwnd):
        """Bring ``hwnd`` to the front again.

        This process has no right to set the foreground on its own, so it joins
        the input of the thread that now holds it, which may, for the moment it
        takes to put the window back.
        """
        current = self._win32gui.GetForegroundWindow()
        mine = self._win32api.GetCurrentThreadId()
        theirs = self._win32process.GetWindowThreadProcessId(current)[0] if current else 0
        attached = bool(theirs) and theirs != mine
        if attached:
            self._win32process.AttachThreadInput(mine, theirs, True)
        try:
            self._win32gui.BringWindowToTop(hwnd)
            self._win32gui.SetForegroundWindow(hwnd)
        finally:
            if attached:
                self._win32process.AttachThreadInput(mine, theirs, False)


def run_keeping_foreground(func, args, kwargs, foreground):
    """Run ``func(*args, **kwargs)`` and hand the foreground back if it took it.

    ``foreground`` is a ``Foreground``, or None to run ``func`` untouched.
    """
    if foreground is None or getattr(func, "shows_powerpoint", False):
        return func(*args, **kwargs)
    try:
        before = foreground.window()
        office_in_front = foreground.process_name(before) in OFFICE_PROCESSES
        touched = foreground.last_input()
    except Exception:  # noqa: BLE001 - noting the front must never stop a tool
        logger.debug("Could not read the foreground before a call", exc_info=True)
        return func(*args, **kwargs)
    try:
        return func(*args, **kwargs)
    finally:
        if before and not office_in_front:
            _put_back_if_taken(foreground, before, touched)


def _put_back_if_taken(foreground, before, touched):
    try:
        now = foreground.window()
        if now == before or foreground.process_name(now) not in OFFICE_PROCESSES:
            return
        if foreground.last_input() != touched:
            # The user did something during the call, perhaps clicked
            # PowerPoint themselves; leave the front where it is.
            return
        if not foreground.is_window(before):
            return
        foreground.put_back(before)
    except Exception:  # noqa: BLE001 - handing back is a courtesy
        logger.debug("Could not hand the foreground back", exc_info=True)
