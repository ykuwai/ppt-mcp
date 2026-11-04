"""The Windows clipboard, and keeping the user's across a copy and paste.

Some operations can only move a shape between presentations through the
clipboard, which belongs to the user. ``kept_clipboard`` saves every format on
it before such a copy and puts them back afterwards, and says which formats it
could not put back rather than lose them silently.

Most formats are a block of global memory and are copied as bytes. A few are
handles to GDI objects. Windows makes ``CF_BITMAP`` from ``CF_DIB``,
``CF_METAFILEPICT`` from ``CF_ENHMETAFILE`` and ``CF_PALETTE`` from a DIB's
colour table on its own, so those are left to it when their source is there.
An enhanced metafile is kept as its bits. What is left, owner drawn formats and
private GDI objects, cannot be copied, and is reported as lost.

The ctypes bindings are shared with ``ppt_com/export.py``. Importing this
module off Windows gives the constants and nothing that touches the system.
"""

import ctypes
import logging
import sys
import threading
import time
from contextlib import contextmanager
from dataclasses import dataclass, field

logger = logging.getLogger(__name__)

CF_TEXT = 1
CF_BITMAP = 2
CF_METAFILEPICT = 3
CF_PALETTE = 9
CF_DIB = 8
CF_ENHMETAFILE = 14
CF_HDROP = 15
CF_DIBV5 = 17
CF_OWNERDISPLAY = 0x0080
CF_DSPBITMAP = 0x0082
CF_DSPMETAFILEPICT = 0x0083
CF_DSPENHMETAFILE = 0x008E
CF_GDIOBJFIRST = 0x0300
CF_GDIOBJLAST = 0x03FF

GMEM_MOVEABLE = 0x0002
GHND = 0x0042  # GMEM_MOVEABLE | GMEM_ZEROINIT

_NAMES = {
    1: "CF_TEXT", 2: "CF_BITMAP", 3: "CF_METAFILEPICT", 4: "CF_SYLK", 5: "CF_DIF",
    6: "CF_TIFF", 7: "CF_OEMTEXT", 8: "CF_DIB", 9: "CF_PALETTE", 10: "CF_PENDATA",
    11: "CF_RIFF", 12: "CF_WAVE", 13: "CF_UNICODETEXT", 14: "CF_ENHMETAFILE",
    15: "CF_HDROP", 16: "CF_LOCALE", 17: "CF_DIBV5", 0x80: "CF_OWNERDISPLAY",
    0x81: "CF_DSPTEXT", 0x82: "CF_DSPBITMAP", 0x83: "CF_DSPMETAFILEPICT",
    0x8E: "CF_DSPENHMETAFILE",
}

# Handles that are no memory block and cannot be copied.
_HANDLES = {CF_BITMAP, CF_METAFILEPICT, CF_PALETTE, CF_OWNERDISPLAY,
            CF_DSPBITMAP, CF_DSPMETAFILEPICT, CF_DSPENHMETAFILE}

if sys.platform == "win32":
    import ctypes.wintypes as wt

    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    gdi32 = ctypes.windll.gdi32

    OpenClipboard = user32.OpenClipboard
    OpenClipboard.argtypes = [wt.HWND]
    OpenClipboard.restype = wt.BOOL
    CloseClipboard = user32.CloseClipboard
    CloseClipboard.restype = wt.BOOL
    EmptyClipboard = user32.EmptyClipboard
    EmptyClipboard.restype = wt.BOOL
    EnumClipboardFormats = user32.EnumClipboardFormats
    EnumClipboardFormats.argtypes = [wt.UINT]
    EnumClipboardFormats.restype = wt.UINT
    GetClipboardData = user32.GetClipboardData
    GetClipboardData.argtypes = [wt.UINT]
    GetClipboardData.restype = wt.HANDLE
    SetClipboardData = user32.SetClipboardData
    SetClipboardData.argtypes = [wt.UINT, wt.HANDLE]
    SetClipboardData.restype = wt.HANDLE
    GetClipboardFormatName = user32.GetClipboardFormatNameW
    GetClipboardFormatName.argtypes = [wt.UINT, wt.LPWSTR, ctypes.c_int]
    GetClipboardFormatName.restype = ctypes.c_int

    GlobalAlloc = kernel32.GlobalAlloc
    GlobalAlloc.argtypes = [wt.UINT, ctypes.c_size_t]
    GlobalAlloc.restype = wt.HGLOBAL
    GlobalLock = kernel32.GlobalLock
    GlobalLock.argtypes = [wt.HGLOBAL]
    GlobalLock.restype = ctypes.c_void_p
    GlobalUnlock = kernel32.GlobalUnlock
    GlobalUnlock.argtypes = [wt.HGLOBAL]
    GlobalFree = kernel32.GlobalFree
    GlobalFree.argtypes = [wt.HGLOBAL]
    GlobalSize = kernel32.GlobalSize
    GlobalSize.argtypes = [wt.HGLOBAL]
    GlobalSize.restype = ctypes.c_size_t

    GetEnhMetaFileBits = gdi32.GetEnhMetaFileBits
    GetEnhMetaFileBits.argtypes = [wt.HANDLE, wt.UINT, ctypes.c_void_p]
    GetEnhMetaFileBits.restype = wt.UINT
    SetEnhMetaFileBits = gdi32.SetEnhMetaFileBits
    SetEnhMetaFileBits.argtypes = [wt.UINT, ctypes.c_void_p]
    SetEnhMetaFileBits.restype = wt.HANDLE
    DeleteEnhMetaFile = gdi32.DeleteEnhMetaFile
    DeleteEnhMetaFile.argtypes = [wt.HANDLE]

    CreateWindowEx = user32.CreateWindowExW
    CreateWindowEx.argtypes = [wt.DWORD, wt.LPCWSTR, wt.LPCWSTR, wt.DWORD,
                               ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                               wt.HWND, wt.HMENU, wt.HINSTANCE, wt.LPVOID]
    CreateWindowEx.restype = wt.HWND
    GetMessage = user32.GetMessageW
    GetMessage.argtypes = [ctypes.POINTER(wt.MSG), wt.HWND, wt.UINT, wt.UINT]
    GetMessage.restype = wt.BOOL
    DispatchMessage = user32.DispatchMessageW
    DispatchMessage.argtypes = [ctypes.POINTER(wt.MSG)]
    IsWindow = user32.IsWindow
    IsWindow.argtypes = [wt.HWND]
    IsWindow.restype = wt.BOOL

# The parent that makes a window message only: never shown, never enumerated.
HWND_MESSAGE = -3

_owner = None
_owner_lock = threading.Lock()


def _run_owner(ready: threading.Event, box: list) -> None:
    hwnd = CreateWindowEx(0, "STATIC", "ppt-mcp clipboard", 0, 0, 0, 0, 0,
                          wt.HWND(HWND_MESSAGE), None, None, None)
    box.append(hwnd)
    ready.set()
    if not hwnd:
        return
    msg = wt.MSG()
    while GetMessage(ctypes.byref(msg), None, 0, 0) > 0:
        DispatchMessage(ctypes.byref(msg))


def owner_window():
    """The window the clipboard is opened with, made once.

    Microsoft documents that a clipboard emptied while opened with no window
    has no owner, and that SetClipboardData then fails. So it is opened with a
    message only window of its own. The window lives on a thread of its own
    that keeps answering messages: Windows sends the owner WM_DESTROYCLIPBOARD
    when another program empties the clipboard, and waits for the answer, so
    an owner on a thread that stopped answering would hang that program.
    Answers None when no window could be made, and the clipboard is then
    opened without one, as before.
    """
    global _owner
    with _owner_lock:
        if _owner is None:
            ready, box = threading.Event(), []
            threading.Thread(target=_run_owner, args=(ready, box), daemon=True,
                             name="clipboard-owner").start()
            ready.wait(5)
            _owner = (box[0] if box else None) or False
            if not _owner:
                logger.warning("No window to own the clipboard; opening it without one")
        return _owner or None


def format_name(fmt: int) -> str:
    """A clipboard format's name, for reports."""
    if fmt in _NAMES:
        return _NAMES[fmt]
    if fmt >= 0xC000 and sys.platform == "win32":
        buffer = ctypes.create_unicode_buffer(256)
        if GetClipboardFormatName(fmt, buffer, 256):
            return buffer.value
    return f"format {fmt}"


@contextmanager
def opened(attempts: int = 20, wait: float = 0.05):
    """Hold the clipboard open. Another program may have it for a moment."""
    owner = owner_window()
    for _ in range(attempts):
        if OpenClipboard(owner):
            break
        time.sleep(wait)
    else:
        raise RuntimeError("The clipboard is in use by another program.")
    try:
        yield
    finally:
        CloseClipboard()


def global_bytes(handle) -> bytes:
    """The contents of a global memory block."""
    size = GlobalSize(handle)
    pointer = GlobalLock(handle)
    if not pointer:
        raise OSError("GlobalLock failed")
    try:
        return ctypes.string_at(pointer, size)
    finally:
        GlobalUnlock(handle)


def global_block(data: bytes):
    """A new global memory block holding ``data``, for SetClipboardData."""
    handle = GlobalAlloc(GHND, max(len(data), 1))
    if not handle:
        raise MemoryError("GlobalAlloc failed")
    pointer = GlobalLock(handle)
    if not pointer:
        GlobalFree(handle)
        raise OSError("GlobalLock failed")
    try:
        ctypes.memmove(pointer, data, len(data))
    finally:
        GlobalUnlock(handle)
    return handle


@dataclass
class SavedClipboard:
    """What was on the clipboard: (format, kind, bytes) in order, and what
    could not be kept."""

    entries: list = field(default_factory=list)
    lost: list = field(default_factory=list)


def plan(formats: list[int]) -> tuple[list[int], list[int]]:
    """Which of ``formats`` to keep, and which cannot be kept.

    Formats Windows makes from another one on the list are neither: it makes
    them again once that one is back.
    """
    present = set(formats)
    keep, lost = [], []
    for fmt in formats:
        if fmt == CF_BITMAP and present & {CF_DIB, CF_DIBV5}:
            continue
        if fmt == CF_PALETTE and present & {CF_DIB, CF_DIBV5}:
            continue
        if fmt == CF_METAFILEPICT and CF_ENHMETAFILE in present:
            continue
        if fmt in _HANDLES or CF_GDIOBJFIRST <= fmt <= CF_GDIOBJLAST:
            lost.append(fmt)
            continue
        keep.append(fmt)
    return keep, lost


def save() -> SavedClipboard:
    """Copy everything on the clipboard that can be put back."""
    saved = SavedClipboard()
    with opened():
        formats = []
        fmt = EnumClipboardFormats(0)
        while fmt:
            formats.append(fmt)
            fmt = EnumClipboardFormats(fmt)
        keep, lost = plan(formats)
        saved.lost.extend(format_name(f) for f in lost)
        for fmt in keep:
            handle = GetClipboardData(fmt)
            if not handle:
                # The program that put it there could not render it.
                saved.lost.append(format_name(fmt))
                continue
            try:
                if fmt == CF_ENHMETAFILE:
                    size = GetEnhMetaFileBits(handle, 0, None)
                    buffer = ctypes.create_string_buffer(size)
                    GetEnhMetaFileBits(handle, size, buffer)
                    saved.entries.append((fmt, "emf", buffer.raw))
                else:
                    saved.entries.append((fmt, "global", global_bytes(handle)))
            except Exception:  # noqa: BLE001 - reported, not raised
                logger.debug("Could not copy clipboard format %s", fmt, exc_info=True)
                saved.lost.append(format_name(fmt))
    return saved


def restore(saved: SavedClipboard) -> list[str]:
    """Put back what ``save`` copied, in place of whatever is there now.

    Answers the formats that could not be kept, at saving or now.
    """
    lost = list(saved.lost)
    with opened():
        # Opened with owner_window(), so the clipboard has an owner once it is
        # emptied and SetClipboardData takes what is handed to it.
        EmptyClipboard()
        for fmt, kind, data in saved.entries:
            if kind == "emf":
                handle = SetEnhMetaFileBits(len(data), data)
                if not handle or not SetClipboardData(fmt, handle):
                    if handle:
                        DeleteEnhMetaFile(handle)
                    lost.append(format_name(fmt))
                continue
            try:
                handle = global_block(data)
            except Exception:  # noqa: BLE001 - reported, not raised
                lost.append(format_name(fmt))
                continue
            if not SetClipboardData(fmt, handle):
                GlobalFree(handle)
                lost.append(format_name(fmt))
    return lost


@contextmanager
def kept_clipboard():
    """Run a copy and paste with the user's clipboard put back afterwards.

    Yields a list that holds, once the block is done, the names of the formats
    that could not be put back; empty when everything was.
    """
    saved = save()
    lost: list[str] = []
    try:
        yield lost
    finally:
        try:
            lost.extend(restore(saved))
        except Exception as e:  # noqa: BLE001 - reported by the caller
            logger.warning("Could not put the clipboard back: %s", e)
            lost.append(f"all ({e})")
