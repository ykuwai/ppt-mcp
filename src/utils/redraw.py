"""Win32 redraw suppression for PowerPoint automation.

PowerPoint has no working ``Application.ScreenUpdating`` (unlike Excel/Word),
so to avoid visible flicker while building/styling shapes we suppress painting
of the PowerPoint frame window at the Win32 level.

We use ``LockWindowUpdate`` — the long-standing community technique for
PowerPoint automation. Unlike ``WM_SETREDRAW`` on a top-level application
window (which causes a jarring full-window WHITE repaint of the whole frame,
ribbon and panes included, when redraw is re-enabled), ``LockWindowUpdate``
accumulates the regions that changed while locked and repaints only those on
unlock — so the ribbon/canvas are not erased to white.

Caveats (per Microsoft docs / Raymond Chen): only ONE window system-wide can be
update-locked at a time, and it is really meant for drag/drop feedback. For our
short, single-threaded COM operations that is acceptable. The lock is always
released in ``__exit__`` so a failure mid-operation never leaves the window
frozen.

If pywin32's GUI modules or the PowerPoint window are unavailable, the context
manager degrades gracefully to a no-op (no lock, no error).
"""

import logging

logger = logging.getLogger(__name__)

try:
    import win32gui

    _WIN32_AVAILABLE = True
except ImportError:  # pragma: no cover - non-Windows / missing pywin32
    _WIN32_AVAILABLE = False

# PowerPoint 2010+ top-level frame window class.
_PPT_FRAME_CLASS = "PPTFrameClass"


def _find_ppt_hwnd():
    """Return the PowerPoint frame window handle, or 0 if not found."""
    if not _WIN32_AVAILABLE:
        return 0
    try:
        return win32gui.FindWindow(_PPT_FRAME_CLASS, None)
    except Exception:
        return 0


class FrozenRedraw:
    """Context manager that suppresses PowerPoint window painting.

    Locks the frame window on enter (``LockWindowUpdate``) and releases it on
    exit. Intermediate states (default theme fill, scroll-to-selection) drawn
    inside the block are therefore not shown — on unlock the system repaints
    only the regions that changed, so the final result appears without the
    incremental flicker and without a full-window white flash.

    Safe to use even when the window can't be located: it becomes a no-op.
    """

    def __init__(self):
        self.hwnd = 0
        self._locked = False

    def __enter__(self):
        # Resolve the window at enter time (not construction) so a PowerPoint
        # launched between construction and use is still found.
        if _WIN32_AVAILABLE:
            self.hwnd = _find_ppt_hwnd()
        if self.hwnd and _WIN32_AVAILABLE:
            try:
                # Returns nonzero on success. Only one window can be locked
                # system-wide; if another lock is active this fails and we
                # simply proceed without suppression.
                self._locked = bool(win32gui.LockWindowUpdate(self.hwnd))
            except Exception:
                logger.warning("Failed to lock PowerPoint window for redraw")
                self._locked = False
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._locked:
            try:
                # Passing 0 unlocks; the system repaints the accumulated
                # invalid regions automatically.
                win32gui.LockWindowUpdate(0)
            except Exception:
                logger.warning("Failed to unlock PowerPoint window for redraw")
            finally:
                self._locked = False
        return False  # never suppress exceptions
