"""Navigation helpers for PowerPoint automation."""

import logging
import os

from backend import IS_MACOS, ppt, vba

logger = logging.getLogger(__name__)


def _follow_edits(value):
    """Whether a write shows the slide it works on, from PPT_FOLLOW_EDITS."""
    value = value.strip().lower()
    if value in ("on", "true", "1", "yes", ""):
        return True
    if value in ("off", "false", "0", "no"):
        return False
    # Read at import, so a typo must not stop the server from starting.
    logger.warning(
        "PPT_FOLLOW_EDITS=%r is not a setting; use 'on' (the default), or 'off' "
        "to leave the window on the slide the user is looking at. Using 'on'.", value)
    return True


# On, each write moves the deck's window to the slide it works on, so the user
# sees where the model is. Off, the window stays where the user left it, for
# working in the same deck at the same time: the edits land either way.
# ppt_goto_slide moves it in both cases, since moving it is what it is for.
FOLLOW_EDITS = _follow_edits(os.getenv("PPT_FOLLOW_EDITS", "on"))


def goto_slide(app, slide_index: int, force: bool = False) -> None:
    """Navigate the target presentation's window to the specified slide.

    Call this at the start of write operations so the user can see
    the slide being edited.  Silently ignores errors (e.g. during
    slideshow mode or when no window is available).

    The window is driven directly rather than via app.ActiveWindow, and is
    deliberately NOT activated: activating it pulls PowerPoint to the
    foreground and steals focus from whatever the user is doing (issue #183).
    The editor still follows along in the background.

    Args:
        app: PowerPoint Application object, kept for call-site
            compatibility. Both platforms resolve the window from the
            session target rather than from this.
        slide_index: 1-based slide index to navigate to.

    Does nothing while ``FOLLOW_EDITS`` is off, unless ``force``: a tool
    whose work needs the slide on screen, such as selecting shapes, which
    PowerPoint does only on the slide the window shows.
    """
    if not FOLLOW_EDITS and not force:
        return
    try:
        if IS_MACOS and not vba.active():
            # Navigate the target deck's own window rather than the frontmost
            # one. There is an `active window` here too, but it raises whenever
            # PowerPoint's start gallery is in front, and it can belong to a
            # different deck than the one being edited.
            from backend.mac_ae import target_window

            pres = ppt._get_pres_impl()
            # A deck can outlive its window, and `document_windows[1]` then
            # raises a bare -1728. This still swallows it, because showing the
            # slide is a courtesy and no tool should fail for want of it, but
            # going through `target_window` means the call sites that do need a
            # window meet the written explanation rather than assuming this
            # worked.
            target_window(pres).view.go_to_slide(number=slide_index)
        else:
            window = ppt._get_target_window_impl()
            if window is None:
                return
            window.View.GotoSlide(slide_index)
    except Exception:
        pass


def shown_slide(window):
    """The slide a deck's window is showing, or None when it cannot be read.

    For the edits that must move the window to work, a paste on the Mac or a
    selection on Windows: they read it first, so ``return_to`` can put it back.
    """
    try:
        if IS_MACOS and not vba.active():
            return window.view.slide.slide_index()
        return window.View.Slide.SlideIndex
    except Exception:
        return None


def return_to(window, slide_index) -> None:
    """Put the window back on the slide it showed, while FOLLOW_EDITS is off.

    On, the window stays on the slide the edit worked on, as every other write
    leaves it.
    """
    if FOLLOW_EDITS or slide_index is None:
        return
    try:
        if IS_MACOS and not vba.active():
            window.view.go_to_slide(number=slide_index)
        else:
            window.View.GotoSlide(slide_index)
    except Exception:
        pass

