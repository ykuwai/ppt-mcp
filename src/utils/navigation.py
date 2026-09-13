"""Navigation helpers for PowerPoint automation."""

import logging

from backend import IS_MACOS, ppt

logger = logging.getLogger(__name__)


def goto_slide(app, slide_index: int) -> None:
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
    """
    try:
        if IS_MACOS:
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
