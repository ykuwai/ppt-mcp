"""Navigation helpers for PowerPoint COM automation."""

import logging

from utils.com_wrapper import ppt

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
        app: PowerPoint Application COM object (kept for call-site
            compatibility; the window is resolved from the session target).
        slide_index: 1-based slide index to navigate to.
    """
    try:
        window = ppt._get_target_window_impl()
        if window is None:
            return
        window.View.GotoSlide(slide_index)
    except Exception:
        pass
