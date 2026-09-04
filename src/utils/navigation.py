"""Navigation helpers for PowerPoint automation."""

import logging

from backend import IS_MACOS, ppt

logger = logging.getLogger(__name__)


def goto_slide(app, slide_index: int) -> None:
    """Navigate the active window to the specified slide.

    Call this at the start of write operations so the user can see
    the slide being edited.  Silently ignores errors (e.g. during
    slideshow mode or when no window is available).

    Args:
        app: PowerPoint Application object.
        slide_index: 1-based slide index to navigate to.
    """
    try:
        if IS_MACOS:
            # Navigate the target deck's own window rather than the frontmost
            # one. There is an `active window` here too, but it raises whenever
            # PowerPoint's start gallery is in front, and it can belong to a
            # different deck than the one being edited.
            pres = ppt._get_pres_impl()
            pres.document_windows[1].view.go_to_slide(number=slide_index)
        else:
            app.ActiveWindow.View.GotoSlide(slide_index)
    except Exception:
        pass
