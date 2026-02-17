"""Navigation helpers for PowerPoint COM automation."""

import logging

logger = logging.getLogger(__name__)


def goto_slide(app, slide_index: int) -> None:
    """Navigate the active window to the specified slide.

    Call this at the start of write operations so the user can see
    the slide being edited.  Silently ignores errors (e.g. during
    slideshow mode or when no window is available).

    Args:
        app: PowerPoint Application COM object.
        slide_index: 1-based slide index to navigate to.
    """
    try:
        app.ActiveWindow.View.GotoSlide(slide_index)
    except Exception:
        pass
