"""Slide show tools, on Apple Events.

Mirrors ``ppt_com/slideshow.py``. Same function names, same signatures, same
returned shapes; what differs is the walk through PowerPoint's object model.

Three things about running a show on this side are worth knowing before reading
on.

**The window ``run slide show`` hands back is not used.** The command answers
with a ``slide show window`` and a reference PowerPoint hands back is not
trusted here, so the window is counted and then indexed instead, exactly as
MACOS_PORT section 5.1 says to. A count of zero afterwards is how a show that
never started announces itself, and that is a refusal rather than a report of
success.

**Next and previous are commands on the slide show view.** Windows drives
``View.Next``. The ``next`` and ``previous`` in this dictionary take a
``presenter tool``, which lives on a ``presenter view window`` that the
application does not list as an element, so nothing can reach it. ``go to next
slide`` and ``go to previous slide`` take the slide show view and are what
these tools use.

**Going to a slide by number is attempted rather than promised.** ``go to
slide`` is declared on ``view``, the class behind a document window, and a
running show has a ``slide show view``, which is a separate class that does not
inherit from it. So the call is made and the position is read back, and a show
that did not move gets a refusal naming the two tools that do work.

Nothing here calls ``goto_slide``. A running show owns the screen, and the
Windows originals do not navigate the editing window either.
"""

import logging
import struct
import time

from appscript import k
from appscript.reference import CommandError

from backend.mac_ae import (
    count,
    count_of,
    error_number,
    ppt,
    windows_constant as _windows_constant,
)
from backend.mac_enums import (
    PpSlideShowRangeType,
    PpSlideShowState,
    PpSlideShowType,
    to_keyword,
)
from backend.unsupported import refusal as _refusal
from ppt_com.constants import (
    SHOW_TYPE_NAMES,
    SLIDESHOW_STATE_NAMES,
    ppShowAll,
    ppShowSlideRange,
    ppShowTypeSpeaker,
)

logger = logging.getLogger(__name__)

_RANGE_TYPES = PpSlideShowRangeType

# The pointer type has no generated table because Windows and macOS name it
# from different vocabularies, but the four macOS enumerators carry the codes
# 0x00d20000 through 0x00d20003 and the Windows constants run 0 through 3, so
# the two agree on the ordinal. Windows' auto arrow and eraser have no macOS
# word and are simply absent, which reads back as None rather than as a guess.
_POINTER_TYPES = {
    0: k.slide_show_pointer_none,            # ppSlideShowPointerNone
    1: k.slide_show_pointer_arrow,           # ppSlideShowPointerArrow
    2: k.slide_show_pointer_pen,             # ppSlideShowPointerPen
    3: k.slide_show_pointer_always_hidden,   # ppSlideShowPointerAlwaysHidden
}

# How long PowerPoint is given to open or close a show window before the count
# is taken as the truth. A show is a window appearing on screen, so it is not
# instant, and reporting a silent no-op for something that simply had not
# happened yet would be worse than waiting.
_SETTLE_SECONDS = 0.5
_SETTLE_STEPS = 5


def _show_window_count(app) -> int:
    """How many slide show windows are open.

    Asked of the application rather than of the collection. ``count_of`` sends
    the safer of the two forms, and the difference between them has already
    cost an open deck once; see ``backend.mac_ae``. Materialising the
    collection is only the fallback for a build that will not answer the first
    question at all.
    """
    try:
        return count_of(app, k.slide_show_window)
    except CommandError:
        return count(app.slide_show_windows)


def _show_window(app):
    """The running show's window, built here rather than taken from a command."""
    if _show_window_count(app) < 1:
        return None
    return app.slide_show_windows[1]


def _settle(app, want_running: bool) -> bool:
    """Wait briefly for a show window to appear or to go, then say what it did."""
    for _ in range(_SETTLE_STEPS):
        if (_show_window_count(app) > 0) == want_running:
            return True
        time.sleep(_SETTLE_SECONDS / _SETTLE_STEPS)
    return (_show_window_count(app) > 0) == want_running


def _state_name(view) -> str:
    """The show's state under the name Windows gives it."""
    word = view.slide_state()
    value = _windows_constant(PpSlideShowState, word)
    if value is None:
        return f"unknown({word})"
    return SLIDESHOW_STATE_NAMES.get(value, f"unknown({value})")


# ---------------------------------------------------------------------------
# Apple Event implementation functions
# ---------------------------------------------------------------------------
def _slideshow_start_impl(start_slide, end_slide, loop, show_type) -> dict:
    # Imported lazily. ppt_com/slideshow.py imports this module at the bottom
    # of its own file, so importing it back at module scope would let an
    # "import ppt_mac.slideshow first" ordering run that swap block against a
    # module that has defined nothing yet, and the swap would silently not
    # happen. By call time both modules are fully loaded.
    from ppt_com.slideshow import SHOW_TYPE_MAP

    app = ppt._get_app_impl()
    if count(app.presentations) == 0:
        raise RuntimeError(
            "No presentation is open. "
            "Use ppt_create_presentation or ppt_open_presentation first."
        )
    pres = ppt._get_pres_impl()

    total_slides = count(pres.slides)
    if total_slides == 0:
        raise RuntimeError("Presentation has no slides.")

    # Everything is validated before anything is written, so a rejected
    # argument cannot leave the deck's show settings half changed.
    if show_type is not None:
        type_key = show_type.lower().strip()
        if type_key not in SHOW_TYPE_MAP:
            raise ValueError(
                f"Unknown show_type '{show_type}'. Supported: {list(SHOW_TYPE_MAP.keys())}"
            )
        wanted_type = SHOW_TYPE_MAP[type_key]
    else:
        wanted_type = ppShowTypeSpeaker

    whole_deck = start_slide is None and end_slide is None
    if whole_deck:
        actual_start, actual_end = 1, total_slides
    else:
        actual_start = start_slide if start_slide is not None else 1
        actual_end = end_slide if end_slide is not None else total_slides
        if actual_start < 1 or actual_start > total_slides:
            raise ValueError(
                f"start_slide {actual_start} out of range (1-{total_slides})"
            )
        if actual_end < actual_start or actual_end > total_slides:
            raise ValueError(
                f"end_slide {actual_end} out of range ({actual_start}-{total_slides})"
            )

    settings = pres.slide_show_settings
    settings.show_type.set(to_keyword(PpSlideShowType, wanted_type, "show type"))

    if whole_deck:
        # No range asked for means the whole deck, which also clears a range
        # left behind by an earlier call.
        settings.range_type.set(
            to_keyword(_RANGE_TYPES, ppShowAll, "slide show range type")
        )
    else:
        settings.range_type.set(
            to_keyword(_RANGE_TYPES, ppShowSlideRange, "slide show range type")
        )
        settings.starting_slide.set(actual_start)
        settings.ending_slide.set(actual_end)

    if loop is not None:
        # A plain boolean here. The msoTrue and msoFalse the Windows code sets
        # are a COM tri-state and mean nothing to Apple Events.
        settings.loop_until_stopped.set(bool(loop))

    try:
        settings.run_slide_show()
    except struct.error:
        # `run slide show` answers with a slide show window, and appscript
        # cannot decode what PowerPoint sends back, so it raises
        # "unpack requires a buffer of 4 bytes" out of `struct` rather than an
        # Apple Event error. The show does start. `duplicate` fails the same
        # way in slides.py, which is where this was recognised.
        #
        # The answer was never used, so nothing is lost by dropping it. What
        # decides whether the show started is the check below, which was
        # already here and simply never got to run.
        logger.debug(
            "run slide show answered with something appscript could not "
            "decode. Checking whether the show opened instead."
        )

    # The window the command answered with is discarded. It is fetched again by
    # counting and indexing, and no window at all is the silent no-op.
    if not _settle(app, want_running=True):
        return _refusal(
            "ppt_slideshow_start",
            "PowerPoint reported success and no slide show window opened, "
            "which is the silent no-op recorded in MACOS_PORT section 5.",
        )
    window = _show_window(app)
    view = window.slideshow_view

    # macOS has a fourth show type, `slide show type presenter`, that Windows
    # has no constant for, so the lookup can come back with nothing. "unknown"
    # is the honest answer for it rather than the nearest Windows name.
    started_as = _windows_constant(PpSlideShowType, settings.show_type())
    return {
        "success": True,
        "show_type": SHOW_TYPE_NAMES.get(started_as, "unknown"),
        "current_slide": view.current_show_position(),
        "total_slides": total_slides,
        "start_slide": actual_start,
        "end_slide": actual_end,
    }


def _slideshow_stop_impl() -> dict:
    app = ppt._get_app_impl()
    window = _show_window(app)
    if window is None:
        return {"success": True, "message": "No slide show was running."}

    window.slideshow_view.exit_slide_show()

    if not _settle(app, want_running=False):
        return _refusal(
            "ppt_slideshow_stop",
            "PowerPoint reported success and the slide show window is still "
            "open, which is the silent no-op recorded in MACOS_PORT section 5. "
            "Pressing Escape in the show ends it.",
        )
    return {"success": True, "message": "Slide show ended."}


def _slideshow_next_impl() -> dict:
    app = ppt._get_app_impl()
    window = _show_window(app)
    if window is None:
        raise RuntimeError("No slide show is running.")

    view = window.slideshow_view
    view.go_to_next_slide()
    return {
        "success": True,
        "current_slide": view.current_show_position(),
        "state": _state_name(view),
    }


def _slideshow_previous_impl() -> dict:
    app = ppt._get_app_impl()
    window = _show_window(app)
    if window is None:
        raise RuntimeError("No slide show is running.")

    view = window.slideshow_view
    view.go_to_previous_slide()
    return {
        "success": True,
        "current_slide": view.current_show_position(),
        "state": _state_name(view),
    }


def _slideshow_goto_impl(slide_index: int) -> dict:
    app = ppt._get_app_impl()
    window = _show_window(app)
    if window is None:
        raise RuntimeError("No slide show is running.")

    view = window.slideshow_view

    # `go to slide` is declared on `view`, which is a document window's view,
    # and `slide show view` is a separate class with no inheritance between
    # them. There is no other way to send a running show to a slide, so it is
    # tried and then the position is read back.
    try:
        view.go_to_slide(number=slide_index)
    except CommandError as exc:
        return _refusal(
            "ppt_slideshow_goto",
            f"PowerPoint answered {error_number(exc)} to `go to slide` against "
            "a running show. The command is declared on `view`, the class "
            "behind a document window, and a show has a `slide show view`, "
            "which is a different class, so the show cannot be sent to a "
            "slide by number.",
            ["ppt_slideshow_next", "ppt_slideshow_previous"],
        )

    landed = view.current_show_position()
    if landed != slide_index:
        return _refusal(
            "ppt_slideshow_goto",
            f"PowerPoint reported success and the show is still on slide "
            f"{landed}, which is the silent no-op recorded in MACOS_PORT "
            "section 5. `go to slide` is declared on a document window's view "
            "rather than on a slide show view.",
            ["ppt_slideshow_next", "ppt_slideshow_previous"],
        )

    return {
        "success": True,
        "current_slide": landed,
        "state": _state_name(view),
    }


def _slideshow_get_status_impl() -> dict:
    app = ppt._get_app_impl()
    window = _show_window(app)
    if window is None:
        return {"running": False}

    view = window.slideshow_view
    state_word = view.slide_state()
    state = _windows_constant(PpSlideShowState, state_word)

    return {
        "running": True,
        "current_slide": view.current_show_position(),
        "state": state,
        "state_name": SLIDESHOW_STATE_NAMES.get(state, f"unknown({state_word})"),
        "pointer_type": _windows_constant(_POINTER_TYPES, view.pointer_type()),
    }
