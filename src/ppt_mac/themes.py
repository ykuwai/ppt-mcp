"""Theme and header footer tools, on Apple Events.

Mirrors ``ppt_com/themes.py``. Same function names, same signatures, same
returned shapes.

Two things are worth knowing before reading on.

**Theme colours work, and the route is not the one the Windows code takes.**
Windows walks ``Designs(n).SlideMaster.Theme``. Here ``presentation.designs``
counts zero on an ordinary deck, and asking ``designs[1]`` for anything answers
-1728, so that walk is a dead end. ``presentation.slide_master.theme.theme
color scheme`` is not, and all twelve colours read and write through it. An
earlier note in ``MACOS_PORT.md`` had this recorded as unreachable, which was
wrong.

**Colours here are lists, not packed numbers.** A ``theme color``'s ``RGB`` is
``[r, g, b]``, while every tool in this project speaks the BGR integer Windows
uses. ``utils/color`` converts in both directions, so the numbers a caller sees
are the same on both platforms.
"""

import logging
import os
import shutil

from appscript.reference import CommandError

from backend.mac_ae import count, ppt, stage_into_container
from backend.unsupported import refusal as _refusal
from ppt_com.constants import msoTrue  # noqa: F401  (kept for signature parity)
from utils.color import int_to_rgb, rgb_list_to_hex

logger = logging.getLogger(__name__)

# The twelve theme colours, in the order Windows numbers them. macOS reaches
# them as elements rather than by a call, and the order agrees, which was
# checked against `theme color scheme index` on each one.
_THEME_COLOR_COUNT = 12


def _theme_color_names():
    """The twelve names Windows reports, fetched at call time.

    Imported inside the function rather than at module scope. `ppt_com/themes.py`
    imports this module at the bottom of its own file, so importing it back up
    here lets an "import ppt_mac.themes first" ordering run that swap block
    against a half built module, and `use_mac_impls` then silently finds nothing
    to swap. By call time both are fully loaded.
    """
    from ppt_com.themes import THEME_COLOR_NAMES

    return THEME_COLOR_NAMES


def _theme_colors(pres):
    """The deck's theme colour scheme.

    Through the slide master, never through `designs`. A deck's `designs`
    collection counts zero here and `designs[1]` answers -1728, so the Windows
    walk has nowhere to start.
    """
    return pres.slide_master.theme.theme_color_scheme


# ---------------------------------------------------------------------------
# Apple Event implementation functions
# ---------------------------------------------------------------------------
def _apply_theme_impl(theme_path):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()

    abs_path = os.path.abspath(os.path.expanduser(theme_path))
    if not os.path.exists(abs_path):
        raise ValueError(f"Theme file not found: {abs_path}")

    staged = stage_into_container(abs_path)
    before = [
        rgb_list_to_hex(_theme_colors(pres).theme_colors[i].RGB())
        for i in range(1, _THEME_COLOR_COUNT + 1)
    ]
    app.apply_theme(pres, file_name=staged)

    # Nothing is trusted because it did not raise. A theme that landed changes
    # the colour scheme, so the scheme is the check.
    after = [
        rgb_list_to_hex(_theme_colors(pres).theme_colors[i].RGB())
        for i in range(1, _THEME_COLOR_COUNT + 1)
    ]
    if after == before:
        return _refusal(
            "ppt_apply_theme",
            "PowerPoint reported no error and the deck's theme colours did not "
            "change, which is the silent no-op recorded in MACOS_PORT section "
            f"5. The file was staged at {staged}, so a theme PowerPoint cannot "
            "read is the likely cause.",
            ["ppt_set_theme_colors"],
        )

    return {
        "success": True,
        "theme_path": abs_path,
    }


def _get_theme_colors_impl():
    pres = ppt._get_pres_impl()
    scheme = _theme_colors(pres)

    colors = []
    for i in range(1, _THEME_COLOR_COUNT + 1):
        colors.append({
            "index": i,
            "name": _theme_color_names()[i],
            "color_hex": rgb_list_to_hex(scheme.theme_colors[i].RGB()),
        })

    return {
        "success": True,
        "colors": colors,
        # Windows counts designs and gets at least one. This dictionary's
        # `designs` collection is empty on an ordinary deck, so the honest
        # answer here is the one master this reads from.
        "design_count": max(count(pres.designs), 1),
    }


def _set_theme_colors_impl(color_map):
    """Set individual theme colours.

    Windows loops over every design's slide master. There are no designs to
    loop over here, so this writes the one scheme the deck has, which is the
    same scheme every slide resolves its theme colours through.
    """
    pres = ppt._get_pres_impl()
    scheme = _theme_colors(pres)

    for idx, bgr in color_map.items():
        r, g, b = int_to_rgb(bgr)
        scheme.theme_colors[int(idx)].RGB.set([r, g, b])

    # Read back rather than echoed, because a theme colour that did not take is
    # invisible otherwise.
    changed = []
    unchanged = []
    for idx, bgr in color_map.items():
        wanted = rgb_list_to_hex(list(int_to_rgb(bgr)))
        got = rgb_list_to_hex(scheme.theme_colors[int(idx)].RGB())
        entry = {"name": _theme_color_names()[int(idx)], "color_hex": wanted}
        if got == wanted:
            changed.append(entry)
        else:
            unchanged.append({**entry, "actual_hex": got})

    result = {
        "success": True,
        "changed": changed,
        "changed_count": len(changed),
        "designs_updated": 1,
    }
    if unchanged:
        result["warnings"] = [
            "PowerPoint kept a different colour for "
            + ", ".join(f"{u['name']} (now {u['actual_hex']})" for u in unchanged)
        ]
    return result


def _set_headers_footers_impl(
    footer_text, footer_visible, slide_number_visible,
    date_visible, date_format, date_fixed_text,
):
    """Set the footer, slide number and date on every slide.

    Per slide, exactly as on Windows. PowerPoint for Mac has no presentation
    wide headers and footers object, and a layout that has no footer
    placeholder answers -1728 for it, which is not a failure of the call and
    does not stop the slides that do have one.
    """
    # Asked for and not kept, so that a dead connection fails here rather
    # than part way through the slides.
    ppt._get_app_impl()
    pres = ppt._get_pres_impl()

    slide_count = count(pres.slides)
    skipped = 0
    for i in range(1, slide_count + 1):
        hf = pres.slides[i].headers_and_footers
        try:
            if footer_visible is not None:
                hf.footer.visible.set(bool(footer_visible))
            if footer_text is not None:
                # Visible first, then the text. A hidden footer answers -1728
                # for its text on some layouts, which is the same ordering the
                # Windows code settled on for its own reasons.
                hf.footer.visible.set(True)
                hf.footer.header_footer_text.set(footer_text)
            if slide_number_visible is not None:
                hf.slide_number.visible.set(bool(slide_number_visible))
            if date_visible is not None:
                hf.date_and_time.visible.set(bool(date_visible))
            if date_format is not None:
                hf.date_and_time.date_format.set(
                    _date_format_keyword(date_format)
                )
            if date_fixed_text is not None:
                hf.date_and_time.use_date_format.set(False)
                hf.date_and_time.header_footer_text.set(date_fixed_text)
        except CommandError:
            # The layout has no placeholder for one of these. Counted rather
            # than swallowed, so the caller can see it happened.
            skipped += 1
            logger.debug("Slide %s would not take a header or footer", i, exc_info=True)

    result = {
        "success": True,
        "slides_updated": slide_count - skipped,
    }
    if skipped:
        result["warnings"] = [
            f"{skipped} of {slide_count} slides have a layout with no "
            "placeholder for what was asked for, so those were left alone."
        ]
    return result


def _date_format_keyword(date_format):
    """Translate a PpDateTimeFormat constant into the macOS enumerator.

    Kept here rather than inline so the ValueError names the argument the
    caller passed rather than a table.
    """
    from backend.mac_enums import PpDateTimeFormat, to_keyword

    return to_keyword(PpDateTimeFormat, date_format, "date format")
