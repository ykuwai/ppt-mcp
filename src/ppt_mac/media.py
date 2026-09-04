"""Media tools, on Apple Events.

Mirrors ``ppt_com/media.py``. All three tools refuse, and the reasons differ
enough to be worth stating separately rather than as one blanket note.

**There is no way to put a media file on a slide.** The dictionary has a
``media object`` class and a ``media2 object`` class, both inheriting ``shape``,
and between them they carry ``file name``, ``link to file`` and
``save with document``. Every one of those is read only, neither class is a
declared element of anything, and there is no ``insert from file``, no
``add media object`` and nothing else that takes a path and makes a shape. So a
media shape that is already in the deck can be read and recognised, and one
that is not cannot be put there.

**``import sound file`` is not the way in either.** It takes a ``sound effect``,
which is the little object hanging off a slide's transition, a shape's action
setting or a shape's animation settings, and all three of those properties are
read only, so the sound it loads belongs to an effect rather than to a shape on
the slide. It is a real capability and it is not the one ``ppt_add_audio``
offers.

**The sandbox would still be in the way.** Even given an insert route,
PowerPoint may only read where it has a grant, so the file would have to be
copied into ``EXPORT_STAGING_DIR`` first, the way ``export.py`` moves its
output the other way. That breaks ``link_to_file`` by construction, because the
link would point at the staged copy rather than at the caller's file.

**Playback settings are refused on purpose, not for want of properties.**
``play settings`` really does carry ``loop until stopped`` and
``hide while not playing``, the two Windows ``PlaySettings`` fields these tools
use. The only route to it is ``shape.animation settings``, and MACOS_PORT
section 5.2 measured writes to four of the five properties on that container
silently flattening every animation on the slide. Volume, mute, trim and fade
have no counterpart anywhere in the dictionary, so honouring the argument that
is reachable would mean risking the deck for two settings out of eight.

Nothing here touches a file, so nothing stages a path, and nothing edits a
slide, so nothing calls ``goto_slide``.
"""

import logging

from backend.unsupported import refusal as _refusal

logger = logging.getLogger(__name__)

# The two sentences every refusal in this module rests on, kept in one place so
# that a reader who meets one of them twice recognises it as the same finding.
_NO_INSERT = (
    "PowerPoint for Mac has no command that puts a media file on a slide. The "
    "`media object` and `media2 object` classes exist but are elements of no "
    "container, their `file name` is read only, and there is no insert from "
    "file. `import sound file` loads a sound into an existing sound effect, "
    "which belongs to a transition or to a shape's action, not to a shape of "
    "its own."
)
_SANDBOX = (
    "PowerPoint is also sandboxed and may only read where it has a grant, so "
    "even given a route the file would have to be copied into its container "
    "first, which is what link_to_file cannot survive."
)


# ---------------------------------------------------------------------------
# Apple Event implementation functions
# ---------------------------------------------------------------------------
def _add_video_impl(slide_index, file_path, left, top, width, height, link_to_file):
    """Refuse, because nothing in the dictionary inserts a movie.

    Windows drives ``Shapes.AddMediaObject2``, which takes a path and a couple
    of flags. There is no such command here, and no writable ``file name`` to
    reach around it with.
    """
    return _refusal(
        "ppt_add_video",
        f"{_NO_INSERT} {_SANDBOX}",
        [
            "Insert the movie once by hand in PowerPoint, then position it "
            "with ppt_update_shape",
            "ppt_add_picture",
        ],
    )


def _add_audio_impl(slide_index, file_path, left, top, width, height, link_to_file):
    """Refuse, for the same reason as video, plus the near miss worth naming.

    ``import sound file`` is close enough to look like the answer and is not,
    so the refusal says what it actually does rather than leaving someone to
    find out.
    """
    return _refusal(
        "ppt_add_audio",
        f"{_NO_INSERT} {_SANDBOX}",
        [
            "Insert the sound once by hand in PowerPoint, then position it "
            "with ppt_update_shape",
            "ppt_set_slide_transition, which is where a sound can be attached "
            "without a shape",
        ],
    )


def _set_media_settings_impl(
    slide_index, shape_name_or_index,
    volume, muted, start_point, end_point,
    fade_in, fade_out, loop, hide_while_not_playing,
):
    """Refuse, because six of the eight settings are absent and two are unsafe.

    The two that exist are ``loop until stopped`` and ``hide while not
    playing`` on ``play settings``, reached only through
    ``shape.animation settings``. Writing to that container was measured to
    flatten the slide's animations, so this returns without asking PowerPoint
    for anything at all.
    """
    return _refusal(
        "ppt_set_media_settings",
        "PowerPoint for Mac exposes no volume, mute, trim or fade for a media "
        "shape anywhere in its dictionary. Loop and hide while not playing do "
        "exist, on the `play settings` object, but the only route to it is "
        "through `shape animation settings`, and writing to that container was "
        "measured to flatten every animation on the slide (MACOS_PORT section "
        "5.2), so they are deliberately left alone rather than paid for with "
        "the rest of the deck.",
        [
            "ppt_get_shape_info, which reports a media shape's type",
            "ppt_update_shape, which moves and resizes it",
        ],
    )
