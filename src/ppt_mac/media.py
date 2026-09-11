"""Media tools, on Apple Events.

Mirrors ``ppt_com/media.py``. Two of the three tools do the job and the third
does part of it, which is a correction: this module used to refuse all three on
the grounds that nothing in the dictionary put a media file on a slide. That
reading was too cautious. ``media2 object`` is not a declared element of
anything, but ``make`` accepts it anyway.

    app.make(new=k.media2_object, at=slide.end,
             with_properties={k.file_name: staged, k.left_position: l,
                              k.top: t})

An ``.aiff`` and an ``.mp4`` both land as ``shape type media``, and both are in
``ppt/media/`` at their original byte size once the deck is saved. What the
measurements settled, so nobody re-derives them:

**The file has to be staged inside PowerPoint's container first.** It is
sandboxed, and a path outside the container raises the Grant Access sheet; in
issue #191 that cascade closed every open document and took PowerPoint with it.
``stage_into_container`` is the one route, the same one pictures take.

**The staged copy can go as soon as the shape exists.** The embed happens at
insert time rather than at save time, which was worth checking rather than
assuming: the two staged files were deleted, the deck was saved afterwards, and
both media parts were in the archive whole.

**Embedding is the only mode.** ``link to file`` reads back ``missing value``
on a shape that was just made, so a caller asking to link is refused by that
argument's name rather than handed an embedded copy it did not ask for.

**PowerPoint sizes media at 25 points tall**, keeping the aspect ratio of the
file: a 320 by 240 movie arrives 33.3 by 25 and a 640 by 360 one 44.4 by 25.
That is not the silent 25 by 25 autoshape of MACOS_PORT section 5, it is where
media really lands, so the size is reported in ``warnings`` rather than treated
as a failure.

**Of the eight playback settings, two exist and only one of them is safe.**
Volume, mute, trim and fade have no words anywhere in the dictionary. ``loop
until stopped`` and ``hide while not playing`` are on ``play settings``, reached
through ``shape.animation settings``, and the difference between them was
measured on a slide holding three effects including an exit:

* writing ``loop until stopped`` three times left all three effects exactly as
  they were, entrance types and exit flag included;
* writing ``hide while not playing`` **once** dropped the exit effect outright
  and turned a bounce into a plain appear. Writing ``False`` over a value that
  was already ``False`` did it too. Reproduced twice.

So that is MACOS_PORT section 5.2's flattening, reached through a property the
old module assumed was as unsafe as its neighbour. ``loop`` is written freely,
``hide_while_not_playing`` only on a slide with no animations, and on any other
slide the whole call is refused by that argument's name so that dropping it and
calling again works.

One edge that count cannot see, said out loud rather than left implicit. It asks
the main sequence how many effects it holds, and an animation triggered by
clicking a shape does not live there. Counting the timeline's own ``sequence``
elements as well was tried and does not work, because inserting a media shape
creates one of those for its playback, so that gate would refuse the write on
every slide this tool is ever called about. The successful write says so in
``warnings`` instead.
"""

import logging
import os

from appscript import k
from appscript.reference import CommandError

from backend.mac_ae import (
    count,
    count_of,
    elements,
    ppt,
    slide_at as _slide,
    stage_into_container,
)
from backend.unsupported import refusal as _refusal
from ppt_mac.shapes import _get_shape, _verify_created
from utils.navigation import goto_slide

logger = logging.getLogger(__name__)

# The height PowerPoint gives a media shape when no size is asked for. The
# width follows the file's aspect ratio; the height is this whatever it is.
_MEDIA_DEFAULT_HEIGHT = 25

# The six Windows `MediaFormat` fields with no counterpart here, in the order
# `ppt_set_media_settings` takes them, so a refusal names them the way the
# caller wrote them.
_NO_WORDS_FOR = ("volume", "muted", "start_point", "end_point", "fade_in", "fade_out")

_LINK_REFUSAL = (
    "PowerPoint for Mac embeds a media file and cannot link to one. `media2 "
    "object` carries `link to file`, and a shape just made from a file answers "
    "`missing value` for it, so there is nothing to set and nothing to read "
    "back. The file is also staged into PowerPoint's container before it is "
    "read, because the sandbox asks the user to grant access to anywhere else, "
    "so a link would point at a copy rather than at the file named here. Call "
    "again without link_to_file to embed it."
)


def _media_shape_name_warning(names_before, shape_name):
    """Say so when the new shape's name is not unique on the slide.

    PowerPoint names a media shape after its file, so two clips called
    ``intro.mp4`` from two folders both arrive as ``intro``, and every tool that
    addresses a shape by name then finds the first of them. Windows never does
    this, so a caller has no reason to expect it.
    """
    if names_before is None:
        return None
    if shape_name not in names_before:
        return None
    return (
        f"The slide already had a shape called '{shape_name}', and PowerPoint "
        "names a media shape after its file rather than numbering it. Address "
        "this one by its index, or rename the file, because a tool given the "
        "name will act on the first of the two."
    )


def _names_on(slide, expected):
    """Every shape name on the slide in one Apple Event, or None.

    Returned only when there is one name per shape. PowerPoint answers a bulk
    property read short rather than raising, and a short answer would make the
    duplicate name check below quietly stop noticing.
    """
    try:
        names = [str(n) for n in elements(slide.shapes.name)]
    except CommandError:
        return None
    return names if len(names) == expected else None


def _add_media(tool_name, what, slide_index, file_path, left, top,
               width, height, link_to_file):
    """Put a movie or a sound on a slide. The body of both tools.

    Video and audio are one call on Windows too; ``AddMediaObject2`` takes
    either. Here they are one ``make`` of a ``media2 object``, and the only
    thing that differs between the two tools is the word in the refusals.
    """
    # Refused before the view moves and before anything is copied anywhere, so
    # a call that cannot be honoured leaves the deck and the container alone.
    if link_to_file:
        return _refusal(
            tool_name,
            _LINK_REFUSAL,
            [f"{tool_name} without link_to_file, which embeds the file"],
            error=f"{tool_name} cannot link to a file on macOS",
        )

    absolute = os.path.abspath(os.path.expanduser(file_path))
    if not os.path.isfile(absolute):
        # The same error Windows raises, and raised here before PowerPoint is
        # asked for anything.
        raise FileNotFoundError(f"{what.capitalize()} file not found: {absolute}")

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)

    before = count(slide.shapes)
    names_before = _names_on(slide, before)

    # Staged, never named where it sits. PowerPoint is sandboxed, and a folder
    # it has no grant for makes macOS ask the user to allow access; in #191
    # that sheet closed every open document. The copy goes away again below,
    # because the embed happens here rather than at save time.
    staged = stage_into_container(absolute)
    try:
        app.make(
            new=k.media2_object,
            at=slide.end,
            with_properties={
                k.file_name: staged,
                k.left_position: left,
                k.top: top,
            },
        )
    finally:
        if os.path.abspath(staged) != os.path.abspath(absolute):
            try:
                os.remove(staged)
            except OSError:
                logger.debug("Could not remove the staged media at %s", staged)

    # Nothing is trusted because it did not raise, and the reference `make`
    # hands back is addressed by subclass, so the shape is found by counting
    # and then indexing rather than by keeping what PowerPoint returned.
    after = count(slide.shapes)
    if after <= before:
        return _refusal(
            tool_name,
            f"PowerPoint reported success and slide {slide_index} still holds "
            f"{after} shapes, which is the silent no-op recorded in "
            "MACOS_PORT section 5. Nothing was added.",
            ["ppt_add_picture"],
        )
    shape = slide.shapes[after]
    # A file PowerPoint cannot read answers -1708 rather than leaving litter
    # behind, but the type is still what proves a media shape arrived, because
    # a `make` it declines quietly comes back as a plain autoshape.
    _verify_created(
        shape, None, None, expected_type=k.shape_type_media, what=what,
    )

    warnings = []
    if width is not None and height is not None:
        shape.lock_aspect_ratio.set(False)
        shape.width.set(width)
        shape.height.set(height)
    elif width is not None:
        shape.lock_aspect_ratio.set(True)
        shape.width.set(width)
    elif height is not None:
        shape.lock_aspect_ratio.set(True)
        shape.height.set(height)

    landed_width = round(shape.width(), 2)
    landed_height = round(shape.height(), 2)
    if width is None and height is None:
        warnings.append(
            f"PowerPoint placed the {what} at {landed_width} by "
            f"{landed_height} points. It sizes media at "
            f"{_MEDIA_DEFAULT_HEIGHT} points tall with the file's own aspect "
            "ratio rather than at its native size, which Windows would have "
            "used. Pass width or height, or call ppt_update_shape."
        )
    else:
        # The size was asked for, so it is read back rather than echoed.
        if width is not None and abs(landed_width - width) > 0.5:
            warnings.append(
                f"A width of {width} was asked for and the {what} is "
                f"{landed_width} points wide."
            )
        if height is not None and abs(landed_height - height) > 0.5:
            warnings.append(
                f"A height of {height} was asked for and the {what} is "
                f"{landed_height} points tall."
            )

    shape_name = shape.name()
    duplicate = _media_shape_name_warning(names_before, shape_name)
    if duplicate:
        warnings.append(duplicate)

    result = {
        # The caller's own path, not the container copy. Where the file was
        # staged on the way in is this module's business and not theirs.
        "success": True,
        "shape_name": shape_name,
        "file_path": absolute,
    }
    if warnings:
        result["warnings"] = warnings
    return result


# ---------------------------------------------------------------------------
# Apple Event implementation functions
# ---------------------------------------------------------------------------
def _add_video_impl(slide_index, file_path, left, top, width, height, link_to_file):
    """Put a movie on a slide, embedded."""
    return _add_media(
        "ppt_add_video", "video",
        slide_index, file_path, left, top, width, height, link_to_file,
    )


def _add_audio_impl(slide_index, file_path, left, top, width, height, link_to_file):
    """Put a sound on a slide, embedded.

    Worth naming the near miss: ``import sound file`` looks like the answer and
    is not. It loads a ``sound effect``, which hangs off a transition or an
    action setting rather than off a shape, so it cannot put a sound on a slide
    of its own. ``media2 object`` can, and an ``.aiff`` was verified through it.
    """
    return _add_media(
        "ppt_add_audio", "audio",
        slide_index, file_path, left, top, width, height, link_to_file,
    )


def _animation_count(slide):
    """How much animation the slide is holding, or None if it will not say.

    Asked of the sequence, never of its effects. ``effects.count()`` and
    ``effects.get()`` both kill PowerPoint with -609; asking the sequence how
    many `effect` elements it holds is one safe round trip. See ``count_of``,
    which carries the measurement.

    The main sequence only, and the timeline's own `sequence` elements are
    deliberately not added to it. Counting those looked like a free way to
    catch the interactive sequence behind a shape triggered animation, and it
    is not: putting a media shape on a slide creates one sequence of its own
    for the playback. It answered 0 on a fresh slide, 0 after a text box, 0
    with one effect in the main sequence, and 1 as soon as a movie was
    inserted, so a gate that counted it would refuse the write on every slide
    that has any media on it, which is every slide this tool is ever called
    about. Worth knowing before anyone tries it again.

    So an animation triggered by clicking a shape is the one case this cannot
    see. `ppt_add_animation` cannot create one and neither can anything else
    here, but a deck built by hand or built on Windows can arrive carrying one,
    and the successful write says so rather than leaving it implied.

    None means the question could not be answered, and the caller here treats
    that the same as "there are animations", because the write it guards is the
    one that would destroy them.
    """
    try:
        return count_of(slide.timeline.main_sequence, k.effect)
    except CommandError as exc:
        logger.debug("Could not count the slide's animation effects: %s", exc)
        return None


def _set_media_settings_impl(
    slide_index, shape_name_or_index,
    volume, muted, start_point, end_point,
    fade_in, fade_out, loop, hide_while_not_playing,
):
    """Set what of a media shape's playback PowerPoint for Mac will hold.

    Two of the eight, and the second of those only on a slide with no
    animations. Both refusals name the argument rather than the tool, so
    dropping it and calling again works.
    """
    absent = [
        name for name, value in zip(
            _NO_WORDS_FOR,
            (volume, muted, start_point, end_point, fade_in, fade_out),
        )
        if value is not None
    ]
    if absent:
        # Refused whole rather than applied in part. A call that sets loop and
        # silently drops the volume is the failure MACOS_PORT section 5 is
        # about, and refusing all of it is what makes "drop the argument and
        # call again" true.
        return _refusal(
            "ppt_set_media_settings",
            "PowerPoint for Mac has no volume, mute, trim or fade for a media "
            "shape anywhere in its dictionary; the words do not exist, so "
            f"{', '.join(absent)} cannot be honoured and nothing was changed. "
            "Loop and hide while not playing do exist and this tool sets both. "
            "Trim and fade have to be done to the file before it is inserted.",
            [
                "ppt_set_media_settings with only loop and "
                "hide_while_not_playing",
            ],
            error=(
                "ppt_set_media_settings cannot set "
                f"{', '.join(absent)} on macOS"
            ),
        )

    ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    shape_type = shape.shape_type()
    if shape_type != k.shape_type_media:
        # Checked before the first write. `play settings` exists on every
        # shape, so writing to a picture would report success and change
        # nothing anybody could ever see.
        raise ValueError(
            f"'{shape_name_or_index}' on slide {slide_index} is not a media "
            "shape, so it has no playback settings. Use ppt_list_shapes to "
            "find the video or sound."
        )

    warnings = []
    if hide_while_not_playing is not None:
        effects = _animation_count(slide)
        if effects is None or effects > 0:
            # Measured, not inferred. One write of `hide while not playing` on
            # a slide holding a fly, a bounce and an exit fade dropped the exit
            # outright and turned the bounce into a plain appear, and writing
            # the value it already held did it too. That is the flattening of
            # MACOS_PORT section 5.2, and there is no way to write this and
            # keep the slide.
            if effects is None:
                found = (
                    f"Slide {slide_index}'s animation effects could not be "
                    "counted"
                )
            else:
                found = (
                    f"Slide {slide_index} holds {effects} animation "
                    f"effect{'' if effects == 1 else 's'}"
                )
            return _refusal(
                "ppt_set_media_settings",
                "`hide while not playing` is reached through the old per shape "
                "animation API, and writing it rewrites every animation on the "
                f"slide. {found}, so the "
                "write was refused rather than paid for with them: measured on "
                "a slide with three effects, one write dropped the exit effect "
                "outright and turned a bounce into a plain appear (MACOS_PORT "
                "section 5.2). Loop can still be set here, and hiding the frame "
                "is safe on a slide with no animations.",
                [
                    "ppt_set_media_settings with loop only",
                    "ppt_clear_animations, after which hide_while_not_playing "
                    "can be set",
                ],
                error=(
                    "ppt_set_media_settings cannot set hide_while_not_playing "
                    f"on slide {slide_index} because it has animations"
                ),
            )

    if loop is not None:
        play = shape.animation_settings.animation_play_settings
        play.loop_until_stopped.set(bool(loop))
        # Read back, because nothing is trusted because it did not raise.
        landed = play.loop_until_stopped()
        if bool(landed) != bool(loop):
            return _refusal(
                "ppt_set_media_settings",
                f"PowerPoint reported success and `loop until stopped` reads "
                f"back as {bool(landed)} rather than {bool(loop)}, which is "
                "the silent no-op recorded in MACOS_PORT section 5.",
                error="ppt_set_media_settings could not set loop",
            )

    if hide_while_not_playing is not None:
        play = shape.animation_settings.animation_play_settings
        play.hide_while_not_playing.set(bool(hide_while_not_playing))
        landed = play.hide_while_not_playing()
        warnings.append(
            "hide_while_not_playing was written because this slide's main "
            "animation sequence is empty. An animation triggered by clicking a "
            "shape lives somewhere else and nothing here can count it, and "
            "writing this setting rewrites every animation on a slide, so "
            "check the slide if it was built by hand with one of those."
        )
        if bool(landed) != bool(hide_while_not_playing):
            return _refusal(
                "ppt_set_media_settings",
                "PowerPoint reported success and `hide while not playing` "
                f"reads back as {bool(landed)} rather than "
                f"{bool(hide_while_not_playing)}, which is the silent no-op "
                "recorded in MACOS_PORT section 5.",
                error=(
                    "ppt_set_media_settings could not set hide_while_not_playing"
                ),
            )

    if loop is None and hide_while_not_playing is None:
        warnings.append(
            "Neither loop nor hide_while_not_playing was given, so nothing was "
            "changed. Those two are the only playback settings PowerPoint for "
            "Mac holds."
        )

    result = {
        "success": True,
        "shape_name": shape.name(),
    }
    if warnings:
        result["warnings"] = warnings
    return result
