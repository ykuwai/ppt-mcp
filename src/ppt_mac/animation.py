"""Animation and transition operations, on Apple Events.

Mirrors ``ppt_com/animation.py``. Same function names, same signatures, same
returned shapes; what differs is the walk through PowerPoint's object model, and
here that walk is narrower than the Windows one in three ways worth knowing
before reading on.

**The effects collection is never materialised.** ``main_sequence.effects.get()``
kills PowerPoint with -609 and takes every open deck with it, on a slide with no
animations at all. It is the sharpest case of the rule in MACOS_PORT section 5.1,
that a reference PowerPoint hands back is not trusted. So effects are counted by
asking the sequence, with ``count_of``, and reached one at a time by index.
Nothing in this module calls ``get`` or ``count`` on a collection of effects,
and nothing should.

**Writing anything through ``animation settings`` rewrites the whole slide.**
This is the trap of the module and nothing about it announces itself. The old
per shape API can only hold one entrance per shape, and writing to it seems to
force the slide back into that model. Three shapes with one entrance each,
and setting ``dim color`` on the first turned all three into a plain appear.
Setting ``animate`` to false on the second removed the second, as asked, and
turned the first into an appear. Put an exit animation on the third and it disappears
outright. ``text unit effect`` and ``animate text in reverse`` do the same.
``animate background`` is the one write that leaves the slide alone.

So this module writes only ``animate background`` and ``animate``, and it writes
``animate`` only from ``ppt_clear_animations``, where the slide is being emptied
anyway and there is nothing left to damage. ``dim_color``, ``after_effect`` and
``animate_in_reverse`` come back as warnings without being attempted, and
``ppt_remove_animation`` refuses outright, because the two routes to it are
-50 and a rewritten slide.

**Timing carries no trigger and no delay.** The ``timing`` class holds
acceleration, autoreverse, deceleration, duration, repeat count, repeat
duration, restart, rewind, smooth end, smooth start and speed, and nothing else.
A trigger can only be chosen at the moment an effect is added, through
``add effect``'s ``trigger`` parameter, and a delay has no home at all. Neither
is silently dropped; both come back in ``warnings``.

Interactive sequences do not exist either. ``timeline.add_sequence()`` returns a
reference and the slide's sequence count stays at zero, so every
``sequence_index`` argument is refused rather than quietly ignored.
"""

import logging

from appscript import k
from appscript.reference import CommandError

from backend.mac_ae import (
    AE_NOT_HANDLED,
    count_of,
    error_number,
    is_missing,
    ppt,
    shape_by_name_or_index as _get_shape,
    shapes_of,
    slide_at as _slide,
    windows_constant as _windows_constant,
)
from backend.mac_enums import (
    MsoAnimAfterEffect,
    MsoAnimDirection,
    MsoAnimEffect,
    MsoAnimTextUnitEffect,
    MsoAnimTriggerType,
    MsoAnimateByLevel,
    PpEntryEffect,
    to_keyword,
)
from backend.unsupported import refusal as _refusal
from ppt_com.constants import (
    AFTER_EFFECT_NAMES,
    ANIMATION_EFFECT_NAMES,
    ANIM_DIRECTION_NAMES,
    BUILD_LEVEL_NAMES,
    TEXT_UNIT_EFFECT_NAMES,
    msoAnimAfterEffectNone,
    msoAnimateLevelNone,
)
from utils.navigation import goto_slide

logger = logging.getLogger(__name__)

# scripts/gen_mac_enums.py pairs by name, so a constant Windows and macOS spell
# differently is left out of the generated table even when both platforms have
# it. Windows says `msoAnimAfterEffectNone` where macOS says `no after effect`,
# and `msoAnimateLevelNone` where macOS says `text by no levels`, so neither
# paired. Both are only ever read back here, and reporting None for a value
# PowerPoint answered plainly would be worse than filling the gap locally. These
# should disappear the next time the table is regenerated.
_AFTER_EFFECTS = dict(MsoAnimAfterEffect)
_AFTER_EFFECTS.setdefault(msoAnimAfterEffectNone, k.no_after_effect)

_BUILD_LEVELS = dict(MsoAnimateByLevel)
_BUILD_LEVELS.setdefault(msoAnimateLevelNone, k.text_by_no_levels)

# What `to_keyword` says when a Windows constant has no macOS word. The message
# reads "PowerPoint for Mac has no <what> matching the Windows constant <n>", so
# each of these is written to finish that sentence and then name the way out.
_WHAT_EFFECT = (
    "animation effect (every effect ppt_add_animation lists is available here, "
    "so a number this rejects is not one of them)"
)
_WHAT_TRANSITION = (
    "slide transition effect (push, wipe, split and reveal each exist here as "
    "four directional variants and in no plain form, so there is nothing to "
    "choose without being told the direction; the other seven all work)"
)
_WHAT_DIRECTION = "animation direction"
_WHAT_TRIGGER = "animation trigger"
_WHAT_LEVEL = "text build level"
_WHAT_TEXT_UNIT = "text unit effect"


def _sequence_refusal(tool_name: str) -> dict:
    """The same answer both tools give to a sequence_index they cannot honour."""
    return _refusal(
        tool_name,
        "Interactive sequences do not exist in PowerPoint for Mac. The timeline "
        "accepts `add sequence` and hands back a reference, but the slide's "
        "sequence count stays at zero afterwards, so a sequence_index can never "
        "name anything real. Leaving it out works on the main sequence, which "
        "is the only one there is.",
        [f"{tool_name} without sequence_index", "ppt_list_animations"],
        error=f"{tool_name} cannot take sequence_index on macOS",
    )


def _main_sequence(slide):
    """The one sequence a slide has here. See the module docstring."""
    return slide.timeline.main_sequence


def _effect_count(seq) -> int:
    """How many effects the sequence holds.

    Asked of the sequence, never of its effects. ``seq.effects.count()`` and
    ``seq.effects.get()`` both kill PowerPoint with -609, while asking the
    sequence how many `effect` elements it holds answers in one round trip. See
    ``count_of``, which carries the measurement.
    """
    return count_of(seq, k.effect)


def _effect_shape_name(eff):
    """The name of the shape an effect belongs to, or None if it cannot be read."""
    try:
        name = eff.shape.name()
    except Exception:
        return None
    return None if is_missing(name) else name


def _effect_shape_names(seq, total: int) -> list:
    """One shape name per effect, in sequence order."""
    return [_effect_shape_name(seq.effects[i]) for i in range(1, total + 1)]


def _read(getter, default=None):
    """Read one property, reporting None rather than inventing a value.

    Every effect property here is optional in practice. PowerPoint answers
    `missing value` for the ones a particular effect does not carry and -1728
    for a few it carries but will not hand over, and neither is worth failing a
    whole listing for.
    """
    try:
        value = getter()
    except Exception:
        return default
    return default if is_missing(value) else value


def _apply_timing(eff, duration, repeat_count, auto_reverse, rewind,
                  smooth_start, smooth_end) -> None:
    """Write the parts of an effect that live on its timing object."""
    timing = eff.timing
    if duration is not None:
        timing.duration.set(duration)
    if repeat_count is not None:
        timing.repeat_count.set(repeat_count)
    if auto_reverse is not None:
        timing.autoreverse.set(bool(auto_reverse))
    if rewind is not None:
        timing.rewind.set(bool(rewind))
    if smooth_start is not None:
        timing.smooth_start.set(bool(smooth_start))
    if smooth_end is not None:
        timing.smooth_end.set(bool(smooth_end))


def _apply_effect_options(
    seq, eff, shape,
    trigger, delay, exit_flag, direction,
    repeat_count, auto_reverse, rewind, smooth_start, smooth_end, duration,
    after_effect, dim_color,
    text_unit_effect, animate_in_reverse, animate_background,
) -> list:
    """Apply everything both add and update can change, and say what did not land.

    Shared so the two tools cannot drift apart. The return value is the list of
    warnings, which the caller puts in its result rather than failing the whole
    call, because an effect that got nine of its ten properties is far more use
    to a reader than an error.

    ``trigger`` is passed here only so it can be reported. Adding is the one
    moment a trigger can be chosen, and ``_add_animation_impl`` has already
    spent it by the time this runs.
    """
    warnings: list[str] = []

    if exit_flag is not None:
        eff.exit_animation.set(bool(exit_flag))

    _apply_timing(
        eff, duration, repeat_count, auto_reverse, rewind, smooth_start, smooth_end,
    )

    if delay is not None:
        warnings.append(
            "delay was not applied. PowerPoint for Mac's timing object carries "
            "no trigger delay, so there is nowhere to put it. Set trigger to "
            "'after_previous' for a sequential feel, or add the pause in "
            "PowerPoint by hand."
        )

    if trigger is not None:
        warnings.append(
            "trigger was not applied. A trigger can only be chosen at the "
            "moment an effect is added here, because the timing object has no "
            "trigger type property. Remove the animation and add it again with "
            "the trigger you want."
        )

    if direction is not None:
        # Reading a direction from an effect that has none answers `missing
        # value`, but writing one to it answers -1708, so this is guarded. A
        # fly or a wipe takes a direction; an appear has nowhere to put one.
        word = to_keyword(MsoAnimDirection, direction, _WHAT_DIRECTION)
        try:
            eff.effect_parameters.direction.set(word)
        except CommandError as exc:
            if error_number(exc) != AE_NOT_HANDLED:
                raise
            warnings.append(
                f"direction '{direction}' was not applied. This effect has no "
                "direction to set. Directions belong to effects that travel, "
                "fly and wipe among them, not to appear or fade."
            )

    if text_unit_effect is not None:
        # The one `convert to` command PowerPoint for Mac has. Windows has four
        # and the other three are rerouted through the shape below.
        seq.convert_to_text_unit_effect(
            Effect=eff,
            unit=to_keyword(MsoAnimTextUnitEffect, text_unit_effect, _WHAT_TEXT_UNIT),
        )

    # Three of the four writes on `animation settings` are destructive, and
    # nothing about them says so. Writing `dim color`, `text unit effect` or
    # `animate text in reverse` on one shape resets **every effect on the
    # slide** to a plain appear, silently and instantly. Three fades on three
    # shapes became three appears after one write, and the shape that was
    # written to was not even the worst of it, the other two lost their
    # animation to a call that never mentioned them. So none of the three is
    # ever written here. `animate background` was checked the same way and is
    # the one that leaves the slide alone.
    settings = shape.animation_settings
    if animate_background is not None:
        settings.animate_background.set(bool(animate_background))
        warnings.append(
            "animate_background belongs to the shape here rather than to one "
            "effect, so it now applies to every animation that shape has."
        )
    if animate_in_reverse is not None:
        warnings.append(
            "animate_in_reverse was not applied. The only way in is the "
            "shape's animation settings, and writing to that resets every "
            "effect on the slide to a plain appear, so it is not attempted. "
            "Set it in PowerPoint by hand."
        )
    if dim_color is not None:
        warnings.append(
            f"dim_color {dim_color} was not applied, for the same reason as "
            "after_effect below and with a sharper edge. Writing a dim colour "
            "wipes every animation on the slide, so it is not attempted."
        )

    if after_effect is not None:
        message = (
            "after_effect was not applied. Writing it fails with Apple Event "
            "error -1708, which is PowerPoint saying it does not implement "
            "that write, and the effect's own after effect information is "
            "read only. So there is no way in on this platform. "
        )
        message += "Set it in PowerPoint by hand instead."
        warnings.append(message)

    return warnings


def _after_effect_report(eff) -> tuple:
    """Read an effect's after-animation behaviour back as (constant, name)."""
    word = _read(eff.effect_information.after_effect_information)
    if word is None:
        return None, None
    value = _windows_constant(_AFTER_EFFECTS, word)
    if value is None:
        return None, None
    return value, AFTER_EFFECT_NAMES.get(value, f"Unknown({value})")


# ---------------------------------------------------------------------------
# Apple Event implementation functions
# ---------------------------------------------------------------------------
def _set_slide_transition_impl(
    slide_index, effect, duration, advance_on_click, advance_on_time, advance_time,
):
    from ppt_com.animation import TRANSITION_EFFECT_MAP

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)

    effect_int = (
        TRANSITION_EFFECT_MAP.get(effect, effect) if isinstance(effect, str) else effect
    )

    transition = slide.slide_show_transition
    # PpEntryEffect covers seven of the transitions Windows names, and the rest
    # raise here rather than landing on something that looks close.
    wanted = to_keyword(PpEntryEffect, effect_int, _WHAT_TRANSITION)
    transition.entry_effect.set(wanted)

    # Nothing is trusted because it did not raise, and this is one read.
    if _read(transition.entry_effect) != wanted:
        return _refusal(
            "ppt_set_slide_transition",
            "PowerPoint reported success but the slide's transition did not "
            "change, which is the silent no-op recorded in MACOS_PORT "
            "section 5.",
        )

    if duration is not None:
        # Windows calls this `Duration`; macOS already uses that word for an
        # effect's own timing, so the slide's is `transition duration`.
        transition.transition_duration.set(duration)
    if advance_on_click is not None:
        transition.advance_on_click.set(bool(advance_on_click))
    if advance_on_time is not None:
        transition.advance_on_time.set(bool(advance_on_time))
    if advance_time is not None:
        transition.advance_time.set(advance_time)

    return {
        "success": True,
        "slide_index": slide_index,
        "effect": effect_int,
    }


def _add_animation_impl(
    slide_index, shape_name_or_index, effect, trigger, duration, delay, exit_flag,
    direction, repeat_count, auto_reverse, rewind, smooth_start, smooth_end,
    trigger_shape, after_effect, dim_color,
    build_level, text_unit_effect, animate_in_reverse, animate_background,
):
    from ppt_com.animation import ANIMATION_EFFECT_MAP, TRIGGER_MAP
    from ppt_com.constants import ANIM_DIRECTION_MAP, BUILD_LEVEL_MAP, TEXT_UNIT_EFFECT_MAP

    # Refused before anything moves, so the user's view is not sent to a slide
    # this call is not going to touch.
    if trigger_shape is not None or trigger == "on_shape_click":
        return _refusal(
            "ppt_add_animation",
            "An animation cannot be triggered by clicking another shape here. "
            "Asking for that trigger fails with Apple Event error -1708, "
            "which is PowerPoint saying it does not implement it, and the "
            "interactive "
            "sequence it would need cannot be created either; the timeline "
            "hands back a sequence and then reports it has none. Use "
            "'on_click', 'with_previous' or 'after_previous' instead.",
            [
                "ppt_add_animation with trigger='on_click'",
                "ppt_add_animation with trigger='after_previous'",
            ],
            error="ppt_add_animation cannot use trigger_shape on macOS",
        )

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, shape_name_or_index)
    shape_name = shape.name()

    effect_int = (
        ANIMATION_EFFECT_MAP.get(effect, effect) if isinstance(effect, str) else effect
    )
    trigger_int = TRIGGER_MAP.get(trigger, 1)
    direction_int = (
        ANIM_DIRECTION_MAP.get(direction, direction)
        if isinstance(direction, str) else direction
    )

    seq = _main_sequence(slide)
    before = _effect_count(seq)

    add_args = {
        "for_": shape,
        "fx": to_keyword(MsoAnimEffect, effect_int, _WHAT_EFFECT),
        "trigger": to_keyword(MsoAnimTriggerType, trigger_int, _WHAT_TRIGGER),
    }
    if build_level is not None and build_level != "none":
        # `text by no levels` is what a plain add already does, and passing it
        # explicitly buys nothing, so only a real level is sent.
        add_args["level"] = to_keyword(
            _BUILD_LEVELS, BUILD_LEVEL_MAP[build_level], _WHAT_LEVEL,
        )
    seq.add_effect(**add_args)

    # Nothing is trusted because it did not raise, and a level other than none
    # can turn one call into several effects, so the sequence is counted again
    # rather than assumed.
    after = _effect_count(seq)
    if after <= before:
        return _refusal(
            "ppt_add_animation",
            "PowerPoint reported success but the slide's animation sequence "
            f"still holds {after} effects, which is the silent no-op recorded "
            "in MACOS_PORT section 5.",
        )

    animation_index = before + 1
    eff = seq.effects[animation_index]

    warnings: list[str] = []
    landed_on = _effect_shape_name(eff)
    if landed_on is not None and landed_on != shape_name:
        # The new effect is expected at the end of the sequence. Saying so when
        # it is not beats quietly editing whichever effect is sitting there.
        warnings.append(
            f"The new effect was expected at position {animation_index} but "
            f"that position belongs to '{landed_on}' rather than "
            f"'{shape_name}'. Everything else in this call was applied to "
            f"position {animation_index}, so check the result with "
            "ppt_list_animations."
        )
    if after - before > 1:
        warnings.append(
            f"build_level='{build_level}' created {after - before} effects, one "
            "per paragraph. The other options were applied to the first of "
            "them."
        )

    warnings += _apply_effect_options(
        seq, eff, shape,
        None, delay, exit_flag, direction_int,
        repeat_count, auto_reverse, rewind, smooth_start, smooth_end, duration,
        after_effect, dim_color,
        TEXT_UNIT_EFFECT_MAP[text_unit_effect] if text_unit_effect else None,
        animate_in_reverse, animate_background,
    )

    result = {
        "success": True,
        "shape_name": shape_name,
        "effect": effect_int,
        "exit": bool(exit_flag),
        "animation_index": animation_index,
    }

    after_effect_val, after_effect_name = _after_effect_report(eff)
    if after_effect_val is not None:
        result["after_effect"] = after_effect_name
    if warnings:
        result["warnings"] = warnings
    return result


def _read_text_anim_info(eff) -> dict:
    """Read the text animation part of an effect, the way Windows reports it."""
    info: dict = {}

    word = _read(eff.effect_information.build_by_level)
    value = _windows_constant(_BUILD_LEVELS, word) if word is not None else None
    info["build_level"] = value
    info["build_level_name"] = (
        BUILD_LEVEL_NAMES.get(value, f"Unknown({value})") if value is not None else None
    )

    word = _read(eff.effect_information.text_unit_effect_information)
    value = _windows_constant(MsoAnimTextUnitEffect, word) if word is not None else None
    info["text_unit_effect"] = value
    info["text_unit_effect_name"] = (
        TEXT_UNIT_EFFECT_NAMES.get(value, f"Unknown({value})")
        if value is not None else None
    )

    reverse = _read(eff.effect_information.animate_text_in_reverse_information)
    info["animate_in_reverse"] = None if reverse is None else bool(reverse)
    background = _read(eff.effect_information.animate_background_information)
    info["animate_background"] = None if background is None else bool(background)
    return info


def _list_animations_impl(slide_index):
    from ppt_com.animation import _get_animation_category

    ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)

    seq = _main_sequence(slide)
    total = _effect_count(seq)

    animations = []
    for i in range(1, total + 1):
        eff = seq.effects[i]

        effect_word = _read(eff.animation_effect_type)
        effect_type = (
            _windows_constant(MsoAnimEffect, effect_word)
            if effect_word is not None else None
        )
        exit_value = _read(eff.exit_animation)
        exit_flag = bool(exit_value) if exit_value is not None else False

        direction_word = _read(eff.effect_parameters.direction)
        direction_val = (
            _windows_constant(MsoAnimDirection, direction_word)
            if direction_word is not None else None
        )

        after_effect_val, after_effect_name = _after_effect_report(eff)

        anim_dict = {
            # No `index` property on an effect here, so the position in the walk
            # is the index, which is the same number every other tool takes.
            "index": i,
            "shape_name": _effect_shape_name(eff),
            "effect_type": effect_type,
            "effect_name": (
                ANIMATION_EFFECT_NAMES.get(effect_type, f"Unknown({effect_type})")
                if effect_type is not None else None
            ),
            # The timing object has no trigger type, and nothing else on the
            # effect carries one, so this is genuinely unreadable rather than
            # merely awkward. None, not a guess at 'on_click'.
            "trigger_type": None,
            "trigger_name": None,
            "duration": _read(eff.timing.duration),
            "exit": exit_flag,
            "category": _get_animation_category(effect_type or 0, exit_flag),
            "direction": direction_val,
            "direction_name": (
                ANIM_DIRECTION_NAMES.get(direction_val, f"Unknown({direction_val})")
                if direction_val is not None else None
            ),
            "after_effect": after_effect_val,
            "after_effect_name": after_effect_name,
        }
        anim_dict.update(_read_text_anim_info(eff))
        animations.append(anim_dict)

    return {
        "success": True,
        "slide_index": slide_index,
        "main_sequence_count": total,
        "animations": animations,
        # Always empty. See the module docstring; a slide here has one sequence.
        "interactive_sequences": [],
        "interactive_count": 0,
    }


def _remove_animation_impl(slide_index, animation_index, sequence_index):
    """Refuse, always. Removing one animation cannot be done without collateral.

    Not a refusal on principle, a refusal on evidence. There are two routes and
    both are worse than doing nothing.

    Deleting the effect answers -50 and changes nothing, through
    ``effects[n].delete()`` and through ``delete effects[n]`` alike.

    Clearing the shape with ``animation settings.animate`` does remove the
    effect, and it damages the rest of the slide on the way past. Three shapes
    with one entrance each, and clearing the second removed the second, as
    asked, while the first turned into a plain appear. Add an exit animation to the third and
    that one disappears outright. Whatever is written through `animation
    settings` seems to rewrite the whole slide through the old one effect per
    shape model, and anything the old model cannot hold is lost.

    So this reports the position and hands the caller the two honest ways
    forward rather than quietly rebuilding their slide.
    """
    if sequence_index is not None:
        return _sequence_refusal("ppt_remove_animation")

    # Validated before refusing, so a caller who passed a bad index hears about
    # that first. Nothing is written, and the view is not moved.
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    seq = _main_sequence(slide)
    total = _effect_count(seq)
    if animation_index < 1 or animation_index > total:
        raise ValueError(
            f"Animation index {animation_index} out of range (1-{total})"
        )

    names = _effect_shape_names(seq, total)
    shape_name = names[animation_index - 1]
    owner = f"'{shape_name}'" if shape_name else "its shape"

    return _refusal(
        "ppt_remove_animation",
        f"Animation {animation_index} belongs to {owner}, and PowerPoint for "
        "Mac offers no way to remove it on its own. Deleting an effect answers "
        "-50 and changes nothing. Clearing the shape does remove it, and it "
        "also flattens earlier effects on the slide to a plain appear and "
        "drops exit animations, so it is not attempted. Clear the slide and "
        "add back the animations you want, which is the only route that leaves "
        "the result predictable.",
        ["ppt_list_animations", "ppt_clear_animations", "ppt_add_animation"],
        error="ppt_remove_animation cannot remove a single animation on macOS",
    )


def _clear_animations_impl(slide_index, clear_transitions):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)

    seq = _main_sequence(slide)
    before = _effect_count(seq)

    # Per shape, because that is the only deletion PowerPoint for Mac offers.
    # A shape with no animation is not an error, so a refusal on one shape does
    # not stop the walk.
    for shape in shapes_of(slide):
        try:
            shape.animation_settings.animate.set(False)
        except CommandError:
            logger.warning(
                "Could not clear the animation on one shape of slide %s",
                slide_index, exc_info=True,
            )

    after = _effect_count(seq)

    if clear_transitions:
        slide.slide_show_transition.entry_effect.set(k.entry_effect_none)

    return {
        "success": True,
        "slide_index": slide_index,
        # Counted before and after rather than reported as intent, so a shape
        # that would not let go of its animation shows up in the number.
        "cleared_count": before - after,
        # Always zero. A slide here has one sequence; see the module docstring.
        "interactive_cleared": 0,
        "remaining_count": after,
    }


def _update_animation_impl(
    slide_index, animation_index, sequence_index,
    effect, trigger, duration, delay, move_to, exit_flag,
    direction, repeat_count, auto_reverse, rewind, smooth_start, smooth_end,
    after_effect, dim_color,
    build_level, text_unit_effect, animate_in_reverse, animate_background,
):
    from ppt_com.animation import ANIMATION_EFFECT_MAP, _get_animation_category
    from ppt_com.constants import ANIM_DIRECTION_MAP, TEXT_UNIT_EFFECT_MAP

    if sequence_index is not None:
        return _sequence_refusal("ppt_update_animation")

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)

    seq = _main_sequence(slide)
    total = _effect_count(seq)
    if animation_index < 1 or animation_index > total:
        raise ValueError(
            f"Animation index {animation_index} out of range (1-{total})"
        )

    eff = seq.effects[animation_index]
    shape_name = _effect_shape_name(eff)
    if shape_name is None:
        return _refusal(
            "ppt_update_animation",
            f"Animation {animation_index} would not say which shape it belongs "
            "to, and several of the properties this tool writes live on the "
            "shape rather than on the effect.",
            ["ppt_list_animations"],
        )
    shape = _get_shape(slide, shape_name)

    warnings: list[str] = []

    if effect is not None:
        effect_int = (
            ANIMATION_EFFECT_MAP.get(effect, effect)
            if isinstance(effect, str) else effect
        )
        eff.animation_effect_type.set(
            to_keyword(MsoAnimEffect, effect_int, _WHAT_EFFECT)
        )

    if build_level is not None:
        # Windows reaches this through ConvertToBuildLevel. PowerPoint for Mac
        # declares one convert command, for text units, and the build level it
        # does expose on an effect is read only, so an existing effect's level
        # cannot be changed. Adding is the moment it is decided.
        warnings.append(
            "build_level was not applied. The build level of an effect is read "
            "only here and PowerPoint for Mac has no command to convert one, so "
            "it can only be chosen when the animation is added. Clear the "
            "animation and add it again with the build_level you want."
        )

    if move_to is not None:
        warnings.append(
            "move_to was not applied. An effect here carries no index and there "
            "is no command to move one, and rebuilding the sequence in the "
            "right order would mean deleting effects one at a time, which "
            "PowerPoint for Mac cannot do. Clear the slide's animations and add "
            "them back in the order you want."
        )

    direction_int = (
        ANIM_DIRECTION_MAP.get(direction, direction)
        if isinstance(direction, str) else direction
    )

    warnings += _apply_effect_options(
        seq, eff, shape,
        trigger, delay, exit_flag, direction_int,
        repeat_count, auto_reverse, rewind, smooth_start, smooth_end, duration,
        after_effect, dim_color,
        TEXT_UNIT_EFFECT_MAP[text_unit_effect] if text_unit_effect else None,
        animate_in_reverse, animate_background,
    )

    # Read back rather than echoed, so a write that did not land shows up. The
    # effect is fetched again by index because converting to a text unit effect
    # replaces it.
    eff = seq.effects[animation_index]
    effect_word = _read(eff.animation_effect_type)
    effect_type = (
        _windows_constant(MsoAnimEffect, effect_word)
        if effect_word is not None else None
    )
    exit_value = _read(eff.exit_animation)
    final_exit = bool(exit_value) if exit_value is not None else False

    result = {
        "success": True,
        "animation_index": animation_index,
        "shape_name": _effect_shape_name(eff) or shape_name,
        "effect_type": effect_type,
        "effect_name": (
            ANIMATION_EFFECT_NAMES.get(effect_type, f"Unknown({effect_type})")
            if effect_type is not None else None
        ),
        # Unreadable rather than unset; see _list_animations_impl.
        "trigger_type": None,
        "trigger_name": None,
        "duration": _read(eff.timing.duration),
        "delay": None,
        "exit": final_exit,
        "category": _get_animation_category(effect_type or 0, final_exit),
    }

    after_effect_val, after_effect_name = _after_effect_report(eff)
    if after_effect_val is not None:
        result["after_effect"] = after_effect_name
    if warnings:
        result["warnings"] = warnings
    return result
