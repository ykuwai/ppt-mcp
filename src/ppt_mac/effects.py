"""Visual effect tools, on Apple Events.

Mirrors ``ppt_com/effects.py``. Same function names, same signatures, same
returned shapes; what differs is the walk through PowerPoint's object model.

Three things about effects on this side are worth knowing before reading on.

**Glow has no transparency.** The whole ``glow format`` class is ``color``,
``color theme index`` and ``radius``, so a glow lands at the colour and size
asked for and at PowerPoint's own opacity. That is reported in ``warnings``
rather than dropped, because the glow itself is still the one the caller wanted.

**Reflection is a preset and nothing else.** ``reflection format`` carries one
property, ``reflection type``, which is the same nine presets the interface
offers. Windows sets blur, offset, size and transparency on top of the preset
and none of those four have anywhere to go here. A call that gives only those
four would change nothing at all, so that one is refused rather than answered
with a success that did not happen.

**Soft edge is a preset too.** ``soft edge format`` carries ``soft edge type``,
six presets and an off, where Windows takes a radius in points. The radius is
snapped to the nearest preset and the result says which preset it landed on. The
point value of each preset is the ladder the interface offers, 1, 2.5, 5, 10, 25
and 50 points, which is read off the interface rather than out of the dictionary
and so is the one assumption in this module.
"""

import logging

from appscript import k
from appscript.reference import CommandError

from backend.mac_ae import ppt, raw, slide_at as _slide
from backend.unsupported import refusal as _refusal
from ppt_mac.shapes import _get_shape
from utils.color import hex_to_rgb_list
from utils.navigation import goto_slide

logger = logging.getLogger(__name__)

# scripts/gen_mac_enums.py pairs each banner section of ppt_com/constants.py
# with an sdef enumeration, and constants.py names neither MsoReflectionType nor
# MsoSoftEdgeType, so neither reached the generated table. Both are in
# PowerPoint's dictionary in full, so they are written out here.
_REFLECTION_TYPES = {
    0: k.reflection_type_none,
    1: k.reflection_type1,
    2: k.reflection_type2,
    3: k.reflection_type3,
    4: k.reflection_type4,
    5: k.reflection_type5,
    6: k.reflection_type6,
    7: k.reflection_type7,
    8: k.reflection_type8,
    9: k.reflection_type9,
}

# Soft edge presets, in points, in the order PowerPoint's own menu lists them.
# The dictionary numbers the presets and says nothing about their size, so these
# are the point values the interface offers beside each one. A radius lands on
# the nearest of them.
_SOFT_EDGE_PRESETS = (
    (0.0, k.no_soft_edge, "none"),
    (1.0, k.soft_edge_type1, "1 point"),
    (2.5, k.soft_edge_type2, "2.5 points"),
    (5.0, k.soft_edge_type3, "5 points"),
    (10.0, k.soft_edge_type4, "10 points"),
    (25.0, k.soft_edge_type5, "25 points"),
    (50.0, k.soft_edge_type6, "50 points"),
)


def _nearest_soft_edge(radius: float):
    """Return the preset closest to a radius in points, with its own name."""
    return min(_SOFT_EDGE_PRESETS, key=lambda preset: abs(preset[0] - radius))


# ---------------------------------------------------------------------------
# Apple Event implementation functions
# ---------------------------------------------------------------------------
def _will_not_draw(shape):
    """The macOS half of ppt_com.effects.will_not_draw.

    `line format` has no `visible` here, the same gap ppt_set_line works
    around, so a border is judged the way that one hides it: weight 0 or full
    transparency reads as no border. Asking for the property that does not
    exist raised, the error was swallowed, and the warning this whole check
    exists for never fired on macOS.
    """
    try:
        if not shape.has_text_frame():
            return False
        if shape.fill_format.visible():
            return False
    except (AttributeError, CommandError):
        return False

    try:
        line = shape.line_format
        return not line.line_weight() or line.transparency() >= 1.0
    except (AttributeError, CommandError):
        # No line format at all, a picture or a placeholder, so there is
        # nothing for a shape effect to be drawn around either.
        return True


def _nothing_drawn_warning(shape, name):
    return (
        f"The {name} was set on shape '{shape.name()}', which has no fill and "
        f"no line, so there is nothing for it to be drawn around and the "
        f"slide will not change. For the halo around the text itself, use "
        f"target='text'."
    )


def _font_effect(shape, name):
    """The effect on the shape's text, or None when the dictionary has none.

    Windows reaches these through TextFrame2's font. Whether PowerPoint for
    Mac's font carries them is not something this code can assume, so it is
    asked rather than guessed, and a tool that cannot get one refuses the
    argument by name instead of setting a shape effect the caller did not ask
    for.
    """
    try:
        return getattr(shape.text_frame.text_range.font, name)
    except (AttributeError, CommandError):
        return None


def _refuse_text_target(tool_name, effect):
    return _refusal(
        tool_name,
        f"PowerPoint for Mac's font has no {effect} in its scripting "
        f"dictionary, so the {effect} around the glyphs cannot be set here. "
        "Nothing was changed. target='shape' works and draws on the shape's "
        "own fill and line.",
        [f"{tool_name} with target='shape'"],
        error=f"{tool_name} cannot take target='text' on macOS",
    )


def _set_glow_impl(slide_index, shape_name_or_index, radius,
                    color, transparency, target="shape") -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    if target == "text":
        glow = _font_effect(shape, "glow_format")
        if glow is None:
            return _refuse_text_target("ppt_set_glow", "glow")
    else:
        glow = shape.glow_format
    glow.radius.set(radius)

    if color is not None:
        glow.color.set(hex_to_rgb_list(color))

    result = {
        "status": "success",
        "shape_name": shape.name(),
        "target": target,
        # Read back rather than echoed, so a radius PowerPoint clamped or
        # ignored shows up in the answer instead of hiding behind the request.
        "glow_radius": glow.radius(),
    }
    warnings = []
    if transparency is not None:
        warnings.append(
            "PowerPoint for Mac's glow format has no transparency property, so "
            f"transparency={transparency} was not applied. The glow is the "
            "radius and colour asked for at PowerPoint's own opacity."
        )
    if target == "shape" and radius and _will_not_draw(shape):
        warnings.append(_nothing_drawn_warning(shape, "glow"))
    if warnings:
        result["warnings"] = warnings
    return result


def _set_reflection_impl(slide_index, shape_name_or_index, reflection_type,
                          blur, offset, size, transparency,
                          target="shape") -> dict:
    # Which of the four Windows extras were asked for, worked out before
    # anything is touched, because whether this call can do anything at all
    # depends on it.
    unsupported = [
        name for name, value in (
            ("blur", blur), ("offset", offset),
            ("size", size), ("transparency", transparency),
        )
        if value is not None
    ]

    if reflection_type is None and unsupported:
        # Nothing would land, and answering success to a call that changed
        # nothing is the failure MACOS_PORT section 5 is about. Refused here,
        # before the view moves, so a refused call leaves the deck alone.
        return _refusal(
            "ppt_set_reflection",
            "PowerPoint for Mac's `reflection format` carries one property, "
            "`reflection type`, which chooses one of the nine presets the "
            f"interface offers. {', '.join(unsupported)} have nowhere to go "
            "here, so this call would have changed nothing. Pass "
            "reflection_type to pick the preset closest to what you want.",
            ["ppt_set_reflection with reflection_type"],
            error=(
                "ppt_set_reflection cannot set "
                f"{', '.join(unsupported)} on macOS"
            ),
        )

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    if target == "text":
        reflection = _font_effect(shape, "reflection_format")
        if reflection is None:
            return _refuse_text_target("ppt_set_reflection", "reflection")
    else:
        reflection = shape.reflection_format

    if reflection_type is not None:
        word = _REFLECTION_TYPES.get(reflection_type)
        if word is None:
            raise ValueError(
                f"Unknown reflection_type {reflection_type}. "
                f"Valid values: {sorted(_REFLECTION_TYPES)}"
            )
        reflection.reflection_type.set(word)

    result = {
        "status": "success",
        "shape_name": shape.name(),
        "target": target,
    }
    warnings = []
    if unsupported:
        warnings.append(
            "PowerPoint for Mac's reflection format is the preset and nothing "
            f"else, so {', '.join(unsupported)} were not applied. The preset "
            "carries its own blur, offset, size and transparency."
        )
    asked_for_something = any(
        value is not None
        for value in (reflection_type, blur, offset, size, transparency)
    )
    if target == "shape" and asked_for_something and _will_not_draw(shape):
        warnings.append(_nothing_drawn_warning(shape, "reflection"))
    if warnings:
        result["warnings"] = warnings
    return result


def _set_soft_edge_impl(slide_index, shape_name_or_index, radius) -> dict:
    points, word, preset_name = _nearest_soft_edge(radius)

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    # By four character code, not by name. `font` also has a property called
    # `soft edge format`, and that one is the enumeration itself rather than an
    # object, so appscript resolves the name to the font's code and the write
    # lands on `soft edge type of soft edge type of shape`, which answers
    # -1708. `raw` addresses the shape's own DSeF property directly.
    raw(shape, b"DSeF").soft_edge_type.set(word)

    result = {
        "status": "success",
        "shape_name": shape.name(),
        "soft_edge_radius": points,
    }
    if points != radius:
        result["warnings"] = [
            "PowerPoint for Mac offers soft edges as presets rather than a "
            f"radius, so radius={radius} was applied as the nearest one, "
            f"{preset_name}."
        ]
    return result
