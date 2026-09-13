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
def _set_glow_impl(slide_index, shape_name_or_index, radius,
                    color, transparency) -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    glow = shape.glow_format
    glow.radius.set(radius)

    if color is not None:
        glow.color.set(hex_to_rgb_list(color))

    result = {
        "status": "success",
        "shape_name": shape.name(),
        # Read back rather than echoed, so a radius PowerPoint clamped or
        # ignored shows up in the answer instead of hiding behind the request.
        "glow_radius": glow.radius(),
    }
    if transparency is not None:
        result["warnings"] = [
            "PowerPoint for Mac's glow format has no transparency property, so "
            f"transparency={transparency} was not applied. The glow is the "
            "radius and colour asked for at PowerPoint's own opacity."
        ]
    return result


def _set_reflection_impl(slide_index, shape_name_or_index, reflection_type,
                          blur, offset, size, transparency) -> dict:
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

    if reflection_type is not None:
        word = _REFLECTION_TYPES.get(reflection_type)
        if word is None:
            raise ValueError(
                f"Unknown reflection_type {reflection_type}. "
                f"Valid values: {sorted(_REFLECTION_TYPES)}"
            )
        shape.reflection_format.reflection_type.set(word)

    result = {
        "status": "success",
        "shape_name": shape.name(),
    }
    if unsupported:
        result["warnings"] = [
            "PowerPoint for Mac's reflection format is the preset and nothing "
            f"else, so {', '.join(unsupported)} were not applied. The preset "
            "carries its own blur, offset, size and transparency."
        ]
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
