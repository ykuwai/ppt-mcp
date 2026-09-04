"""Shape grouping tools, on Apple Events.

Mirrors ``ppt_com/groups.py``. Same function names, same signatures, same
returned shapes.

Three things about groups on this side are worth knowing before reading on.

**Nothing can be gathered into a shape range, so nothing can be grouped.**
``group`` takes a shape range, and the only shape range PowerPoint for Mac hands
out is the one already selected in a window. Its dictionary has an ``unselect``
command and no ``select``, so a script has no way to put shapes into a selection
to act on them. ``ppt_group_shapes`` refuses on exactly the grounds
``ppt_merge_shapes`` does, and borrows layout.py's wording so the two agree.

**A group's members cannot be read.** The dictionary gives ``shape`` a ``shape``
element, which reads as though a group's members are reached the way a slide's
shapes are. They are not. A group of two text boxes answers 0 for its ``shapes``,
0 for its ``text boxes``, 0 for every other subclass collection, and -1728 for
``shapes[1]``. ``has child`` answers ``missing value``. So
``ppt_get_group_items`` refuses rather than returning an empty list, which would
read as a group with nothing in it.

**Ungrouping works, and it is the only thing here that does.** ``ungroup`` is
declared to take a shape range like ``group`` is, but its description reads "the
specified shape or range of shapes" and a single group shape is accepted.
Verified live, on a group of two that came apart with both members appearing on
the slide by name. The check is the group's own disappearance and the slide
growing, because the members cannot be listed beforehand to compare against.
"""

import logging

from appscript.reference import CommandError

from backend.mac_ae import ppt
from backend.unsupported import refusal as _refusal
from ppt_com.constants import msoGroup
from ppt_mac.layout import _NO_SHAPE_RANGE
from ppt_mac.shapes import (
    _WIN_SHAPE_TYPE,
    _get_shape,
    _shape_names,
    _slide,
    _win_constant,
)
from utils.navigation import goto_slide

logger = logging.getLogger(__name__)


def _require_group(shape, verb: str):
    """Raise unless a shape really is a group, and say what it is instead.

    The same guard Windows puts in front of both group tools, reading the type
    through the generated table so the number in the message is the Windows one
    a caller already knows.
    """
    type_val = _win_constant(_WIN_SHAPE_TYPE, shape.shape_type())
    if type_val != msoGroup:
        raise ValueError(
            f"Shape '{shape.name()}' is not a group (type={type_val}). "
            f"Only group shapes (type={msoGroup}) can be {verb}."
        )


# ---------------------------------------------------------------------------
# Apple Event implementation functions
# ---------------------------------------------------------------------------
def _group_shapes_impl(slide_index, shape_names):
    """Refuse, because there is no way to build the shape range grouping needs."""
    return _refusal(
        "ppt_group_shapes",
        "PowerPoint for Mac's `group` command takes a shape range and there is "
        "no way for a script to build one. " + _NO_SHAPE_RANGE,
        [
            "Group the shapes by hand in PowerPoint",
            "ppt_align_shapes",
            "ppt_distribute_shapes",
        ],
    )


def _ungroup_shapes_impl(slide_index, shape_name_or_index):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, shape_name_or_index)
    _require_group(shape, "ungrouped")

    # The members cannot be listed first, so the before picture is the slide.
    group_name = shape.name()
    before = _shape_names(slide)

    failure = None
    try:
        shape.ungroup()
    except CommandError as exc:
        # Held rather than raised. PowerPoint hands back a shape range here and
        # a reply appscript cannot unpack looks exactly like a command that was
        # refused, so the slide decides which happened, not the exception.
        failure = exc

    # A group that came apart is gone from the slide, and its members stand
    # where it stood. The group's own disappearance is the whole test. Counting
    # is not, because PowerPoint will let a group hold one shape and that one
    # comes back out leaving the slide exactly as long as it was.
    after = _shape_names(slide)
    came_apart = group_name not in after
    if not came_apart:
        return _refusal(
            "ppt_ungroup_shapes",
            f"PowerPoint reported no error and '{group_name}' is still on "
            f"slide {slide_index} as one shape, so nothing came apart. This is "
            "the silent no-op recorded in MACOS_PORT section 5."
            + (f" PowerPoint answered {failure}." if failure else ""),
            ["Ungroup the shape by hand in PowerPoint"],
        )

    members = [name for name in after if name not in before]
    return {
        "success": True,
        "ungrouped_count": len(members),
        "shape_names": members,
    }


def _get_group_items_impl(slide_index, shape_name_or_index):
    """Refuse, because a group will not say what is inside it.

    The dictionary reads as though this works. ``shape`` has a ``shape``
    element, and a group is a shape. Against a real group of two text boxes,
    every route answers nothing. ``shapes`` counts 0, so does ``text boxes``
    and every other subclass collection, ``shapes[1]`` answers -1728, and
    ``has child`` answers ``missing value``. An empty list would read as a
    group with nothing in it, which is worse than saying so.
    """
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, shape_name_or_index)
    _require_group(shape, "inspected")

    return _refusal(
        "ppt_get_group_items",
        f"'{shape.name()}' is a group and PowerPoint for Mac will not say what "
        "is in it. Its `shapes` collection counts zero, so does every subclass "
        "collection, and indexing into it fails with Apple Event error "
        "-1728, which is PowerPoint saying the reference does not resolve. "
        "Ungrouping it does "
        "work, and the members can be read individually once they are on the "
        "slide in their own right.",
        ["ppt_ungroup_shapes", "ppt_list_shapes"],
    )
