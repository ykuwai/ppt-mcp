"""Freeform path tools, on Apple Events.

Mirrors ``ppt_com/freeform.py``. Same function names, same signatures, and the
same wire form, which in this one module is a JSON string rather than a dict,
because the freeform tools return what ``ppt.execute`` hands back without
encoding it again.

All seven tools refuse. Four near misses were checked before concluding that,
and each is named below so nobody spends an afternoon re-deriving them.

**There is no node anywhere in the object model.** Parsing ``PowerPoint.sdef``
into the tables appscript builds gives a reference table of every property,
element and command PowerPoint answers to, and it holds no ``node``, no
``vertex``, no ``vertices``, no ``segment`` on a shape and no ``build
freeform``. The word
``node`` survives only as enumerator values, and those belong to SmartArt. So
``Shapes.BuildFreeform`` has no counterpart and neither does ``Shape.Nodes``,
which is what the other six tools walk.

**``line shape`` is real and it is not a freeform.** It is a declared class
inheriting ``shape``, with ``begin line X``, ``begin line Y``, ``end line X``
and ``end line Y`` all writable, and ``ppt_add_line`` already makes one. Two
points and a straight segment is genuinely reachable. What it is not is a path.
Each segment would be its own shape, so a five segment outline is five shapes
rather than one, it cannot be closed into a region, it cannot be filled, and it
cannot be moved or scaled as a unit. Drawing one silently and calling it a
freeform would be the substitution this port exists to refuse, so
``ppt_build_freeform`` names ``ppt_add_line`` as a route the caller can choose
rather than taking it for them.

**``motion effect`` has a writable ``path`` and it draws nothing.** It is the
only place in the dictionary where a path is handed over as data, a text
property reached through an animation behaviour on an animation effect. What it
describes is where a shape travels during an animation, not what the shape looks
like, and PowerPoint also ships sixty odd ``animation type ... path``
enumerators for the same purpose. Close enough to look like the answer, and not
the answer.

**``text frame`` has a ``path format`` and that is not it either.** Four preset
curves for text to follow, chosen from an enumeration. It bends the text inside
a shape and leaves the shape's outline where it was.

**``adjustment`` is the one piece of shape geometry that is writable.** ``shape``
really does have an ``adjustment`` element, and an ``adjustment`` really does
carry a writable ``adjustment value``, which is why ``ppt_update_shape`` can
already drag an autoshape's yellow handles. It is not a way in here. Adjustments
parametrise a built-in shape, an arrow's head width or a rounded rectangle's
corner, and a freeform has none, because its geometry is its node list and there
is no word for that.

**What is not missing is the shape.** ``shape type free form`` is a real
enumerator, code ``0x008c0005``, which is the Windows ``msoFreeform`` constant 5
in the low byte, so a freeform already in the deck reports itself as one and can
be found, named, moved, resized, read and deleted through the ordinary shape
tools. Its outline is what cannot be read or edited, not the shape.

Nothing here edits a slide, so nothing calls ``goto_slide``, and a refused call
leaves the user's view where it was.
"""

import json
import logging

from backend.mac_ae import ppt, slide_at as _slide
from backend.unsupported import refusal as _refusal
from ppt_com.constants import msoFreeform
from ppt_mac.shapes import _WIN_SHAPE_TYPE, _get_shape, _win_constant

logger = logging.getLogger(__name__)

# The one finding every refusal in this module rests on, kept in one place so
# that a reader who meets it twice recognises it as the same finding.
_NO_NODES = (
    "PowerPoint for Mac's Apple Event dictionary has no `node`, no `vertex`, "
    "no `vertices` and no `segment` on a shape, and no `build freeform` "
    "command. Windows walks Shape.Nodes and Shape.Vertices for this, and "
    "neither has a counterpart here."
)

# What survives, said the same way each time. An existing freeform is a shape
# like any other, and these are the tools that treat it as one.
_SHAPE_TOOLS = [
    "ppt_get_shape_info, which reports a freeform's type, name and box",
    "ppt_update_shape, which moves and resizes it",
    "ppt_list_shapes",
    "ppt_delete_shape",
]


def _check_freeform(shape):
    """Raise unless a shape really is a freeform, and say what it is instead.

    The guard Windows puts in front of all six node tools. The message is the
    one Windows gives, word for word, because a caller who moves between the
    two should not have to learn it twice.
    """
    type_val = _win_constant(_WIN_SHAPE_TYPE, shape.shape_type())
    if type_val != msoFreeform:
        raise ValueError(
            f"Shape '{shape.name()}' is not a freeform (type={type_val}). "
            "Only freeform shapes (type=5) support node operations."
        )


def _refuse_for_node_tool(tool_name, slide_index, shape_name, shape_index, detail):
    """Find the freeform, then refuse, naming it.

    Resolving the shape first is what separates a caller who pointed at the
    wrong shape from a caller who picked a platform that cannot help. The first
    hears about the shape. Nothing is written and the view is not moved, so a
    refusal costs the deck nothing.

    Returns a JSON string, which is what this module's tools return.
    """
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, None, shape_name=shape_name, shape_index=shape_index)
    _check_freeform(shape)

    return json.dumps(_refusal(
        tool_name,
        f"'{shape.name()}' on slide {slide_index} is a freeform and {detail} "
        f"{_NO_NODES}",
        [
            "Edit the path by hand in PowerPoint, with Edit Points",
        ] + _SHAPE_TOOLS,
    ))


# ---------------------------------------------------------------------------
# Apple Event implementation functions
# ---------------------------------------------------------------------------
def _build_freeform_impl(slide_index, start_et_int, start_x, start_y, nodes_data, close_path, shape_name):
    """Refuse, because there is no builder and no honest stand-in for one.

    The segment count and whether the path closes both go into the message.
    They are what decides whether the ``ppt_add_line`` route is worth taking,
    so a caller can tell from the refusal alone rather than by trying it.
    """
    logger.debug("refusing a freeform of %d segment(s)", len(nodes_data))

    return json.dumps(_refusal(
        "ppt_build_freeform",
        f"{_NO_NODES} The path asked for here has {len(nodes_data)} segment(s) "
        f"and close_path={close_path}, and none of it can be drawn as one "
        "shape. `line shape` does exist and is writable, so a straight segment "
        "can be drawn on its own, but each one is a separate shape, so a "
        "polyline built that way cannot be closed, filled, moved or scaled as "
        "a unit, and calling it a freeform would be untrue. `motion effect` "
        "has a writable `path`, and that is the route an animation takes "
        "across the slide rather than an outline.",
        [
            "Draw the path once by hand in PowerPoint, then position it with "
            "ppt_update_shape",
            "ppt_add_line, one call per straight segment, when separate lines "
            "are acceptable",
            "ppt_add_shape, which reaches every built-in autoshape including "
            "the arrows, stars and flowchart outlines",
        ],
    ))


def _get_shape_nodes_impl(slide_index, shape_name, shape_index):
    """Refuse, because a freeform will not say what its outline is.

    Worth its own wording rather than the shared one. A caller reading nodes is
    usually trying to find out what is on the slide, and the useful answer is
    that the shape itself reads fine and only its outline does not.
    """
    return _refuse_for_node_tool(
        "ppt_get_shape_nodes",
        slide_index, shape_name, shape_index,
        "its outline cannot be read. The shape answers for its name, type, "
        "position, size, fill and line, and there is no property anywhere that "
        "reports the points those are drawn between.",
    )


def _set_node_position_impl(slide_index, shape_name, shape_index, node_index, x, y):
    """Refuse, because there is no node to move."""
    return _refuse_for_node_tool(
        "ppt_set_node_position",
        slide_index, shape_name, shape_index,
        f"node {node_index} cannot be moved, or found, or counted. Windows "
        "calls Shape.Nodes.SetPosition for this.",
    )


def _insert_node_impl(slide_index, shape_name, shape_index, after_index, seg_int, et_int, x1, y1, x2, y2, x3, y3):
    """Refuse, because there is nothing to insert into."""
    return _refuse_for_node_tool(
        "ppt_insert_node",
        slide_index, shape_name, shape_index,
        f"nothing can be inserted after node {after_index}. Windows calls "
        "Shape.Nodes.Insert for this.",
    )


def _delete_node_impl(slide_index, shape_name, shape_index, node_index):
    """Refuse, because there is no node to delete.

    A silent no-op would be at its worst here. This tool is destructive on
    Windows, so a caller who believes it worked will go on to re-index the
    remaining nodes against a path that never changed.
    """
    return _refuse_for_node_tool(
        "ppt_delete_node",
        slide_index, shape_name, shape_index,
        f"node {node_index} cannot be deleted. Windows calls "
        "Shape.Nodes.Delete for this, and the outline is untouched here rather "
        "than partly edited.",
    )


def _set_node_editing_type_impl(slide_index, shape_name, shape_index, node_index, et_int):
    """Refuse, because a node's editing type has no representation.

    Nothing in the dictionary pairs with ``MsoEditingType``. The nearest thing
    is ``adjustment value``, which is a number on a built-in shape rather than
    a corner or smooth flag on a point.
    """
    return _refuse_for_node_tool(
        "ppt_set_node_editing_type",
        slide_index, shape_name, shape_index,
        f"the editing type of node {node_index} cannot be set. Nothing in the "
        "dictionary pairs with MsoEditingType, so corner, smooth and symmetric "
        "have no spelling here.",
    )


def _set_segment_type_impl(slide_index, shape_name, shape_index, node_index, seg_int):
    """Refuse, because a segment has no representation either."""
    return _refuse_for_node_tool(
        "ppt_set_segment_type",
        slide_index, shape_name, shape_index,
        f"the segment after node {node_index} cannot be switched between line "
        "and curve. Nothing in the dictionary pairs with MsoSegmentType.",
    )
