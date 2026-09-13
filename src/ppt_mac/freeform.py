"""Freeform path tools, on Apple Events and the clipboard.

Mirrors ``ppt_com/freeform.py``. Same function names, same signatures, and the
same wire form, which in this one module is a JSON string rather than a dict,
because the freeform tools return what ``ppt.execute`` hands back without
encoding it again.

**There is no node anywhere in the dictionary.** No ``node``, no ``vertex``,
no ``vertices``, no ``segment`` on a shape and no ``build freeform``. The
near misses were checked: ``line shape`` is two points and not a path,
``motion effect`` has a writable ``path`` that is an animation route rather
than an outline, ``text frame``'s ``path format`` bends text, and
``adjustment`` parametrises a built-in shape and a freeform has none.

**The clipboard carries what the dictionary does not.** A freeform copied with
``copy shape`` lands on the pasteboard as a DrawingML package whose
``a:custGeom`` holds every point, and ``paste object`` takes such a package
back. So ``ppt_build_freeform`` writes the path as ``a:custGeom`` and pastes
it, and ``ppt_get_shape_nodes`` copies the shape and reads the path out. The
numbering is the Windows one, one node per straight segment and three per
curve; see ``gvml/freeform.py``. The procedure and its checks are in
``ppt_mac/gvml_paste.py``.

**The five tools that edit nodes still refuse.** They are the design's third
tier (docs/gvml-design.md section 6): copy, edit the path, paste, delete the
original, put position and z order back. The parts are all here now, and
those tools are not yet written. ``ppt_set_node_editing_type`` will not be:
corner, smooth and symmetric are not in the dictionary and not in the XML
either. PowerPoint's UI derives them from where the handles sit, and there is
nowhere to write one.
"""

import json
import logging

from backend.mac_ae import ppt, slide_at as _slide
from backend.mac_enums import MsoShapeType
from backend.unsupported import refusal as _refusal
from gvml import PackageError, build
from gvml import canvas as _canvas
from gvml import freeform as _gvml_freeform
from gvml import shapes as _gvml_shapes
from ppt_com.constants import msoFreeform
from ppt_mac.gvml_paste import (
    Clipboard,
    Refused,
    copy_shape_package,
    paste_package,
    unused_name,
    with_warnings,
)
from ppt_mac.shapes import _WIN_SHAPE_TYPE, _get_shape, _shape_names, _win_constant

logger = logging.getLogger(__name__)

# The one finding every remaining refusal in this module rests on.
_NO_NODES = (
    "PowerPoint for Mac's Apple Event dictionary has no `node`, no `vertex`, "
    "no `vertices` and no `segment` on a shape. Windows walks Shape.Nodes for "
    "this, and it has no counterpart here. The path can be read with "
    "ppt_get_shape_nodes, which goes through the clipboard; editing it the "
    "same way is not yet written."
)

_BUILD_ALTERNATIVES = [
    "ppt_add_shape, which reaches every built-in autoshape",
    "ppt_add_line, one call per straight segment",
    "Draw the path by hand in PowerPoint, then position it with ppt_update_shape",
]

# What survives, said the same way each time.
_SHAPE_TOOLS = [
    "ppt_get_shape_nodes, which reads the outline",
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
    """Write the path as ``a:custGeom`` and paste it.

    ``start_et_int`` is accepted for signature parity and has no effect: the
    editing type of a node is not stored in the XML, on either platform, and
    Windows only uses it to choose between the two ``AddNodes`` forms.
    Control points for an ``auto`` curve are computed here (a Catmull-Rom
    spline through the anchors) where Windows computes its own.
    """
    logger.debug("building a freeform of %d segment(s), start editing type %r",
                 len(nodes_data), start_et_int)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)

    geometry, x, y, cx, cy = _gvml_freeform.build_geometry(start_x, start_y, nodes_data, close_path)
    name = shape_name or unused_name(_shape_names(slide), "Freeform")
    drawing = _canvas.wrap(_gvml_shapes.shape_xml(2, name, x, y, cx, cy, geometry), x, y, cx, cy)
    raw = build(drawing)

    clip = Clipboard.take()
    try:
        try:
            pasted = paste_package(
                pres, slide, slide_index, raw, clip, "ppt_build_freeform",
                MsoShapeType[msoFreeform], _canvas.pt(x), _canvas.pt(y),
                _BUILD_ALTERNATIVES,
            )
        except Refused as refused:
            return json.dumps(refused.payload)
    finally:
        clip.restore()

    shape = pasted.shape
    return json.dumps(with_warnings({
        "success": True,
        "shape_name": pasted.name,
        "shape_index": shape.z_order_position(),
        "left": round(shape.left_position(), 2),
        "top": round(shape.top(), 2),
        "width": round(shape.width(), 2),
        "height": round(shape.height(), 2),
    }, clip, pasted.warnings))


def _get_shape_nodes_impl(slide_index, shape_name, shape_index):
    """Copy the freeform and read its path out of the package.

    Nothing on the slide changes and the view is not moved. The clipboard is
    put back afterwards.
    """
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, None, shape_name=shape_name, shape_index=shape_index)
    _check_freeform(shape)
    name = shape.name()

    clip = Clipboard.take()
    try:
        try:
            package = copy_shape_package(shape, clip, "ppt_get_shape_nodes")
        except Refused as refused:
            return json.dumps(refused.payload)
        try:
            element = _gvml_shapes.only_child(
                _canvas.children(_canvas.parse(package.drawing())), "freeform"
            )
            nodes = _gvml_freeform.read_nodes(element)
        except PackageError as exc:
            return json.dumps(_refusal(
                "ppt_get_shape_nodes",
                f"'{name}' reports itself as a freeform, but the path in the "
                f"package PowerPoint wrote for it could not be read: {exc}",
                _SHAPE_TOOLS[1:],
            ))
    finally:
        clip.restore()

    return json.dumps(with_warnings({
        "shape_name": name,
        "node_count": len(nodes),
        "nodes": nodes,
    }, clip))


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
    """Refuse, because a node's editing type has no representation anywhere.

    Nothing in the dictionary pairs with ``MsoEditingType``, and nothing in
    the XML does either. ``a:custGeom`` stores points; corner, smooth and
    symmetric are what PowerPoint's UI calls the geometry of the two handles
    at a point, and there is no attribute to write one into. So this one
    stays refused even now that the path can be rewritten through the
    clipboard.
    """
    return _refuse_for_node_tool(
        "ppt_set_node_editing_type",
        slide_index, shape_name, shape_index,
        f"the editing type of node {node_index} cannot be set. Nothing in the "
        "dictionary pairs with MsoEditingType, and the DrawingML the clipboard "
        "carries has no such attribute either: corner, smooth and symmetric "
        "are read off the handle geometry, not stored. Move the handles "
        "instead, once ppt_set_node_position is written for this platform.",
    )


def _set_segment_type_impl(slide_index, shape_name, shape_index, node_index, seg_int):
    """Refuse, because a segment cannot yet be rewritten here."""
    return _refuse_for_node_tool(
        "ppt_set_segment_type",
        slide_index, shape_name, shape_index,
        f"the segment after node {node_index} cannot be switched between line "
        "and curve. Nothing in the dictionary pairs with MsoSegmentType.",
    )
