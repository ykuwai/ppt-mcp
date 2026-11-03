"""SmartArt tools, on Apple Events.

Mirrors ``ppt_com/smartart.py``. ``ppt_modify_smartart`` writes, adds and
deletes nodes (``set_text``, ``add_node``, ``delete_node``) through the
graphic's clipboard package; everything else here refuses.

**There is no smart art in this dictionary, and the gap is a strange shape.**
Parsing ``PowerPoint.sdef`` into the tables appscript builds gives a reference
table of every property, element and command PowerPoint answers to, and nothing
in it mentions SmartArt. The two words that come close are ``smart cut paste``
and ``smart quotes``, both autocorrect preferences. There is no ``smart art``
class, no ``smart art layouts`` on ``application``, no ``nodes`` on ``shape``.
So the whole of ``SmartArt.AllNodes``, which is what all eight actions of
``ppt_modify_smartart`` walk, has nothing to walk.

**The vocabulary for it was kept and nothing was left to take it.** The type
table has ``MsoSmartArtNodePosition`` with ``after node``, ``before node``,
``above node`` and ``below node``, ``MsoSmartArtNodeType`` with ``default node``
and ``assistant node``, and ``MsoOrgChartLayoutType`` with its six
organisation chart layouts. Every one of those enumerators is declared exactly
once and appears in no
property, no parameter and no command. They are the words for inserting a node
into a SmartArt graphic, with nothing left that accepts them. That is the
clearest evidence in the whole port that this was removed rather than never
built.

**What is not missing is the shape.** ``shape type smartart graphic`` is a real
enumerator, code ``0x008c0018``, which is the Windows ``msoSmartArt`` constant
24 in the low byte, so a SmartArt graphic already in the deck reports itself as
one. It can be found, named, moved, resized, read and deleted through the
ordinary shape tools, which is what every refusal here points at. Whether its
child shapes can be reached is a separate question and the honest answer is
that it has not been measured. The nearest measurement is groups.py's, where a
group answers 0 for its ``shapes`` and -1728 for ``shapes[1]``, and it is a
group rather than a SmartArt graphic, so it is not quoted here as though it
were.

``ppt_list_smartart_layouts`` was weighed rather than refused out of hand,
because this module already carries a static table of layout names. See its
implementation for why serving from it would be worse than refusing.

**A node's text is reachable another way (#189).** ``copy shape`` on a
SmartArt graphic writes the whole diagram to the clipboard package: its data
model, layout, style, colours and the drawing PowerPoint laid out. Pasting the
package back makes a SmartArt graphic again, and a node's text rewritten in it
shows (PowerPoint 16.114). So the three node actions copy the graphic,
change its data model (``gvml.smartart``), and paste it back over the
original with ``replace_shape``, which reads the paste back and checks every
node's text before it deletes anything. A node added or deleted only in the
data model is laid out again by PowerPoint on the paste, the drawing it
carried notwithstanding. A layout, colour scheme, style or node format would
have to come from another graphic's package, and none ships here, so those
five actions still refuse.

A refused call edits nothing, so it does not call ``goto_slide``, and leaves
the user's view where it was.
"""

import logging

from backend.mac_ae import ppt, slide_at as _slide
from backend.mac_enums import MsoShapeType
from backend.unsupported import refusal as _refusal
from ppt_com.constants import SHAPE_TYPE_NAMES, msoSmartArt
from ppt_mac.shapes import _get_shape, _shape_index, _win_constant

logger = logging.getLogger(__name__)

_SHAPE_TYPES = MsoShapeType

# The table runs Windows constant to macOS keyword, which is the direction a
# write needs. Reading a shape's type needs the way back.
_WIN_SHAPE_TYPE = {word: number for number, word in _SHAPE_TYPES.items()}

# The one finding every refusal in this module rests on, kept in one place so
# that a reader who meets it twice recognises it as the same finding.
_NO_SMART_ART = (
    "PowerPoint for Mac has no `smart art` class in its Apple Event "
    "dictionary, no `smart art layouts` on the application and no `nodes` on a "
    "shape. The enumerators for placing a node, `after node`, `before node`, "
    "`above node`, `below node`, `assistant node`, are all still declared and "
    "no property, parameter or command anywhere takes them."
)

# What survives, said the same way each time. An existing SmartArt graphic is a
# shape like any other, and these are the tools that treat it as one.
_SHAPE_TOOLS = [
    "ppt_get_shape_info, which reports a SmartArt shape's name, box and type",
    "ppt_update_shape, which moves and resizes it",
    "ppt_list_shapes",
    "ppt_delete_shape",
]

# The eight actions ppt_modify_smartart accepts, in the order its error message
# lists them, so the two cannot drift apart.
_ACTIONS = (
    "set_text", "add_node", "delete_node",
    "change_color", "change_style", "change_layout",
    "format_node", "format_all_nodes",
)

_LIST_TYPES = ("layouts", "colors", "styles", "categories")


def _require_smartart(shape):
    """Raise unless a shape really is a SmartArt graphic, and say what it is.

    The guard Windows puts in front of ``ppt_modify_smartart``, reading the
    type through the generated table so the number in the message is the
    Windows one a caller already knows.
    """
    type_val = _win_constant(_WIN_SHAPE_TYPE, shape.shape_type())
    if type_val != msoSmartArt:
        raise ValueError(
            f"Shape '{shape.name()}' is not a SmartArt graphic "
            f"(type={SHAPE_TYPE_NAMES.get(type_val, type_val)})"
        )


# ---------------------------------------------------------------------------
# Apple Event implementation functions
# ---------------------------------------------------------------------------
def _add_smartart_impl(slide_index, layout_name, layout_index, left, top, width, height,
                       node_texts, color_index, style_index, font_name, font_size, bold, font_color):
    """Refuse, because there is no layout to build from and nothing to build.

    Windows resolves a layout out of ``Application.SmartArtLayouts`` and hands
    it to ``Shapes.AddSmartArt``. Neither the collection nor the command is in
    the dictionary, and there is no ``make new smart art`` to reach around
    them with.
    """
    return _refusal(
        "ppt_add_smartart",
        f"{_NO_SMART_ART} Windows builds one by handing a layout from "
        "Application.SmartArtLayouts to Shapes.AddSmartArt, and neither the "
        "collection nor the command exists here. A SmartArt graphic inserted "
        "by hand is a different matter. It reports `shape type smartart "
        "graphic` and behaves as an ordinary shape from then on.",
        [
            "Insert the SmartArt once by hand in PowerPoint, then position it "
            "with ppt_update_shape",
            "ppt_add_shape with ppt_add_connector, which draws the same "
            "diagram out of parts this platform can make",
            "ppt_add_table, for a list or matrix layout",
        ],
    )


def _modify_smartart_impl(slide_index, shape_name_or_index, action,
                          node_index, text,
                          layout_name, layout_index,
                          color_index, style_index,
                          font_name, font_size, bold, font_color,
                          fill_color, line_color, line_width):
    """Refuse, after checking the shape and the action the way Windows does.

    All eight actions go through ``SmartArt.AllNodes``, so there is no action
    worth honouring partially and no argument worth singling out with
    ``error=``. What is worth keeping is the order. The shape is resolved and
    checked first, then the action name, so a caller who named the wrong shape
    or misspelled the action hears about that rather than about the platform.
    """
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, shape_name_or_index)
    _require_smartart(shape)

    if action not in _ACTIONS:
        raise ValueError(
            f"Unknown action '{action}'. Supported: "
            "'set_text', 'add_node', 'delete_node', "
            "'change_color', 'change_style', 'change_layout', "
            "'format_node', 'format_all_nodes'"
        )

    from gvml import smartart as gvml_smartart

    if action == "set_text":
        if node_index is None:
            raise ValueError("node_index is required for 'set_text' action")
        if text is None:
            raise ValueError("text is required for 'set_text' action")
        return _edit_nodes(pres, slide, slide_index, shape, action,
                           lambda data, drawing: gvml_smartart.set_node_text(
                               data, drawing, node_index, text))
    if action == "add_node":
        return _edit_nodes(pres, slide, slide_index, shape, action,
                           lambda data, drawing: gvml_smartart.add_node(data, node_index, text))
    if action == "delete_node":
        if node_index is None:
            raise ValueError("node_index is required for 'delete_node' action")
        return _edit_nodes(pres, slide, slide_index, shape, action,
                           lambda data, drawing: gvml_smartart.delete_node(data, node_index))

    return _refusal(
        "ppt_modify_smartart",
        f"'{shape.name()}' on slide {slide_index} is a SmartArt graphic and "
        f"the action '{action}' cannot be carried out on it. Nodes can be "
        "written, added and deleted (action='set_text', 'add_node', "
        "'delete_node'), through the graphic's clipboard package; a layout, "
        "colour scheme, style or node format would have to come from another "
        "graphic's package, and none ships with the server. "
        f"{_NO_SMART_ART}",
        [
            "ppt_modify_smartart with action='set_text', 'add_node' or 'delete_node'",
            "Edit the SmartArt by hand in PowerPoint",
            "ppt_ungroup_shapes, which does work and turns a converted "
            "diagram into shapes the text and formatting tools can reach",
        ] + _SHAPE_TOOLS,
        error=f"ppt_modify_smartart cannot do '{action}' on macOS",
    )


def _diagram_parts(package):
    """The names of the package's data model and drawing parts."""
    data = [p for p in package.parts if "/diagrams/data" in p and p.endswith(".xml")]
    drawing = [p for p in package.parts if "/diagrams/drawing" in p and p.endswith(".xml")]
    return (data[0] if data else None), (drawing[0] if drawing else None)


def _edit_nodes(pres, slide, slide_index, shape, action, edit) -> dict:
    """Change the graphic's nodes through its clipboard package.

    ``edit(data, drawing)`` changes the data model, and the drawing where it
    keeps up (a node's text), in place; a ValueError from it is passed on
    with nothing changed. The graphic is pasted back over the original with
    ``replace_shape``, and every node's text is read back before the
    original goes. Adding or deleting a node leaves the drawing as it was:
    PowerPoint lays the diagram out again from the data model when it is
    pasted (checked live on 16.114, a node added and one deleted).
    """
    import xml.etree.ElementTree as ET

    from gvml.package import PackageError, register_prefixes
    from gvml import smartart as gvml_smartart
    from gvml.canvas import XML_DECL
    from ppt_mac.gvml_paste import (
        Clipboard, Refused, copy_shape_package, replace_shape, with_warnings)

    name = shape.name()
    index = _shape_index(slide, name)
    left, top = shape.left_position(), shape.top()
    clip = Clipboard.take()
    try:
        try:
            package = copy_shape_package(shape, clip, "ppt_modify_smartart")
        except Refused as refused:
            return refused.payload
        data_part, drawing_part = _diagram_parts(package)
        if data_part is None:
            return _refusal(
                "ppt_modify_smartart",
                f"The package PowerPoint wrote for '{name}' holds no SmartArt "
                "data model, so there is no node to write. Nothing was changed.",
                ["Edit the SmartArt by hand in PowerPoint"] + _SHAPE_TOOLS,
            )
        for part in (data_part, drawing_part):
            if part is not None:
                register_prefixes(package.parts[part])
        data = ET.fromstring(package.parts[data_part])
        drawing = ET.fromstring(package.parts[drawing_part]) if drawing_part else None
        edit(data, drawing)
        wanted = gvml_smartart.node_texts(data)
        package.parts[data_part] = (XML_DECL + ET.tostring(data, encoding="unicode")).encode("utf-8")
        if drawing is not None:
            package.parts[drawing_part] = (
                XML_DECL + ET.tostring(drawing, encoding="unicode")).encode("utf-8")

        def check(readback):
            back, _ = _diagram_parts(readback)
            if back is None:
                return "the pasted graphic carries no SmartArt data model"
            try:
                got = gvml_smartart.node_texts(ET.fromstring(readback.parts[back]))
            except (PackageError, ET.ParseError) as exc:
                return f"its data model could not be read ({exc})"
            if got != wanted:
                return f"its nodes read back as {got!r}, not {wanted!r}"
            return None

        try:
            replaced = replace_shape(
                pres, slide, slide_index, index, package.to_bytes(), clip,
                "ppt_modify_smartart", _SHAPE_TYPES[msoSmartArt], left, top,
                "SmartArt graphic", check,
                ["Edit the SmartArt by hand in PowerPoint"],
            )
        except Refused as refused:
            return refused.payload
    finally:
        clip.restore()

    return with_warnings({
        "success": True,
        "action": action,
        "total_nodes": len(wanted),
        "shape_name": replaced.name,
    }, clip, replaced.warnings)


def _list_smartart_options_impl(list_type, category, keyword, include_description):
    """Refuse, rather than serve a catalogue nothing can consume.

    This one was worth a second look. ``ppt_com/smartart.py`` carries
    ``SMARTART_ENGLISH_ALIASES``, a static table of layout ids and English
    names, so a layouts listing could be answered without asking PowerPoint
    anything. It should not be.

    The payload's ``index`` is a 1-based position in
    ``Application.SmartArtLayouts``, which is the number ``layout_index`` feeds
    back to ``ppt_add_smartart``. There is no such collection here, so any
    index served from a static dict would be invented, and it would be invented
    for a tool that refuses. A ``success: True`` answer full of numbers that
    mean nothing is exactly the plausible looking no-op this port is built to
    avoid, so the catalogue stays behind the refusal and the English names stay
    where they are useful, which is on Windows.

    The ``list_type`` is still checked first, so a caller who wrote 'layout'
    hears about the spelling.
    """
    if list_type not in _LIST_TYPES:
        raise ValueError(
            f"Unknown list_type '{list_type}'. "
            "Use: 'layouts', 'colors', 'styles', or 'categories'"
        )

    return _refusal(
        "ppt_list_smartart_layouts",
        f"{_NO_SMART_ART} The layouts, colour schemes and quick styles this "
        "lists all live on Application.SmartArtLayouts, SmartArtColors and "
        "SmartArtQuickStyles, none of which exist here. The English layout "
        "names are still in the source, but the index this tool returns is a "
        "position in a collection macOS does not have, so answering from the "
        "static table would hand back numbers that mean nothing and that "
        "ppt_add_smartart could not use in any case.",
        [
            "Insert the SmartArt by hand in PowerPoint, where the layout "
            "gallery is the same one this tool lists",
            "ppt_add_shape and ppt_add_connector, for a diagram built out of "
            "parts this platform can make",
        ],
    )
