"""SmartArt tools, on Apple Events.

Mirrors ``ppt_com/smartart.py``. All three tools work, through clipboard
packages rather than Apple Events: ``ppt_modify_smartart`` edits a graphic's
package and pastes it back, ``ppt_add_smartart`` builds one and pastes it,
and ``ppt_list_smartart_layouts`` reads the layouts, colour schemes and
styles PowerPoint for Mac installs (``smartart_defs``).

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

**A node's text is reachable another way (#189).** ``copy shape`` on a
SmartArt graphic writes the whole diagram to the clipboard package: its data
model, layout, style, colours and the drawing PowerPoint laid out. Pasting the
package back makes a SmartArt graphic again, and a node's text rewritten in it
shows (PowerPoint 16.114). So the three node actions copy the graphic,
change its data model (``gvml.smartart``), and paste it back over the
original with ``replace_shape``, which reads the paste back and checks every
node's text before it deletes anything. A node added or deleted only in the
data model is laid out again by PowerPoint on the paste, the drawing it
carried notwithstanding. A node's own fill, outline and font are written
where PowerPoint for Windows keeps them, on the node's point.

**A layout, colour scheme or style is a whole part of the package**, and
PowerPoint for Mac installs every one it offers as a file of that same XML
(``smartart_defs``). ``change_layout``, ``change_color`` and ``change_style``
put the installed file in place of the part and name it on the data model's
document point; PowerPoint lays the graphic out again on the paste. A new
graphic is the same parts around a data model made from the layout's own
sample nodes, with no drawing at all, which PowerPoint draws on the paste.

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
# A new graphic's colour scheme and style when none is asked for, the ones
# PowerPoint gives a graphic inserted from the gallery.
_DEFAULT_COLORS = "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2"
_DEFAULT_STYLE = "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1"


def _no_definitions(tool, exc):
    return _refusal(
        tool,
        f"{exc} The layouts, colour schemes and styles are read from there, "
        "and nothing was changed.",
        ["Insert or restyle the SmartArt by hand in PowerPoint"] + _SHAPE_TOOLS,
        error=f"{tool} cannot find PowerPoint for Mac's SmartArt definitions",
    )


def _add_smartart_impl(slide_index, layout_name, layout_index, left, top, width, height,
                       node_texts, color_index, style_index, font_name, font_size, bold, font_color):
    """Build the graphic's package from the installed definitions and paste it.

    The data model starts from the layout's sample nodes, as a graphic
    inserted from the gallery does; ``node_texts`` fills them and adds more
    past the last. Style and colour scheme are both written into the package,
    so Windows' "style before colour" order has nothing to reset here. The
    paste is read back, and removed again unless every node's text is there.
    """
    import xml.etree.ElementTree as ET

    from gvml import Package, canvas
    from gvml import smartart as gvml_smartart
    from ppt_mac import smartart_defs as defs
    from ppt_mac.gvml_paste import (
        Clipboard, Refused, copy_shape_package, paste_package, unused_name, with_warnings)
    from ppt_mac.shapes import _shape_names
    from utils.color import hex_to_rgb

    if font_color is not None:
        hex_to_rgb(font_color)
    try:
        if layout_name:
            layout = defs.resolve_layout(layout_name)
        else:
            layout = defs.by_index("layouts", layout_index if layout_index else 1)
        colors = (defs.by_index("colors", color_index) if color_index is not None
                  else defs.by_id("colors", _DEFAULT_COLORS) or defs.by_index("colors", 1))
        style = (defs.by_index("styles", style_index) if style_index is not None
                 else defs.by_id("styles", _DEFAULT_STYLE) or defs.by_index("styles", 1))
    except defs.DefinitionsMissing as exc:
        return _no_definitions("ppt_add_smartart", exc)

    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    name = unused_name(_shape_names(slide), "Diagram")
    ids = {"layouts": (layout.id, layout.category), "colors": (colors.id, colors.category),
           "styles": (style.id, style.category)}
    raw = gvml_smartart.new_graphic(
        layout.xml(), colors.xml(), style.xml(), ids, node_texts,
        canvas.emu(left), canvas.emu(top), canvas.emu(width), canvas.emu(height), name,
        fonts={"font_name": font_name, "font_size": font_size, "bold": bold,
               "font_color": font_color},
    )
    built = Package.from_bytes(raw)
    wanted = gvml_smartart.node_texts(ET.fromstring(built.parts[_diagram_parts(built)[0]]))
    alternatives = ["Insert the SmartArt by hand in PowerPoint"] + _SHAPE_TOOLS

    clip = Clipboard.take()
    try:
        try:
            pasted = paste_package(
                pres, slide, slide_index, raw, clip, "ppt_add_smartart",
                _SHAPE_TYPES[msoSmartArt], left, top, alternatives,
            )
            back = copy_shape_package(pasted.shape, clip, "ppt_add_smartart")
        except Refused as refused:
            return refused.payload
        data_part, _ = _diagram_parts(back)
        got = (gvml_smartart.node_texts(ET.fromstring(back.parts[data_part]))
               if data_part else None)
        if got != wanted:
            pasted.shape.delete()
            return _refusal(
                "ppt_add_smartart",
                f"PowerPoint pasted the {layout.name} graphic, but its nodes read "
                f"back as {got!r}, not {wanted!r}, so it was removed again.",
                alternatives,
                error="ppt_add_smartart could not confirm the graphic's nodes",
            )
    finally:
        clip.restore()

    return with_warnings({
        "success": True,
        "shape_name": pasted.name,
        "node_count": len(got),
        "layout_name": layout.name,
    }, clip, pasted.warnings)


def _modify_smartart_impl(slide_index, shape_name_or_index, action,
                          node_index, text,
                          layout_name, layout_index,
                          color_index, style_index,
                          font_name, font_size, bold, font_color,
                          fill_color, line_color, line_width):
    """Carry out an action through the graphic's clipboard package.

    The shape is resolved and checked first, then the action name, as on
    Windows, so a caller who named the wrong shape or misspelled the action
    hears about that first.
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
    if action in ("format_node", "format_all_nodes"):
        if action == "format_node" and node_index is None:
            raise ValueError("node_index is required for 'format_node' action")
        from utils.color import hex_to_rgb

        for colour in (fill_color, line_color, font_color):
            if colour is not None:
                hex_to_rgb(colour)  # the same refusal Windows gives, before any copy
        which = node_index if action == "format_node" else None
        return _edit_nodes(pres, slide, slide_index, shape, action,
                           lambda data, drawing: gvml_smartart.format_nodes(
                               data, which, fill_color=fill_color, line_color=line_color,
                               line_width=line_width, font_name=font_name,
                               font_size=font_size, bold=bold, font_color=font_color))

    from ppt_mac import smartart_defs as defs

    try:
        if action == "change_layout":
            if layout_name:
                swaps = [("layouts", defs.resolve_layout(layout_name))]
            elif layout_index:
                swaps = [("layouts", defs.by_index("layouts", layout_index))]
            else:
                raise ValueError(
                    "layout_name or layout_index is required for 'change_layout' action")
        elif action == "change_color":
            if color_index is None:
                raise ValueError("color_index is required for 'change_color' action")
            swaps = [("colors", defs.by_index("colors", color_index))]
        else:  # change_style
            if style_index is None:
                raise ValueError("style_index is required for 'change_style' action")
            swaps = [("styles", defs.by_index("styles", style_index))]
            if color_index is not None:
                swaps.append(("colors", defs.by_index("colors", color_index)))
    except defs.DefinitionsMissing as exc:
        return _no_definitions("ppt_modify_smartart", exc)
    return _edit_nodes(pres, slide, slide_index, shape, action,
                       lambda data, drawing: None, swaps)


def _diagram_parts(package):
    """The names of the package's data model and drawing parts."""
    data = [p for p in package.parts if "/diagrams/data" in p and p.endswith(".xml")]
    drawing = [p for p in package.parts if "/diagrams/drawing" in p and p.endswith(".xml")]
    return (data[0] if data else None), (drawing[0] if drawing else None)


def _edit_nodes(pres, slide, slide_index, shape, action, edit, swaps=()) -> dict:
    """Change the graphic's nodes through its clipboard package.

    ``edit(data, drawing)`` changes the data model, and the drawing where it
    keeps up (a node's text), in place; a ValueError from it is passed on
    with nothing changed. The graphic is pasted back over the original with
    ``replace_shape``, and every node's text is read back before the
    original goes. Adding or deleting a node leaves the drawing as it was:
    PowerPoint lays the diagram out again from the data model when it is
    pasted (checked live on 16.114, a node added and one deleted).

    ``swaps`` is ``(kind, smartart_defs.Definition)`` pairs: each puts the
    installed definition in place of the package's layout, colours or
    quickStyle part and names it on the document point, and is checked on
    the read-back too. PowerPoint lays the graphic out again with it.
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
        for kind, definition in swaps:
            stem = "/diagrams/" + gvml_smartart.DEFINITIONS[kind][0]
            part = next((p for p in package.parts if stem in p and p.endswith(".xml")), None)
            if part is None:
                return _refusal(
                    "ppt_modify_smartart",
                    f"The package PowerPoint wrote for '{name}' holds no "
                    f"{kind[:-1]} part to replace. Nothing was changed.",
                    ["Edit the SmartArt by hand in PowerPoint"] + _SHAPE_TOOLS,
                )
            package.parts[part] = gvml_smartart.definition_part(
                kind, definition.xml(), definition.id)
            gvml_smartart.set_definition(data, kind, definition.id, definition.category)
        wanted_ids = {kind: definition.id for kind, definition in swaps}
        wanted = gvml_smartart.node_texts(data)
        wanted_formats = gvml_smartart.node_formats(data)
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
            formats = gvml_smartart.node_formats(ET.fromstring(readback.parts[back]))
            if formats != wanted_formats:
                return "its nodes' formatting did not read back as written"
            ids = gvml_smartart.definition_ids(ET.fromstring(readback.parts[back]))
            for kind, uid in wanted_ids.items():
                if ids.get(kind) != uid:
                    return f"its {kind[:-1]} reads back as {ids.get(kind)}, not {uid}"
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

    result = {
        "success": True,
        "action": action,
        "total_nodes": len(wanted),
        "shape_name": replaced.name,
    }
    for kind, definition in swaps:
        result[f"{kind[:-1]}_name"] = definition.name
    return with_warnings(result, clip, replaced.warnings)


def _list_smartart_options_impl(list_type, category, keyword, include_description):
    """List what PowerPoint for Mac installs, in the shape Windows lists it.

    Colour schemes and styles come in Windows' order, so their indices are
    the Windows ones. Layouts are numbered in this list's own order (see
    ``smartart_defs``), which is what ``layout_index`` takes here. A layout's
    ``english_name`` is its name, which is already English on the Mac.
    """
    if list_type not in _LIST_TYPES:
        raise ValueError(
            f"Unknown list_type '{list_type}'. "
            "Use: 'layouts', 'colors', 'styles', or 'categories'"
        )
    from ppt_mac import smartart_defs as defs

    try:
        layouts = defs.definitions("layouts")
        if list_type == "categories":
            counts = {}
            for d in layouts:
                if d.category:
                    counts[d.category] = counts.get(d.category, 0) + 1
            return {
                "success": True,
                "list_type": "categories",
                "total_layouts": len(layouts),
                "categories": [{"category": k, "count": v} for k, v in counts.items()],
            }
        items_all = defs.definitions(list_type)
    except defs.DefinitionsMissing as exc:
        return _no_definitions("ppt_list_smartart_layouts", exc)

    cat_lower = category.lower() if category else None
    kw_lower = keyword.lower() if keyword else None
    items = []
    for d in items_all:
        english_name = d.name if list_type == "layouts" else None
        if cat_lower and list_type == "layouts" and cat_lower not in d.category.lower():
            continue
        if kw_lower and kw_lower not in d.name.lower() and not (
                english_name and kw_lower in english_name.lower()):
            continue
        entry = {"index": d.index, "name": d.name}
        if list_type == "layouts":
            entry["english_name"] = english_name
            entry["category"] = d.category
        if include_description:
            entry["description"] = d.description
        items.append(entry)

    return {
        "success": True,
        "list_type": list_type,
        "total_count": len(items_all),
        "filtered_count": len(items),
        list_type: items,
    }
