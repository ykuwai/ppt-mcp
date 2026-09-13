"""Shapes on a canvas: writing ``a:sp`` and ``a:grpSp``, reading any of them.

Writing is by template. The ``a:style`` block is the one PowerPoint gives a
shape made with ``make new shape``, so a freeform built here is filled and
outlined like a rectangle drawn beside it. The ``a:txSp`` block is what makes
the shape able to take text afterwards through the ordinary text tools.

Reading covers every shape kind a canvas can hold, because a group copied off
a slide holds whatever was grouped: text boxes, pictures, charts, lines,
freeforms, other groups. ``describe`` answers the same four numbers and the
Windows type constant for each, which is what ``ppt_get_group_items`` returns.
"""

import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import Iterable, List, Optional, Tuple

from gvml.canvas import pt
from gvml.package import NS_A, NS_C, NS_R, PackageError

# Windows MsoShapeType constants, spelled here so gvml imports nothing from
# ppt_com (design section 1: that line is the one that matters).
MSO_AUTO_SHAPE = 1
MSO_CHART = 3
MSO_FREEFORM = 5
MSO_GROUP = 6
MSO_LINE = 9
MSO_PICTURE = 13
MSO_TEXT_BOX = 17
MSO_TABLE = 19

_TYPE_NAMES = {
    MSO_AUTO_SHAPE: "AutoShape", MSO_CHART: "Chart", MSO_FREEFORM: "Freeform",
    MSO_GROUP: "Group", MSO_LINE: "Line", MSO_PICTURE: "Picture",
    MSO_TEXT_BOX: "TextBox", MSO_TABLE: "Table",
}

_A = f"{{{NS_A}}}"

# The a:style PowerPoint writes for a shape it made itself. Theme references,
# so the shape follows the deck's palette.
STYLE_XML = (
    '<a:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef>'
    '<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>'
    '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>'
    '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></a:style>'
)

# An empty text body, so the shape can take text later.
TEXT_XML = (
    '<a:txSp><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></a:txBody>'
    '<a:useSpRect/></a:txSp>'
)

RECT_GEOMETRY_XML = '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'


def _escape(text: str) -> str:
    return (
        text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def shape_xml(
    shape_id: int, name: str, x: int, y: int, cx: int, cy: int,
    geometry_xml: str = RECT_GEOMETRY_XML,
) -> str:
    """One ``a:sp`` at ``x, y`` EMU with the given geometry.

    Element order: nvSpPr, spPr, txSp, style. Copied from PowerPoint's own
    output and pinned by the golden fixture test.
    """
    return (
        f'<a:sp><a:nvSpPr><a:cNvPr id="{shape_id}" name="{_escape(name)}"/><a:cNvSpPr/></a:nvSpPr>'
        f'<a:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>{geometry_xml}</a:spPr>'
        f'{TEXT_XML}{STYLE_XML}</a:sp>'
    )


def group_xml(group_id: int, name: str, x: int, y: int, cx: int, cy: int, children_xml: Iterable[str]) -> str:
    """One ``a:grpSp`` whose children keep their slide coordinates.

    ``chOff``/``chExt`` equal ``off``/``ext``, so no child is rescaled: each
    child's own ``a:off`` is where it was on the slide, and the group's box is
    their union.
    """
    return (
        f'<a:grpSp><a:nvGrpSpPr><a:cNvPr id="{group_id}" name="{_escape(name)}"/><a:cNvGrpSpPr/></a:nvGrpSpPr>'
        f'<a:grpSpPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/>'
        f'<a:chOff x="{x}" y="{y}"/><a:chExt cx="{cx}" cy="{cy}"/></a:xfrm></a:grpSpPr>'
        + "".join(children_xml)
        + "</a:grpSp>"
    )


# ---------------------------------------------------------------------------
# Reading
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class ShapeInfo:
    """What a canvas child says about itself, in EMU."""

    name: str
    type: int
    x: int
    y: int
    cx: int
    cy: int
    element: ET.Element

    @property
    def type_name(self) -> str:
        return _TYPE_NAMES.get(self.type, f"Unknown({self.type})")


_NON_VISUAL = {
    f"{_A}sp": f"{_A}nvSpPr",
    f"{_A}grpSp": f"{_A}nvGrpSpPr",
    f"{_A}pic": f"{_A}nvPicPr",
    f"{_A}graphicFrame": f"{_A}nvGraphicFramePr",
    f"{_A}cxnSp": f"{_A}nvCxnSpPr",
}


def name_element(element: ET.Element) -> ET.Element:
    """The ``a:cNvPr`` of a canvas child, where its name lives."""
    nv = _NON_VISUAL.get(element.tag)
    if nv is None:
        raise PackageError(f"{element.tag} is not a shape element")
    cnv = element.find(f"{nv}/{_A}cNvPr")
    if cnv is None:
        raise PackageError(f"{element.tag} has no cNvPr")
    return cnv


def name_of(element: ET.Element) -> str:
    return name_element(element).get("name", "")


def rename(element: ET.Element, name: str) -> None:
    name_element(element).set("name", name)


def xfrm_of(element: ET.Element) -> Tuple[int, int, int, int]:
    """``x, y, cx, cy`` of a canvas child, in EMU."""
    if element.tag == f"{_A}grpSp":
        xfrm = element.find(f"{_A}grpSpPr/{_A}xfrm")
    elif element.tag == f"{_A}graphicFrame":
        xfrm = element.find(f"{_A}xfrm")
    else:
        xfrm = element.find(f"{_A}spPr/{_A}xfrm")
    if xfrm is None:
        raise PackageError(f"{element.tag} has no transform")
    off = xfrm.find(f"{_A}off")
    ext = xfrm.find(f"{_A}ext")
    if off is None or ext is None:
        raise PackageError(f"{element.tag} has a transform without offset or extent")
    return int(off.get("x", 0)), int(off.get("y", 0)), int(ext.get("cx", 0)), int(ext.get("cy", 0))


def type_of(element: ET.Element) -> int:
    """The Windows shape type constant a canvas child corresponds to."""
    tag = element.tag
    if tag == f"{_A}grpSp":
        return MSO_GROUP
    if tag == f"{_A}pic":
        return MSO_PICTURE
    if tag == f"{_A}cxnSp":
        return MSO_LINE
    if tag == f"{_A}graphicFrame":
        data = element.find(f"{_A}graphic/{_A}graphicData")
        uri = data.get("uri", "") if data is not None else ""
        if uri == NS_C:
            return MSO_CHART
        if uri.endswith("/table"):
            return MSO_TABLE
        return MSO_CHART if data is not None and data.find(f"{{{NS_C}}}chart") is not None else MSO_AUTO_SHAPE
    if tag == f"{_A}sp":
        if element.find(f"{_A}spPr/{_A}custGeom") is not None:
            return MSO_FREEFORM
        cnv = element.find(f"{_A}nvSpPr/{_A}cNvSpPr")
        if cnv is not None and cnv.get("txBox") == "1":
            return MSO_TEXT_BOX
        return MSO_AUTO_SHAPE
    raise PackageError(f"{tag} is not a shape element")


def describe(element: ET.Element) -> ShapeInfo:
    x, y, cx, cy = xfrm_of(element)
    return ShapeInfo(name_of(element), type_of(element), x, y, cx, cy, element)


def bounding_box(infos: Iterable[ShapeInfo]) -> Tuple[int, int, int, int]:
    """The union of several boxes, as ``x, y, cx, cy``."""
    infos = list(infos)
    if not infos:
        raise ValueError("no shapes to bound")
    left = min(i.x for i in infos)
    top = min(i.y for i in infos)
    right = max(i.x + i.cx for i in infos)
    bottom = max(i.y + i.cy for i in infos)
    return left, top, right - left, bottom - top


def group_items(group: ET.Element) -> List[dict]:
    """A group's members as ``ppt_get_group_items`` reports them, in points.

    A member's ``a:off`` is in the group's child space, which is what
    ``chOff``/``chExt`` describe, and the group's ``off``/``ext`` say where
    that space sits on the slide. A group that has been resized since its
    members were placed has the two differ, and the scale between them is what
    turns a member's box into slide coordinates.
    """
    if group.tag != f"{_A}grpSp":
        raise PackageError(f"expected a group, found {group.tag}")
    xfrm = group.find(f"{_A}grpSpPr/{_A}xfrm")
    if xfrm is None:
        raise PackageError("the group has no transform")
    off = xfrm.find(f"{_A}off")
    ext = xfrm.find(f"{_A}ext")
    ch_off = xfrm.find(f"{_A}chOff")
    ch_ext = xfrm.find(f"{_A}chExt")
    if None in (off, ext, ch_off, ch_ext):
        raise PackageError("the group transform is incomplete")
    gx, gy = int(off.get("x", 0)), int(off.get("y", 0))
    gcx, gcy = int(ext.get("cx", 0)), int(ext.get("cy", 0))
    cx0, cy0 = int(ch_off.get("x", 0)), int(ch_off.get("y", 0))
    ccx, ccy = int(ch_ext.get("cx", 0)), int(ch_ext.get("cy", 0))
    sx = gcx / ccx if ccx else 1.0
    sy = gcy / ccy if ccy else 1.0

    items = []
    for child in group:
        if child.tag not in _NON_VISUAL:
            continue
        info = describe(child)
        items.append({
            "name": info.name,
            "type": info.type,
            "type_name": info.type_name,
            "left": pt(gx + (info.x - cx0) * sx),
            "top": pt(gy + (info.y - cy0) * sy),
            "width": pt(info.cx * sx),
            "height": pt(info.cy * sy),
        })
    return items


def strip_creation_ids(element: ET.Element) -> ET.Element:
    """Drop the ``a16:creationId`` extension PowerPoint stamps on every shape.

    Two shapes carrying the same creation id inside one paste is not something
    PowerPoint was seen to object to, but a member that is pasted back while
    its original is still on the slide is exactly that case, and the id is
    worth nothing to us. Removing it costs nothing and removes the doubt.
    """
    for cnv in element.iter(f"{_A}cNvPr"):
        for ext_lst in cnv.findall(f"{_A}extLst"):
            cnv.remove(ext_lst)
    return element


def only_child(canvas_children: List[ET.Element], what: str) -> ET.Element:
    """The one shape a canvas copied off a slide holds, or a clear complaint."""
    if len(canvas_children) != 1:
        raise PackageError(
            f"expected the copied {what} to arrive as one shape on the canvas, "
            f"found {len(canvas_children)}"
        )
    return canvas_children[0]


def first_of(canvas_children: List[ET.Element], tag: str) -> Optional[ET.Element]:
    for child in canvas_children:
        if child.tag == f"{_A}{tag}":
            return child
    return None


def chart_frame_xml(frame_id: int, name: str, x: int, y: int, cx: int, cy: int, rel_id: str = "rId1") -> str:
    """One ``a:graphicFrame`` holding a chart, in the order that pastes.

    ``nvGraphicFramePr``, then ``graphic``, then ``xfrm``. That is the order
    PowerPoint writes, and the other order (``xfrm`` before ``graphic``) is
    dropped by ``paste object`` with no error (design section 0, probe 6 H2).
    The golden fixture test pins it.
    """
    return (
        f'<a:graphicFrame><a:nvGraphicFramePr><a:cNvPr id="{frame_id}" name="{_escape(name)}"/>'
        '<a:cNvGraphicFramePr/></a:nvGraphicFramePr>'
        f'<a:graphic><a:graphicData uri="{NS_C}">'
        f'<c:chart xmlns:c="{NS_C}" xmlns:r="{NS_R}" r:id="{rel_id}"/>'
        '</a:graphicData></a:graphic>'
        f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '</a:graphicFrame>'
    )
