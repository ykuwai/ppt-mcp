"""The ``lockedCanvas`` wrapper, and the unit conversion everything else uses.

A GVML drawing is one ``a:graphic`` whose data is a ``lc:lockedCanvas``, which
is a group: a non-visual header, a transform, then the shapes. The transform
is the part worth knowing. ``off``/``ext`` are always ``0,0`` and the size of
the content, and ``chOff``/``chExt`` are where the content sits on the slide,
so a child's own ``a:off`` is written in slide EMU and is what the child's
position was when it was copied. Where a pasted shape lands is not governed by
any of this; PowerPoint drops it in the middle of the view and the caller
writes the position afterwards (design section 4).
"""

import xml.etree.ElementTree as ET
from typing import List, Tuple

from gvml.package import NS_A, NS_LC, NS_R, XML_DECL, PackageError

EMU_PER_PT = 12700

# The children a canvas can hold that stand for a shape on the slide. Anything
# else under the canvas is a property, not a shape.
SHAPE_TAGS = (
    f"{{{NS_A}}}sp",
    f"{{{NS_A}}}grpSp",
    f"{{{NS_A}}}pic",
    f"{{{NS_A}}}graphicFrame",
    f"{{{NS_A}}}cxnSp",
)


def emu(points: float) -> int:
    """Points to EMU, rounded to the integer OOXML requires."""
    return int(round(points * EMU_PER_PT))


def pt(emus) -> float:
    """EMU to points, rounded to two places the way every tool here reports."""
    return round(int(emus) / EMU_PER_PT, 2)


def wrap(body_xml: str, x: int, y: int, cx: int, cy: int) -> str:
    """Put shape XML on a canvas whose content occupies ``x, y, cx, cy`` EMU.

    The order of the elements is PowerPoint's own, taken from a package it
    wrote, and is not to be rearranged.
    """
    return (
        XML_DECL
        + f'<a:graphic xmlns:a="{NS_A}" xmlns:r="{NS_R}">'
        + f'<a:graphicData uri="{NS_LC}"><lc:lockedCanvas xmlns:lc="{NS_LC}">'
        + '<a:nvGrpSpPr><a:cNvPr id="0" name=""/><a:cNvGrpSpPr/></a:nvGrpSpPr>'
        + '<a:grpSpPr><a:xfrm>'
        + f'<a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/>'
        + f'<a:chOff x="{x}" y="{y}"/><a:chExt cx="{cx}" cy="{cy}"/>'
        + '</a:xfrm></a:grpSpPr>'
        + body_xml
        + "</lc:lockedCanvas></a:graphicData></a:graphic>"
    )


def parse(drawing_xml: bytes) -> ET.Element:
    """The ``lc:lockedCanvas`` element of a drawing part."""
    try:
        root = ET.fromstring(drawing_xml)
    except ET.ParseError as exc:
        raise PackageError(f"the drawing is not well formed XML: {exc}") from exc
    return canvas_of(root)


def canvas_of(root: ET.Element) -> ET.Element:
    """Find the canvas under a parsed ``a:graphic``, or say what is there instead."""
    if root.tag != f"{{{NS_A}}}graphic":
        raise PackageError(f"the drawing's root is {root.tag}, not a:graphic")
    locked = root.find(f"{{{NS_A}}}graphicData/{{{NS_LC}}}lockedCanvas")
    if locked is None:
        raise PackageError("the drawing holds no lockedCanvas")
    return locked


def children(canvas: ET.Element) -> List[ET.Element]:
    """The shape elements directly on a canvas, in z order."""
    return [child for child in canvas if child.tag in SHAPE_TAGS]


def bounds(canvas: ET.Element) -> Tuple[int, int, int, int]:
    """Where the canvas content sits on the slide, as ``x, y, cx, cy`` EMU."""
    xfrm = canvas.find(f"{{{NS_A}}}grpSpPr/{{{NS_A}}}xfrm")
    if xfrm is None:
        raise PackageError("the canvas has no transform")
    off = xfrm.find(f"{{{NS_A}}}chOff")
    ext = xfrm.find(f"{{{NS_A}}}chExt")
    if off is None or ext is None:
        raise PackageError("the canvas transform has no child offset or extent")
    return int(off.get("x", 0)), int(off.get("y", 0)), int(ext.get("cx", 0)), int(ext.get("cy", 0))


def serialize(element: ET.Element) -> str:
    """One element back to a string, with the prefixes PowerPoint uses."""
    return ET.tostring(element, encoding="unicode")


def serialize_drawing(root: ET.Element) -> str:
    """A whole drawing back to bytes-ready text, declaration included."""
    return XML_DECL + ET.tostring(root, encoding="unicode")
