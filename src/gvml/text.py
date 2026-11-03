"""A canvas shape's text, read and rewritten in its ``a:txBody``.

For the members of a group, which Apple Events cannot reach (#220): the group
is copied, a member's text is rewritten here, and the group is pasted back.

The text a caller writes follows the tools' convention: ``\\n`` between
paragraphs, ``\\v`` for a line break inside one. A paragraph keeps the
paragraph properties of the first paragraph that was there, and its runs the
character properties of the first run, so the member keeps its look.
"""

import copy
import xml.etree.ElementTree as ET
from typing import List, Optional

from gvml.package import NS_A, PackageError
from gvml.shapes import name_of

_A = f"{{{NS_A}}}"


def find_member(group: ET.Element, path: List[str]) -> Optional[ET.Element]:
    """The member of ``group`` named by ``path``, or None.

    ``path`` is the names from the group down, as Windows writes
    "Group/Child": one name finds a member at any depth, several walk down
    through nested groups.
    """
    if len(path) > 1:
        for child in group:
            if child.tag == f"{_A}grpSp" and name_of(child) == path[0]:
                return find_member(child, path[1:])
        return None
    for child in group.iter():
        if child is group or child.tag not in (f"{_A}sp", f"{_A}grpSp", f"{_A}pic",
                                               f"{_A}graphicFrame", f"{_A}cxnSp"):
            continue
        if name_of(child) == path[0]:
            return child
    return None


def _body(shape: ET.Element) -> Optional[ET.Element]:
    return shape.find(f"{_A}txSp/{_A}txBody")


def has_text_body(shape: ET.Element) -> bool:
    return _body(shape) is not None


def text_of(shape: ET.Element) -> str:
    """The shape's text, ``\\n`` between paragraphs and ``\\v`` for breaks."""
    body = _body(shape)
    if body is None:
        raise PackageError("the shape has no text body")
    paragraphs = []
    for p in body.findall(f"{_A}p"):
        parts = []
        for piece in p:
            if piece.tag in (f"{_A}r", f"{_A}fld"):
                t = piece.find(f"{_A}t")
                parts.append(t.text or "" if t is not None else "")
            elif piece.tag == f"{_A}br":
                parts.append("\v")
        paragraphs.append("".join(parts))
    return "\n".join(paragraphs)


def set_text(shape: ET.Element, text: str) -> None:
    """Replace the shape's paragraphs with ``text``, keeping their look."""
    body = _body(shape)
    if body is None:
        raise PackageError("the shape has no text body")
    old = body.findall(f"{_A}p")
    first = old[0] if old else None
    ppr = first.find(f"{_A}pPr") if first is not None else None
    rpr = None
    end = None
    for p in old:
        run = p.find(f"{_A}r")
        if run is not None and run.find(f"{_A}rPr") is not None:
            rpr = run.find(f"{_A}rPr")
            break
    if first is not None:
        end = first.find(f"{_A}endParaRPr")
    for p in old:
        body.remove(p)

    def run_props(tag):
        source = rpr if rpr is not None else end
        if source is None:
            return None
        made = copy.deepcopy(source)
        made.tag = f"{_A}{tag}"
        return made

    for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        p = ET.SubElement(body, f"{_A}p")
        if ppr is not None:
            p.append(copy.deepcopy(ppr))
        pieces = line.split("\v")
        for i, piece in enumerate(pieces):
            if i:
                br = ET.SubElement(p, f"{_A}br")
                props = run_props("rPr")
                if props is not None:
                    br.append(props)
            if piece:
                r = ET.SubElement(p, f"{_A}r")
                props = run_props("rPr")
                if props is not None:
                    r.append(props)
                ET.SubElement(r, f"{_A}t").text = piece
        props = run_props("endParaRPr")
        if props is not None:
            p.append(props)
