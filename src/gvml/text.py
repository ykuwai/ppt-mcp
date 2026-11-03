"""A canvas shape's text, read and rewritten in its ``a:txBody``.

For the members of a group, which Apple Events cannot reach (#220): the group
is copied, a member's text or its runs are rewritten here, and the group is
pasted back.

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


def _paragraph_text(paragraph: ET.Element) -> str:
    parts = []
    for piece in paragraph:
        if piece.tag in (f"{_A}r", f"{_A}fld"):
            t = piece.find(f"{_A}t")
            parts.append(t.text or "" if t is not None else "")
        elif piece.tag == f"{_A}br":
            parts.append("\v")
    return "".join(parts)


def replace_span(shape: ET.Element, start: int, length: int, new_text: str) -> None:
    """Replace ``length`` characters from 1-based ``start`` with ``new_text``,
    as ``text_of`` counts them.

    The span has to lie inside one line of one paragraph, and the new text
    must not break one; anything else raises ValueError, naming the whole
    text as the way to do it. A span inside one run keeps that run's
    formatting; one across runs takes the first run's, and typing at the
    boundary between two runs goes on in the run before, as in PowerPoint.
    """
    body = _body(shape)
    if body is None:
        raise PackageError("the shape has no text body")
    full = text_of(shape)
    end = start - 1 + length
    if (start < 1 or end > len(full) or any(c in full[start - 1:end] for c in "\n\v")
            or any(c in new_text for c in "\n\r\v")):
        raise ValueError(
            "Inside a group, part of the text can be replaced only within one "
            "line of one paragraph. Pass the whole text instead."
        )
    offset = 0
    for paragraph in body.findall(f"{_A}p"):
        size = len(_paragraph_text(paragraph))
        if offset <= start - 1 and end <= offset + size:
            _replace_in_paragraph(paragraph, start - 1 - offset, length, new_text)
            return
        offset += size + 1


def _replace_in_paragraph(paragraph: ET.Element, at: int, length: int, new_text: str) -> None:
    runs, pos = [], 0
    for piece in paragraph:
        if piece.tag in (f"{_A}r", f"{_A}fld"):
            t = piece.find(f"{_A}t")
            n = len(t.text or "") if t is not None else 0
            if piece.tag == f"{_A}r":
                runs.append((pos, pos + n, piece))
            pos += n
        elif piece.tag == f"{_A}br":
            pos += 1
    if not runs:
        end_rpr = paragraph.find(f"{_A}endParaRPr")
        run = ET.Element(f"{_A}r")
        if end_rpr is not None:
            run.append(ET.Element(f"{_A}rPr", end_rpr.attrib))
        ET.SubElement(run, f"{_A}t").text = new_text
        index = list(paragraph).index(end_rpr) if end_rpr is not None else len(paragraph)
        paragraph.insert(index, run)
        return

    stop = at + length
    first = next((r for r in runs if r[0] <= at < r[1]), None)
    if first is None or (length == 0 and first[0] == at and first is not runs[0]):
        first = next((r for r in reversed(runs) if r[0] < at <= r[1]), runs[0])
    s0, e0, el0 = first
    t0 = el0.find(f"{_A}t")
    text0 = t0.text or ""
    after = text0[stop - s0:] if stop <= e0 else ""
    t0.text = text0[:at - s0] + new_text + after
    for s, e, el in runs:
        if el is el0 or e <= at or s >= stop:
            continue
        t = el.find(f"{_A}t")
        rest = (t.text or "")[max(stop, s) - s:]
        if rest:
            t.text = rest
        else:
            paragraph.remove(el)


# The theme colour names the tools take, as a run's a:schemeClr spells them.
SCHEME_COLORS = {
    "dark1": "dk1", "light1": "lt1", "dark2": "dk2", "light2": "lt2",
    "accent1": "accent1", "accent2": "accent2", "accent3": "accent3",
    "accent4": "accent4", "accent5": "accent5", "accent6": "accent6",
    "hyperlink": "hlink", "followed_hyperlink": "folHlink",
}

# CT_TextCharacterProperties' children in schema order. PowerPoint drops a
# child that is out of order rather than saying so.
_RPR_ORDER = [
    "ln", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill",
    "effectLst", "effectDag", "highlight", "uLnTx", "uLn", "uFillTx", "uFill",
    "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst",
]
_FILLS = {"noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"}


def _put(rpr: ET.Element, element: ET.Element) -> None:
    """Put ``element`` into ``rpr`` where the schema has it, replacing one of
    its kind (any fill, for a fill)."""
    local = element.tag[len(_A):]
    same = _FILLS if local in _FILLS else {local}
    for child in list(rpr):
        if child.tag[len(_A):] in same:
            rpr.remove(child)
    rank = _RPR_ORDER.index(local)
    index = len(rpr)
    for i, child in enumerate(rpr):
        name = child.tag[len(_A):]
        if name in _RPR_ORDER and _RPR_ORDER.index(name) > rank:
            index = i
            break
    rpr.insert(index, element)


def format_runs(shape: ET.Element, *, font_name=None, font_name_fareast=None,
                font_size=None, bold=None, italic=None, underline=None,
                color=None) -> int:
    """Set character formatting on every run, and on each paragraph's end so
    text typed there later matches. ``color`` is a hex value or a theme name
    as ``utils.color.theme_color_name`` spells it. Answers the runs formatted."""
    body = _body(shape)
    if body is None:
        raise PackageError("the shape has no text body")
    runs = list(body.iter(f"{_A}r"))
    targets = []
    for run in runs:
        rpr = run.find(f"{_A}rPr")
        if rpr is None:
            rpr = ET.Element(f"{_A}rPr")
            run.insert(0, rpr)
        targets.append(rpr)
    targets += [br.find(f"{_A}rPr") for br in body.iter(f"{_A}br")
                if br.find(f"{_A}rPr") is not None]
    targets += list(body.iter(f"{_A}endParaRPr"))

    for rpr in targets:
        if font_size is not None:
            rpr.set("sz", str(int(round(font_size * 100))))
        if bold is not None:
            rpr.set("b", "1" if bold else "0")
        if italic is not None:
            rpr.set("i", "1" if italic else "0")
        if underline is not None:
            rpr.set("u", "sng" if underline else "none")
        if color is not None:
            fill = ET.Element(f"{_A}solidFill")
            scheme = SCHEME_COLORS.get(color)
            if scheme is not None:
                ET.SubElement(fill, f"{_A}schemeClr", {"val": scheme})
            else:
                ET.SubElement(fill, f"{_A}srgbClr", {"val": color.lstrip("#").upper()})
            _put(rpr, fill)
        if font_name is not None:
            _put(rpr, ET.Element(f"{_A}latin", {"typeface": font_name}))
        east = font_name_fareast if font_name_fareast is not None else font_name
        if east is not None:
            _put(rpr, ET.Element(f"{_A}ea", {"typeface": east}))
    return len(runs)


def run_formats(shape: ET.Element) -> List[dict]:
    """Each run's size, bold, italic, underline, colour (hex or scheme name)
    and typeface, as read back. For checking a paste."""
    body = _body(shape)
    if body is None:
        return []
    out = []
    for run in body.iter(f"{_A}r"):
        rpr = run.find(f"{_A}rPr")
        rpr = rpr if rpr is not None else ET.Element(f"{_A}rPr")
        fill = rpr.find(f"{_A}solidFill")
        latin = rpr.find(f"{_A}latin")
        out.append({
            "sz": rpr.get("sz"), "b": rpr.get("b"), "i": rpr.get("i"), "u": rpr.get("u"),
            "color": fill[0].get("val") if fill is not None and len(fill) else None,
            "latin": latin.get("typeface") if latin is not None else None,
        })
    return out
