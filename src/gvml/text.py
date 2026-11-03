"""The text of a group's member, read and rewritten in the clipboard package.

Nothing reaches a group's members over Apple Events (#220): ``shapes`` on a
group counts 0 and ``shapes[1]`` answers -1728. The package ``copy shape``
writes for the group carries every member, and a member's text is ordinary
DrawingML under ``a:txSp/a:txBody``: paragraphs (``a:p``) of runs (``a:r``),
each run with its character formatting (``a:rPr``) and its text (``a:t``),
and ``a:br`` for a line break. So a member is found by name here, its text
or its runs are rewritten, and the group is pasted back over the original
(``ppt_mac.gvml_paste.replace_shape``).

Text is read and written the way the text tools see it: paragraphs joined by
``\\n`` and a line break inside a paragraph as ``\\v``.

A member is named the way Windows names one (``ppt_com.shape_lookup``): by its
own name when that is the only member of that name in the group, or by its
path from the group down, ``"Group 4/TextBox 3"``, which is matched against
the joined names, so a name that itself holds a ``/`` still matches.
"""

import copy
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Tuple

from gvml.package import NS_A, PackageError
from gvml.shapes import _NON_VISUAL, name_of

_A = f"{{{NS_A}}}"
SEPARATOR = "/"
PARAGRAPH = "\n"
LINE_BREAK = "\v"

# The theme colour names the tools take, as a run's a:schemeClr spells them.
SCHEME_COLORS = {
    "dark1": "dk1", "light1": "lt1", "dark2": "dk2", "light2": "lt2",
    "accent1": "accent1", "accent2": "accent2", "accent3": "accent3",
    "accent4": "accent4", "accent5": "accent5", "accent6": "accent6",
    "hyperlink": "hlink", "followed_hyperlink": "folHlink",
}

# The order a:rPr's children take in the schema (CT_TextCharacterProperties).
# PowerPoint drops what is out of order rather than saying so, so a child is
# put where the schema has it.
_RPR_ORDER = [
    "ln", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill",
    "effectLst", "effectDag", "highlight", "uLnTx", "uLn", "uFillTx", "uFill",
    "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst",
]
_FILLS = {"noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"}


class MemberNotFound(PackageError):
    """No member of that name or path in the group."""


class AmbiguousMember(PackageError):
    """The bare name belongs to more than one member. A path says which."""


def members(group: ET.Element, group_name: str) -> List[Tuple[str, ET.Element]]:
    """Every member of the group at any depth, with its path from the group."""
    found = []

    def walk(element, path):
        for child in element:
            if child.tag not in _NON_VISUAL:
                continue
            child_path = f"{path}{SEPARATOR}{name_of(child)}"
            found.append((child_path, child))
            if child.tag == f"{_A}grpSp":
                walk(child, child_path)

    walk(group, group_name)
    return found


def find_member(group: ET.Element, group_name: str, identifier: str) -> Tuple[str, ET.Element]:
    """The member ``identifier`` names, and its path. Raises when there is none,
    or when a bare name belongs to more than one member."""
    everything = members(group, group_name)
    exact = [(path, el) for path, el in everything if path == identifier]
    if exact:
        return exact[0]
    by_name = [(path, el) for path, el in everything if name_of(el) == identifier]
    if len(by_name) == 1:
        return by_name[0]
    if by_name:
        raise AmbiguousMember(
            f"'{identifier}' names {len(by_name)} members of '{group_name}': "
            + ", ".join(path for path, _ in by_name) + ". Pass the path of one."
        )
    raise MemberNotFound(f"'{identifier}' is not a member of '{group_name}'")


def text_body(member: ET.Element) -> Optional[ET.Element]:
    """The member's ``a:txBody``, or None for a member that takes no text."""
    return member.find(f"{_A}txSp/{_A}txBody")


def _paragraphs(body: ET.Element) -> List[ET.Element]:
    return body.findall(f"{_A}p")


def _paragraph_text(paragraph: ET.Element) -> str:
    parts = []
    for child in paragraph:
        if child.tag in (f"{_A}r", f"{_A}fld"):
            t = child.find(f"{_A}t")
            parts.append(t.text or "" if t is not None else "")
        elif child.tag == f"{_A}br":
            parts.append(LINE_BREAK)
    return "".join(parts)


def read_text(member: ET.Element) -> str:
    """The member's text as the text tools return it."""
    body = text_body(member)
    if body is None:
        raise PackageError(f"'{name_of(member)}' takes no text")
    return PARAGRAPH.join(_paragraph_text(p) for p in _paragraphs(body))


def _split_paragraphs(text: str) -> List[str]:
    return text.replace("\r\n", PARAGRAPH).replace("\r", PARAGRAPH).split(PARAGRAPH)


def _run(rpr: Optional[ET.Element], text: str) -> ET.Element:
    run = ET.Element(f"{_A}r")
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    t = ET.SubElement(run, f"{_A}t")
    t.text = text
    return run


def write_text(member: ET.Element, text: str) -> None:
    """Replace the member's whole text.

    Each new paragraph takes the paragraph settings (``a:pPr``) and the
    character formatting of the first run of the old paragraph at the same
    place, or of the last old one when there are more new paragraphs than
    old, as PowerPoint does when text is typed over a selection.
    """
    body = text_body(member)
    if body is None:
        raise PackageError(f"'{name_of(member)}' takes no text")
    old = _paragraphs(body)
    insert_at = list(body).index(old[0]) if old else len(body)
    for p in old:
        body.remove(p)

    for i, para_text in enumerate(_split_paragraphs(text)):
        template = old[min(i, len(old) - 1)] if old else None
        p = ET.Element(f"{_A}p")
        ppr = rpr = end = None
        if template is not None:
            ppr = template.find(f"{_A}pPr")
            first = template.find(f"{_A}r")
            rpr = first.find(f"{_A}rPr") if first is not None else None
            end = template.find(f"{_A}endParaRPr")
            if rpr is None and end is not None:
                rpr = ET.Element(f"{_A}rPr", end.attrib)
                rpr.extend(copy.deepcopy(list(end)))
        if ppr is not None:
            p.append(copy.deepcopy(ppr))
        for j, line in enumerate(para_text.split(LINE_BREAK)):
            if j:
                br = ET.SubElement(p, f"{_A}br")
                if rpr is not None:
                    br.append(copy.deepcopy(rpr))
            if line:
                p.append(_run(rpr, line))
        if end is not None:
            p.append(copy.deepcopy(end))
        body.insert(insert_at + i, p)


def replace_span(member: ET.Element, start: int, length: int, new_text: str) -> None:
    """Replace ``length`` characters from 1-based ``start`` with ``new_text``.

    The span has to lie inside one paragraph and must not hold a line break,
    and the new text must not start a paragraph either; anything else is
    refused, with the whole text as the way to do it. A span inside one run
    keeps that run's formatting; one across runs takes the first run's.
    """
    body = text_body(member)
    if body is None:
        raise PackageError(f"'{name_of(member)}' takes no text")
    full = read_text(member)
    end = start - 1 + length
    if PARAGRAPH in full[start - 1:end] or LINE_BREAK in full[start - 1:end] \
            or PARAGRAPH in new_text or "\r" in new_text or LINE_BREAK in new_text:
        raise ValueError(
            "Inside a group, a part of the text can be replaced only within one "
            "line of one paragraph. Pass the whole text instead."
        )

    # Find the paragraph, then the runs, by counting characters.
    offset = 0
    for paragraph in _paragraphs(body):
        para_len = len(_paragraph_text(paragraph))
        if offset + para_len >= end and start - 1 >= offset:
            _replace_in_paragraph(paragraph, start - 1 - offset, length, new_text)
            return
        offset += para_len + 1
    raise ValueError(f"Characters {start} to {end} are not in the text ({len(full)} characters).")


def _replace_in_paragraph(paragraph: ET.Element, at: int, length: int, new_text: str) -> None:
    pieces = []  # (start, end, element) for every run and break, in order
    pos = 0
    for child in paragraph:
        if child.tag in (f"{_A}r", f"{_A}fld"):
            t = child.find(f"{_A}t")
            n = len(t.text or "") if t is not None else 0
            pieces.append((pos, pos + n, child))
            pos += n
        elif child.tag == f"{_A}br":
            pieces.append((pos, pos + 1, child))
            pos += 1

    runs = [(s, e, el) for s, e, el in pieces if el.tag == f"{_A}r"]
    if not runs:
        # An empty paragraph: a run to hold the text, formatted as its end.
        end_rpr = paragraph.find(f"{_A}endParaRPr")
        rpr = ET.Element(f"{_A}rPr", end_rpr.attrib) if end_rpr is not None else None
        run = _run(rpr, new_text)
        index = list(paragraph).index(end_rpr) if end_rpr is not None else len(paragraph)
        paragraph.insert(index, run)
        return

    stop = at + length
    # The run the edit starts in: the one holding `at`, or for an insertion at
    # a boundary, the run before it (typing goes on in the run you are at).
    first = next((r for r in runs if r[0] <= at < r[1]), None)
    if first is None or (length == 0 and first[0] == at and first is not runs[0]):
        first = next((r for r in reversed(runs) if r[0] < at <= r[1]), runs[0])

    s0, _, el0 = first
    t0 = el0.find(f"{_A}t")
    text0 = t0.text or ""
    keep_after = text0[min(stop, first[1]) - s0:] if stop <= first[1] else ""
    t0.text = text0[:at - s0] + new_text + keep_after
    for s, e, el in runs:
        if el is el0 or e <= at or s >= stop:
            continue
        t = el.find(f"{_A}t")
        rest = (t.text or "")[max(stop, s) - s:]
        if rest:
            t.text = rest
        else:
            paragraph.remove(el)


def _rpr_of(run: ET.Element) -> ET.Element:
    rpr = run.find(f"{_A}rPr")
    if rpr is None:
        rpr = ET.Element(f"{_A}rPr")
        run.insert(0, rpr)
    return rpr


def _set_child(rpr: ET.Element, element: ET.Element) -> None:
    """Put ``element`` into ``rpr`` where the schema has it, replacing one of
    the same kind (any fill, for a fill)."""
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


def _color_element(color: str) -> ET.Element:
    fill = ET.Element(f"{_A}solidFill")
    scheme = SCHEME_COLORS.get(color)
    if scheme is not None:
        ET.SubElement(fill, f"{_A}schemeClr", {"val": scheme})
    else:
        ET.SubElement(fill, f"{_A}srgbClr", {"val": color.lstrip("#").upper()})
    return fill


def format_runs(member: ET.Element, *, font_name=None, font_name_fareast=None,
                font_size=None, bold=None, italic=None, underline=None,
                color=None) -> int:
    """Set character formatting on every run of the member, and on each
    paragraph's end, so text typed there later matches. ``color`` is a hex
    value or a theme colour name as ``utils.color.theme_color_name`` spells
    it. Answers how many runs were formatted."""
    body = text_body(member)
    if body is None:
        raise PackageError(f"'{name_of(member)}' takes no text")
    targets = [_rpr_of(run) for run in body.iter(f"{_A}r")]
    formatted = len(targets)
    targets += [br.find(f"{_A}rPr") for br in body.iter(f"{_A}br") if br.find(f"{_A}rPr") is not None]
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
            _set_child(rpr, _color_element(color))
        if font_name is not None:
            _set_child(rpr, ET.Element(f"{_A}latin", {"typeface": font_name}))
        east = font_name_fareast if font_name_fareast is not None else font_name
        if east is not None:
            _set_child(rpr, ET.Element(f"{_A}ea", {"typeface": east}))
    return formatted


def run_formats(member: ET.Element) -> List[Dict[str, Optional[str]]]:
    """Each run's formatting as read back: size, bold, italic, underline,
    colour (hex or scheme name) and typeface. For checking a paste."""
    body = text_body(member)
    if body is None:
        return []
    out = []
    for run in body.iter(f"{_A}r"):
        rpr = run.find(f"{_A}rPr")
        rpr = rpr if rpr is not None else ET.Element(f"{_A}rPr")
        fill = rpr.find(f"{_A}solidFill")
        color = None
        if fill is not None and len(fill):
            color = fill[0].get("val")
        latin = rpr.find(f"{_A}latin")
        out.append({
            "sz": rpr.get("sz"), "b": rpr.get("b"), "i": rpr.get("i"), "u": rpr.get("u"),
            "color": color, "latin": latin.get("typeface") if latin is not None else None,
        })
    return out
