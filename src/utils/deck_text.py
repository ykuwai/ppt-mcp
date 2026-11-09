"""The text of a saved .pptx, for ppt_get_all_text.

Asking PowerPoint for a deck's text shape by shape costs thousands of calls:
on a 50-slide deck, 10 s over COM, 4 s over Apple Events on PowerPoint for Mac
16.97 and 112 s on 16.114. The deck written out once with SaveCopyAs and read
here takes well under a second on all three, and leaves the clipboard alone.

A saved slide holds only what differs from its layout, so a placeholder's box,
and a paragraph's bullet and a run's emphasis, are looked up the way PowerPoint
does: the shape, then the layout's placeholder, the master's, and the master's
text styles. Needs nothing past the standard library.
"""

import zipfile
from xml.etree import ElementTree

from utils.package import _NS, _rels

_EMU = 12700
_A = "{%s}" % _NS["a"]
_P = "{%s}" % _NS["p"]
_MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
_MATH_TEXT = "{http://schemas.openxmlformats.org/officeDocument/2006/math}t"

_TITLES = ("title", "ctrTitle")
_SKIPPED = ("dt", "ftr", "sldNum", "hdr")
_TRUE = ("1", "true")
_FALSE = ("0", "false")


def _children(parent):
    """An element's children, with mc:AlternateContent replaced by its first
    choice: what PowerPoint itself shows. It can wrap a shape, a run, or, for
    an equation PowerPoint saved, a whole paragraph."""
    for child in parent:
        if child.tag == _MC + "AlternateContent":
            chosen = child.find(_MC + "Choice")
            if chosen is None:
                chosen = child.find(_MC + "Fallback")
            if chosen is not None:
                yield from _children(chosen)
        else:
            yield child


def _placeholder(element):
    nv = element.find("p:nvSpPr/p:nvPr/p:ph", _NS)
    if nv is None:
        nv = element.find("p:nvGraphicFramePr/p:nvPr/p:ph", _NS)
    return nv


class _Part:
    """A layout or master: its placeholders, by idx and by type."""

    def __init__(self, root, path, archive):
        self.root = root
        self.rels = _rels(archive, path)
        self.by_idx, self.by_type = {}, {}
        tree = root.find("p:cSld/p:spTree", _NS)
        for element in _children(tree) if tree is not None else ():
            ph = _placeholder(element)
            if ph is None:
                continue
            if ph.get("idx") is not None:
                self.by_idx.setdefault(ph.get("idx"), element)
            self.by_type.setdefault(ph.get("type", "body"), element)

    def related(self, folder):
        return next((p for p in self.rels.values() if f"/{folder}/" in p), None)


def _master_type(kind):
    if kind in _TITLES:
        return "title"
    return kind if kind in _SKIPPED else "body"


class _Package:
    """The parts a slide inherits from, each parsed once."""

    def __init__(self, archive):
        self.archive = archive
        self.parts = {}
        presentation = ElementTree.fromstring(archive.read("ppt/presentation.xml"))
        self.default_style = presentation.find("p:defaultTextStyle", _NS)
        self.rels = _rels(archive, "ppt/presentation.xml")
        self.order = [self.rels.get(s.get("{%s}id" % _NS["r"]))
                      for s in presentation.findall("p:sldIdLst/p:sldId", _NS)]

    def part(self, path):
        if path is None:
            return None
        if path not in self.parts:
            self.parts[path] = _Part(ElementTree.fromstring(self.archive.read(path)), path, self.archive)
        return self.parts[path]

    def inherited(self, slide_rels, ph):
        """(the layout's placeholder, the master's, the master's text style)
        for a placeholder on a slide."""
        layout = self.part(next((p for p in slide_rels.values() if "/slideLayouts/" in p), None))
        master = self.part(layout.related("slideMasters")) if layout is not None else None
        kind = ph.get("type", "body")
        in_layout = None
        if layout is not None:
            in_layout = layout.by_idx.get(ph.get("idx")) if ph.get("idx") is not None else None
            if in_layout is None:
                in_layout = layout.by_type.get(kind)
        in_master, style = None, None
        if master is not None:
            in_master = master.by_type.get(_master_type(kind))
            styles = master.root.find("p:txStyles", _NS)
            if styles is not None:
                name = "p:titleStyle" if kind in _TITLES else "p:bodyStyle"
                style = styles.find(name, _NS) if kind not in _SKIPPED else styles.find("p:otherStyle", _NS)
        return in_layout, in_master, style


def _box(element, frame=False):
    """(left, top, width, height) in EMU, or None when the shape sets none."""
    xfrm = element.find("p:xfrm", _NS) if frame else element.find("p:spPr/a:xfrm", _NS)
    if xfrm is None:
        return None
    off, ext = xfrm.find("a:off", _NS), xfrm.find("a:ext", _NS)
    if off is None or ext is None:
        return None
    return int(off.get("x", 0)), int(off.get("y", 0)), int(ext.get("cx", 0)), int(ext.get("cy", 0))


def _inner_map(group, outer):
    """``outer``, a function from a group's frame to the slide's, carried one
    group further in: its members are placed in the group's child space."""
    xfrm = group.find("p:grpSpPr/a:xfrm", _NS)
    if xfrm is None:
        return outer
    off, ext = xfrm.find("a:off", _NS), xfrm.find("a:ext", _NS)
    ch_off, ch_ext = xfrm.find("a:chOff", _NS), xfrm.find("a:chExt", _NS)
    if off is None or ext is None or ch_off is None or ch_ext is None:
        return outer
    x, y, cx, cy = int(off.get("x", 0)), int(off.get("y", 0)), int(ext.get("cx", 0)), int(ext.get("cy", 0))
    cx0, cy0 = int(ch_off.get("x", 0)), int(ch_off.get("y", 0))
    ccx, ccy = int(ch_ext.get("cx", 0)), int(ch_ext.get("cy", 0))
    sx, sy = cx / ccx if ccx else 1.0, cy / ccy if ccy else 1.0

    def mapped(box):
        left, top, width, height = box
        return outer((x + (left - cx0) * sx, y + (top - cy0) * sy, width * sx, height * sy))

    return mapped


def _level_props(style, level):
    return style.find(f"a:lvl{level}pPr", _NS) if style is not None else None


def _resolved(chain, level, read):
    """The first answer ``read`` gives along the styles in ``chain``."""
    for style in chain:
        props = _level_props(style, level)
        if props is None:
            continue
        answer = read(props)
        if answer is not None:
            return answer
    return None


def _bullet_of(props):
    """"bullet", "number", "none", or None when ``props`` does not say."""
    if props.find("a:buNone", _NS) is not None:
        return "none"
    if props.find("a:buAutoNum", _NS) is not None:
        return "number"
    if props.find("a:buChar", _NS) is not None or props.find("a:buBlip", _NS) is not None:
        return "bullet"
    return None


def _flag(props, name):
    value = props.get(name) if props is not None else None
    if value in _TRUE:
        return True
    if value in _FALSE:
        return False
    return None


def _run(element, text, chain, level, paragraph_default):
    props = element.find("a:rPr", _NS)
    emphasis = {}
    for name, key in (("b", "bold"), ("i", "italic")):
        own = _flag(props, name)
        if own is None:
            own = _flag(paragraph_default, name)
        if own is None:
            own = _resolved(chain, level, lambda p, n=name: _flag(p.find("a:defRPr", _NS), n))
        emphasis[key] = bool(own)
    return {"text": text, **emphasis}


def _body_paragraphs(body):
    return [child for child in _children(body) if child.tag == _A + "p"]


def _is_math(element):
    return element.tag.endswith(("}m", "}oMathPara", "}oMath"))


def _math_text(element):
    return "".join(t.text or "" for t in element.iter(_MATH_TEXT))


def _paragraphs(body, chain):
    out = []
    for paragraph in _body_paragraphs(body):
        props = paragraph.find("a:pPr", _NS)
        level = int(props.get("lvl", "0")) + 1 if props is not None else 1
        bullet = _bullet_of(props) if props is not None else None
        if bullet is None:
            bullet = _resolved(chain, level, _bullet_of)
        default = props.find("a:defRPr", _NS) if props is not None else None
        runs = []
        for child in _children(paragraph):
            if child.tag in (_A + "r", _A + "fld"):
                t = child.find("a:t", _NS)
                text = (t.text or "") if t is not None else ""
                if text:
                    runs.append(_run(child, text, chain, level, default))
            elif child.tag == _A + "br":
                runs.append(_run(child, "\n", chain, level, default))
            elif _is_math(child):
                text = _math_text(child)
                if text:
                    runs.append({"text": text, "bold": False, "italic": False})
        out.append({"runs": runs, "level": level, "bullet": None if bullet == "none" else bullet})
    return out


def _cell_text(cell):
    body = cell.find("a:txBody", _NS)
    if body is None:
        return ""
    lines = []
    for paragraph in _body_paragraphs(body):
        line = []
        for child in _children(paragraph):
            if child.tag in (_A + "r", _A + "fld"):
                t = child.find("a:t", _NS)
                line.append((t.text or "") if t is not None else "")
            elif child.tag == _A + "br":
                line.append(" ")
            elif _is_math(child):
                line.append(_math_text(child))
        lines.append("".join(line))
    return " ".join(lines)


def _table(frame):
    """A table's rows of text. A merged cell answers its text at every place
    it covers, as ``Table.Cell`` does."""
    table = frame.find("a:graphic/a:graphicData/a:tbl", _NS)
    if table is None:
        return None
    rows = []
    for tr in table.findall("a:tr", _NS):
        row = []
        for column, cell in enumerate(tr.findall("a:tc", _NS)):
            if cell.get("hMerge") in _TRUE and row:
                row.append(row[-1])
            elif cell.get("vMerge") in _TRUE and rows and column < len(rows[-1]):
                row.append(rows[-1][column])
            else:
                row.append(_cell_text(cell))
        rows.append(row)
    return rows


def _info(box, kind, **rest):
    left, top, width, height = box if box is not None else (0, 0, 0, 0)
    return {"top": top / _EMU, "left": left / _EMU, "width": width / _EMU, "height": height / _EMU,
            "is_title": kind in _TITLES, "is_subtitle": kind == "subTitle", **rest}


def _walk(parent, package, slide_rels, to_slide, found):
    for element in _children(parent):
        name = element.tag
        if name == _P + "grpSp":
            _walk(element, package, slide_rels, _inner_map(element, to_slide), found)
        elif name == _P + "graphicFrame":
            rows = _table(element)
            if rows is None:
                continue
            ph = _placeholder(element)
            kind = ph.get("type", "body") if ph is not None else None
            if kind in _SKIPPED:
                continue
            box = _box(element, frame=True)
            found.append(_info(to_slide(box) if box else None, kind, has_table=True, rows=rows))
        elif name == _P + "sp":
            _shape(element, package, slide_rels, to_slide, found)


def _shape(element, package, slide_rels, to_slide, found):
    ph = _placeholder(element)
    kind = ph.get("type", "body") if ph is not None else None
    if kind in _SKIPPED:
        return
    body = element.find("p:txBody", _NS)
    box = _box(element)
    chain = [body.find("a:lstStyle", _NS)] if body is not None else []
    if ph is not None:
        in_layout, in_master, style = package.inherited(slide_rels, ph)
        for inherited in (in_layout, in_master):
            if inherited is None:
                continue
            if box is None:
                box = _box(inherited)
            own = inherited.find("p:txBody/a:lstStyle", _NS)
            if own is not None:
                chain.append(own)
        chain.append(style)
    else:
        chain.append(package.default_style)
    if body is None:
        return
    paragraphs = _paragraphs(body, [c for c in chain if c is not None])
    if not any(run["text"] for p in paragraphs for run in p["runs"]):
        return
    found.append(_info(to_slide(box) if box else None, kind, has_table=False, paragraphs=paragraphs))


def _language_of(props):
    default = props.find("a:defRPr", _NS)
    return default.get("lang") if default is not None else None


def inherited_language(shape, level, package, slide_rels):
    """The language of a shape's text whose run and paragraph name none.

    The shape's list style, then for a placeholder the layout's, the master's
    and the master's text styles. Measured on Windows, 2026-11-10: the deck's
    default reaches a text box but not a placeholder, whose text has no
    language at all when none of those gives one, so that is None.
    """
    body = shape.find("p:txBody", _NS)
    chain = [body.find("a:lstStyle", _NS)] if body is not None else []
    ph = _placeholder(shape)
    if ph is not None:
        in_layout, in_master, style = package.inherited(slide_rels, ph)
        chain += [part.find("p:txBody/a:lstStyle", _NS)
                  for part in (in_layout, in_master) if part is not None]
        chain.append(style)
        return _resolved(chain, level, _language_of)
    deck = package.default_style
    fallback = deck.find("a:defPPr", _NS) if deck is not None else None
    return (_resolved([*chain, deck], level, _language_of)
            or (_language_of(fallback) if fallback is not None else None))


def read_slides(path, slide_indices=None):
    """(how many slides the deck has, {slide index: {"hidden", "shapes"}}).

    ``slide_indices`` None is every slide; one the deck does not have is left
    out. Each shape is a dict with its box in points, ``is_title``,
    ``is_subtitle``, and either ``paragraphs`` or, with ``has_table``, ``rows``.
    """
    slides = {}
    with zipfile.ZipFile(path) as archive:
        package = _Package(archive)
        total = len(package.order)
        wanted = range(1, total + 1) if slide_indices is None else slide_indices
        for index in wanted:
            if not 1 <= index <= total or index in slides:
                continue
            part = package.order[index - 1]
            root = ElementTree.fromstring(archive.read(part))
            found = []
            tree = root.find("p:cSld/p:spTree", _NS)
            if tree is not None:
                _walk(tree, package, _rels(archive, part), lambda box: box, found)
            slides[index] = {"hidden": root.get("show") in _FALSE, "shapes": found}
    return total, slides


def format_runs(raw) -> str:
    """Runs as Markdown, consecutive ones with the same emphasis merged.

    A run's spaces at either end go outside its markers: emphasis that
    starts or ends with a space is not emphasis in Markdown, so " sharply"
    in italics written as "* sharply*" showed its asterisks, and
    "**fell*** sharply*" lost the italics. A run of spaces alone is left
    plain.
    """
    merged = []
    for r in raw:
        if merged and merged[-1]["bold"] == r["bold"] and merged[-1]["italic"] == r["italic"]:
            merged[-1]["text"] += r["text"]
        else:
            merged.append(dict(r))

    parts = []
    for m in merged:
        text = m["text"]
        marks = "***" if m["bold"] and m["italic"] else "**" if m["bold"] else "*" if m["italic"] else ""
        core = text.strip()
        if not marks or not core:
            parts.append(text)
            continue
        start = text.index(core)
        parts.append(f"{text[:start]}{marks}{core}{marks}{text[start + len(core):]}")
    return "".join(parts)


def _paragraph_text(paragraph) -> str:
    return "".join(run["text"] for run in paragraph["runs"])


def _heading(paragraphs, marker) -> str:
    text = " ".join(_paragraph_text(p).replace("\n", " ") for p in paragraphs).strip()
    return f"{marker} {text}" if text else ""


def _table_markdown(rows) -> str:
    lines = []
    for number, row in enumerate(rows):
        cells = [cell.replace("|", "\\|").strip() for cell in row]
        lines.append("| " + " | ".join(cells) + " |")
        if number == 0:
            lines.append("| " + " | ".join(["---"] * len(cells)) + " |")
    return "\n".join(lines)


def shape_markdown(info, subheading_level="##") -> str:
    """One shape from ``read_slides`` as ppt_get_all_text writes it."""
    if info["has_table"]:
        return _table_markdown(info["rows"])
    paragraphs = info["paragraphs"]
    if info["is_title"]:
        return _heading(paragraphs, "#")
    if not info["is_subtitle"]:
        words = [run for p in paragraphs for run in p["runs"] if run["text"].strip()]
        if words and all(run["bold"] for run in words):
            return _heading(paragraphs, subheading_level)
    # An empty last paragraph is not one PowerPoint counts.
    if len(paragraphs) > 1 and not _paragraph_text(paragraphs[-1]):
        paragraphs = paragraphs[:-1]
    lines = []
    for paragraph in paragraphs:
        text = format_runs(paragraph["runs"]).strip()
        if not text:
            lines.append("")
            continue
        indent = "  " * max(0, paragraph["level"] - 1)
        prefix = {"number": f"{indent}1. ", "bullet": f"{indent}- "}.get(paragraph["bullet"], "")
        lines.append(f"{prefix}{text}")
    return "\n".join(lines)
