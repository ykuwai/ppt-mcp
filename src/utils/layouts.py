"""Slide layouts as a saved .pptx names them, for PowerPoint for Mac.

PowerPoint for Mac's dictionary declares no ``name`` on ``custom layout`` and
no ``designs`` on a presentation, and a layout reached through the slide
master (``slide master``'s ``custom layouts``) cannot be given to a slide:
setting one, or making a slide with one, answers "Application could not
handle this command" (16.114). So ``ppt_add_slide`` cannot find a template's
layout by name through Apple Events at all.

The deck's file can. Its masters list their layouts in order, each layout part carries the
name Windows reports as ``CustomLayout.Name``, and each saved slide names
the layout it uses. This module reads that,
and writes a copy of the deck holding a single new slide on a chosen layout,
which the Mac side opens and pastes from; a slide pasted between two decks
with the same master lands on the destination's layout of the same name.

Standard library only, like ``utils.package``.
"""

import posixpath
import re
import zipfile
from dataclasses import dataclass, field
from typing import Dict, List, Optional
from xml.etree import ElementTree

_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
_NS = {"p": _P, "a": _A, "r": _R, "rel": _REL}
_RID = "{%s}id" % _R

_REL_SLIDE = _R + "/slide"
_REL_LAYOUT = _R + "/slideLayout"
_CT_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"

# Placeholders a new slide does not get from its layout, as PowerPoint adds
# them: the date, footer and slide number come from the master when shown.
_NOT_COPIED = {"dt", "ftr", "sldNum"}


@dataclass
class Layout:
    name: str
    part: str          # e.g. "ppt/slideLayouts/slideLayout6.xml"
    type: str          # the layout's ST_SlideLayoutType, "obj" when not given


@dataclass
class Master:
    index: int         # 1-based, the order PowerPoint lists its designs in
    name: str
    part: str
    layouts: List[Layout] = field(default_factory=list)


def _rels(archive, part) -> Dict[str, tuple]:
    """A part's relationships: id to (type, target path in the package)."""
    folder, name = posixpath.split(part)
    path = posixpath.join(folder, "_rels", name + ".rels")
    if path not in archive.namelist():
        return {}
    root = ElementTree.fromstring(archive.read(path))
    return {
        rel.get("Id"): (rel.get("Type"),
                        posixpath.normpath(posixpath.join(folder, rel.get("Target"))))
        for rel in root.findall("rel:Relationship", _NS)
        if rel.get("TargetMode") != "External"
    }


def _cSld_name(root) -> str:
    c = root.find("p:cSld", _NS)
    return (c.get("name") if c is not None else None) or ""


def masters(path: str) -> List[Master]:
    """Every master in the deck, its layouts in the order PowerPoint lists them."""
    with zipfile.ZipFile(path) as archive:
        presentation = ElementTree.fromstring(archive.read("ppt/presentation.xml"))
        prels = _rels(archive, "ppt/presentation.xml")
        out = []
        for i, mid in enumerate(presentation.findall("p:sldMasterIdLst/p:sldMasterId", _NS), 1):
            part = prels.get(mid.get(_RID), (None, None))[1]
            if not part or part not in archive.namelist():
                continue
            root = ElementTree.fromstring(archive.read(part))
            mrels = _rels(archive, part)
            name = _cSld_name(root)
            if not name:
                theme = next((t for kind, t in mrels.values() if kind.endswith("/theme")), None)
                if theme and theme in archive.namelist():
                    name = ElementTree.fromstring(archive.read(theme)).get("name") or ""
            master = Master(i, name, part)
            for lid in root.findall("p:sldLayoutIdLst/p:sldLayoutId", _NS):
                lpart = mrels.get(lid.get(_RID), (None, None))[1]
                if not lpart or lpart not in archive.namelist():
                    continue
                lroot = ElementTree.fromstring(archive.read(lpart))
                master.layouts.append(Layout(_cSld_name(lroot), lpart, lroot.get("type") or "obj"))
            out.append(master)
        return out


def find(path: str, name: str, design_index: Optional[int] = None) -> List[tuple]:
    """``(master, layout)`` for each master with a layout named exactly ``name``,
    as Windows' ``_find_layout_matches`` compares (case and spaces count)."""
    found = []
    for master in masters(path):
        if design_index is not None and master.index != design_index:
            continue
        layout = next((lay for lay in master.layouts if lay.name == name), None)
        if layout is not None:
            found.append((master, layout))
    return found


def slide_layouts(path: str) -> Dict[int, str]:
    """Each saved slide's id (``sldId``'s ``id``, PowerPoint's ``slide ID``)
    to the layout part it uses."""
    with zipfile.ZipFile(path) as archive:
        presentation = ElementTree.fromstring(archive.read("ppt/presentation.xml"))
        prels = _rels(archive, "ppt/presentation.xml")
        out = {}
        for sid in presentation.findall("p:sldIdLst/p:sldId", _NS):
            part = prels.get(sid.get(_RID), (None, None))[1]
            if not part or part not in archive.namelist():
                continue
            layout = next((t for kind, t in _rels(archive, part).values()
                           if kind == _REL_LAYOUT), None)
            if layout:
                out[int(sid.get("id"))] = layout
        return out


def _new_slide_xml(layout_root) -> bytes:
    """A slide carrying the layout's placeholders, empty, as a new slide gets them."""
    shapes = []
    next_id = 2
    tree = layout_root.find("p:cSld/p:spTree", _NS)
    for sp in (tree.findall("p:sp", _NS) if tree is not None else []):
        ph = sp.find("p:nvSpPr/p:nvPr/p:ph", _NS)
        if ph is None or ph.get("type") in _NOT_COPIED:
            continue
        cnv = sp.find("p:nvSpPr/p:cNvPr", _NS)
        name = cnv.get("name") if cnv is not None else f"Placeholder {next_id}"
        attrs = "".join(f' {k}="{_escape(v)}"' for k, v in ph.attrib.items())
        shapes.append(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{next_id}" name="{_escape(name)}"/>'
            '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            f'<p:nvPr><p:ph{attrs}/></p:nvPr></p:nvSpPr><p:spPr/>'
            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p>'
            '</p:txBody></p:sp>'
        )
        next_id += 1
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<p:sld xmlns:a="{_A}" xmlns:r="{_R}" xmlns:p="{_P}"><p:cSld><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr/>' + "".join(shapes) +
        '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
    ).encode("utf-8")


def _escape(text: str) -> str:
    return (text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


def one_slide_copy(path: str, layout_part: str, out_path: str) -> None:
    """Write a copy of the deck at ``out_path`` whose only slide is a new one
    on ``layout_part``.

    The deck's other slides are left out of the slide list (their parts stay
    in the package, unreferenced), so opening the copy shows just the one.
    ``presentation.xml`` is edited as text: re-serialising it can drop the
    namespace declarations ``mc:Ignorable`` names, and PowerPoint then offers
    to repair the file.
    """
    with zipfile.ZipFile(path) as src:
        names = src.namelist()
        if layout_part not in names:
            raise ValueError(f"{layout_part} is not in {path}")
        layout_root = ElementTree.fromstring(src.read(layout_part))

        n = 1
        while f"ppt/slides/slide{n}.xml" in names:
            n += 1
        slide_part = f"ppt/slides/slide{n}.xml"

        prels_path = "ppt/_rels/presentation.xml.rels"
        prels = src.read(prels_path).decode("utf-8")
        prels = re.sub(r'<Relationship\b[^>]*Type="%s"[^>]*/>' % re.escape(_REL_SLIDE), "", prels)
        used = set(re.findall(r'Id="([^"]+)"', prels))
        rid = "rId1"
        k = 1
        while rid in used:
            k += 1
            rid = f"rId{k}"
        prels = prels.replace(
            "</Relationships>",
            f'<Relationship Id="{rid}" Type="{_REL_SLIDE}" Target="slides/slide{n}.xml"/>'
            "</Relationships>")

        presentation = src.read("ppt/presentation.xml").decode("utf-8")
        sld_list = f'<p:sldIdLst><p:sldId id="256" r:id="{rid}"/></p:sldIdLst>'
        if re.search(r"<p:sldIdLst\b", presentation):
            presentation = re.sub(r"<p:sldIdLst\b.*?</p:sldIdLst>|<p:sldIdLst\s*/>",
                                  sld_list, presentation, count=1, flags=re.S)
        else:
            # The list comes after the masters and the notes and handout
            # master lists, and before the slide size.
            anchor = None
            for tag in ("p:handoutMasterIdLst", "p:notesMasterIdLst", "p:sldMasterIdLst"):
                m = re.search(r"</%s>|<%s\b[^>]*/>" % (tag, tag), presentation)
                if m:
                    anchor = m.end()
                    break
            if anchor is None:
                raise ValueError("the deck's presentation.xml lists no slide master")
            presentation = presentation[:anchor] + sld_list + presentation[anchor:]

        types = src.read("[Content_Types].xml").decode("utf-8").replace(
            "</Types>",
            f'<Override PartName="/{slide_part}" ContentType="{_CT_SLIDE}"/></Types>')

        slide_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<Relationships xmlns="{_REL}"><Relationship Id="rId1" Type="{_REL_LAYOUT}" '
            f'Target="../slideLayouts/{posixpath.basename(layout_part)}"/></Relationships>'
        )

        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as dst:
            for item in src.infolist():
                if item.filename == prels_path:
                    dst.writestr(item, prels)
                elif item.filename == "ppt/presentation.xml":
                    dst.writestr(item, presentation)
                elif item.filename == "[Content_Types].xml":
                    dst.writestr(item, types)
                else:
                    dst.writestr(item, src.read(item.filename))
            dst.writestr(slide_part, _new_slide_xml(layout_root))
            dst.writestr(f"ppt/slides/_rels/slide{n}.xml.rels", slide_rels)
