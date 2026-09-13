"""The zip around a GVML canvas: content types, relationships, parts.

A package PowerPoint writes has five parts, and three of them are all it needs
back: ``[Content_Types].xml``, ``_rels/.rels`` pointing at the drawing, and the
drawing itself at ``clipboard/drawings/drawing1.xml``. A chart adds
``clipboard/charts/chart1.xml`` and a relationship from the drawing to it, a
picture adds ``clipboard/media/image1.png`` the same way, and the theme part
PowerPoint includes is optional in both directions.

``validate`` is the check the design (docs/gvml-design.md section 8) puts in
front of every paste. PowerPoint will not report a broken package, so this is
the only place a mistake in our own assembly can be caught before it turns
into a silent no-op on the slide.
"""

import io
import posixpath
import re
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional

GVML_UTI = "com.microsoft.Art--GVML-ClipFormat"

# Where PowerPoint puts things, kept so a package we write is laid out like one
# it wrote. The names are not required by anything; the relationships are.
DRAWING_PART = "clipboard/drawings/drawing1.xml"
CHART_PART = "clipboard/charts/chart1.xml"
CONTENT_TYPES_PART = "[Content_Types].xml"
ROOT_RELS_PART = "_rels/.rels"

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_LC = "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas"
NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_PR = "http://schemas.openxmlformats.org/package/2006/relationships"

REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
REL_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
REL_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

CT_RELS = "application/vnd.openxmlformats-package.relationships+xml"
CT_XML = "application/xml"
CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"
CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
CT_THEME = "application/vnd.openxmlformats-officedocument.theme+xml"

XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

# Serialising through etree renames every prefix to ns0 unless the prefixes
# are registered first, and PowerPoint's acceptance of the registered form was
# measured (design section 11, settled: it pastes). Registering is global and
# harmless, so it happens at import.
for _prefix, _uri in (("a", NS_A), ("lc", NS_LC), ("c", NS_C), ("r", NS_R)):
    ET.register_namespace(_prefix, _uri)
ET.register_namespace("a16", "http://schemas.microsoft.com/office/drawing/2014/main")


class PackageError(ValueError):
    """A package that would not survive a paste, found before the paste."""


_XMLNS = re.compile(rb'xmlns:([A-Za-z_][\w.-]*)="([^"]+)"')


def register_prefixes(xml_bytes: bytes) -> None:
    """Register every ``xmlns:prefix`` a document declares, before parsing it.

    The four prefixes registered at import cover what this package writes.
    A ``chart1.xml`` PowerPoint wrote declares more (``mc``, ``c14``, ``c16r2``
    and whatever else the version added), and ``mc:Choice Requires="c14"``
    names a prefix by its spelling, so etree renaming it to ``ns3`` on the way
    out would leave the choice pointing at nothing. Reading the declarations
    off the bytes and registering each one keeps the spelling PowerPoint
    used. Registering is global and repeating it is harmless.
    """
    for prefix, uri in _XMLNS.findall(xml_bytes):
        try:
            ET.register_namespace(prefix.decode("ascii"), uri.decode("ascii"))
        except (ValueError, UnicodeDecodeError):
            # A prefix etree reserves (ns0 style) or one it cannot spell; the
            # document keeps working, the prefix is just renamed on output.
            continue


@dataclass(frozen=True)
class Relationship:
    """One ``<Relationship>``, with its target already resolved to a part name."""

    id: str
    type: str
    target: str


def _rels_part_for(part: str) -> str:
    """``clipboard/drawings/drawing1.xml`` -> ``clipboard/drawings/_rels/drawing1.xml.rels``."""
    if part == "":
        return ROOT_RELS_PART
    head, tail = posixpath.split(part)
    return posixpath.join(head, "_rels", tail + ".rels")


def _resolve_target(source_part: str, target: str) -> str:
    """Resolve a relationship target against the part that holds the rels."""
    if target.startswith("/"):
        return target.lstrip("/")
    base = posixpath.dirname(source_part) if source_part else ""
    return posixpath.normpath(posixpath.join(base, target)) if base else posixpath.normpath(target)


def _relative_target(source_part: str, target_part: str) -> str:
    base = posixpath.dirname(source_part) if source_part else ""
    if not base:
        return target_part
    return posixpath.relpath(target_part, base)


def _parse_xml(data: bytes, what: str) -> ET.Element:
    try:
        return ET.fromstring(data)
    except ET.ParseError as exc:
        raise PackageError(f"{what} is not well formed XML: {exc}") from exc


class Package:
    """A GVML package held in memory as a dict of part name to bytes.

    Part names carry no leading slash. ``[Content_Types].xml`` is a part like
    any other here, and is parsed rather than trusted when a content type is
    asked for.
    """

    def __init__(self, parts: Dict[str, bytes]):
        self.parts: Dict[str, bytes] = dict(parts)

    # -- reading -----------------------------------------------------------

    @classmethod
    def from_bytes(cls, raw: bytes) -> "Package":
        """Unpack a zip. Raises ``PackageError`` when the bytes are not one."""
        try:
            with zipfile.ZipFile(io.BytesIO(raw)) as zf:
                return cls({name: zf.read(name) for name in zf.namelist()
                            if not name.endswith("/")})
        except zipfile.BadZipFile as exc:
            raise PackageError("the clipboard data is not a zip archive") from exc

    def to_bytes(self) -> bytes:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            # Content types first, as every OPC writer does. Nothing here
            # depends on it, but a package that reads like PowerPoint's is
            # easier to diff against one.
            names = sorted(self.parts, key=lambda n: (n != CONTENT_TYPES_PART, n))
            for name in names:
                zf.writestr(name, self.parts[name])
        return buf.getvalue()

    def content_type(self, part: str) -> Optional[str]:
        """The content type ``[Content_Types].xml`` gives a part, or None."""
        ct = self.parts.get(CONTENT_TYPES_PART)
        if ct is None:
            return None
        root = _parse_xml(ct, CONTENT_TYPES_PART)
        for override in root.findall(f"{{{NS_CT}}}Override"):
            if override.get("PartName", "").lstrip("/") == part:
                return override.get("ContentType")
        ext = part.rsplit(".", 1)[-1].lower() if "." in part else ""
        for default in root.findall(f"{{{NS_CT}}}Default"):
            if default.get("Extension", "").lower() == ext:
                return default.get("ContentType")
        return None

    def relationships(self, part: str = "") -> List[Relationship]:
        """The relationships of a part (the package root when ``part`` is empty)."""
        rels_name = _rels_part_for(part)
        data = self.parts.get(rels_name)
        if data is None:
            return []
        root = _parse_xml(data, rels_name)
        found = []
        for rel in root.findall(f"{{{NS_PR}}}Relationship"):
            found.append(Relationship(
                id=rel.get("Id", ""),
                type=rel.get("Type", ""),
                target=_resolve_target(part, rel.get("Target", "")),
            ))
        return found

    def drawing_part(self) -> str:
        """The drawing the root relationship points at. Raises when there is none."""
        for rel in self.relationships(""):
            if rel.type == REL_DRAWING:
                return rel.target
        raise PackageError("the package root has no drawing relationship")

    def drawing(self) -> bytes:
        part = self.drawing_part()
        try:
            return self.parts[part]
        except KeyError:
            raise PackageError(f"the drawing relationship points at {part}, which is not in the package") from None

    def chart_part(self) -> Optional[str]:
        """The chart part the drawing relates to, if any."""
        drawing = self.drawing_part()
        for rel in self.relationships(drawing):
            if rel.type == REL_CHART:
                return rel.target
        return None

    def chart(self) -> Optional[bytes]:
        part = self.chart_part()
        return None if part is None else self.parts.get(part)

    def remove_relationship(self, source_part: str, rel_id: str) -> Optional[str]:
        """Drop one relationship of a part, and the part it pointed at.

        Returns the name of the part removed, or None when the relationship
        was not there. The target is removed only when nothing else in the
        package relates to it. Used to let go of a chart's embedded workbook
        once the caches it was written from no longer match it.
        """
        rels = self.relationships(source_part)
        dropped = [r for r in rels if r.id == rel_id]
        if not dropped:
            return None
        target = dropped[0].target
        self.parts[_rels_part_for(source_part)] = _rels_xml(
            source_part, [r for r in rels if r.id != rel_id]
        )
        still_used = any(
            r.target == target
            for part in [""] + [n for n in self.parts if not n.endswith(".rels")]
            for r in self.relationships(part)
        )
        if not still_used:
            self.parts.pop(target, None)
            self.parts.pop(_rels_part_for(target), None)
        return target

    # -- checking ----------------------------------------------------------

    def validate(self) -> None:
        """Raise ``PackageError`` for anything PowerPoint would silently drop.

        Checks the things the design lists (section 7): every part has a
        content type, every relationship target exists, the drawing is reached
        from the root relationship, it parses, and its canvas holds at least
        one shape. Element order inside the canvas is not checked here; that
        is what the golden fixtures in the tests are for.
        """
        if CONTENT_TYPES_PART not in self.parts:
            raise PackageError(f"{CONTENT_TYPES_PART} is missing")
        _parse_xml(self.parts[CONTENT_TYPES_PART], CONTENT_TYPES_PART)
        for name in self.parts:
            if name == CONTENT_TYPES_PART:
                continue
            if self.content_type(name) is None:
                raise PackageError(f"{name} has no content type")
        for source in [""] + [n for n in self.parts if not n.endswith(".rels")]:
            for rel in self.relationships(source):
                if rel.target not in self.parts:
                    raise PackageError(
                        f"relationship {rel.id} of {source or 'the package root'} "
                        f"points at {rel.target}, which is not in the package"
                    )
        drawing_part = self.drawing_part()
        if self.content_type(drawing_part) != CT_DRAWING:
            raise PackageError(f"{drawing_part} does not carry the drawing content type")
        # Imported here rather than at the top so package.py stays the bottom
        # of the dependency order inside gvml.
        from gvml import canvas as _canvas

        root = _parse_xml(self.drawing(), drawing_part)
        locked = _canvas.canvas_of(root)
        if not _canvas.children(locked):
            raise PackageError("the canvas holds no shape at all")
        chart_part = self.chart_part()
        if chart_part is not None:
            _parse_xml(self.parts[chart_part], chart_part)
            if self.content_type(chart_part) != CT_CHART:
                raise PackageError(f"{chart_part} does not carry the chart content type")


# ---------------------------------------------------------------------------
# Writing
# ---------------------------------------------------------------------------

def _content_types_xml(defaults: Dict[str, str], overrides: Dict[str, str]) -> bytes:
    body = "".join(
        f'<Default Extension="{ext}" ContentType="{ct}"/>'
        for ext, ct in sorted(defaults.items())
    ) + "".join(
        f'<Override PartName="/{part}" ContentType="{ct}"/>'
        for part, ct in overrides.items()
    )
    return (XML_DECL + f'<Types xmlns="{NS_CT}">' + body + "</Types>").encode()


def _rels_xml(source_part: str, rels: Iterable[Relationship]) -> bytes:
    body = "".join(
        f'<Relationship Id="{rel.id}" Type="{rel.type}" '
        f'Target="{_relative_target(source_part, rel.target)}"/>'
        for rel in rels
    )
    return (XML_DECL + f'<Relationships xmlns="{NS_PR}">' + body + "</Relationships>").encode()


def build(
    drawing_xml: str,
    parts: Optional[Dict[str, bytes]] = None,
    drawing_rels: Optional[List[Relationship]] = None,
    overrides: Optional[Dict[str, str]] = None,
    defaults: Optional[Dict[str, str]] = None,
    nested_rels: Optional[Dict[str, List[Relationship]]] = None,
) -> bytes:
    """Assemble a package around one drawing and return the zip bytes.

    Args:
        drawing_xml: The whole ``drawing1.xml``, declaration included.
        parts: Extra parts by name, a chart or an image for instance.
        drawing_rels: Relationships from the drawing to those parts.
        overrides: Content type overrides for the extra parts, by part name.
        defaults: Content type defaults by extension, for media.
        nested_rels: Relationships of extra parts (a chart's own rels).
    """
    all_parts: Dict[str, bytes] = dict(parts or {})
    all_parts[DRAWING_PART] = drawing_xml.encode("utf-8")
    all_defaults = {"rels": CT_RELS, "xml": CT_XML}
    all_defaults.update(defaults or {})
    all_overrides = {DRAWING_PART: CT_DRAWING}
    all_overrides.update(overrides or {})
    all_parts[CONTENT_TYPES_PART] = _content_types_xml(all_defaults, all_overrides)
    all_parts[ROOT_RELS_PART] = _rels_xml("", [Relationship("rId1", REL_DRAWING, DRAWING_PART)])
    if drawing_rels:
        all_parts[_rels_part_for(DRAWING_PART)] = _rels_xml(DRAWING_PART, drawing_rels)
    for part, rels in (nested_rels or {}).items():
        all_parts[_rels_part_for(part)] = _rels_xml(part, rels)
    return Package(all_parts).to_bytes()


def read(raw: bytes) -> Package:
    """Unpack and check a package PowerPoint (or we) wrote."""
    package = Package.from_bytes(raw)
    package.validate()
    return package


def validate(raw: bytes) -> None:
    """Raise ``PackageError`` unless ``raw`` is a package worth pasting."""
    Package.from_bytes(raw).validate()


# ---------------------------------------------------------------------------
# Grafting: carrying the parts of several packages into one
# ---------------------------------------------------------------------------

@dataclass
class Graft:
    """Parts and relationships gathered from other packages for a new one.

    Grouping copies each member off the slide as its own package, and a member
    that is a picture or a chart brings a part with it that the group's package
    has to carry as well. Part names collide across packages (every picture is
    ``image1.png``), so each carried part is renamed as it comes in and the
    relationship ids inside the member's XML are rewritten to match.
    """

    parts: Dict[str, bytes]
    drawing_rels: List[Relationship]
    overrides: Dict[str, str]
    defaults: Dict[str, str]
    nested_rels: Dict[str, List[Relationship]]
    _counter: int = 0

    @classmethod
    def empty(cls) -> "Graft":
        return cls({}, [], {}, {}, {})

    def _fresh_name(self, part: str) -> str:
        head, tail = posixpath.split(part)
        stem, dot, ext = tail.rpartition(".")
        if not dot:
            stem, ext = tail, ""
        stem = re.sub(r"\d+$", "", stem) or "part"
        self._counter += 1
        candidate = posixpath.join(head, f"{stem}{self._counter}{dot}{ext}")
        while candidate in self.parts:
            self._counter += 1
            candidate = posixpath.join(head, f"{stem}{self._counter}{dot}{ext}")
        return candidate

    def _carry(self, source: Package, part: str) -> str:
        """Copy a part and everything it relates to, returning its new name."""
        new_name = self._fresh_name(part)
        self.parts[new_name] = source.parts[part]
        ct = source.content_type(part)
        ext = part.rsplit(".", 1)[-1].lower() if "." in part else ""
        if ct is not None:
            if ext in ("xml", "rels") or not ext:
                self.overrides[new_name] = ct
            else:
                self.defaults.setdefault(ext, ct)
        nested = []
        for rel in source.relationships(part):
            if rel.type == REL_THEME or rel.target not in source.parts:
                continue
            nested.append(Relationship(rel.id, rel.type, self._carry(source, rel.target)))
        if nested:
            self.nested_rels[new_name] = nested
        return new_name

    def take(self, source: Package, element: ET.Element) -> ET.Element:
        """Carry the parts ``element`` refers to and rewrite its ids to match.

        ``element`` is one child of ``source``'s canvas. Every relationship of
        the source drawing except the theme is carried, whether or not this
        element uses it; a canvas holds one shape when it comes off a slide,
        so there is nothing else it could belong to.
        """
        drawing = source.drawing_part()
        renames: Dict[str, str] = {}
        for rel in source.relationships(drawing):
            if rel.type == REL_THEME or rel.target not in source.parts:
                continue
            new_id = f"rId{len(self.drawing_rels) + 1}"
            self.drawing_rels.append(Relationship(new_id, rel.type, self._carry(source, rel.target)))
            renames[rel.id] = new_id
        if renames:
            for node in element.iter():
                for key, value in list(node.attrib.items()):
                    if key.startswith(f"{{{NS_R}}}") and value in renames:
                        node.set(key, renames[value])
        return element


__all__ = [
    "GVML_UTI", "DRAWING_PART", "CHART_PART", "NS_A", "NS_LC", "NS_C", "NS_R",
    "REL_CHART", "REL_IMAGE", "REL_THEME", "CT_CHART", "XML_DECL",
    "Package", "PackageError", "Relationship", "Graft", "build", "read", "validate",
    "register_prefixes",
]
