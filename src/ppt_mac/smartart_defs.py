"""The SmartArt layouts, colour schemes and styles PowerPoint for Mac installs.

Windows lists them from ``Application.SmartArtLayouts``, ``SmartArtColors``
and ``SmartArtQuickStyles``. The Apple Event dictionary has none of the three,
but PowerPoint for Mac ships the definitions those collections are built from
as plain files inside the app::

    SmartArt.framework/Versions/A/Resources/lo/*.glo   layoutDef
    SmartArt.framework/Versions/A/Resources/cs/*.gcs   colorsDef
    SmartArt.framework/Versions/A/Resources/qs/*.gqs   styleDef

Each is the XML a SmartArt graphic's own ``layout``, ``colors`` and
``quickStyle`` parts hold, with a ``uniqueId``, an English name in a comment
after its empty ``<title val=""/>``, and its categories with a priority. They
are read here when asked for, from the installed app, and nothing is shipped.

**The order.** A colour scheme's or style's index is its position in the
Windows collection: sorting by the category order Windows uses, then by each
file's ``pri``, gives exactly Windows' order for all 38 colour schemes and 14
styles (PowerPoint for Mac 16.114 against Windows 16.0.20326). Layouts are
sorted the same way, by category then ``pri``, which puts Basic Block List
first as on Windows, but the two installs do not carry the same layouts (158
shared, 18 only on Windows, 4 only on the Mac), so a layout index here is a
position in this list and should be read from ``ppt_list_smartart_layouts``.

**The names.** 27 layout files carry no name comment; their name is the last
part of the id, split at its capitals ("BendingPictureCaption" becomes
"Bending Picture Caption"), which is the Windows name for nearly all of them.
The few that differ from what PowerPoint for Windows calls the same layout
are given the Windows name (``_WINDOWS_NAMES``), so every layout both
installs carry has one name on both, and no two layouts share one.

The names are already English, so a layout is looked up by them alone, and
``ppt_com.smartart.SMARTART_ENGLISH_ALIASES``, which gives a Windows install
in another language its English names, is not needed here.
"""

import glob
import os
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from functools import lru_cache

APP = "/Applications/Microsoft PowerPoint.app"
_RESOURCES = "Contents/Frameworks/SmartArt.framework/Versions/A/Resources"
NS_DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"

KINDS = {"layouts": ("lo", "glo"), "colors": ("cs", "gcs"), "styles": ("qs", "gqs")}

# The order Windows lists each kind's categories in. A category not named here
# goes after these, in name order.
_CATEGORY_ORDER = {
    "layouts": ["list", "process", "cycle", "hierarchy", "relationship", "matrix",
                "pyramid", "picture", "timeline", "meettheteam", "textcard"],
    "colors": ["mainscheme", "colorful", "accent1", "accent2", "accent3",
               "accent4", "accent5", "accent6"],
    "styles": ["simple", "3d"],
}

# Layouts both installs carry whose Mac file says, or whose id spells, a
# different name from PowerPoint for Windows (16.0.20326).
_WINDOWS_NAMES = {
    "urn:microsoft.com/office/officeart/2009/3/layout/BlockDescendingList": "Descending Block List",
    "urn:microsoft.com/office/officeart/2009/3/layout/SubStepProcess": "Sub-Step Process",
    "urn:microsoft.com/office/officeart/2008/layout/NameandTitleOrganizationalChart":
        "Name and Title Organization Chart",
    "urn:microsoft.com/office/officeart/2008/layout/PictureAccentList": "Titled Picture Accent List",
    "urn:microsoft.com/office/officeart/2024/layout/NumberedTitleCardList": "Numbered Card List",
}

_NAME_COMMENT = re.compile(r'<title\s+val=""\s*/>\s*<!--\s*"([^"]*)"\s*-->')
_DESC_COMMENT = re.compile(r'<desc\s+val=""\s*/>\s*<!--\s*"([^"]*)"\s*-->')


class DefinitionsMissing(RuntimeError):
    """The installed SmartArt definitions could not be found."""


@dataclass(frozen=True)
class Definition:
    index: int
    id: str
    name: str
    category: str
    description: str
    path: str

    @property
    def category_id(self) -> str:
        """The ``loCatId``, ``csCatId`` or ``qsCatId`` a data model records."""
        return self.category

    def xml(self) -> bytes:
        with open(self.path, "rb") as f:
            return f.read()


def resources_dir() -> str:
    """The folder holding ``lo``, ``cs`` and ``qs``, or DefinitionsMissing."""
    folder = os.path.join(os.environ.get("PPT_MCP_POWERPOINT_APP", APP), _RESOURCES)
    if not os.path.isdir(os.path.join(folder, "lo")):
        raise DefinitionsMissing(
            f"PowerPoint for Mac's SmartArt definitions were not found in {folder}. "
            "They are installed with Microsoft PowerPoint in /Applications; if "
            "PowerPoint lives elsewhere, set PPT_MCP_POWERPOINT_APP to its .app path."
        )
    return folder


def _name_from_id(uid: str) -> str:
    tail = uid.rsplit("/", 1)[-1]
    return re.sub(r"(?<=[a-z0-9])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])", " ", tail)


def _read(path: str):
    with open(path, "rb") as f:
        raw = f.read()
    root = ET.fromstring(raw)
    text = raw.decode("utf-8", "replace")
    uid = root.get("uniqueId") or ""
    named = _NAME_COMMENT.search(text)
    desc = _DESC_COMMENT.search(text)
    cats = [(c.get("type") or "", int(c.get("pri") or 0))
            for c in root.iter(f"{{{NS_DGM}}}cat")]
    category, pri = cats[0] if cats else ("", 0)
    return {
        "id": uid,
        "name": _WINDOWS_NAMES.get(uid) or (
            named.group(1) if named and named.group(1) else _name_from_id(uid)),
        "description": desc.group(1) if desc else "",
        "category": category,
        "pri": pri,
        "path": path,
    }


@lru_cache(maxsize=None)
def _load(folder: str, kind: str) -> tuple:
    sub, ext = KINDS[kind]
    rows = []
    for path in glob.glob(os.path.join(folder, sub, f"*.{ext}")):
        try:
            rows.append(_read(path))
        except (OSError, ET.ParseError):
            continue
    order = _CATEGORY_ORDER[kind]

    def key(row):
        cat = row["category"].lower()
        rank = order.index(cat) if cat in order else len(order)
        return (rank, cat, row["pri"], row["id"])

    rows.sort(key=key)
    return tuple(
        Definition(i, r["id"], r["name"], r["category"], r["description"], r["path"])
        for i, r in enumerate(rows, 1)
    )


def definitions(kind: str) -> tuple:
    """Every installed definition of a kind ('layouts', 'colors', 'styles'), in order."""
    defs = _load(resources_dir(), kind)
    if not defs:
        raise DefinitionsMissing(
            f"No SmartArt {kind} were found under {resources_dir()}."
        )
    return defs


def by_index(kind: str, index: int) -> Definition:
    defs = definitions(kind)
    if not 1 <= index <= len(defs):
        raise ValueError(
            f"SmartArt {kind[:-1]} index {index} is out of range: PowerPoint for "
            f"Mac has {len(defs)} (1-{len(defs)}). Use ppt_list_smartart_layouts "
            f"with list_type='{kind}' to see them."
        )
    return defs[index - 1]


def by_id(kind: str, uid: str):
    return next((d for d in definitions(kind) if d.id == uid), None)


def resolve_layout(layout_name: str) -> Definition:
    """A layout by name, the way Windows' ``_resolve_layout`` finds one: a
    case-insensitive "contains" match on the name, in list order."""
    wanted = layout_name.lower()
    for d in definitions("layouts"):
        if wanted in d.name.lower():
            return d
    raise ValueError(
        f"SmartArt layout '{layout_name}' not found. "
        "Use ppt_list_smartart_layouts to see available layouts with English names."
    )
