"""Facts read out of a saved .pptx, for what PowerPoint will not say.

PowerPoint gives no way to ask a picture for its own proportions short of
resetting its scale, which changes the deck. The saved file holds them: each
picture's image is in the package, and its crop is in the slide. Reading
them needs nothing past the standard library.
"""

import posixpath
import struct
import unicodedata
import zipfile
from xml.etree import ElementTree

_NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
_EMBED = "{%s}embed" % _NS["r"]
_ID = "{%s}id" % _NS["r"]


def image_size(data: bytes):
    """An image's width and height in pixels, or None.

    PNG, JPEG, GIF, BMP and WebP. A JPEG its camera marked as turned is left
    out, since whether PowerPoint turns it as well is not known here.
    """
    if data[:8] == b"\x89PNG\r\n\x1a\n" and len(data) >= 24:
        return struct.unpack(">II", data[16:24])
    if data[:6] in (b"GIF87a", b"GIF89a") and len(data) >= 10:
        return struct.unpack("<HH", data[6:10])
    if data[:2] == b"BM" and len(data) >= 26:
        width, height = struct.unpack("<ii", data[18:26])
        return abs(width), abs(height)
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP" and len(data) >= 30:
        kind = data[12:16]
        if kind == b"VP8 ":
            width, height = struct.unpack("<HH", data[26:30])
            return width & 0x3FFF, height & 0x3FFF
        if kind == b"VP8L":
            bits = int.from_bytes(data[21:25], "little")
            return (bits & 0x3FFF) + 1, ((bits >> 14) & 0x3FFF) + 1
        if kind == b"VP8X":
            return (int.from_bytes(data[24:27], "little") + 1,
                    int.from_bytes(data[27:30], "little") + 1)
        return None
    if data[:2] == b"\xff\xd8":
        return _jpeg_size(data)
    return None


def _jpeg_size(data):
    i, size, turned = 2, None, False
    while i + 4 <= len(data):
        if data[i] != 0xFF:
            i += 1
            continue
        marker = data[i + 1]
        if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7:
            i += 2
            continue
        length = struct.unpack(">H", data[i + 2:i + 4])[0]
        segment = data[i + 4:i + 2 + length]
        if marker == 0xE1 and segment[:6] == b"Exif\x00\x00":
            turned = _exif_turned(segment[6:])
        elif 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC) and len(segment) >= 5:
            height, width = struct.unpack(">HH", segment[1:5])
            size = (width, height)
            break
        i += 2 + length
    return None if turned else size


def _exif_turned(tiff):
    """Whether an EXIF block says the picture is turned a quarter."""
    if len(tiff) < 8 or tiff[:2] not in (b"II", b"MM"):
        return False
    order = "<" if tiff[:2] == b"II" else ">"
    offset = struct.unpack(order + "I", tiff[4:8])[0]
    if offset + 2 > len(tiff):
        return False
    count = struct.unpack(order + "H", tiff[offset:offset + 2])[0]
    for n in range(count):
        entry = offset + 2 + 12 * n
        if entry + 12 > len(tiff):
            break
        tag, = struct.unpack(order + "H", tiff[entry:entry + 2])
        if tag == 0x0112:
            value, = struct.unpack(order + "H", tiff[entry + 8:entry + 10])
            return value >= 5
    return False


def _rels(archive, part):
    """A part's relationships, id to the target's path in the package."""
    folder, name = posixpath.split(part)
    path = posixpath.join(folder, "_rels", name + ".rels")
    if path not in archive.namelist():
        return {}
    root = ElementTree.fromstring(archive.read(path))
    return {
        rel.get("Id"): posixpath.normpath(posixpath.join(folder, rel.get("Target")))
        for rel in root.findall("rel:Relationship", _NS)
        if rel.get("TargetMode") != "External"
    }


def _slides(archive, slide_indices):
    """Each wanted slide's index, parsed XML and relationships."""
    presentation = ElementTree.fromstring(archive.read("ppt/presentation.xml"))
    targets = _rels(archive, "ppt/presentation.xml")
    order = [targets.get(s.get(_ID))
             for s in presentation.findall("p:sldIdLst/p:sldId", _NS)]
    for index in slide_indices:
        if not 1 <= index <= len(order) or not order[index - 1]:
            continue
        part = order[index - 1]
        yield index, ElementTree.fromstring(archive.read(part)), _rels(archive, part)


def _unique(found, name, value, seen):
    """Keep ``value`` under ``name`` unless the name was seen before."""
    if name in seen:
        found.pop(name, None)
        return
    seen.add(name)
    if value is not None:
        found[name] = value


def _picture_aspect(archive, pic, media):
    fill = pic.find("p:blipFill", _NS)
    if fill is None:
        return None
    blip = fill.find("a:blip", _NS)
    if blip is None or fill.find("a:tile", _NS) is not None:
        return None
    image = media.get(blip.get(_EMBED))
    if not image or image not in archive.namelist():
        return None
    size = image_size(archive.read(image))
    if not size or not size[0] or not size[1]:
        return None
    crop = fill.find("a:srcRect", _NS)
    cut = {side: 0.0 for side in "ltrb"}
    if crop is not None:
        for side in "ltrb":
            cut[side] = int(crop.get(side, "0")) / 100000
    width = size[0] * (1 - cut["l"] - cut["r"])
    height = size[1] * (1 - cut["t"] - cut["b"])
    return width / height if width > 0 and height > 0 else None


def _chart(archive, frame, parts):
    from gvml import PackageError
    from gvml.charts import read_chart

    chart = frame.find("a:graphic/a:graphicData/c:chart", {**_NS, "c": _CHART})
    if chart is None:
        return None
    part = parts.get(chart.get(_ID))
    if not part or part not in archive.namelist():
        return None
    try:
        return read_chart(archive.read(part))
    except (PackageError, ElementTree.ParseError):
        return None


_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"


def package_facts(path, slide_indices):
    """What the saved file says that PowerPoint will not.

    ``{"pictures": {slide: {name: aspect}}, "charts": {slide: {name: data}}}``.
    A picture's aspect is its image's width over its height, as cropped; a
    chart's data is what ``gvml.charts.read_chart`` reads. A picture whose
    image cannot be measured, one filled as a tile, and names used twice on
    a slide are left out, so an absent entry always means unknown.
    """
    pictures, charts = {}, {}
    with zipfile.ZipFile(path) as archive:
        for index, slide, parts in _slides(archive, slide_indices):
            aspects, data, seen = {}, {}, set()
            for element in slide.iter():
                if element.tag == "{%s}pic" % _NS["p"]:
                    props = element.find("p:nvPicPr/p:cNvPr", _NS)
                    if props is not None:
                        _unique(aspects, props.get("name"),
                                _picture_aspect(archive, element, parts), seen)
                elif element.tag == "{%s}graphicFrame" % _NS["p"]:
                    props = element.find("p:nvGraphicFramePr/p:cNvPr", _NS)
                    if props is not None:
                        _unique(data, props.get("name"), _chart(archive, element, parts), seen)
            pictures[index], charts[index] = aspects, data
    return {"pictures": pictures, "charts": charts}


def picture_aspects(path, slide_indices):
    """Each picture's own proportions: ``package_facts``'s pictures."""
    return package_facts(path, slide_indices)["pictures"]


# ---------------------------------------------------------------------------
# Layouts
# ---------------------------------------------------------------------------
_EMU = 12700
_SCHEME = ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4",
           "accent5", "accent6", "hlink", "folHlink")
_TITLES = ("title", "ctrTitle")

# The shape types the layout checks know, as PowerPoint numbers them.
_AUTOSHAPE, _GROUP, _LINE, _PICTURE = 1, 6, 9, 13


def _related(_, archive, part, folder):
    """The part a part's relationships point to in ``folder``, or None."""
    for target in _rels(archive, part).values():
        if f"/{folder}/" in target:
            return target
    return None


def _scheme(archive, master):
    """The master's theme colours by the names shapes use: tx1, bg1, accent1."""
    theme = _related(None, archive, master, "theme")
    colours = {}
    if theme and theme in archive.namelist():
        root = ElementTree.fromstring(archive.read(theme))
        scheme = root.find(".//a:clrScheme", _NS)
        for name in _SCHEME:
            entry = scheme.find(f"a:{name}", _NS) if scheme is not None else None
            if entry is None:
                continue
            rgb = entry.find("a:srgbClr", _NS)
            system = entry.find("a:sysClr", _NS)
            value = rgb.get("val") if rgb is not None else (
                system.get("lastClr") if system is not None else None)
            if value:
                colours[name] = "#" + value.upper()
    mapping = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2"}
    root = ElementTree.fromstring(archive.read(master))
    clr_map = root.find("p:clrMap", _NS)
    if clr_map is not None:
        mapping.update({key: clr_map.get(key) for key in mapping if clr_map.get(key)})
    for alias, name in mapping.items():
        if name in colours:
            colours[alias] = colours[name]
    return colours


def _theme_fonts(archive, master):
    """The theme's heading and body fonts, Latin and right to left, as
    ppt_com.layout_check._theme_fonts gives them."""
    fonts = {"latin": [], "script": []}
    theme = _related(None, archive, master, "theme")
    if not theme or theme not in archive.namelist():
        return fonts
    scheme = ElementTree.fromstring(archive.read(theme)).find(".//a:fontScheme", _NS)
    for part in ("a:majorFont", "a:minorFont"):
        found = scheme.find(part, _NS) if scheme is not None else None
        for kind, tag in (("latin", "a:latin"), ("script", "a:cs")):
            font = found.find(tag, _NS) if found is not None else None
            if font is not None and font.get("typeface"):
                fonts[kind].append(font.get("typeface"))
    return fonts


def _colour(element, scheme):
    """The colour a fill element names, or None when it is changed by tints
    or shades, which are not worked out here."""
    for kind in ("a:srgbClr", "a:schemeClr"):
        found = element.find(kind, _NS)
        if found is None:
            continue
        if len(found):
            return None
        if kind == "a:srgbClr":
            return "#" + found.get("val", "").upper()
        return scheme.get(found.get("val"))
    return None


def _fill(props, style, scheme):
    """The fill facts the layout checks read, from a shape's properties."""
    facts = {"visible": False, "solid": True, "color_hex": None, "transparency": 0.0}
    if props is None:
        return facts
    if props.find("a:noFill", _NS) is not None:
        return facts
    solid = props.find("a:solidFill", _NS)
    if solid is not None:
        facts.update(visible=True, color_hex=_colour(solid, scheme))
        return facts
    if props.find("a:gradFill", _NS) is not None or props.find("a:blipFill", _NS) is not None:
        facts.update(visible=True, solid=False)
        return facts
    ref = style.find("a:fillRef", _NS) if style is not None else None
    if ref is not None and ref.get("idx", "0") != "0":
        facts.update(visible=True, color_hex=_colour(ref, scheme))
    return facts


def _box(props):
    xfrm = props.find("a:xfrm", _NS) if props is not None else None
    if xfrm is None:
        return None
    off, ext = xfrm.find("a:off", _NS), xfrm.find("a:ext", _NS)
    if off is None or ext is None:
        return None
    return ([round(int(off.get(k)) / _EMU, 2) for k in ("x", "y")]
            + [round(int(ext.get(k)) / _EMU, 2) for k in ("cx", "cy")],
            round(int(xfrm.get("rot", "0")) / 60000, 2))


def _layout_shapes(root, scheme):
    """The graphics and the title box a layout or a master draws."""
    graphics, title = [], None
    tree = root.find("p:cSld/p:spTree", _NS)
    for element in list(tree) if tree is not None else []:
        tag = element.tag.rsplit("}", 1)[-1]
        if tag not in ("sp", "pic", "cxnSp", "grpSp"):
            continue
        nv = next((child for child in element if child.tag.endswith("Pr")
                   and child.tag.rsplit("}", 1)[-1].startswith("nv")), None)
        props_name = {"sp": "p:spPr", "pic": "p:spPr", "cxnSp": "p:spPr", "grpSp": "p:grpSpPr"}[tag]
        props = element.find(props_name, _NS)
        name = nv.find("p:cNvPr", _NS).get("name") if nv is not None else ""
        hidden = nv is not None and _hidden(nv.find("p:cNvPr", _NS))
        placeholder = nv.find("p:nvPr/p:ph", _NS) if nv is not None else None
        placed = _box(props)
        if placeholder is not None:
            if placeholder.get("type", "body") in _TITLES and placed:
                title = placed[0]
            continue
        if not placed:
            continue
        (left, top, width, height), rotation = placed
        geometry = props.find("a:prstGeom", _NS) if props is not None else None
        line = tag == "cxnSp" or (geometry is not None and geometry.get("prst") == "line")
        kind = {"pic": _PICTURE, "grpSp": _GROUP}.get(tag, _LINE if line else _AUTOSHAPE)
        outline = props.find("a:ln", _NS) if props is not None else None
        graphics.append({
            "name": name, "type": kind, "left": left, "top": top, "width": width,
            "height": height, "rotation": rotation, "visible": not hidden, "connector": tag == "cxnSp",
            "placeholder_type": None,
            "fill": None if kind in (_GROUP, _LINE) else _fill(props, element.find("p:style", _NS), scheme),
            "line_visible": None if outline is None else outline.find("a:noFill", _NS) is None,
            "text": None,
        })
    return graphics, title


def layout_facts(path, slide_indices):
    """What each slide's layout draws and where it puts the title.

    ``{slide_index: {"graphics": [...], "title": [left, top, width, height],
    "shows_master": bool, "theme_fonts": {...}}}``, the graphics in the shape
    the layout checks read, the title as the layout places it, or as its
    master does when the layout leaves it where the master put it, and the
    fonts of the master's theme.
    """
    found = {}
    with zipfile.ZipFile(path) as archive:
        for index, _, parts in _slides(archive, slide_indices):
            layout = next((p for p in parts.values() if "/slideLayouts/" in p), None)
            if not layout or layout not in archive.namelist():
                continue
            master = _related(None, archive, layout, "slideMasters")
            if not master or master not in archive.namelist():
                continue
            scheme = _scheme(archive, master)
            root = ElementTree.fromstring(archive.read(layout))
            graphics, title = _layout_shapes(root, scheme)
            if title is None:
                _, title = _layout_shapes(ElementTree.fromstring(archive.read(master)), scheme)
            found[index] = {"graphics": graphics, "title": title,
                            "shows_master": root.get("showMasterSp", "1") not in ("0", "false"),
                            "theme_fonts": _theme_fonts(archive, master)}
    return found


# ---------------------------------------------------------------------------
# What is inside groups
# ---------------------------------------------------------------------------
def _hidden(props):
    return props is not None and props.get("hidden") in ("1", "true")


def _right_to_left(text):
    return any(unicodedata.bidirectional(ch) in ("R", "AL") for ch in text)


def _level_directions(styles):
    """The rtl setting of each outline level a list of styles sets, {1: True}."""
    found = {}
    for level in range(1, 10):
        props = styles.find(f"a:lvl{level}pPr", _NS) if styles is not None else None
        if props is not None and props.get("rtl") is not None:
            found[level] = props.get("rtl") in ("1", "true")
    return found


def _text_facts(body, inherited=None):
    """The text facts the content checks read, from a shape's text body.

    ``inherited`` is the direction each outline level takes from the deck's
    default text style, overridden by the shape's own list styles.
    """
    paragraphs, runs, directions, scripts = [], set(), [], set()
    levels = {**(inherited or {}), **_level_directions(body.find("a:lstStyle", _NS))}
    for paragraph in body.findall("a:p", _NS):
        props = paragraph.find("a:pPr", _NS)
        first = paragraph.find("a:r/a:rPr", _NS)
        language = (first.get("lang") or "").split("-")[0].lower() if first is not None else ""
        level = int(props.get("lvl", "0")) + 1 if props is not None else 1
        rtl = props.get("rtl") if props is not None else None
        # Set right to left, here or by its style, or in a language PowerPoint
        # draws from the right.
        directions.append("rtl" if (rtl in ("1", "true") if rtl is not None else levels.get(level))
                          or language in ("he", "iw", "ar", "fa", "ur", "yi", "syr", "dv")
                          else "ltr")
        words = []
        for run in paragraph:
            tag = run.tag.rsplit("}", 1)[-1]
            if tag == "br":
                # A line break inside the paragraph, as PowerPoint hands it back.
                words.append("\v")
                continue
            if tag not in ("r", "fld"):
                continue
            text = run.findtext("a:t", default="", namespaces=_NS)
            words.append(text)
            rpr = run.find("a:rPr", _NS)
            if not text.strip() or rpr is None:
                continue
            latin = rpr.find("a:latin", _NS)
            cs = rpr.find("a:cs", _NS)
            size = rpr.get("sz")
            runs.add((latin.get("typeface") if latin is not None else None,
                      int(size) / 100 if size else None, rpr.get("b") in ("1", "true")))
            # Only Hebrew or Arabic is drawn in the right-to-left font.
            if cs is not None and _right_to_left(text):
                scripts.add(cs.get("typeface"))
        paragraphs.append("".join(words))
    content = "\r".join(paragraphs)
    return {
        "chars": len(content.strip()), "content": content,
        "left": None, "top": None, "width": None, "height": None, "alignment": None,
        "runs": [{"color_hex": None, "size": size, "bold": bold, "font": font, "theme": None}
                 for font, size, bold in sorted(runs, key=lambda r: (r[0] or "", r[1] or 0))],
        "margins": None, "word_wrap": None, "autofit": None, "orientation": None,
        "vertical_anchor": None, "directions": directions,
        "script_font": next(iter(scripts)) if len(scripts) == 1 else None,
        # Set through the tools only on Windows.
        "script_font_settable": False,
    }


def group_texts(path, slide_indices):
    """The text of every shape inside a group, ``{slide: [facts]}``.

    Each is named by the path the tools take, "Group 4/TextBox 2", and
    carries what its words and fonts are, for the checks that need no more:
    where PowerPoint draws them is not in the file.
    """
    found = {}
    with zipfile.ZipFile(path) as archive:
        presentation = ElementTree.fromstring(archive.read("ppt/presentation.xml"))
        inherited = _level_directions(presentation.find("p:defaultTextStyle", _NS))
        for index, slide, _ in _slides(archive, slide_indices):
            inner = []

            def walk(group, prefix, owners, hidden, into):
                for element in group:
                    tag = element.tag.rsplit("}", 1)[-1]
                    nv = next((c for c in element if c.tag.rsplit("}", 1)[-1].startswith("nv")), None)
                    props = nv.find("p:cNvPr", _NS) if nv is not None else None
                    if props is None:
                        continue
                    name = prefix + props.get("name", "")
                    if tag == "grpSp":
                        walk(element, name + "/", owners + [props.get("name", "")],
                             hidden or _hidden(props), into)
                        continue
                    body = element.find("p:txBody", _NS)
                    if tag != "sp" or body is None:
                        continue
                    into.append({
                        "name": name, "z": 0, "type": _AUTOSHAPE,
                        "left": 0.0, "top": 0.0, "width": 0.0, "height": 0.0,
                        "rotation": 0.0, "visible": not (hidden or _hidden(props)),
                        "connector": False,
                        "placeholder_type": None, "fill": None, "line_visible": None,
                        "text": _text_facts(body, inherited), "in_group": owners[0],
                        # Outermost first: each is ungrouped in turn to reach it.
                        "groups": owners,
                        # Read from the file because the tools cannot reach it.
                        "reachable": False,
                    })

            tree = slide.find("p:cSld/p:spTree", _NS)
            for group in tree.findall("p:grpSp", _NS) if tree is not None else []:
                props = group.find("p:nvGrpSpPr/p:cNvPr", _NS)
                if props is not None:
                    walk(group, props.get("name", "") + "/", [props.get("name", "")],
                         _hidden(props), inner)
            found[index] = inner
    return found
