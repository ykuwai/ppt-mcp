"""Freeform paths: the node list Windows speaks, to and from ``a:custGeom``.

Windows builds a freeform from a node list in which a straight segment is one
node and a Bézier segment is three, two control points and an end point, and
reads it back the same way, with the control points reporting their metadata
as ``inaccessible``. ``a:lnTo`` and ``a:cubicBezTo`` have exactly that shape,
so the numbering here is the Windows numbering and a caller who counts nodes
on one platform can count them on the other.

Coordinates in a path are in the path's own space, ``a:path w/h``, and are
mapped onto the shape's ``a:ext``. PowerPoint writes the two equal for a
freeform it made, and a freeform resized afterwards keeps its path and
changes only the extent, so the reader scales rather than assuming. A
freeform is written here with the two equal, at 12700 EMU to the point, so
its path coordinates are its slide coordinates less its offset.
"""

import xml.etree.ElementTree as ET
from typing import List, Optional, Sequence, Tuple

from gvml.canvas import emu, pt
from gvml.package import NS_A, PackageError

# MsoSegmentType and MsoEditingType, as Windows numbers them. Spelled here so
# gvml imports nothing from ppt_com.
SEGMENT_LINE = 0
SEGMENT_CURVE = 1
EDITING_AUTO = 0
EDITING_CORNER = 1
EDITING_SMOOTH = 2
EDITING_SYMMETRIC = 3

SEGMENT_NAMES = {SEGMENT_LINE: "line", SEGMENT_CURVE: "curve"}
EDITING_NAMES = {
    EDITING_AUTO: "auto", EDITING_CORNER: "corner",
    EDITING_SMOOTH: "smooth", EDITING_SYMMETRIC: "symmetric",
}

INACCESSIBLE_NOTE = "Metadata not accessible via COM (control point or closing node)."

_A = f"{{{NS_A}}}"
Point = Tuple[float, float]


# ---------------------------------------------------------------------------
# Writing
# ---------------------------------------------------------------------------

def _auto_controls(anchors: Sequence[Point], i: int, closed: bool) -> Tuple[Point, Point]:
    """Control points for the curve from ``anchors[i]`` to ``anchors[i+1]``.

    Windows chooses the control points itself for a curve added with
    ``msoEditingAuto``, and does not say how. This is a Catmull-Rom spline,
    which passes through every anchor with a tangent parallel to the chord
    between its neighbours, the usual answer to "a smooth curve through these
    points". At an open end the missing neighbour is the end point itself, so
    the curve arrives straight. On a closed path the neighbours wrap.
    """
    n = len(anchors)
    p1, p2 = anchors[i], anchors[i + 1]
    if closed and n > 2:
        # anchors[-1] repeats anchors[0], so the wrap skips the duplicate.
        p0 = anchors[i - 1] if i > 0 else anchors[n - 2]
        p3 = anchors[i + 2] if i + 2 < n else anchors[1]
    else:
        p0 = anchors[i - 1] if i > 0 else p1
        p3 = anchors[i + 2] if i + 2 < n else p2
    c1 = (p1[0] + (p2[0] - p0[0]) / 6.0, p1[1] + (p2[1] - p0[1]) / 6.0)
    c2 = (p2[0] - (p3[0] - p1[0]) / 6.0, p2[1] - (p3[1] - p1[1]) / 6.0)
    return c1, c2


def build_geometry(
    start_x: float, start_y: float, nodes_data: Sequence[dict], close_path: bool,
) -> Tuple[str, int, int, int, int]:
    """``a:custGeom`` XML for a node list, plus the shape box in EMU.

    ``nodes_data`` is the list the Windows tool passes to its ``_impl``: dicts
    with ``seg_int``, ``et_int``, ``x1``, ``y1`` and, for a corner curve,
    ``x2``, ``y2``, ``x3``, ``y3``. Coordinates are in points on the slide.

    Returns ``(geometry_xml, x, y, cx, cy)``.
    """
    # First pass: the anchors, so auto curves can look at their neighbours.
    anchors: List[Point] = [(float(start_x), float(start_y))]
    for nd in nodes_data:
        if nd["seg_int"] == SEGMENT_CURVE and nd["et_int"] != EDITING_AUTO:
            anchors.append((float(nd["x3"]), float(nd["y3"])))
        else:
            anchors.append((float(nd["x1"]), float(nd["y1"])))
    if close_path:
        anchors.append(anchors[0])

    # Second pass: the segments, in points.
    segments: List[tuple] = []
    for i, nd in enumerate(nodes_data):
        if nd["seg_int"] == SEGMENT_LINE:
            segments.append(("line", anchors[i + 1]))
        elif nd["et_int"] == EDITING_AUTO:
            c1, c2 = _auto_controls(anchors, i, close_path)
            segments.append(("curve", c1, c2, anchors[i + 1]))
        else:
            segments.append((
                "curve",
                (float(nd["x1"]), float(nd["y1"])),
                (float(nd["x2"]), float(nd["y2"])),
                anchors[i + 1],
            ))
    if close_path:
        # Windows adds an explicit straight segment back to the start, and so
        # does this, so the node count agrees on both sides.
        segments.append(("line", anchors[0]))

    points = [anchors[0]] + [p for seg in segments for p in seg[1:]]
    left = min(p[0] for p in points)
    top = min(p[1] for p in points)
    right = max(p[0] for p in points)
    bottom = max(p[1] for p in points)
    x, y = emu(left), emu(top)
    # A path with no width or no height (all points on one line) is still a
    # shape; one EMU keeps the extent non-zero without moving anything.
    cx = max(emu(right) - x, 1)
    cy = max(emu(bottom) - y, 1)

    def local(p: Point) -> str:
        return f'<a:pt x="{emu(p[0]) - x}" y="{emu(p[1]) - y}"/>'

    body = f"<a:moveTo>{local(anchors[0])}</a:moveTo>"
    for seg in segments:
        if seg[0] == "line":
            body += f"<a:lnTo>{local(seg[1])}</a:lnTo>"
        else:
            body += f"<a:cubicBezTo>{local(seg[1])}{local(seg[2])}{local(seg[3])}</a:cubicBezTo>"
    if close_path:
        body += "<a:close/>"

    geometry = (
        '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
        '<a:rect l="0" t="0" r="r" b="b"/>'
        f'<a:pathLst><a:path w="{cx}" h="{cy}">{body}</a:path></a:pathLst></a:custGeom>'
    )
    return geometry, x, y, cx, cy


# ---------------------------------------------------------------------------
# Reading
# ---------------------------------------------------------------------------

def _pt_of(element: ET.Element) -> Tuple[int, int]:
    return int(element.get("x", 0)), int(element.get("y", 0))


def _editing_type(incoming: Optional[Point], anchor: Point, outgoing: Optional[Point]) -> str:
    """Corner, smooth or symmetric, from the geometry of the two handles.

    OOXML stores no editing type; PowerPoint's UI derives it from where the
    handles sit, and this does the same. Collinear handles on opposite sides
    are smooth, and smooth with equal lengths is symmetric. Anything else,
    including an anchor with a handle on one side only, is a corner.
    """
    if incoming is None or outgoing is None:
        return EDITING_NAMES[EDITING_CORNER]
    ax, ay = anchor
    ux, uy = incoming[0] - ax, incoming[1] - ay
    vx, vy = outgoing[0] - ax, outgoing[1] - ay
    lu = (ux * ux + uy * uy) ** 0.5
    lv = (vx * vx + vy * vy) ** 0.5
    if lu == 0 or lv == 0:
        return EDITING_NAMES[EDITING_CORNER]
    cross = abs(ux * vy - uy * vx) / (lu * lv)
    dot = (ux * vx + uy * vy) / (lu * lv)
    if cross > 0.02 or dot > -0.98:
        return EDITING_NAMES[EDITING_CORNER]
    if abs(lu - lv) <= 0.02 * max(lu, lv):
        return EDITING_NAMES[EDITING_SYMMETRIC]
    return EDITING_NAMES[EDITING_SMOOTH]


def read_nodes(sp: ET.Element) -> List[dict]:
    """The nodes of an ``a:sp`` with custom geometry, Windows numbered.

    Each dict has ``index``, ``x``, ``y`` in points on the slide, and
    ``editing_type`` and ``segment_type``. ``segment_type`` on an anchor is
    the segment that follows it, which is what ``ShapeNode.SegmentType``
    reports; a control point, and the last node of a path, which nothing
    follows, report ``inaccessible`` with the note Windows gives.
    """
    x0, y0, cx, cy = _xfrm(sp)
    cust = sp.find(f"{_A}spPr/{_A}custGeom")
    if cust is None:
        raise PackageError("the shape has no custom geometry, so it has no nodes to read")
    paths = cust.findall(f"{_A}pathLst/{_A}path")
    if not paths:
        raise PackageError("the custom geometry holds no path")

    nodes: List[dict] = []
    for path in paths:
        w = int(path.get("w", 0)) or cx
        h = int(path.get("h", 0)) or cy
        sx = cx / w if w else 1.0
        sy = cy / h if h else 1.0

        def slide_pt(p: Tuple[int, int]) -> Point:
            return (pt(x0 + p[0] * sx), pt(y0 + p[1] * sy))

        # Every node in this path, with what arrives at it and leaves it.
        # entries: [kind, point, in_control, out_control, follows]
        entries: List[dict] = []
        for cmd in path:
            tag = cmd.tag[len(_A):] if cmd.tag.startswith(_A) else cmd.tag
            pts = [slide_pt(_pt_of(p)) for p in cmd.findall(f"{_A}pt")]
            if tag == "moveTo":
                entries.append({"kind": "anchor", "p": pts[0], "in": None, "out": None, "follows": None})
            elif tag == "lnTo":
                if entries:
                    entries[-1]["follows"] = "line"
                entries.append({"kind": "anchor", "p": pts[0], "in": None, "out": None, "follows": None})
            elif tag in ("cubicBezTo", "quadBezTo"):
                if tag == "quadBezTo":
                    # Degree elevation, so a quadratic reads as the cubic it
                    # equals and the node count stays three per curve.
                    p0 = entries[-1]["p"] if entries else pts[0]
                    q, p1 = pts[0], pts[1]
                    pts = [
                        (p0[0] + 2 * (q[0] - p0[0]) / 3, p0[1] + 2 * (q[1] - p0[1]) / 3),
                        (p1[0] + 2 * (q[0] - p1[0]) / 3, p1[1] + 2 * (q[1] - p1[1]) / 3),
                        p1,
                    ]
                if entries:
                    entries[-1]["follows"] = "curve"
                    entries[-1]["out"] = pts[0]
                entries.append({"kind": "control", "p": pts[0]})
                entries.append({"kind": "control", "p": pts[1]})
                entries.append({"kind": "anchor", "p": pts[2], "in": pts[1], "out": None, "follows": None})
            elif tag == "close":
                continue
            elif tag == "arcTo":
                raise PackageError(
                    "the path holds an arcTo segment, which has no node form on "
                    "Windows and is not read here"
                )
            else:
                raise PackageError(f"the path holds an unexpected {tag} element")

        for entry in entries:
            index = len(nodes) + 1
            x, y = round(entry["p"][0], 2), round(entry["p"][1], 2)
            if entry["kind"] == "control" or entry["follows"] is None:
                nodes.append({
                    "index": index, "x": x, "y": y,
                    "editing_type": "inaccessible",
                    "segment_type": "inaccessible",
                    "note": INACCESSIBLE_NOTE,
                })
            else:
                nodes.append({
                    "index": index, "x": x, "y": y,
                    "editing_type": _editing_type(entry["in"], entry["p"], entry["out"]),
                    "segment_type": entry["follows"],
                })
    return nodes


def _xfrm(sp: ET.Element) -> Tuple[int, int, int, int]:
    xfrm = sp.find(f"{_A}spPr/{_A}xfrm")
    if xfrm is None:
        raise PackageError("the shape has no transform")
    off = xfrm.find(f"{_A}off")
    ext = xfrm.find(f"{_A}ext")
    if off is None or ext is None:
        raise PackageError("the shape's transform has no offset or extent")
    return int(off.get("x", 0)), int(off.get("y", 0)), int(ext.get("cx", 0)), int(ext.get("cy", 0))
