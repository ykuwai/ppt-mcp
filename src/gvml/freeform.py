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
from dataclasses import dataclass, field
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

    path_xml, x, y, cx, cy = _path_xml(anchors[0], segments, close_path)
    geometry = (
        '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
        '<a:rect l="0" t="0" r="r" b="b"/>'
        f'<a:pathLst>{path_xml}</a:pathLst></a:custGeom>'
    )
    return geometry, x, y, cx, cy


def _box(start: Point, segments: Sequence[tuple]) -> Tuple[int, int, int, int]:
    """The EMU box round every point of a path, control points included."""
    points = [start] + [p for seg in segments for p in seg[1:]]
    left = min(p[0] for p in points)
    top = min(p[1] for p in points)
    right = max(p[0] for p in points)
    bottom = max(p[1] for p in points)
    x, y = emu(left), emu(top)
    # A path with no width or no height (all points on one line) is still a
    # shape; one EMU keeps the extent non-zero without moving anything.
    cx = max(emu(right) - x, 1)
    cy = max(emu(bottom) - y, 1)
    return x, y, cx, cy


def _path_xml(
    start: Point, segments: Sequence[tuple], closed: bool, attrs: str = "",
) -> Tuple[str, int, int, int, int]:
    """One ``a:path`` for segments given in slide points, plus its box.

    ``attrs`` carries any attributes of the path other than ``w`` and ``h``
    (``fill="none"`` on an open path PowerPoint drew, for instance), already
    serialised.
    """
    x, y, cx, cy = _box(start, segments)

    def local(p: Point) -> str:
        return f'<a:pt x="{emu(p[0]) - x}" y="{emu(p[1]) - y}"/>'

    body = f"<a:moveTo>{local(start)}</a:moveTo>"
    for seg in segments:
        if seg[0] == "line":
            body += f"<a:lnTo>{local(seg[1])}</a:lnTo>"
        else:
            body += f"<a:cubicBezTo>{local(seg[1])}{local(seg[2])}{local(seg[3])}</a:cubicBezTo>"
    if closed:
        body += "<a:close/>"
    return f'<a:path w="{cx}" h="{cy}"{attrs}>{body}</a:path>', x, y, cx, cy


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


@dataclass
class Path:
    """One ``a:path`` in slide points: a start, its segments, and whether it
    closes. A segment is ``("line", end)`` or ``("curve", c1, c2, end)``.

    ``attrs`` keeps the path's attributes other than ``w`` and ``h`` so an
    edited path is written back with the same ones.
    """

    start: Point
    segments: List[tuple] = field(default_factory=list)
    closed: bool = False
    attrs: dict = field(default_factory=dict)

    def copy(self) -> "Path":
        return Path(self.start, [tuple(s) for s in self.segments], self.closed, dict(self.attrs))


def _parse_paths(sp: ET.Element) -> List[Path]:
    """Every path of a shape's custom geometry, in slide points."""
    x0, y0, cx, cy = _xfrm(sp)
    cust = sp.find(f"{_A}spPr/{_A}custGeom")
    if cust is None:
        raise PackageError("the shape has no custom geometry, so it has no nodes to read")
    paths = cust.findall(f"{_A}pathLst/{_A}path")
    if not paths:
        raise PackageError("the custom geometry holds no path")

    out: List[Path] = []
    for path in paths:
        w = int(path.get("w", 0)) or cx
        h = int(path.get("h", 0)) or cy
        sx = cx / w if w else 1.0
        sy = cy / h if h else 1.0

        def slide_pt(p: Tuple[int, int]) -> Point:
            return (pt(x0 + p[0] * sx), pt(y0 + p[1] * sy))

        parsed: Optional[Path] = None
        for cmd in path:
            tag = cmd.tag[len(_A):] if cmd.tag.startswith(_A) else cmd.tag
            pts = [slide_pt(_pt_of(p)) for p in cmd.findall(f"{_A}pt")]
            if tag == "moveTo":
                if parsed is not None:
                    raise PackageError(
                        "the path holds a second moveTo, which is a subpath "
                        "Windows numbers as one list and this cannot rewrite"
                    )
                parsed = Path(pts[0], [], False, {k: v for k, v in path.attrib.items() if k not in ("w", "h")})
            elif parsed is None:
                raise PackageError("the path does not begin with a moveTo")
            elif tag == "lnTo":
                parsed.segments.append(("line", pts[0]))
            elif tag in ("cubicBezTo", "quadBezTo"):
                if tag == "quadBezTo":
                    # Degree elevation, so a quadratic reads as the cubic it
                    # equals and the node count stays three per curve.
                    p0 = parsed.segments[-1][-1] if parsed.segments else parsed.start
                    q, p1 = pts[0], pts[1]
                    pts = [
                        (p0[0] + 2 * (q[0] - p0[0]) / 3, p0[1] + 2 * (q[1] - p0[1]) / 3),
                        (p1[0] + 2 * (q[0] - p1[0]) / 3, p1[1] + 2 * (q[1] - p1[1]) / 3),
                        p1,
                    ]
                parsed.segments.append(("curve", pts[0], pts[1], pts[2]))
            elif tag == "close":
                parsed.closed = True
            elif tag == "arcTo":
                raise PackageError(
                    "the path holds an arcTo segment, which has no node form on "
                    "Windows and is not read here"
                )
            else:
                raise PackageError(f"the path holds an unexpected {tag} element")
        if parsed is not None:
            out.append(parsed)
    if not out:
        raise PackageError("the custom geometry holds no path with a moveTo")
    return out


def _slots(path: Path) -> List[tuple]:
    """Windows node index minus one -> ``(segment index, role)``.

    Role is ``"start"`` for the moveTo (segment index -1), ``"c1"`` and
    ``"c2"`` for a curve's control points, and ``"end"`` for the point a
    segment ends at. One entry per node, in Windows order.
    """
    slots = [(-1, "start")]
    for i, seg in enumerate(path.segments):
        if seg[0] == "curve":
            slots.append((i, "c1"))
            slots.append((i, "c2"))
        slots.append((i, "end"))
    return slots


def _point_at(path: Path, slot: tuple) -> Point:
    i, role = slot
    if role == "start":
        return path.start
    seg = path.segments[i]
    if role == "end":
        return seg[-1]
    return seg[1] if role == "c1" else seg[2]


def nodes_of(paths: Sequence[Path]) -> List[dict]:
    """The Windows node list of one or more paths, numbered continuously."""
    nodes: List[dict] = []
    for path in paths:
        slots = _slots(path)
        for k, slot in enumerate(slots):
            i, role = slot
            p = _point_at(path, slot)
            index = len(nodes) + 1
            x, y = round(p[0], 2), round(p[1], 2)
            follows = path.segments[i + 1] if i + 1 < len(path.segments) else None
            if role in ("c1", "c2") or follows is None:
                nodes.append({
                    "index": index, "x": x, "y": y,
                    "editing_type": "inaccessible",
                    "segment_type": "inaccessible",
                    "note": INACCESSIBLE_NOTE,
                })
            else:
                incoming = path.segments[i][2] if i >= 0 and path.segments[i][0] == "curve" else None
                outgoing = follows[1] if follows[0] == "curve" else None
                nodes.append({
                    "index": index, "x": x, "y": y,
                    "editing_type": _editing_type(incoming, p, outgoing),
                    "segment_type": follows[0],
                })
    return nodes


def read_nodes(sp: ET.Element) -> List[dict]:
    """The nodes of an ``a:sp`` with custom geometry, Windows numbered.

    Each dict has ``index``, ``x``, ``y`` in points on the slide, and
    ``editing_type`` and ``segment_type``. ``segment_type`` on an anchor is
    the segment that follows it, which is what ``ShapeNode.SegmentType``
    reports; a control point, and the last node of a path, which nothing
    follows, report ``inaccessible`` with the note Windows gives.
    """
    return nodes_of(_parse_paths(sp))


# ---------------------------------------------------------------------------
# Editing one path
# ---------------------------------------------------------------------------
# What the four node tools do to a path, in slide points, before it is written
# back with ``write_path``. Each takes the Windows node index and returns a new
# Path, raising ``ValueError`` with the Windows wording for an index that is
# out of range. What each does to the neighbouring points is what PowerPoint's
# own Edit Points does, since Windows does not document its Nodes methods any
# more closely than that.

def read_path(sp: ET.Element) -> Path:
    """The one path of a shape, or a complaint that it has more than one."""
    paths = _parse_paths(sp)
    if len(paths) != 1:
        raise PackageError(
            f"the custom geometry holds {len(paths)} paths; Windows numbers "
            "them as one node list and this rewrites only a single path"
        )
    return paths[0]


def node_count(path: Path) -> int:
    return len(_slots(path))


def _check_index(path: Path, index: int, what: str = "node_index") -> tuple:
    slots = _slots(path)
    if index < 1 or index > len(slots):
        raise ValueError(f"{what} {index} out of range (shape has {len(slots)} nodes).")
    return slots[index - 1]


def _with_point(seg: tuple, role: str, p: Point) -> tuple:
    if seg[0] == "line":
        return ("line", p)
    c1, c2, end = seg[1], seg[2], seg[3]
    if role == "c1":
        c1 = p
    elif role == "c2":
        c2 = p
    else:
        end = p
    return ("curve", c1, c2, end)


def set_node_position(path: Path, index: int, x: float, y: float) -> Path:
    """Move one node. An anchor takes the handles attached to it along,
    which is how Edit Points drags a vertex; a control point moves alone."""
    i, role = _check_index(path, index)
    new = path.copy()
    p = (float(x), float(y))
    if role in ("c1", "c2"):
        new.segments[i] = _with_point(new.segments[i], role, p)
        return new
    old = _point_at(path, (i, role))
    dx, dy = p[0] - old[0], p[1] - old[1]
    if role == "start":
        new.start = p
    else:
        seg = new.segments[i]
        if seg[0] == "curve":
            seg = ("curve", seg[1], (seg[2][0] + dx, seg[2][1] + dy), p)
        else:
            seg = ("line", p)
        new.segments[i] = seg
    following = i + 1
    if following < len(new.segments) and new.segments[following][0] == "curve":
        seg = new.segments[following]
        new.segments[following] = ("curve", (seg[1][0] + dx, seg[1][1] + dy), seg[2], seg[3])
    return new


def insert_node(
    path: Path, after_index: int, seg_int: int, et_int: int,
    x1: float, y1: float, x2=None, y2=None, x3=None, y3=None,
) -> Path:
    """Add a segment after an anchor, ending at the new point; the segment
    that used to follow the anchor now starts from the new point."""
    i, role = _check_index(path, after_index, "after_index")
    if role in ("c1", "c2"):
        raise ValueError(
            f"after_index {after_index} is a curve's control point; insert "
            "after a vertex (a node whose segment_type is not 'inaccessible')."
        )
    new = path.copy()
    at = i + 1
    if seg_int == SEGMENT_LINE:
        seg = ("line", (float(x1), float(y1)))
    elif et_int == EDITING_AUTO:
        start = _point_at(path, (i, role))
        end = (float(x1), float(y1))
        before = _point_at(path, (i - 1, "end")) if i >= 1 else (path.start if i == 0 else None)
        after = path.segments[at][-1] if at < len(path.segments) else None
        anchors = [a for a in (before, start, end, after) if a is not None]
        k = anchors.index(start)
        c1, c2 = _auto_controls(anchors, k, False)
        seg = ("curve", c1, c2, end)
    else:
        seg = ("curve", (float(x1), float(y1)), (float(x2), float(y2)), (float(x3), float(y3)))
    new.segments.insert(at, seg)
    return new


def delete_node(path: Path, index: int) -> Path:
    """Remove a node and the segment after it.

    A control point stands for its own curve, and the segment after a
    control point is that curve, so deleting one deletes the whole Bézier
    segment, its other control point and its end point with it. That is
    what ``ppt_delete_node`` promises and what Windows does; leaving a
    straight line between the curve's ends would put an edge in the outline
    that the caller never asked for.
    """
    i, role = _check_index(path, index)
    new = path.copy()
    if role in ("c1", "c2"):
        del new.segments[i]
    elif role == "start":
        if not new.segments:
            raise ValueError("node_index 1 is the only node; a path cannot lose it.")
        new.start = new.segments[0][-1]
        del new.segments[0]
    elif i + 1 < len(new.segments):
        following = new.segments[i + 1]
        new.segments[i] = _with_point(new.segments[i], "end", following[-1])
        del new.segments[i + 1]
    else:
        del new.segments[i]
    if not new.segments:
        raise ValueError(
            f"node_index {index} cannot be deleted; a freeform needs at least two nodes."
        )
    return new


def set_segment_type(path: Path, index: int, seg_int: int) -> Path:
    """Switch the segment after a node between line and curve. A control
    point stands for its own curve. A new curve puts its handles a third
    and two thirds of the way along the chord, so the outline does not
    move until a handle does."""
    i, role = _check_index(path, index)
    target = i if role in ("c1", "c2") else i + 1
    if target >= len(path.segments):
        raise ValueError(f"node_index {index} is the last node; no segment follows it.")
    new = path.copy()
    seg = new.segments[target]
    if seg_int == SEGMENT_LINE:
        new.segments[target] = ("line", seg[-1])
    elif seg[0] == "line":
        start = _point_at(path, (target - 1, "end")) if target >= 1 else path.start
        end = seg[1]
        c1 = (start[0] + (end[0] - start[0]) / 3.0, start[1] + (end[1] - start[1]) / 3.0)
        c2 = (start[0] + 2.0 * (end[0] - start[0]) / 3.0, start[1] + 2.0 * (end[1] - start[1]) / 3.0)
        new.segments[target] = ("curve", c1, c2, end)
    return new


def write_path(sp: ET.Element, path: Path) -> Tuple[int, int, int, int]:
    """Put a path back into a shape's ``a:sp``, resizing the shape to it.

    The path list is rewritten and the transform's offset and extent are set
    to the box round every point, with ``a:path w/h`` equal to the extent, so
    a resized freeform's scale is baked in. Everything else on the shape,
    fill, line, text, style, is left as it was. Returns the new box in EMU.
    """
    attrs = "".join(f' {k}="{v}"' for k, v in path.attrs.items())
    path_xml, x, y, cx, cy = _path_xml(path.start, path.segments, path.closed, attrs)
    cust = sp.find(f"{_A}spPr/{_A}custGeom")
    path_lst = cust.find(f"{_A}pathLst") if cust is not None else None
    if path_lst is None:
        raise PackageError("the shape has no path list to write into")
    for old in list(path_lst):
        path_lst.remove(old)
    path_lst.append(ET.fromstring(f'<w xmlns:a="{NS_A}">{path_xml}</w>')[0])
    xfrm = sp.find(f"{_A}spPr/{_A}xfrm")
    xfrm.find(f"{_A}off").set("x", str(x))
    xfrm.find(f"{_A}off").set("y", str(y))
    xfrm.find(f"{_A}ext").set("cx", str(cx))
    xfrm.find(f"{_A}ext").set("cy", str(cy))
    return x, y, cx, cy


def _xfrm(sp: ET.Element) -> Tuple[int, int, int, int]:
    xfrm = sp.find(f"{_A}spPr/{_A}xfrm")
    if xfrm is None:
        raise PackageError("the shape has no transform")
    off = xfrm.find(f"{_A}off")
    ext = xfrm.find(f"{_A}ext")
    if off is None or ext is None:
        raise PackageError("the shape's transform has no offset or extent")
    return int(off.get("x", 0)), int(off.get("y", 0)), int(ext.get("cx", 0)), int(ext.get("cy", 0))
