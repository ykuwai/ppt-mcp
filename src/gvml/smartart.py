"""A SmartArt graphic's nodes, read and rewritten in its clipboard package.

PowerPoint for Mac has no ``smart art`` class and no ``nodes`` (#189), but
``copy shape`` on a SmartArt graphic writes the whole diagram to the
clipboard package, and pasting that package back makes a SmartArt graphic
again (checked on PowerPoint 16.114). Two parts of it hold a node's text:

- the data model (``diagrams/data*.xml``): a node is a ``dgm:pt`` with no
  ``type``, its text in ``dgm:t``; ``dgm:cxn`` elements with no ``type`` say
  which node is under which, in ``srcOrd`` order, from the ``doc`` point down;
- the drawing PowerPoint laid out from it (``diagrams/drawing*.xml``): a
  ``dsp:sp`` per drawn shape, whose ``modelId`` is the node's presentation
  point, the ``destId`` of the ``presOf`` connection from the node.

A node is numbered as Windows numbers ``SmartArt.AllNodes``: depth first from
the top, each node's children in order. Its text is written in both parts,
so the graphic shows it as soon as it is pasted.
"""

import xml.etree.ElementTree as ET
from typing import Dict, List, Optional

from gvml.package import PackageError
from gvml.text import body_text, write_body

NS_DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
NS_DSP = "http://schemas.microsoft.com/office/drawing/2008/diagram"
_D = f"{{{NS_DGM}}}"
_S = f"{{{NS_DSP}}}"


def _points(data: ET.Element) -> Dict[str, ET.Element]:
    lst = data.find(f"{_D}ptLst")
    if lst is None:
        raise PackageError("the SmartArt data model has no point list")
    return {pt.get("modelId"): pt for pt in lst.findall(f"{_D}pt")}


def _connections(data: ET.Element) -> List[ET.Element]:
    lst = data.find(f"{_D}cxnLst")
    return [] if lst is None else lst.findall(f"{_D}cxn")


def node_ids(data: ET.Element) -> List[str]:
    """The node points' ids in ``AllNodes`` order."""
    points = _points(data)
    children: Dict[str, List[tuple]] = {}
    for cxn in _connections(data):
        if cxn.get("type", "parOf") == "parOf":
            children.setdefault(cxn.get("srcId"), []).append(
                (int(cxn.get("srcOrd", "0")), cxn.get("destId")))
    roots = [mid for mid, pt in points.items() if pt.get("type") == "doc"]
    if not roots:
        raise PackageError("the SmartArt data model has no document point")

    order: List[str] = []

    def walk(mid):
        for _, child in sorted(children.get(mid, [])):
            pt = points.get(child)
            if pt is not None and pt.get("type", "node") == "node":
                order.append(child)
                walk(child)

    walk(roots[0])
    return order


def node_texts(data: ET.Element) -> List[str]:
    """Each node's text, in ``AllNodes`` order."""
    points = _points(data)
    out = []
    for mid in node_ids(data):
        t = points[mid].find(f"{_D}t")
        out.append("" if t is None else body_text(t))
    return out


def _drawn_shape(data: ET.Element, drawing: Optional[ET.Element], node_id: str):
    """The ``dsp:sp`` PowerPoint drew a node's text in, or None."""
    if drawing is None:
        return None
    pres = [c.get("destId") for c in _connections(data)
            if c.get("type") == "presOf" and c.get("srcId") == node_id]
    for sp in drawing.iter(f"{_S}sp"):
        if sp.get("modelId") in pres and sp.find(f"{_S}txBody") is not None:
            return sp
    return None


def set_node_text(data: ET.Element, drawing: Optional[ET.Element], node_index: int,
                  text: str) -> None:
    """Write node ``node_index`` (1-based, ``AllNodes`` order)'s text in the
    data model and, where it was drawn, in the drawing. Raises ValueError for
    an index out of range."""
    ids = node_ids(data)
    if not 1 <= node_index <= len(ids):
        raise ValueError(f"node_index {node_index} out of range (1-{len(ids)})")
    pt = _points(data)[ids[node_index - 1]]
    t = pt.find(f"{_D}t")
    if t is None:
        t = ET.SubElement(pt, f"{_D}t")
    write_body(t, text)
    # An empty node is marked as a placeholder, which PowerPoint shows as
    # "[Text]" while editing; a node with text is not one.
    pr = pt.find(f"{_D}prSet")
    if pr is not None and "phldr" in pr.attrib:
        del pr.attrib["phldr"]
    drawn = _drawn_shape(data, drawing, ids[node_index - 1])
    if drawn is not None:
        write_body(drawn.find(f"{_S}txBody"), text)


def _parent_link(data: ET.Element, node_id: str) -> ET.Element:
    """The ``parOf`` connection that puts ``node_id`` under its parent."""
    for cxn in _connections(data):
        if cxn.get("type", "parOf") == "parOf" and cxn.get("destId") == node_id:
            return cxn
    raise PackageError("a node has no connection to its parent")


def _new_id() -> str:
    import uuid

    return "{" + str(uuid.uuid4()).upper() + "}"


def add_node(data: ET.Element, after_index: Optional[int], text: Optional[str]) -> int:
    """Add a node after node ``after_index`` (1-based, ``AllNodes`` order) at
    its level, or after the last node when it is None, as Windows'
    ``AllNodes(i).AddNode()`` does. Answers the new node's index.

    The new node, and the two transition points every connection carries,
    are made from the neighbour's, so they have the shape PowerPoint gave
    its own; the later siblings move one place on. Only the data model is
    written: PowerPoint lays the diagram out again when it is pasted.
    """
    import copy

    ids = node_ids(data)
    if after_index is None:
        after_index = len(ids)
    if not 1 <= after_index <= len(ids):
        raise ValueError(f"node_index {after_index} out of range (1-{len(ids)})")
    points = _points(data)
    neighbour = ids[after_index - 1]
    link = _parent_link(data, neighbour)
    parent, order = link.get("srcId"), int(link.get("srcOrd", "0"))

    for cxn in _connections(data):
        if (cxn.get("type", "parOf") == "parOf" and cxn.get("srcId") == parent
                and int(cxn.get("srcOrd", "0")) > order):
            cxn.set("srcOrd", str(int(cxn.get("srcOrd")) + 1))

    node_id, cxn_id = _new_id(), _new_id()
    node = copy.deepcopy(points[neighbour])
    node.set("modelId", node_id)
    for child in list(node):
        if child.tag not in (f"{_D}prSet", f"{_D}spPr", f"{_D}t"):
            node.remove(child)
    point_list = data.find(f"{_D}ptLst")
    point_list.append(node)
    transitions = {}
    for kind in ("parTransId", "sibTransId"):
        source = points.get(link.get(kind))
        if source is None:
            continue
        made = copy.deepcopy(source)
        transitions[kind] = _new_id()
        made.set("modelId", transitions[kind])
        made.set("cxnId", cxn_id)
        point_list.append(made)
    new_link = copy.deepcopy(link)
    new_link.set("modelId", cxn_id)
    new_link.set("destId", node_id)
    new_link.set("srcOrd", str(order + 1))
    for kind, value in transitions.items():
        new_link.set(kind, value)
    data.find(f"{_D}cxnLst").append(new_link)

    index = node_ids(data).index(node_id) + 1
    t = node.find(f"{_D}t")
    if t is None:
        t = ET.SubElement(node, f"{_D}t")
    write_body(t, text or "")
    pr = node.find(f"{_D}prSet")
    if pr is None:
        pr = ET.Element(f"{_D}prSet")
        node.insert(0, pr)
    if text:
        pr.attrib.pop("phldr", None)
    else:
        pr.set("phldr", "1")
    return index


def delete_node(data: ET.Element, node_index: int) -> int:
    """Delete node ``node_index`` and the nodes under it, as Windows'
    ``AllNodes(i).Delete()`` does, with their transition and presentation
    points and every connection to them; the later siblings move one place
    back. Answers how many nodes went. Deleting the only node is refused,
    since a graphic with none has nothing to lay out."""
    ids = node_ids(data)
    if not 1 <= node_index <= len(ids):
        raise ValueError(f"node_index {node_index} out of range (1-{len(ids)})")
    connections = _connections(data)
    first = ids[node_index - 1]

    subtree = [first]
    for current in subtree:
        subtree += [c.get("destId") for c in connections
                    if c.get("type", "parOf") == "parOf" and c.get("srcId") == current]
    if len(subtree) >= len(ids):
        raise ValueError(
            "That would delete every node, and a SmartArt graphic needs one. "
            "Set the last node's text instead."
        )

    link = _parent_link(data, first)
    parent, order = link.get("srcId"), int(link.get("srcOrd", "0"))
    gone = set(subtree)
    for node_id in subtree:
        own = _parent_link(data, node_id)
        gone |= {own.get("parTransId"), own.get("sibTransId")} - {None}
        gone |= {c.get("destId") for c in connections
                 if c.get("type") == "presOf" and c.get("srcId") == node_id}

    point_list = data.find(f"{_D}ptLst")
    for pt in list(point_list):
        if pt.get("modelId") in gone:
            point_list.remove(pt)
    connection_list = data.find(f"{_D}cxnLst")
    for cxn in list(connection_list):
        if cxn.get("srcId") in gone or cxn.get("destId") in gone:
            connection_list.remove(cxn)
        elif (cxn.get("type", "parOf") == "parOf" and cxn.get("srcId") == parent
              and int(cxn.get("srcOrd", "0")) > order):
            cxn.set("srcOrd", str(int(cxn.get("srcOrd")) - 1))
    return len(subtree)
