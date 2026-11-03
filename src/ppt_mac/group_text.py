"""Text and formatting on a member of a group, by way of the group's package.

Nothing reaches a group's member over Apple Events (#220), so a tool given
the name of one used to answer that the name was not on the slide, and the
only way in was ``ppt_ungroup_shapes``, which renames every member and drops
the group's animations. Here the group is copied, the member is found in the
package by name or ``Group/Child`` path (``gvml.text``), its text or its runs
are rewritten, and the group is pasted back over the original with
``replace_shape``, which checks the paste before it deletes anything. The
members keep their names, because the package carries them.

What that costs is said, as for every edit through this route: the group is
a new object with the old name, and animations on it are gone with the
original.
"""

import logging
import xml.etree.ElementTree as ET
from typing import Callable, Optional

from backend.mac_ae import ppt
from backend.mac_enums import MsoShapeType
from gvml import PackageError
from gvml import canvas as _canvas
from gvml import shapes as _gvml_shapes
from gvml import text as _gvml_text
from ppt_com.constants import msoGroup
from ppt_mac.gvml_paste import (
    Clipboard,
    Refused,
    copy_shape_package,
    replace_shape,
    with_warnings,
)
from ppt_mac.shapes import _shape_names
from utils.navigation import goto_slide

logger = logging.getLogger(__name__)

_ALTERNATIVES = ["ppt_get_group_items", "ppt_ungroup_shapes"]


def is_top_level(slide, identifier) -> bool:
    """Whether the tools' ordinary lookup will find ``identifier``."""
    return not isinstance(identifier, str) or identifier in _shape_names(slide)


def _groups(slide, identifier: str):
    """The top-level groups that could hold ``identifier``, as (position,
    name): the one a path starts with, or every group for a bare name."""
    names = _shape_names(slide)
    groups = []
    for position, name in enumerate(names, start=1):
        try:
            if slide.shapes[position].shape_type() == MsoShapeType[msoGroup]:
                groups.append((position, name))
        except Exception:  # noqa: BLE001 - a shape that will not say is not one
            continue
    starts = [g for g in groups if identifier.startswith(g[1] + _gvml_text.SEPARATOR)]
    return starts or groups


def _locate(slide, identifier: str, clip, tool_name: str):
    """Copy each group that could hold the member until it is found.

    Answers (position, group name, package, drawing root, group element,
    path, member), or None when no group holds it. A bare name held by
    members of two groups is refused, with both paths.
    """
    found = []
    for position, group_name in _groups(slide, identifier):
        package = copy_shape_package(slide.shapes[position], clip, tool_name)
        root = ET.fromstring(package.drawing())
        group = _gvml_shapes.only_child(
            _canvas.children(_canvas.canvas_of(root)), "group")
        try:
            path, member = _gvml_text.find_member(group, group_name, identifier)
        except _gvml_text.MemberNotFound:
            continue
        found.append((position, group_name, package, root, group, path, member))
        if path == identifier:
            break
    if len(found) > 1:
        raise _gvml_text.AmbiguousMember(
            f"'{identifier}' names members of {len(found)} groups: "
            + ", ".join(f[5] for f in found) + ". Pass the path of one."
        )
    return found[0] if found else None


def not_found(slide_index: int, identifier) -> ValueError:
    """The error for a name that is neither on the slide nor in a group on it."""
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]
    names = _shape_names(slide)
    groups = [name for _, name in _groups(slide, str(identifier))]
    where = f" or in its groups ({', '.join(groups)})" if groups else ""
    return ValueError(
        f"Shape '{identifier}' is not on slide {slide_index}{where}. "
        f"On the slide: {', '.join(names) if names else 'nothing'}."
    )


def edit_member(slide_index: int, identifier: str, tool_name: str,
                edit: Callable, check: Callable) -> Optional[dict]:
    """Rewrite one member of a group and paste the group back.

    ``edit(member)`` changes the member element in place and answers the
    fields to report, or raises ValueError, which is passed on with nothing
    changed. ``check(member)`` is handed the member as the pasted group reads
    back and answers what is missing from it, or None. Answers None when no
    group on the slide holds ``identifier``, so the caller's own not-found
    error stands; otherwise a result dict, a success or a refusal.
    """
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]

    clip = Clipboard.take()
    try:
        try:
            located = _locate(slide, identifier, clip, tool_name)
        except Refused as refused:
            return refused.payload
        if located is None:
            return None
        position, group_name, package, root, group, path, member = located
        if _gvml_text.text_body(member) is None:
            raise ValueError(
                f"'{path}' is a member of '{group_name}' that takes no text."
            )
        fields = edit(member)
        _gvml_shapes.strip_creation_ids(group)
        package.parts[package.drawing_part()] = _canvas.serialize_drawing(root).encode("utf-8")

        def verify(readback):
            try:
                back = _gvml_shapes.only_child(
                    _canvas.children(_canvas.parse(readback.drawing())), "group")
                _, back_member = _gvml_text.find_member(back, group_name, path)
            except PackageError as exc:
                return f"the member could not be found in the pasted group ({exc})"
            return check(back_member)

        group_shape = slide.shapes[position]
        left, top = group_shape.left_position(), group_shape.top()
        try:
            replaced = replace_shape(
                pres, slide, slide_index, position, package.to_bytes(), clip, tool_name,
                MsoShapeType[msoGroup], left, top, "group", verify, _ALTERNATIVES,
            )
        except Refused as refused:
            return refused.payload
    finally:
        clip.restore()

    body = {"slide_index": slide_index, "shape_name": path.rsplit(_gvml_text.SEPARATOR, 1)[-1],
            "path": path, "group_name": replaced.name}
    body.update(fields)
    return with_warnings(body, clip, replaced.warnings)
