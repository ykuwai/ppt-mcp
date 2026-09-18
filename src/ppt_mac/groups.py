"""Shape grouping tools, on Apple Events and the clipboard.

Mirrors ``ppt_com/groups.py``. Same function names, same signatures, same
returned shapes.

Three things about groups on this side are worth knowing before reading on.

**Nothing can be gathered into a shape range, so ``group`` cannot be sent.**
``group`` takes a shape range, and the only shape range PowerPoint for Mac
hands out is the one already selected in a window. Its dictionary has an
``unselect`` command and no ``select``. So the group is not made by PowerPoint
at all. Each member is copied with ``copy shape``, which puts a DrawingML
package on the clipboard, the members are wrapped in one ``a:grpSp`` and
pasted back as a single group, and only once the group is on the slide and
verified are the originals deleted. See ``ppt_mac/gvml_paste.py`` for the
procedure and ``docs/gvml-design.md`` for the measurements behind it.

**A group's members cannot be read through the dictionary.** ``shape`` has a
``shape`` element, which reads as though a group's members are reached the
way a slide's shapes are. They are not. A group of two text boxes answers 0
for its ``shapes`` and -1728 for ``shapes[1]``. The same clipboard package
carries them, though, with names, positions and sizes, and that is where
``ppt_get_group_items`` reads them from.

**Ungrouping works through the dictionary.** ``ungroup`` is declared to take a
shape range but accepts a single group shape. Verified live, on a group of two
that came apart with both members appearing on the slide by name. The check is
the group's own disappearance and the slide growing.
"""

import logging

from appscript.reference import CommandError

from backend.mac_ae import ppt
from backend.mac_enums import MsoShapeType
from backend.unsupported import refusal as _refusal
from gvml import PackageError, build
from gvml import canvas as _canvas
from gvml import shapes as _gvml_shapes
from gvml.package import Graft
from ppt_com.constants import msoGroup
from ppt_mac.gvml_paste import (
    Clipboard,
    Refused,
    copy_shape_package,
    paste_package,
    unused_name,
    with_warnings,
)
from ppt_mac.shapes import (
    _WIN_SHAPE_TYPE,
    _get_shape,
    _shape_index,
    _shape_names,
    _slide,
    _win_constant,
)
from utils.navigation import goto_slide

logger = logging.getLogger(__name__)

_GROUP_ALTERNATIVES = [
    "Group the shapes by hand in PowerPoint",
    "ppt_align_shapes",
    "ppt_distribute_shapes",
]


def _require_group(shape, verb: str):
    """Raise unless a shape really is a group, and say what it is instead.

    The same guard Windows puts in front of both group tools, reading the type
    through the generated table so the number in the message is the Windows one
    a caller already knows.
    """
    type_val = _win_constant(_WIN_SHAPE_TYPE, shape.shape_type())
    if type_val != msoGroup:
        raise ValueError(
            f"Shape '{shape.name()}' is not a group (type={type_val}). "
            f"Only group shapes (type={msoGroup}) can be {verb}."
        )


# ---------------------------------------------------------------------------
# Apple Event implementation functions
# ---------------------------------------------------------------------------
def _group_shapes_impl(slide_index, shape_names):
    ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)

    # Every name is checked before anything is copied, the same message as
    # Windows, so a misspelling costs nothing.
    names_on_slide = _shape_names(slide)
    for name in shape_names:
        if name not in names_on_slide:
            raise ValueError(f"Shape '{name}' not found on slide {slide_index}")

    clip = Clipboard.take()
    try:
        # Each member comes off the slide as its own package. Anything a
        # member refers to (a picture's image, a chart's XML) is carried into
        # the group's package by the graft, with the ids rewritten to match.
        graft = Graft.empty()
        infos = []
        children_xml = []
        for name in shape_names:
            member = _get_shape(slide, name)
            package = copy_shape_package(member, clip, "ppt_group_shapes")
            try:
                element = _gvml_shapes.only_child(
                    _canvas.children(_canvas.parse(package.drawing())), f"shape '{name}'"
                )
                element = graft.take(package, _gvml_shapes.strip_creation_ids(element))
                infos.append(_gvml_shapes.describe(element))
            except PackageError as exc:
                return _refusal(
                    "ppt_group_shapes",
                    f"The package PowerPoint wrote for '{name}' could not be "
                    f"read as one shape: {exc}. Nothing was changed.",
                    _GROUP_ALTERNATIVES,
                )
            children_xml.append(_canvas.serialize(element))

        x, y, cx, cy = _gvml_shapes.bounding_box(infos)
        group_name = unused_name(names_on_slide, "Group")
        drawing = _canvas.wrap(
            _gvml_shapes.group_xml(2, group_name, x, y, cx, cy, children_xml), x, y, cx, cy,
        )
        raw = build(
            drawing,
            parts=graft.parts,
            drawing_rels=graft.drawing_rels,
            overrides=graft.overrides,
            defaults=graft.defaults,
            nested_rels=graft.nested_rels,
        )

        # Paste, verify, and only then delete the originals. The other order
        # leaves nothing behind when the paste is silently dropped.
        try:
            pasted = paste_package(
                pres, slide, slide_index, raw, clip, "ppt_group_shapes",
                MsoShapeType[msoGroup], _canvas.pt(x), _canvas.pt(y),
                _GROUP_ALTERNATIVES,
            )
        except Refused as refused:
            return refused.payload

        left_behind = []
        for name in shape_names:
            try:
                slide.shapes[_shape_index(slide, name)].delete()
            except (CommandError, ValueError) as exc:
                logger.warning("Could not delete '%s' after grouping: %s", name, exc)
            if name in _shape_names(slide):
                left_behind.append(name)
        if left_behind:
            return _refusal(
                "ppt_group_shapes",
                f"The group '{pasted.name}' was pasted onto slide {slide_index} "
                f"and verified, but the original shape(s) {left_behind} could "
                "not be deleted afterwards, so both are on the slide now. "
                "Delete one or the other with ppt_delete_shape.",
                ["ppt_delete_shape"],
                error="ppt_group_shapes left both the group and its originals on the slide",
            )
    finally:
        clip.restore()

    group = _get_shape(slide, pasted.name)
    return with_warnings({
        "success": True,
        "group_name": pasted.name,
        "shape_index": group.z_order_position(),
    }, clip, pasted.warnings)


def _ungroup_shapes_impl(slide_index, shape_name_or_index):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, shape_name_or_index)
    _require_group(shape, "ungrouped")

    # The members cannot be listed first, so the before picture is the slide.
    group_name = shape.name()
    before = _shape_names(slide)

    failure = None
    try:
        shape.ungroup()
    except CommandError as exc:
        # Held rather than raised. PowerPoint hands back a shape range here and
        # a reply appscript cannot unpack looks exactly like a command that was
        # refused, so the slide decides which happened, not the exception.
        failure = exc

    # A group that came apart is gone from the slide, and its members stand
    # where it stood. The group's own disappearance is the whole test. Counting
    # is not, because PowerPoint will let a group hold one shape and that one
    # comes back out leaving the slide exactly as long as it was.
    after = _shape_names(slide)
    came_apart = group_name not in after
    if not came_apart:
        return _refusal(
            "ppt_ungroup_shapes",
            f"PowerPoint reported no error and '{group_name}' is still on "
            f"slide {slide_index} as one shape, so nothing came apart. This is "
            "the silent no-op recorded in MACOS_PORT section 5."
            + (f" PowerPoint answered {failure}." if failure else ""),
            ["Ungroup the shape by hand in PowerPoint"],
        )

    members = [name for name in after if name not in before]

    # PowerPoint renames the members on the way out, and every tool here tells
    # callers to address shapes by name rather than index because indices
    # shift. So a caller who noted the names before ungrouping finds none of
    # them afterwards, and the ones it hands back are the only ones that work.
    # Pictures move furthest: one went from `Picture 17` to `Picture 42` to
    # `Picture 54` across two rounds, and on a Japanese system text boxes come
    # back as `テキスト ボックス 18`.
    return {
        "success": True,
        "ungrouped_count": len(members),
        "shape_names": members,
        "warnings": [
            "PowerPoint renamed the members as it ungrouped them, so any name "
            "noted before this call no longer matches anything. Use the names "
            "in shape_names. Grouping them again renames them once more."
        ],
    }


def _get_group_items_impl(slide_index, shape_name_or_index):
    """Read a group's members out of the package ``copy shape`` writes.

    The dictionary reads as though ``shape.shapes`` works on a group, and
    against a real group every route answers nothing: ``shapes`` counts 0,
    ``shapes[1]`` answers -1728, ``has child`` answers ``missing value``. The
    clipboard package carries every member with its name, type and box, so
    the group is copied and the package is read. Nothing on the slide changes;
    the clipboard is put back afterwards.
    """
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, shape_name_or_index)
    _require_group(shape, "inspected")
    group_name = shape.name()

    clip = Clipboard.take()
    try:
        try:
            package = copy_shape_package(shape, clip, "ppt_get_group_items")
        except Refused as refused:
            return refused.payload
        try:
            group = _gvml_shapes.only_child(
                _canvas.children(_canvas.parse(package.drawing())), "group"
            )
            items = _gvml_shapes.group_items(group)
        except PackageError as exc:
            return _refusal(
                "ppt_get_group_items",
                f"'{group_name}' is a group, but the package PowerPoint wrote "
                f"for it could not be read as one: {exc}",
                ["ppt_ungroup_shapes", "ppt_list_shapes"],
            )
    finally:
        clip.restore()

    # Windows reports the "Group/Child" string that reaches a member from
    # every other tool. Nothing reaches a group's member here, so there is no
    # such string and saying None is the honest version of the same key.
    items = [dict(item, path=None) for item in items]

    return with_warnings({
        "success": True,
        "group_name": group_name,
        "items": items,
    }, clip)
