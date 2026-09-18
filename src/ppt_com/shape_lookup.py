"""Finding a shape on a slide, including the ones inside a group.

Thirteen modules each carried their own copy of this lookup, and every copy
saw only the top level. ppt_get_group_items would hand a caller the name of a
shape inside a group and no other tool would accept it, which reads as a bug
the first time it happens. The copies now all come here.

A group's children are reachable by their own name when that name is the only
one like it on the slide, which PowerPoint's generated names are in practice.
When it is not, `"Group 20/Rounded Rectangle 22"` names one exactly.

A path is matched one step at a time, against the direct members of the group
named so far. A stale `"Outer/Deep"` for a shape that is really at
`"Outer/Inner/Deep"` finds nothing rather than quietly editing the shape two
levels down. And because a person may well call a shape `"A/B"`, a segment is
allowed to contain the separator: every way of reading the string is tried,
so `"G/A/B"` finds the child `"A/B"` of `"G"` when that is what exists.

COM only. macOS keeps its own resolver, because a group there answers no
members over Apple Events at all and the honest thing is to say so.
"""

from ppt_com.constants import msoGroup

PATH_SEPARATOR = "/"


class ShapeNotFound(ValueError):
    """No shape of that name, at the top level or in any group."""


class AmbiguousShape(ValueError):
    """The bare name belongs to more than one child. A path says which."""


def _direct_children(shape):
    """The members of a group, one level down. Empty for anything else."""
    try:
        if shape.Type != msoGroup:
            return []
        items = shape.GroupItems
        count = items.Count
    except Exception:
        return []

    children = []
    for i in range(1, count + 1):
        try:
            children.append(items(i))
        except Exception:
            continue
    return children


def _top_level_shapes(slide):
    shapes = []
    for i in range(1, slide.Shapes.Count + 1):
        shapes.append(slide.Shapes(i))
    return shapes


def walk_group_children(shape, prefix=""):
    """Yield (child, path) for everything inside a group, depth first.

    `path` is the child's name with every group it sits in in front of it,
    which is the string resolve_shape takes back.
    """
    for child in _direct_children(shape):
        try:
            name = child.Name
        except Exception:
            continue
        path = prefix + name
        yield child, path
        yield from walk_group_children(child, path + PATH_SEPARATOR)


def _top_level(slide, name):
    for shape in _top_level_shapes(slide):
        if shape.Name == name:
            return shape
    return None


def _children(slide, name):
    """Every shape inside any group on the slide whose name is `name`."""
    found = []
    for parent in _top_level_shapes(slide):
        prefix = ""
        try:
            prefix = parent.Name + PATH_SEPARATOR
        except Exception:
            pass
        for child, path in walk_group_children(parent, prefix):
            if child.Name == name:
                found.append((child, path))
    return found


def _match_path(candidates, name):
    """Read `name` as a path over `candidates` and their members.

    Returns (shape, [segment names]) or None. A whole segment is matched
    before the string is split any further, so a shape called "A/B" is found
    as itself rather than read as a step into "A".
    """
    for shape in candidates:
        try:
            own = shape.Name
        except Exception:
            continue
        if own == name:
            return shape, [own]

    for shape in candidates:
        try:
            own = shape.Name
        except Exception:
            continue
        head = own + PATH_SEPARATOR
        if not name.startswith(head):
            continue
        found = _match_path(_direct_children(shape), name[len(head):])
        if found is not None:
            deeper, segments = found
            return deeper, [own] + segments
    return None


def group_names(slide):
    """The groups on the slide, for an error that says where to look."""
    names = []
    for shape in _top_level_shapes(slide):
        try:
            if shape.Type == msoGroup:
                names.append(shape.Name)
        except Exception:
            pass
    return names


def resolve_with_path(slide, name_or_index):
    """Find a shape, and say where it turned out to live.

    Returns (shape, path). The path is the name for a shape at the top level
    and "Group/Child" for one inside a group, which is the string that reaches
    it again.
    """
    if isinstance(name_or_index, int):
        if name_or_index < 1 or name_or_index > slide.Shapes.Count:
            raise ShapeNotFound(
                f"Shape index {name_or_index} out of range "
                f"(1-{slide.Shapes.Count})"
            )
        shape = slide.Shapes(name_or_index)
        return shape, shape.Name

    name = name_or_index

    shape = _top_level(slide, name)
    if shape is not None:
        return shape, name

    matches = _children(slide, name)
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        paths = ", ".join(sorted(path for _, path in matches))
        raise AmbiguousShape(
            f"Shape '{name}' is inside more than one group on this slide: "
            f"{paths}. Pass one of those paths instead of the bare name."
        )

    if PATH_SEPARATOR in name:
        found = _match_path(_top_level_shapes(slide), name)
        if found is not None:
            shape, segments = found
            return shape, PATH_SEPARATOR.join(segments)

    message = f"Shape '{name}' not found on slide"
    groups = group_names(slide)
    if groups:
        message += (
            ". Shapes inside a group are reachable by their own name, or as "
            f"'{groups[0]}{PATH_SEPARATOR}<child name>'; ppt_get_group_items "
            "lists what is in " + ", ".join(f"'{g}'" for g in groups)
        )
    raise ShapeNotFound(message)


def resolve_shape(slide, name_or_index):
    """Find a shape by name (str) or 1-based index (int).

    A name is looked for at the top level first, then inside the groups, then
    read as a Group/Child path.
    """
    return resolve_with_path(slide, name_or_index)[0]


def require_top_level(slide, names, slide_index, tool_name):
    """Check every name is a shape at the top level, and say why when not.

    COM's `Shapes.Range` only addresses the top level, so aligning,
    distributing, merging and grouping genuinely cannot take a child of a
    group. They used to report that the shape does not exist, which is not
    what happened and sends the caller looking in the wrong place.
    """
    for name in names:
        if _top_level(slide, name) is not None:
            continue
        try:
            _, path = resolve_with_path(slide, name)
        except ShapeNotFound:
            raise ValueError(
                f"Shape '{name}' not found on slide {slide_index}"
            ) from None
        raise ValueError(
            f"Shape '{name}' is on slide {slide_index} but inside group "
            f"'{path.split(PATH_SEPARATOR)[0]}'. {tool_name} works on shapes "
            "at the top level only, because that is all PowerPoint will take "
            "a range of. Ungroup it first with ppt_ungroup_shapes."
        )
