"""Finding a shape on a slide, including the ones inside a group.

Thirteen modules each carried their own copy of this lookup, and every copy
saw only the top level. ppt_get_group_items would hand a caller the name of a
shape inside a group and no other tool would accept it, which reads as a bug
the first time it happens. The copies now all come here.

A group's children are reachable by their own name when that name is the only
one like it on the slide, which PowerPoint's generated names are in practice.
When it is not, `"Group 20/Rounded Rectangle 22"` names one exactly. A shape
at the top level always wins over a child with the same name, because that is
what the caller who has not thought about groups meant.

COM only. macOS keeps its own resolver, because a group there answers no
members over Apple Events at all and the honest thing is to say so.
"""

from ppt_com.constants import msoGroup

PATH_SEPARATOR = "/"


def walk_group_children(shape, prefix=""):
    """Yield (child, path) for everything inside a group, depth first.

    `path` is the child's name with every group it sits in in front of it,
    which is the string resolve_shape takes back.
    """
    try:
        if shape.Type != msoGroup:
            return
        items = shape.GroupItems
        count = items.Count
    except Exception:
        return

    for i in range(1, count + 1):
        try:
            child = items(i)
            name = child.Name
        except Exception:
            continue
        path = prefix + name
        yield child, path
        yield from walk_group_children(child, path + PATH_SEPARATOR)


def _top_level(slide, name):
    for i in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes(i)
        if shape.Name == name:
            return shape
    return None


def _children(slide, name):
    """Every shape inside any group on the slide whose name is `name`."""
    found = []
    for i in range(1, slide.Shapes.Count + 1):
        parent = slide.Shapes(i)
        prefix = ""
        try:
            prefix = parent.Name + PATH_SEPARATOR
        except Exception:
            pass
        for child, path in walk_group_children(parent, prefix):
            if child.Name == name:
                found.append((child, path))
    return found


def _by_path(slide, path):
    """Walk a Group/Child path, one segment at a time."""
    segments = path.split(PATH_SEPARATOR)
    shape = _top_level(slide, segments[0])
    if shape is None:
        return None
    for segment in segments[1:]:
        match = None
        for child, _ in walk_group_children(shape):
            if child.Name == segment:
                match = child
                break
        if match is None:
            return None
        shape = match
    return shape


def group_names(slide):
    """The groups on the slide, for an error that says where to look."""
    names = []
    for i in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes(i)
        try:
            if shape.Type == msoGroup:
                names.append(shape.Name)
        except Exception:
            pass
    return names


def find_in_groups(slide, name):
    """The path of `name` inside a group, or None when it is not in one.

    For the tools that cannot work on a child and want to say why rather than
    report that the shape does not exist.
    """
    matches = _children(slide, name)
    return matches[0][1] if len(matches) == 1 else None


def resolve_shape(slide, name_or_index):
    """Find a shape by name (str) or 1-based index (int).

    A name is looked for at the top level first, then inside the groups, then
    read as a Group/Child path. Trying the whole string as a name before
    splitting it means a shape actually called "A/B" is found as itself.
    """
    if isinstance(name_or_index, int):
        if name_or_index < 1 or name_or_index > slide.Shapes.Count:
            raise ValueError(
                f"Shape index {name_or_index} out of range "
                f"(1-{slide.Shapes.Count})"
            )
        return slide.Shapes(name_or_index)

    name = name_or_index

    shape = _top_level(slide, name)
    if shape is not None:
        return shape

    matches = _children(slide, name)
    if len(matches) == 1:
        return matches[0][0]
    if len(matches) > 1 and PATH_SEPARATOR not in name:
        paths = ", ".join(sorted(path for _, path in matches))
        raise ValueError(
            f"Shape '{name}' is inside more than one group on this slide: "
            f"{paths}. Pass one of those paths instead of the bare name."
        )

    if PATH_SEPARATOR in name:
        shape = _by_path(slide, name)
        if shape is not None:
            return shape

    message = f"Shape '{name}' not found on slide"
    groups = group_names(slide)
    if groups:
        message += (
            ". Shapes inside a group are reachable by their own name, or as "
            f"'{groups[0]}{PATH_SEPARATOR}<child name>'; ppt_get_group_items "
            "lists what is in " + ", ".join(f"'{g}'" for g in groups)
        )
    raise ValueError(message)


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
        path = find_in_groups(slide, name)
        if path is None:
            raise ValueError(f"Shape '{name}' not found on slide {slide_index}")
        raise ValueError(
            f"Shape '{name}' is on slide {slide_index} but inside group "
            f"'{path.split(PATH_SEPARATOR)[0]}'. {tool_name} works on shapes "
            "at the top level only, because that is all PowerPoint will take "
            "a range of. Ungroup it first with ppt_ungroup_shapes."
        )
