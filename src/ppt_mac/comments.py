"""Slide comment tools, on Apple Events.

Mirrors ``ppt_com/comments.py``. Same function names, same signatures, same
returned shapes; what differs is the walk through PowerPoint's object model.

Three things about comments on this side are worth knowing before reading on.

**A comment is a shape, and nothing else.** The whole declaration is
``<class name="comment" code="cD09" inherits="shape" plural="comments"/>``, with
no properties of its own at all. So a comment has a name, a position, a z order
and a text frame, and it has no author, no initials and no date. Everything the
Windows tools read out of ``Comment.Author``, ``Comment.AuthorInitials`` and
``Comment.DateTime`` is unreachable here, and the honest answer for those three
is nothing rather than a plausible guess.

**A slide has no comments collection.** ``comment`` is a declared element of
``shape`` and of nothing else, so ``slide.comments`` is not in the dictionary.
Comments are found by walking the slide's shapes through ``shapes_of`` and
keeping the ones whose ``shape type`` is ``shape type comment``, which is the
same trusted positional route every other module uses. It also settles the
ordering question, because ``comment_index`` is then a position among the
comment shapes in z order, which is the order ``ppt_list_comments`` reports.
The two agree with each other even where they disagree with the numbering
PowerPoint prints in its own comment pane.

**Adding one is attempted rather than assumed.** ``make new comment`` has no
declared location to go to, since ``slide`` does not hold comments, so the two
documented outcomes are an error or a stray autoshape left on the slide. Both
are handled. The shapes are counted before and after, a shape that arrived and
is not a comment is deleted again, and anything else comes back as a refusal
that says what happened.
"""

import logging

from appscript import k
from appscript.reference import CommandError

from backend.mac_ae import count, error_number, is_missing, ppt, shapes_of
from backend.unsupported import refusal as _refusal
from utils.navigation import goto_slide

logger = logging.getLogger(__name__)

# Ink comments are a different shape type carrying a drawing rather than text,
# and none of the Windows tools can read or write one, so they are left out
# rather than listed as comments nobody can act on.
_COMMENT_SHAPE_TYPE = k.shape_type_comment

# Said the same way wherever author and initials are dropped, because the cause
# is the same one line of the dictionary every time.
_NO_AUTHOR = (
    "PowerPoint for Mac's `comment` class inherits everything it has from "
    "`shape` and declares no properties of its own, so it carries no author, "
    "no initials and no date."
)


def _is_comment(shape) -> bool:
    """True when a shape is a comment.

    A shape whose type cannot be read is not a comment for these purposes.
    Guessing the other way would put a shape nobody can identify into the
    comment list and, worse, into reach of ppt_delete_comment.
    """
    try:
        return shape.shape_type() == _COMMENT_SHAPE_TYPE
    except CommandError:
        return False


def _comment_shapes(slide, shapes=None) -> list:
    """The comment shapes on a slide, in z order.

    ``shapes`` is there so a caller that has already paid for ``shapes_of`` can
    filter the same list again without a second walk.
    """
    return [shape for shape in (shapes if shapes is not None else shapes_of(slide))
            if _is_comment(shape)]


def _comment_text(shape) -> str:
    """Read a comment's text out of its text frame.

    The text frame is the only part of a comment that holds anything a reader
    wants, and a comment with none reads as empty rather than as a failure.
    """
    try:
        if not shape.has_text_frame():
            return ""
        content = shape.text_frame.text_range.content()
    except CommandError:
        return ""
    return "" if is_missing(content) else content


def _rounded(reference):
    """Read a coordinate, or None where PowerPoint will not answer for one."""
    try:
        value = reference()
    except CommandError:
        return None
    return None if is_missing(value) else round(value, 2)


# ---------------------------------------------------------------------------
# Apple Event implementation functions
# ---------------------------------------------------------------------------
def _add_comment_impl(slide_index, text, author, author_initials, left, top) -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]

    before_shapes = shapes_of(slide)
    before_comments = len(_comment_shapes(slide, before_shapes))

    # No properties are handed to `make`. Position is set afterwards instead,
    # so that a property PowerPoint dislikes cannot be what turns a comment
    # that would have been created into an error.
    try:
        app.make(new=k.comment, at=slide.end)
    except CommandError as exc:
        logger.warning("Making a comment was refused", exc_info=True)
        return _refusal(
            "ppt_add_comment",
            "PowerPoint refused to make a comment (Apple Event error "
            f"{error_number(exc)}). `comment` is a declared element of "
            "`shape` and of nothing else, so a slide is not a place a comment "
            "can be made, and there is no add comment command in the "
            "dictionary either.",
            ["ppt_set_slide_notes", "ppt_add_textbox"],
        )

    # Nothing is trusted because it did not raise.
    after_shapes = shapes_of(slide)
    after_comments = _comment_shapes(slide, after_shapes)

    if len(after_comments) != before_comments + 1:
        stray = len(after_shapes) - len(before_shapes)
        if stray > 0:
            # A `make` that falls through leaves an empty autoshape behind, as
            # MACOS_PORT section 5 records. It is removed again rather than
            # left on the slide for the user to find.
            return _refusal(
                "ppt_add_comment",
                _remove_stray(slide, after_shapes[-1], len(before_shapes)),
                ["ppt_set_slide_notes", "ppt_add_textbox"],
            )
        return _refusal(
            "ppt_add_comment",
            "PowerPoint reported success but the slide gained no comment, "
            "which is the silent no-op recorded in MACOS_PORT section 5. "
            "Nothing was added and nothing was left behind.",
            ["ppt_set_slide_notes", "ppt_add_textbox"],
        )

    # `make`'s own return value is not used. PowerPoint hands back a reference
    # that does not resolve, so the new comment is fetched again by the route
    # that found the others.
    comment = after_comments[-1]

    warnings = [f"author and author_initials were ignored. {_NO_AUTHOR}"]

    # PowerPoint's own line separator, the same substitution the table and text
    # tools make, so a multi line comment does not arrive as one run-on line.
    wanted = text.replace("\n", "\r")
    try:
        comment.text_frame.text_range.content.set(wanted)
    except CommandError:
        logger.warning("Could not write the comment's text", exc_info=True)

    try:
        comment.left_position.set(left)
        comment.top.set(top)
    except CommandError:
        warnings.append(
            "PowerPoint would not move the comment, so it sits where it was "
            "placed rather than at the requested left and top."
        )

    written = _comment_text(comment)
    if written != wanted:
        warnings.append(
            "The comment was created but its text did not stay. It reads as "
            f"{written!r} rather than as it was given."
        )

    return {
        "success": True,
        # Read back rather than echoed, so a write that did not land shows up.
        "text": written,
        # The key stays so the answer reads the same on both platforms, and the
        # value is nothing because nothing is what macOS can store.
        "author": None,
        "warnings": warnings,
    }


def _remove_stray(slide, shape, expected: int) -> str:
    """Delete the autoshape a fallen through `make` left behind, and report.

    Returns the sentence the refusal carries, which differs depending on
    whether the slide is back to ``expected`` shapes afterwards. A shape nobody
    can remove is worth naming precisely, since the user has to find it by hand.
    """
    name = "the new shape"
    try:
        name = shape.name()
    except CommandError:
        logger.warning("The shape the make left behind has no readable name", exc_info=True)
    try:
        shape.delete()
    except CommandError:
        logger.warning("Could not remove the shape the make left behind", exc_info=True)

    if count(slide.shapes) == expected:
        return (
            "PowerPoint made a shape rather than a comment, which is the "
            "fallen through `make` recorded in MACOS_PORT section 5. It was "
            "deleted again, so the slide is as it was."
        )
    return (
        "PowerPoint made a shape rather than a comment, which is the fallen "
        "through `make` recorded in MACOS_PORT section 5, and it could not be "
        f"deleted again. Remove the shape named {name!r} by hand."
    )


def _list_comments_impl(slide_index) -> dict:
    ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]

    comments = []
    for index, shape in enumerate(_comment_shapes(slide), 1):
        comments.append({
            "index": index,
            # Nothing to read these from. See _NO_AUTHOR.
            "author": None,
            "author_initials": None,
            "text": _comment_text(shape),
            # The Windows side answers with an empty string when the date will
            # not read, so the same empty string is used rather than a null
            # nobody else produces.
            "datetime": "",
            "left": _rounded(shape.left_position),
            "top": _rounded(shape.top),
        })

    return {
        "slide_index": slide_index,
        "comments_count": len(comments),
        "comments": comments,
        "note": (
            f"Author, initials and date are always empty here. {_NO_AUTHOR} "
            "The text and the position are read from the comment's shape."
        ),
    }


def _delete_comment_impl(slide_index, comment_index) -> dict:
    # Nothing here can be settled before goto_slide. Both answers below need
    # the slide's comments counted first, and reading them is itself the round
    # trip that moving the view would have saved.
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]

    comments = _comment_shapes(slide)
    if comment_index < 1 or comment_index > len(comments):
        raise ValueError(
            f"Comment index {comment_index} out of range (1-{len(comments)})"
        )

    before = count(slide.shapes)
    comments[comment_index - 1].delete()

    if count(slide.shapes) != before - 1:
        return _refusal(
            "ppt_delete_comment",
            "PowerPoint reported success but the slide still holds the same "
            "number of shapes, so the comment is still there. Deleting a "
            "comment is not scriptable here.",
        )

    return {
        "success": True,
    }
