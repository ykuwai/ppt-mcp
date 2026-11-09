"""Slide comment tools for PowerPoint COM automation.

Handles adding, listing, and deleting comments on slides.
"""

import json
import logging

from pydantic import BaseModel, Field, ConfigDict

from utils.offload import run_offloaded
from backend import ppt
from utils.navigation import goto_slide

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class AddCommentInput(BaseModel):
    """Input for adding a comment to a slide."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    text: str = Field(..., description="Comment text")
    author: str = Field(
        default="AI Agent",
        description="Comment author name. PowerPoint signs a modern comment as the signed-in Office user instead; "
                    "the answer says who signed it.",
    )
    author_initials: str = Field(
        default="AI", description="Comment author initials, as for author"
    )
    left: float = Field(
        default=0, description="Horizontal position in points"
    )
    top: float = Field(
        default=0, description="Vertical position in points"
    )


class ListCommentsInput(BaseModel):
    """Input for listing comments on a slide."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    slide_index: int = Field(..., ge=1, description="1-based slide index")


class DeleteCommentInput(BaseModel):
    """Input for deleting a comment from a slide."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    comment_index: int = Field(
        ..., ge=1, description="1-based comment index"
    )


# ---------------------------------------------------------------------------
# COM implementation functions
# ---------------------------------------------------------------------------
def _add_comment_impl(slide_index, text, author, author_initials,
                       left, top) -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    comments_col = slide.Comments
    before = comments_col.Count

    # Try Add2 first (newer PowerPoint versions), fall back to Add
    try:
        comments_col.Add2(left, top, author, author_initials, text, "AD", "")
    except Exception:
        try:
            comments_col.Add(left, top, author, author_initials, text)
        except Exception as e:
            raise RuntimeError(f"Failed to add comment: {e}")

    try:
        unchanged = comments_col.Count == before
    except Exception:  # noqa: BLE001 - not known, so not reported as not added
        unchanged = False
    if unchanged:
        raise RuntimeError(f"PowerPoint did not add the comment to slide {slide_index}. Nothing was changed.")

    result = {"success": True, "text": text, "author": author}
    try:
        _report_what_powerpoint_made(result, comments_col, before, text, author, author_initials, left, top)
    except Exception as exc:  # noqa: BLE001 - the comment is there: a retry would add a second
        logger.warning("Could not read back the comment just added: %s", exc)
        result["note"] = "The comment was added, but PowerPoint would not say how it signed or placed it."
    return result


def _report_what_powerpoint_made(result, comments_col, before, text, author, author_initials, left, top) -> None:
    """Put what PowerPoint made of a new comment in ``result``. A modern
    comment is signed by the Office user, whatever author it is given, and on
    the Mac it belongs to the slide as a whole rather than to a point on it
    (measured 2026-11-08)."""
    added = _added_comment(comments_col, before, text)
    if added is None:
        return
    signer, initials = added.Author, added.AuthorInitials
    placed_left, placed_top = round(added.Left, 2), round(added.Top, 2)
    result.update(author=signer, author_initials=initials, left=placed_left, top=placed_top)
    notes = []
    if (signer, initials) != (author, author_initials):
        result["requested_author"] = author
        notes.append(
            f"PowerPoint signed the comment as {signer!r} ({initials}), the "
            f"Office user, not {author!r}: modern comments take their author from the "
            "signed-in account. Put a name in the text if it has to show.")
    if abs(placed_left - left) > 1 or abs(placed_top - top) > 1:
        notes.append(
            f"PowerPoint placed it at {placed_left}, {placed_top}, not at {left}, {top}; "
            "a new comment can belong to the slide as a whole, as it always does on the Mac.")
    if notes:
        result["note"] = " ".join(notes)


def _same_text(held, asked) -> bool:
    """Whether a comment holds this text; PowerPoint keeps a line break as \\r."""
    return held.replace("\r", "\n") == asked.replace("\r", "\n")


def _added_comment(comments_col, before, text):
    """The comment just added: the last one with this text, once the count
    has grown by one, or None."""
    count = comments_col.Count
    if count != before + 1:
        return None
    for i in range(count, 0, -1):
        comment = comments_col(i)
        if _same_text(comment.Text, text):
            return comment
    return None


def _list_comments_impl(slide_index) -> dict:
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)

    comments_col = slide.Comments
    count = comments_col.Count
    comments = []

    for i in range(1, count + 1):
        comment = comments_col(i)
        try:
            dt_str = str(comment.DateTime)
        except Exception:
            dt_str = ""
        comments.append({
            "index": i,
            "author": comment.Author,
            "author_initials": comment.AuthorInitials,
            "text": comment.Text,
            "datetime": dt_str,
            "left": round(comment.Left, 2),
            "top": round(comment.Top, 2),
        })

    return {
        "slide_index": slide_index,
        "comments_count": count,
        "comments": comments,
    }


def _delete_comment_impl(slide_index, comment_index) -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)

    comments_col = slide.Comments
    if comment_index < 1 or comment_index > comments_col.Count:
        raise ValueError(
            f"Comment index {comment_index} out of range "
            f"(1-{comments_col.Count})"
        )

    comments_col(comment_index).Delete()

    return {
        "success": True,
    }


# ---------------------------------------------------------------------------
# MCP tool functions
# ---------------------------------------------------------------------------
def add_comment(params: AddCommentInput) -> str:
    """Add a comment to a slide."""
    try:
        result = ppt.execute(
            _add_comment_impl,
            params.slide_index, params.text, params.author,
            params.author_initials, params.left, params.top,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def list_comments(params: ListCommentsInput) -> str:
    """List all comments on a slide."""
    try:
        result = ppt.execute(
            _list_comments_impl,
            params.slide_index,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def delete_comment(params: DeleteCommentInput) -> str:
    """Delete a comment from a slide."""
    try:
        result = ppt.execute(
            _delete_comment_impl,
            params.slide_index, params.comment_index,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register all comment tools with the MCP server."""

    @mcp.tool(
        name="ppt_add_comment",
        annotations={
            "title": "Add Slide Comment",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_ppt_add_comment(params: AddCommentInput) -> str:
        """Add a comment to a slide.

        Provide text, author name, and optional position.
        Uses Add2 for modern PowerPoint, falls back to Add for older versions.
        """
        return await run_offloaded(add_comment, params)

    @mcp.tool(
        name="ppt_list_comments",
        annotations={
            "title": "List Slide Comments",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_ppt_list_comments(params: ListCommentsInput) -> str:
        """List all comments on a slide.

        Returns index, author, text, datetime, and position for each comment.
        """
        return await run_offloaded(list_comments, params)

    @mcp.tool(
        name="ppt_delete_comment",
        annotations={
            "title": "Delete Slide Comment",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_ppt_delete_comment(params: DeleteCommentInput) -> str:
        """Delete a comment from a slide by its 1-based index.

        Use ppt_list_comments to find the comment index first.
        """
        return await run_offloaded(delete_comment, params)


# ---------------------------------------------------------------------------
# macOS
# ---------------------------------------------------------------------------
# The implementations above walk COM. Their Apple Event counterparts have the
# same names and signatures, so on macOS they simply take their place; nothing
# else in this module changes.
from backend import IS_MACOS, use_mac_impls  # noqa: E402

if IS_MACOS:  # pragma: no cover - platform specific
    from ppt_mac import comments as _mac_comments

    use_mac_impls(globals(), _mac_comments)
