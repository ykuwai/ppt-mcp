"""Slide comment tools for PowerPoint COM automation.

Handles adding, listing, and deleting comments on slides.
"""

import json
import logging

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class AddCommentInput(BaseModel):
    """Input for adding a comment to a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    text: str = Field(..., description="Comment text")
    author: str = Field(
        default="AI Agent", description="Comment author name"
    )
    author_initials: str = Field(
        default="AI", description="Comment author initials"
    )
    left: float = Field(
        default=0, description="Horizontal position in points"
    )
    top: float = Field(
        default=0, description="Vertical position in points"
    )


class ListCommentsInput(BaseModel):
    """Input for listing comments on a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")


class DeleteCommentInput(BaseModel):
    """Input for deleting a comment from a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

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
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    # Try Add2 first (newer PowerPoint versions), fall back to Add
    try:
        slide.Comments.Add2(left, top, author, author_initials, text, "AD", "")
    except Exception:
        try:
            slide.Comments.Add(left, top, author, author_initials, text)
        except Exception as e:
            raise RuntimeError(f"Failed to add comment: {e}")

    return {
        "success": True,
        "text": text,
        "author": author,
    }


def _list_comments_impl(slide_index) -> dict:
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
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
    pres = app.ActivePresentation
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
        return add_comment(params)

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
        return list_comments(params)

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
        return delete_comment(params)
