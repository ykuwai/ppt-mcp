"""Document property operations for PowerPoint COM automation.

Handles getting and setting built-in document properties such as
title, author, subject, keywords, comments, category, and company.
"""

import json
import logging
from typing import Optional

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt

logger = logging.getLogger(__name__)

# Property names are always English strings regardless of Office language
WRITABLE_PROPERTIES = [
    "Title", "Author", "Subject", "Keywords",
    "Comments", "Category", "Company",
]

READABLE_PROPERTIES = [
    "Title", "Author", "Subject", "Keywords",
    "Comments", "Category", "Company",
    "Last Author", "Creation Date", "Last Save Time",
]


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class SetPropertiesInput(BaseModel):
    """Input for setting document properties."""
    model_config = ConfigDict(str_strip_whitespace=True)

    title: Optional[str] = Field(default=None, description="Document title")
    author: Optional[str] = Field(default=None, description="Author name")
    subject: Optional[str] = Field(default=None, description="Document subject")
    keywords: Optional[str] = Field(default=None, description="Keywords (comma-separated)")
    comments: Optional[str] = Field(default=None, description="Document comments")
    category: Optional[str] = Field(default=None, description="Document category")
    company: Optional[str] = Field(default=None, description="Company name")


# ---------------------------------------------------------------------------
# COM implementation functions (run on COM thread via ppt.execute)
# ---------------------------------------------------------------------------
def _set_properties_impl(title, author, subject, keywords, comments, category, company):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    props = pres.BuiltInDocumentProperties

    # Map field names to COM property names
    field_map = {
        "Title": title,
        "Author": author,
        "Subject": subject,
        "Keywords": keywords,
        "Comments": comments,
        "Category": category,
        "Company": company,
    }

    properties_set = 0
    set_names = []
    for prop_name, value in field_map.items():
        if value is not None:
            props(prop_name).Value = value
            properties_set += 1
            set_names.append(prop_name)

    return {
        "success": True,
        "properties_set": properties_set,
        "set_names": set_names,
    }


def _get_properties_impl():
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    props = pres.BuiltInDocumentProperties

    result = {}
    for prop_name in READABLE_PROPERTIES:
        try:
            value = props(prop_name).Value
            # Convert COM date objects to string
            if hasattr(value, "strftime"):
                value = value.strftime("%Y-%m-%d %H:%M:%S")
            elif not isinstance(value, (str, int, float, bool)):
                value = str(value)
            result[prop_name] = value
        except Exception:
            result[prop_name] = None

    return {
        "success": True,
        "properties": result,
    }


# ---------------------------------------------------------------------------
# MCP tool functions (async wrappers that delegate to COM thread)
# ---------------------------------------------------------------------------
def set_properties(params: SetPropertiesInput) -> str:
    """Set built-in document properties.

    Args:
        params: Properties to set. Only provided (non-None) values are updated.

    Returns:
        JSON with count of properties set and their names.
    """
    try:
        result = ppt.execute(
            _set_properties_impl,
            params.title, params.author, params.subject,
            params.keywords, params.comments, params.category,
            params.company,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to set properties: {str(e)}"})


def get_properties() -> str:
    """Get built-in document properties.

    Returns:
        JSON with all readable document properties.
    """
    try:
        result = ppt.execute(_get_properties_impl)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to get properties: {str(e)}"})


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register all document property tools with the MCP server."""

    @mcp.tool(
        name="ppt_set_properties",
        annotations={
            "title": "Set Document Properties",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_set_properties(params: SetPropertiesInput) -> str:
        """Set built-in document properties.

        Updates title, author, subject, keywords, comments, category,
        and/or company. Only provided values are changed.
        """
        return set_properties(params)

    @mcp.tool(
        name="ppt_get_properties",
        annotations={
            "title": "Get Document Properties",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_get_properties() -> str:
        """Get built-in document properties.

        Returns title, author, subject, keywords, comments, category,
        company, last author, creation date, and last save time.
        Properties that have never been set return null.
        """
        return get_properties()
