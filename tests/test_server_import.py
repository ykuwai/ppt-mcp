"""Import smoke tests for the MCP server entry point.

Every other test module imports individual tool modules but never `server.py`
itself, so a dependency that renames or drops a symbol `server.py` imports goes
undetected — mcp 2.0 removing `mcp.server.fastmcp` (issue #176) broke the server
at module load for every user before any test noticed.

Importing the server does not launch PowerPoint: the COM connection is
established lazily on the first tool call (issue #148).
"""

from __future__ import annotations

import asyncio
import json
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


@pytest.fixture(scope="module")
def server():
    pytest.importorskip("win32com", reason="pywin32 required to import the server")
    import src.server as server_module

    return server_module


def test_server_module_imports(server):
    """server.py loads and exposes a server instance and an entry point."""
    assert server.mcp is not None
    assert callable(server.main)


def test_tools_are_registered(server):
    """All tools register successfully under whichever mcp major is installed."""
    tools = asyncio.run(server.mcp.list_tools())
    assert len(tools) > 100
    assert all(t.name.startswith("ppt_") for t in tools)


def test_annotations_serialize_with_protocol_field_names(server):
    """Annotations must reach the wire in camelCase, as MCP specifies.

    Tools declare annotations as plain dicts (`{"readOnlyHint": ...}`). Both mcp
    majors coerce those into a `ToolAnnotations` model whose fields are
    snake_case, so the protocol names survive only via serialization aliases.
    """
    tools = asyncio.run(server.mcp.list_tools())
    annotated = [t for t in tools if t.annotations is not None]
    assert annotated, "expected tools to carry annotations"

    dumped = annotated[0].annotations.model_dump(by_alias=True, exclude_none=True)
    assert "readOnlyHint" in dumped
    assert "read_only_hint" not in dumped


def _input_schema(tool) -> dict:
    """Wire schema of a tool under either mcp major (`input_schema` / `inputSchema`)."""
    schema = getattr(tool, "input_schema", None)
    if schema is None:
        schema = getattr(tool, "inputSchema", None)
    return schema or {}


def test_tool_schemas_are_self_contained(server):
    """Wire schemas carry no `$ref`/`$defs`; nested `params` models are inlined.

    Clients such as Claude Desktop do not dereference `$ref`, so a tool whose
    `params` model was only referenced showed up with an empty `params` object
    and callers had to guess the field names from validation errors.
    """
    tools = asyncio.run(server.mcp.list_tools())
    schemas = {t.name: _input_schema(t) for t in tools}

    with_refs = sorted(
        name for name, schema in schemas.items()
        if "$ref" in json.dumps(schema) or "$defs" in schema
    )
    assert with_refs == [], f"unresolved $ref/$defs in: {with_refs[:10]}"

    # Every `params` must be an inlined object schema with its fields visible.
    # (A model without fields legitimately inlines to `"properties": {}`.)
    opaque_params = sorted(
        name for name, schema in schemas.items()
        if "params" in schema.get("properties", {})
        and (
            schema["properties"]["params"].get("type") != "object"
            or "properties" not in schema["properties"]["params"]
        )
    )
    assert opaque_params == [], f"params not inlined in: {opaque_params[:10]}"

    params = schemas["ppt_add_section"]["properties"]["params"]
    assert params["properties"]["slide_index"]["minimum"] == 1
    assert "slide_index" in params["required"]
