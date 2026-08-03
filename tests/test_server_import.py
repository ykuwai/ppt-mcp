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
