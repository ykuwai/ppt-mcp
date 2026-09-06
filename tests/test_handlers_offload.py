"""Every tool handler must hand its blocking work to a thread (issue #198).

`ppt.execute()` blocks until the COM worker answers. Called straight from an
`async def` handler it stops the whole event loop, so the server cannot answer
a ping, a cancellation or even `tools/list` while one COM call is in flight.
From the client's side that is indistinguishable from a hang.

`utils.offload.run_offloaded` keeps the loop free and the call cancellable.
This module is the guard: it fails the moment a new tool is added the old way.

The checks look for an awaited call to `run_offloaded`, not merely for the name
appearing somewhere in the function, so a handler that references the helper
without awaiting it does not slip through.
"""

from __future__ import annotations

import ast
import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
_src_dir = str(_root / "src")
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)

SRC = _root / "src"


def _tool_handlers():
    """Yield (path, ast node) for every `async def tool_*` under src/."""
    for path in sorted(SRC.rglob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if isinstance(node, ast.AsyncFunctionDef) and node.name.startswith("tool"):
                yield path, node


def _awaits_run_offloaded(node) -> bool:
    """True if the handler awaits a call to run_offloaded(...)."""
    for sub in ast.walk(node):
        if not isinstance(sub, ast.Await):
            continue
        call = sub.value
        if isinstance(call, ast.Call) and isinstance(call.func, ast.Name):
            if call.func.id == "run_offloaded":
                return True
    return False


def _calls_to(node, obj, attr):
    """Every Call node in `node` of the form `obj.attr(...)`."""
    for sub in ast.walk(node):
        if (isinstance(sub, ast.Call)
                and isinstance(sub.func, ast.Attribute)
                and sub.func.attr == attr
                and isinstance(sub.func.value, ast.Name)
                and sub.func.value.id == obj):
            yield sub


def test_every_handler_offloads_to_a_thread():
    handlers = list(_tool_handlers())
    assert len(handlers) > 100, "handler discovery is broken, not the handlers"

    offenders = [
        f"{path.relative_to(_root)}:{node.lineno} {node.name}"
        for path, node in handlers
        if not _awaits_run_offloaded(node)
    ]
    assert offenders == [], (
        "these handlers run their work on the event loop; wrap the call in "
        f"await run_offloaded(...): {offenders}"
    )


def test_no_handler_calls_ppt_execute_directly():
    """`ppt.execute` blocks. Inside a handler it must go through a thread,
    which means it appears as an argument to run_offloaded, never as the call."""
    offenders = [
        f"{path.relative_to(_root)}:{call.lineno} {node.name}"
        for path, node in _tool_handlers()
        for call in _calls_to(node, "ppt", "execute")
    ]
    assert offenders == [], (
        "call await run_offloaded(ppt.execute, ...) instead of "
        f"ppt.execute(...): {offenders}"
    )


def test_modules_with_handlers_import_the_helper():
    modules = {path for path, _ in _tool_handlers()}
    missing = []
    for path in sorted(modules):
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        imported = {
            alias.name
            for node in ast.walk(tree)
            if isinstance(node, ast.ImportFrom) and node.module == "utils.offload"
            for alias in node.names
        }
        if "run_offloaded" not in imported:
            missing.append(str(path.relative_to(_root)))
    assert missing == [], (
        f"handlers use run_offloaded but the module never imports it: {missing}"
    )


def test_handlers_do_not_call_anyio_directly():
    """One helper owns the threading and the cancellation semantics. Calling
    anyio.to_thread from a handler would bypass the future cancellation."""
    offenders = []
    for path, node in _tool_handlers():
        for sub in ast.walk(node):
            if isinstance(sub, ast.Attribute) and sub.attr == "run_sync":
                offenders.append(f"{path.relative_to(_root)}:{sub.lineno} {node.name}")
    assert offenders == [], (
        f"use run_offloaded rather than anyio.to_thread.run_sync directly: {offenders}"
    )
