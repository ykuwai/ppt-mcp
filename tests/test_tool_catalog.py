"""Catalog mode, the default: three tools listed instead of 157.

The mode is read when the server module loads, so each check runs in a fresh
interpreter with the variable set. Nothing here reaches PowerPoint: the calls
are refused by argument validation, which runs before any tool body.
"""

import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]

_PROBE = r'''
import asyncio, json, sys
sys.path.insert(0, ".")
import src.server as s

async def main():
    out = {}
    out["listed"] = [t.name for t in await s.mcp.list_tools()]
    out["registered"] = len(s.mcp._tool_manager.list_tools())
    out["instructions"] = s.mcp.instructions or ""

    async def text(name, args):
        try:
            res = await s.mcp.call_tool(name, args)
        except Exception as exc:
            return "ERROR: " + str(exc)
        try:
            blocks = s._content_of(res)
        except Exception as exc:
            return "ERROR: " + str(exc)
        return "".join(getattr(b, "text", "") or "" for b in blocks)

    out["catalog"] = await text("ppt_describe", {})
    out["described"] = await text("ppt_describe", {"names": ["ppt_add_textbox", "ppt_nope"]})
    out["unknown"] = await text("ppt_run", {"tool": "ppt_add_texbox", "params": {}})
    out["invalid"] = await text("ppt_run", {"tool": "ppt_add_textbox", "params": {
        "slide_index": 0, "left": 0, "top": 0, "width": 10, "height": 10}})
    out["steps"] = await text("ppt_run_steps", {"steps": [
        {"tool": "ppt_add_textbox", "params": {"slide_index": 1, "left": 0, "top": 0,
                                                "width": 10, "height": 10, "wrong": 1}},
        {"tool": "ppt_delete_shape", "params": {"slide_index": 1, "shape_name": "$prev.shape_name"}}]})
    out["refs"] = s._resolve_refs({"a": "$prev.shape_name", "b": ["$1.x.0"], "c": "$keep"},
                                  [{"x": [7]}, {"shape_name": "TextBox 2"}])
    print(json.dumps(out))

asyncio.run(main())
'''


def _probe(mode):
    try:
        import win32com  # noqa: F401
    except ImportError:
        if sys.platform != "darwin":
            pytest.skip("the server needs pywin32 on Windows")
    env = dict(os.environ)
    env.pop("PPT_TOOL_MODE", None)
    if mode is not None:
        env["PPT_TOOL_MODE"] = mode
    done = subprocess.run([sys.executable, "-c", _PROBE], cwd=ROOT, env=env, check=False,
                          capture_output=True, text=True, timeout=180)
    assert done.returncode == 0, done.stderr[-2000:]
    return json.loads(done.stdout.strip().splitlines()[-1])


@pytest.fixture(scope="module")
def catalog():
    # No setting at all: catalog mode is the default.
    return _probe(None)


def test_three_tools_are_listed_and_all_stay_registered(catalog):
    assert sorted(catalog["listed"]) == ["ppt_describe", "ppt_run", "ppt_run_steps"]
    assert catalog["registered"] > 150


def test_the_catalog_names_every_tool_on_a_line_of_its_own(catalog):
    lines = catalog["catalog"].splitlines()
    assert len(lines) == catalog["registered"] - 3
    assert any(line.startswith("ppt_add_textbox: ") for line in lines)
    assert not any(line.startswith(("ppt_run", "ppt_describe")) for line in lines)


def test_describe_gives_the_params_schema_and_names_the_unknown(catalog):
    described = json.loads(catalog["described"])
    params = described["tools"]["ppt_add_textbox"]["params"]
    assert "slide_index" in params["properties"]
    assert described["unknown"] == ["ppt_nope"]


def test_an_unknown_tool_suggests_the_close_one(catalog):
    assert "No tool called 'ppt_add_texbox'" in catalog["unknown"]
    assert "ppt_add_textbox" in catalog["unknown"]


def test_a_tool_run_through_the_catalog_validates_as_before(catalog):
    assert "slide_index" in catalog["invalid"]


def test_steps_stop_at_the_first_error_and_say_which(catalog):
    steps = json.loads(catalog["steps"])
    assert steps["ran"] == 1 and steps["of"] == 2
    assert steps["steps"][0]["tool"] == "ppt_add_textbox"
    assert "wrong" in steps["steps"][0]["error"]


def test_references_read_earlier_results(catalog):
    assert catalog["refs"] == {"a": "TextBox 2", "b": [7], "c": "$keep"}


def test_catalog_mode_puts_its_note_first_in_the_instructions(catalog):
    assert catalog["instructions"].startswith("Catalog mode: call ppt_describe()")
    # The rest of the instructions follow, unchanged.
    assert "## Getting started" in catalog["instructions"]


def test_full_mode_lists_every_tool_but_the_three():
    full = _probe("full")
    assert "Catalog mode" not in full["instructions"]
    assert len(full["listed"]) == full["registered"] - 3
    assert not {"ppt_describe", "ppt_run", "ppt_run_steps"} & set(full["listed"])


def test_a_setting_that_is_not_one_falls_back_to_catalog():
    assert sorted(_probe("ful")["listed"]) == ["ppt_describe", "ppt_run", "ppt_run_steps"]
