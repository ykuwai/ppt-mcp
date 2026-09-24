"""Tests for refusing arguments a tool does not take (#244, #194 item 5).

Every input model used to accept unknown keys and drop them, so a misspelled
or outdated argument answered success while nothing was applied: `notes`
instead of `notes_text`, `rotation` on a tool that has no rotation,
`font_color` inside a `ranges` entry whose field is `color`. Now:

* every model reachable from a registered tool's arguments forbids extra
  keys, nested models included, so a new model without it fails here;
* a key next to `params` that the tool does not declare is refused by
  _PowerPointServer.call_tool, since the SDK's own argument model ignores it;
* the refusal is one readable message naming the tool, where the key was
  passed, the likely intended name and the valid arguments, instead of
  pydantic's multi-line report.

All of it is checked without COM: validation fails before any tool body runs.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import List, Optional

import anyio
import pytest
from pydantic import BaseModel, ConfigDict, Field

_src_dir = str(Path(__file__).resolve().parents[1] / "src")
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)

from utils.arguments import _flatten, suggest  # noqa: E402


def _server():
    import server

    return server


def _refusal(name, arguments):
    """The message call_tool refuses `arguments` with, on either SDK major."""
    server = _server()
    with pytest.raises(server.ToolError) as caught:
        anyio.run(server.mcp.call_tool, name, arguments)
    return str(caught.value)


# ---------------------------------------------------------------------------
# Every model forbids extra keys
# ---------------------------------------------------------------------------
def _nested_models(annotation, seen):
    for tp in _flatten(annotation):
        if isinstance(tp, type) and issubclass(tp, BaseModel):
            if tp not in seen:
                seen.add(tp)
                for field in tp.model_fields.values():
                    _nested_models(field.annotation, seen)
            continue
        for arg in getattr(tp, "__args__", ()) or ():
            _nested_models(arg, seen)


def test_every_tool_input_model_forbids_unknown_keys():
    server = _server()
    tools = server.mcp._tool_manager.list_tools()
    seen: set[type[BaseModel]] = set()
    for tool in tools:
        arg_model = tool.fn_metadata.arg_model
        # `params`, or nothing at all for a tool such as ppt_get_app_info.
        assert set(arg_model.model_fields) <= {"params"}, tool.name
        for field in arg_model.model_fields.values():
            _nested_models(field.annotation, seen)
    lax = sorted(m.__qualname__ for m in seen if m.model_config.get("extra") != "forbid")
    assert not lax, f"input models that still drop unknown keys: {lax}"
    names = {m.__name__ for m in seen}
    # The nested ones the issue named, so the walk is known to reach them.
    assert {"TextRangeSpec", "RunSpec", "NodeSpec", "BatchOperation"} <= names
    assert len(seen) >= 140


# ---------------------------------------------------------------------------
# The reporter's cases, through the real server's call_tool
# ---------------------------------------------------------------------------
def test_notes_instead_of_notes_text_is_refused():
    said = _refusal("ppt_set_slide_notes", {"params": {"slide_index": 97, "notes": "x"}})
    assert said.startswith("ppt_set_slide_notes: unknown argument 'notes'")
    assert "(did you mean 'notes_text'?)" in said
    assert "in params." in said
    assert "Valid arguments: slide_index, notes_text," in said
    assert "Nothing was applied." in said


def test_rotation_on_add_textbox_is_refused():
    said = _refusal("ppt_add_textbox", {"params": {
        "slide_index": 1, "left": 10, "top": 10, "width": 100, "height": 40,
        "text": "a", "rotation": 358,
    }})
    assert said.startswith("ppt_add_textbox: unknown argument 'rotation' in params.")
    assert "Valid arguments: slide_index, left, top, width, height, text," in said
    assert "Extra inputs are not permitted" not in said, "the raw pydantic report is gone"


def test_font_color_inside_ranges_is_refused_with_its_path():
    said = _refusal("ppt_format_text_range", {"params": {
        "slide_index": 1, "shape_name_or_index": "Title",
        "ranges": [
            {"start": 1, "length": 2, "color": "#000000"},
            {"start": 34, "length": 10, "font_size": 44, "font_color": "#0070C0"},
        ],
    }})
    assert "unknown argument 'font_color'" in said
    assert "in params.ranges[1]." in said
    # difflib alone would answer font_color_theme, a different thing.
    assert "(did you mean 'color'?)" in said
    assert "search_text" in said, "the valid list is the range's, not the tool's"


def test_an_unknown_key_on_set_glow_is_refused():
    said = _refusal("ppt_set_glow", {"params": {
        "slide_index": 1, "shape_name_or_index": 1, "radius": 5, "glow_color": "#FF0000",
    }})
    assert said.startswith("ppt_set_glow: unknown argument 'glow_color'")
    assert "Valid arguments: slide_index, shape_name_or_index, radius, color, transparency, target." in said


def test_target_on_set_glow_is_valid_now():
    server = _server()
    arg_model = server.mcp._tool_manager.get_tool("ppt_set_glow").fn_metadata.arg_model
    arg_model.model_validate({"params": {
        "slide_index": 1, "shape_name_or_index": 1, "radius": 5, "target": "text",
    }})


def test_valid_calls_still_validate():
    server = _server()
    manager = server.mcp._tool_manager
    manager.get_tool("ppt_set_slide_notes").fn_metadata.arg_model.model_validate(
        {"params": {"slide_index": 1, "notes_text": "  kept  "}})
    manager.get_tool("ppt_format_text_range").fn_metadata.arg_model.model_validate(
        {"params": {"slide_index": 1, "shape_name_or_index": 1, "ranges": [
            {"search_text": "a", "color": "#0070C0", "font_size": 44}]}})
    manager.get_tool("ppt_batch_apply_formatting").fn_metadata.arg_model.model_validate(
        {"params": {"slide_index": 1, "shapes": ["A", 2], "operations": [
            {"tool": "set_fill", "params": {"fill_type": "solid", "color": "#FFFFFF"}}]}})


def test_several_unknown_keys_share_one_valid_list():
    said = _refusal("ppt_build_freeform", {"params": {
        "slide_index": 1, "start_x": 0, "start_y": 0,
        "nodes": [{"segment_type": "line", "x1": 1, "y1": 1, "colour": 1, "z": 2}],
    }})
    assert "unknown arguments 'colour', 'z' in params.nodes[0]." in said
    assert said.count("Valid arguments:") == 1


def test_other_validation_errors_are_readable_too():
    said = _refusal("ppt_set_glow", {"params": {
        "slide_index": 0, "shape_name_or_index": 1, "radius": "big",
    }})
    assert said.startswith("ppt_set_glow: 2 problems with the arguments, nothing was applied.")
    assert "- params.slide_index: Input should be greater than or equal to 1 (got 0)" in said
    assert "- params.radius: " in said and "(got 'big')" in said
    assert "validation error for" not in said


def test_a_validator_message_loses_the_value_error_prefix():
    said = _refusal("ppt_format_text_range", {"params": {
        "slide_index": 1, "shape_name_or_index": 1,
        "ranges": [{"search_text": "a", "start": 1, "length": 1}],
    }})
    assert "params.ranges[0]: " in said
    assert "Value error," not in said


def test_a_missing_argument_is_named():
    said = _refusal("ppt_set_slide_notes", {"params": {"notes_text": "x"}})
    assert said == (
        "ppt_set_slide_notes: missing required argument 'slide_index' in params. "
        "Nothing was applied."
    )


# ---------------------------------------------------------------------------
# Keys next to `params`
# ---------------------------------------------------------------------------
def test_an_unknown_top_level_key_is_refused():
    said = _refusal("ppt_set_glow", {
        "params": {"slide_index": 1, "shape_name_or_index": 1, "radius": 5},
        "presntation": "deck.pptx",
    })
    assert said == (
        "ppt_set_glow: unknown argument 'presntation' (did you mean 'presentation'?). "
        "Valid arguments: params, presentation. Nothing was applied."
    )


def test_a_flattened_call_is_told_where_its_keys_belong():
    said = _refusal("ppt_set_slide_notes", {"slide_index": 1, "notes_text": "x"})
    assert "'slide_index', 'notes_text' belong inside params, not next to it." in said


def test_presentation_is_still_dropped_where_it_means_nothing():
    # ppt_list_templates takes no presentation; passing one stays harmless,
    # but a different stray key there is refused.
    said = _refusal("ppt_list_templates", {"params": {}, "presentation": "x", "extra": 1})
    assert "unknown argument 'extra'" in said
    assert "presentation" not in said.split("Valid arguments:")[1]


def test_a_tool_without_arguments_says_so():
    said = _refusal("ppt_get_app_info", {"params": {"verbose": True}})
    assert said == (
        "ppt_get_app_info: unknown argument 'params'. "
        "This tool takes no arguments. Nothing was applied."
    )


def test_an_empty_params_on_a_tool_without_arguments_is_harmless():
    server = _server()
    srv = server._PowerPointServer("probe")

    @srv.tool(name="ppt_bare")
    async def bare() -> str:
        return "ran"

    for arguments in ({}, {"params": {}}, {"params": None}):
        assert _text_of(anyio.run(srv.call_tool, "ppt_bare", arguments)) == "ran"


# ---------------------------------------------------------------------------
# A probe server, for what the real tools cannot show without COM
# ---------------------------------------------------------------------------
class _Item(BaseModel):
    model_config = ConfigDict(extra="forbid")
    color: Optional[str] = None


class _ProbeInput(BaseModel):
    model_config = ConfigDict(extra="forbid")
    value: int = Field(0, ge=0)
    items: List[_Item] = Field(default_factory=list)


def _probe_server():
    server = _server()
    srv = server._PowerPointServer("probe")

    @srv.tool(name="ppt_probe")
    async def probe(params: _ProbeInput) -> str:
        return f"{params.value}|{[i.color for i in params.items]}"

    @srv.tool(name="ppt_inner")
    async def inner(params: _ProbeInput) -> str:
        # A model built inside the body fails on its own terms; that is not
        # the caller's arguments and is not rewritten.
        _Item(colour="x")
        return "unreachable"

    return srv


def _text_of(result):
    content = getattr(result, "content", result)
    if isinstance(content, tuple):
        content = content[0]
    return "".join(getattr(block, "text", "") for block in content)


def test_probe_valid_call_is_unaffected():
    srv = _probe_server()
    result = anyio.run(srv.call_tool, "ppt_probe",
                       {"params": {"value": 2, "items": [{"color": "#000000"}]},
                        "presentation": "deck.pptx"})
    assert _text_of(result) == "2|['#000000']"


def test_probe_nested_unknown_key():
    server = _server()
    srv = _probe_server()
    with pytest.raises(server.ToolError) as caught:
        anyio.run(srv.call_tool, "ppt_probe", {"params": {"items": [{"font_color": "#000000"}]}})
    assert str(caught.value) == (
        "ppt_probe: unknown argument 'font_color' (did you mean 'color'?) in "
        "params.items[0]. Valid arguments: color. Nothing was applied."
    )


def test_probe_unknown_top_level_key_lists_presentation():
    server = _server()
    srv = _probe_server()
    with pytest.raises(server.ToolError) as caught:
        anyio.run(srv.call_tool, "ppt_probe", {"params": {}, "prams": {}})
    assert str(caught.value) == (
        "ppt_probe: unknown argument 'prams' (did you mean 'params'?). "
        "Valid arguments: params, presentation. Nothing was applied."
    )


def test_a_validation_error_from_inside_a_tool_is_left_alone():
    server = _server()
    srv = _probe_server()
    with pytest.raises(server.ToolError) as caught:
        anyio.run(srv.call_tool, "ppt_inner", {"params": {}})
    said = str(caught.value)
    # 1.x carries pydantic's text, 2.x keeps a crash's text on the server;
    # either way it is the SDK's message, not a refusal of the arguments.
    assert said.startswith("Error executing tool ppt_inner")
    assert "unknown argument" not in said


# ---------------------------------------------------------------------------
# suggest
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("key, known, expected", [
    ("notes", ["slide_index", "notes_text"], "notes_text"),
    ("font_color", ["color", "font_color_theme"], "color"),
    ("fill_color", ["color", "fill_type"], "color"),
    ("colour", ["color", "radius"], "color"),
    ("rotation", ["slide_index", "left", "top"], None),
    # A sibling name only counts where its target exists.
    ("line_weight", ["dash_style", "text"], None),
])
def test_suggest(key, known, expected):
    assert suggest(key, known) == expected
