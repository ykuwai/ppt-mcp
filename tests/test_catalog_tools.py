"""The catalog tools themselves: ppt_describe, ppt_run and ppt_run_steps.

They reach the other tools through the server's own call_tool, so here they
are pointed at a stand-in server whose tools need no PowerPoint: one makes a
"shape" and answers its name, one echoes what it was given, one fails, one
answers plain text, one an image, and one takes no arguments. A crash's text
reaches the caller under mcp 1.x and stays on the server under 2.x; a ToolError's
reaches it under both, and the probes check each as a direct call would answer. The mode itself,
read when the server loads, is in test_tool_catalog.py.
"""

import base64
import json
import logging

import anyio
import pytest
from pydantic import BaseModel, ConfigDict

from utils.com_wrapper import call_presentation

PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)


def _server():
    import server

    return server


class _Name(BaseModel):
    model_config = ConfigDict(extra="forbid")
    name: str = "Shape 1"


class _Echo(BaseModel):
    model_config = ConfigDict(extra="forbid")
    value: object = None
    slide_index: int = 1


@pytest.fixture
def probes(monkeypatch):
    """The catalog code, pointed at a server of stand-in tools."""
    server = _server()
    srv = server._PowerPointServer("probes")

    @srv.tool(name="ppt_make")
    async def make(params: _Name) -> str:
        """Make a shape and answer its name."""
        return json.dumps({"success": True, "shape_name": params.name,
                           "deck": call_presentation.get(), "nested": {"list": [{"id": 7}]}})

    @srv.tool(name="ppt_echo")
    async def echo(params: _Echo) -> str:
        """Answer what it was given. Then more words that are not in the summary."""
        return json.dumps({"value": params.value, "slide_index": params.slide_index,
                           "deck": call_presentation.get()})

    @srv.tool(name="ppt_fail")
    async def fail(params: _Echo) -> str:
        """Always fails!"""
        raise RuntimeError("the shape is gone")

    @srv.tool(name="ppt_refuse")
    async def refuse(params: _Echo) -> str:
        """Refuse, saying why."""
        raise server.ToolError("slide 9 does not exist")

    @srv.tool(name="ppt_text")
    async def text(params: _Echo) -> str:
        """Answer plain text?"""
        return "not json"

    @srv.tool(name="ppt_picture")
    async def picture(params: _Echo):
        """Answer a preview image."""
        return [json.dumps({"slide_index": params.slide_index}), server.Image(data=PNG, format="png")]

    @srv.tool(name="ppt_bare")
    async def bare() -> str:
        """Take no arguments at all."""
        return json.dumps({"ok": True})

    # As the real server does at load: params schemas inlined, no $ref.
    server._inline_schema_defs(srv)
    monkeypatch.setattr(server, "mcp", srv)
    return server


def run(fn, *args, **kwargs):
    return anyio.run(lambda: fn(*args, **kwargs))


def steps(server, items, **kwargs):
    """Run ppt_run_steps and answer its report and the images after it."""
    out = run(server.ppt_run_steps, items, **kwargs)
    if isinstance(out, str):
        return json.loads(out), []
    return json.loads(out[0]), out[1:]


def crash_text_reaches_caller(server):
    """Whether a direct call to a crashing tool answers the crash's text."""
    try:
        run(server.mcp.call_tool, "ppt_fail", {"params": {}})
    except Exception as exc:
        return "the shape is gone" in str(exc)
    raise AssertionError("ppt_fail did not fail")


def text_of(blocks):
    return "".join(getattr(b, "text", "") or "" for b in blocks)


# ---------------------------------------------------------------------------
# ppt_describe
# ---------------------------------------------------------------------------
class TestDescribe:
    def test_no_names_answers_the_list_one_line_a_tool(self, probes):
        lines = run(probes.ppt_describe).splitlines()
        assert lines == sorted(lines, key=lambda line: [
            "ppt_make", "ppt_echo", "ppt_fail", "ppt_refuse", "ppt_text", "ppt_picture", "ppt_bare",
        ].index(line.split(":")[0]))
        assert "ppt_echo: Answer what it was given." in lines

    def test_an_empty_list_is_the_same_as_no_names(self, probes):
        assert run(probes.ppt_describe, []) == run(probes.ppt_describe)

    def test_the_line_is_the_first_sentence_whatever_ends_it(self, probes):
        lines = dict(line.split(": ", 1) for line in run(probes.ppt_describe).splitlines())
        assert lines["ppt_fail"] == "Always fails!"
        assert lines["ppt_text"] == "Answer plain text?"
        assert "more words" not in lines["ppt_echo"]

    def test_the_catalog_tools_are_not_in_the_list(self, probes):
        text = run(probes.ppt_describe)
        for name in ("ppt_describe", "ppt_run", "ppt_run_steps"):
            assert not any(line.startswith(name + ":") for line in text.splitlines())

    def test_names_answer_the_description_and_the_params_schema(self, probes):
        found = json.loads(run(probes.ppt_describe, ["ppt_make"]))
        tool = found["tools"]["ppt_make"]
        assert tool["description"].startswith("Make a shape")
        assert "name" in tool["params"]["properties"]
        assert "unknown" not in found

    def test_presentation_is_not_part_of_the_params_schema(self, probes):
        tool = json.loads(run(probes.ppt_describe, ["ppt_echo"]))["tools"]["ppt_echo"]
        assert "presentation" not in tool["params"].get("properties", {})

    def test_a_tool_without_params_answers_none_for_them(self, probes):
        tool = json.loads(run(probes.ppt_describe, ["ppt_bare"]))["tools"]["ppt_bare"]
        assert tool["params"] is None

    def test_unknown_names_are_listed_and_the_known_still_answered(self, probes):
        found = json.loads(run(probes.ppt_describe, ["ppt_make", "ppt_nope", "ppt_run"]))
        assert list(found["tools"]) == ["ppt_make"]
        assert found["unknown"] == ["ppt_nope", "ppt_run"]

    def test_several_names_answer_each_once(self, probes):
        found = json.loads(run(probes.ppt_describe, ["ppt_make", "ppt_echo", "ppt_make"]))
        assert sorted(found["tools"]) == ["ppt_echo", "ppt_make"]


# ---------------------------------------------------------------------------
# ppt_run
# ---------------------------------------------------------------------------
class TestRun:
    def test_runs_the_tool_and_answers_what_it_answers(self, probes):
        out = json.loads(text_of(run(probes.ppt_run, "ppt_echo", {"value": "hi", "slide_index": 3})))
        assert out == {"value": "hi", "slide_index": 3, "deck": None}

    def test_params_left_out_take_the_defaults(self, probes):
        out = json.loads(text_of(run(probes.ppt_run, "ppt_make")))
        assert out["shape_name"] == "Shape 1"

    def test_presentation_reaches_the_tool_for_that_call_only(self, probes):
        out = json.loads(text_of(run(probes.ppt_run, "ppt_echo", {}, "deck.pptx")))
        assert out["deck"] == "deck.pptx"
        assert call_presentation.get() is None

    def test_a_tool_without_params_runs_with_none(self, probes):
        assert json.loads(text_of(run(probes.ppt_run, "ppt_bare"))) == {"ok": True}

    def test_a_tool_without_params_refuses_some(self, probes):
        with pytest.raises(Exception, match="ppt_bare takes no params"):
            run(probes.ppt_run, "ppt_bare", {"x": 1})

    def test_an_unknown_tool_names_the_close_one(self, probes):
        with pytest.raises(Exception, match=r"No tool called 'ppt_ech'.*ppt_echo"):
            run(probes.ppt_run, "ppt_ech", {})

    def test_an_unknown_tool_with_nothing_close_says_where_the_list_is(self, probes):
        with pytest.raises(Exception, match="ppt_describe") as caught:
            run(probes.ppt_run, "zzzz", {})
        assert "Did you mean" not in str(caught.value)

    def test_the_catalog_tools_do_not_run_themselves(self, probes):
        for name in ("ppt_run", "ppt_run_steps", "ppt_describe"):
            with pytest.raises(Exception, match=f"No tool called '{name}'"):
                run(probes.ppt_run, name, {})

    def test_an_unknown_argument_is_refused_by_name(self, probes):
        with pytest.raises(Exception, match="valeu"):
            run(probes.ppt_run, "ppt_echo", {"valeu": 1})

    def test_a_wrong_type_is_refused(self, probes):
        with pytest.raises(Exception, match="slide_index"):
            run(probes.ppt_run, "ppt_echo", {"slide_index": "three"})

    def test_a_failing_tool_fails_the_run(self, probes):
        with pytest.raises(Exception, match="Error executing tool ppt_fail") as raised:
            run(probes.ppt_run, "ppt_fail", {})
        # The reason is there exactly when a direct call would give it.
        assert ("the shape is gone" in str(raised.value)) == crash_text_reaches_caller(probes)

    def test_a_refusal_keeps_its_reason(self, probes):
        with pytest.raises(Exception, match="slide 9 does not exist"):
            run(probes.ppt_run, "ppt_refuse", {})

    def test_an_image_comes_back_as_an_image(self, probes):
        blocks = run(probes.ppt_run, "ppt_picture", {"slide_index": 2})
        assert [getattr(b, "type", "") for b in blocks] == ["text", "image"]


# ---------------------------------------------------------------------------
# ppt_run_steps
# ---------------------------------------------------------------------------
class TestRunSteps:
    def test_steps_run_in_order_and_each_is_reported(self, probes):
        report, images = steps(probes, [
            {"tool": "ppt_make", "params": {"name": "A"}},
            {"tool": "ppt_echo", "params": {"value": 1}},
            {"tool": "ppt_bare"},
        ])
        assert (report["ran"], report["of"]) == (3, 3)
        assert [s["step"] for s in report["steps"]] == [1, 2, 3]
        assert [s["tool"] for s in report["steps"]] == ["ppt_make", "ppt_echo", "ppt_bare"]
        assert report["steps"][0]["result"]["shape_name"] == "A"
        assert images == []

    def test_prev_is_the_previous_steps_field(self, probes):
        report, _ = steps(probes, [
            {"tool": "ppt_make", "params": {"name": "Rectangle 3"}},
            {"tool": "ppt_echo", "params": {"value": "$prev.shape_name"}},
        ])
        assert report["steps"][1]["result"]["value"] == "Rectangle 3"

    def test_a_number_names_any_earlier_step(self, probes):
        report, _ = steps(probes, [
            {"tool": "ppt_make", "params": {"name": "First"}},
            {"tool": "ppt_make", "params": {"name": "Second"}},
            {"tool": "ppt_echo", "params": {"value": "$1.shape_name"}},
        ])
        assert report["steps"][2]["result"]["value"] == "First"

    def test_a_path_reaches_into_dicts_and_lists(self, probes):
        report, _ = steps(probes, [
            {"tool": "ppt_make", "params": {}},
            {"tool": "ppt_echo", "params": {"value": "$prev.nested.list.0.id"}},
        ])
        assert report["steps"][1]["result"]["value"] == 7

    def test_references_inside_lists_and_dicts_are_resolved(self, probes):
        report, _ = steps(probes, [
            {"tool": "ppt_make", "params": {"name": "A"}},
            {"tool": "ppt_echo", "params": {"value": {"names": ["$prev.shape_name", "fixed"]}}},
        ])
        assert report["steps"][1]["result"]["value"] == {"names": ["A", "fixed"]}

    def test_a_reference_keeps_its_type(self, probes):
        report, _ = steps(probes, [
            {"tool": "ppt_echo", "params": {"slide_index": 4}},
            {"tool": "ppt_echo", "params": {"slide_index": "$prev.slide_index"}},
        ])
        assert report["steps"][1]["result"]["slide_index"] == 4

    def test_strings_that_only_look_like_references_are_left_alone(self, probes):
        report, _ = steps(probes, [
            {"tool": "ppt_make", "params": {}},
            {"tool": "ppt_echo", "params": {"value": ["$5", "$", "$prevx", "$total.cost", "US$ 3.50"]}},
        ])
        assert report["steps"][1]["result"]["value"] == ["$5", "$", "$prevx", "$total.cost", "US$ 3.50"]

    def test_a_reference_to_a_missing_field_is_an_error_of_that_step(self, probes):
        report, _ = steps(probes, [
            {"tool": "ppt_make", "params": {}},
            {"tool": "ppt_echo", "params": {"value": "$prev.colour"}},
        ])
        assert report["ran"] == 2
        assert "has no 'colour'" in report["steps"][1]["error"]

    def test_a_reference_to_a_later_step_is_an_error(self, probes):
        report, _ = steps(probes, [
            {"tool": "ppt_echo", "params": {"value": "$2.value"}},
            {"tool": "ppt_echo", "params": {}},
        ])
        assert "refers to a step that has no result" in report["steps"][0]["error"]

    def test_prev_on_the_first_step_is_an_error(self, probes):
        report, _ = steps(probes, [{"tool": "ppt_echo", "params": {"value": "$prev.value"}}])
        assert "no result" in report["steps"][0]["error"]

    def test_the_first_error_stops_the_run_and_says_which_step(self, probes):
        report, _ = steps(probes, [
            {"tool": "ppt_echo", "params": {}},
            {"tool": "ppt_fail", "params": {}},
            {"tool": "ppt_echo", "params": {}},
        ])
        assert (report["ran"], report["of"]) == (2, 3)
        assert report["steps"][1] == {"step": 2, "tool": "ppt_fail",
                                      "error": report["steps"][1]["error"]}
        assert "Error executing tool ppt_fail" in report["steps"][1]["error"]

    def test_a_refusal_in_a_step_keeps_its_reason(self, probes):
        report, _ = steps(probes, [{"tool": "ppt_refuse", "params": {}}])
        assert "slide 9 does not exist" in report["steps"][0]["error"]

    def test_without_stopping_every_step_runs_and_each_failure_is_reported(self, probes):
        report, _ = steps(probes, [
            {"tool": "ppt_fail", "params": {}},
            {"tool": "ppt_nope", "params": {}},
            {"tool": "ppt_echo", "params": {"value": 2}},
        ], stop_on_error=False)
        assert (report["ran"], report["of"]) == (3, 3)
        assert "error" in report["steps"][0] and "error" in report["steps"][1]
        assert report["steps"][2]["result"]["value"] == 2

    def test_a_reference_to_a_failed_step_is_an_error(self, probes):
        report, _ = steps(probes, [
            {"tool": "ppt_fail", "params": {}},
            {"tool": "ppt_echo", "params": {"value": "$1.value"}},
        ], stop_on_error=False)
        assert "no result" in report["steps"][1]["error"]

    def test_an_unknown_argument_is_reported_with_the_step(self, probes):
        report, _ = steps(probes, [{"tool": "ppt_echo", "params": {"valeu": 1}}])
        assert "valeu" in report["steps"][0]["error"]

    def test_a_step_with_no_tool_says_so(self, probes):
        report, _ = steps(probes, [{"params": {}}, {"tool": 3}], stop_on_error=False)
        assert report["steps"][0]["tool"] is None
        assert report["steps"][0]["error"].startswith("A step needs a 'tool'")
        assert report["steps"][1]["error"].startswith("A step needs a 'tool'")

    def test_a_step_that_is_not_a_dict_says_so(self, probes):
        report, _ = steps(probes, [None, "ppt_echo"], stop_on_error=False)
        assert all("A step needs a 'tool'" in s["error"] for s in report["steps"])

    def test_plain_text_answers_are_kept_as_text(self, probes):
        report, _ = steps(probes, [{"tool": "ppt_text", "params": {}}])
        assert report["steps"][0]["result"] == {"text": "not json"}

    def test_images_follow_the_report_in_step_order(self, probes):
        report, images = steps(probes, [
            {"tool": "ppt_picture", "params": {"slide_index": 1}},
            {"tool": "ppt_echo", "params": {}},
            {"tool": "ppt_picture", "params": {"slide_index": 2}},
        ])
        assert len(images) == 2
        assert all(getattr(i, "type", "") == "image" for i in images)
        assert report["steps"][0]["result"] == {"slide_index": 1}

    def test_no_steps_is_an_empty_report(self, probes):
        report, images = steps(probes, [])
        assert report == {"steps": [], "ran": 0, "of": 0}
        assert images == []

    def test_the_presentation_applies_to_every_step(self, probes):
        report, _ = steps(probes, [{"tool": "ppt_echo"}, {"tool": "ppt_make"}],
                          presentation="a.pptx")
        assert [s["result"]["deck"] for s in report["steps"]] == ["a.pptx", "a.pptx"]

    def test_a_step_can_name_another_presentation(self, probes):
        report, _ = steps(probes, [
            {"tool": "ppt_echo"},
            {"tool": "ppt_echo", "presentation": "b.pptx"},
        ], presentation="a.pptx")
        assert [s["result"]["deck"] for s in report["steps"]] == ["a.pptx", "b.pptx"]

    def test_no_presentation_leaves_none_behind(self, probes):
        steps(probes, [{"tool": "ppt_echo"}], presentation="a.pptx")
        assert call_presentation.get() is None

    def test_non_ascii_text_is_kept_as_it_is(self, probes):
        report, _ = steps(probes, [{"tool": "ppt_echo", "params": {"value": "שלום · مرحبا · 日本"}}])
        assert report["steps"][0]["result"]["value"] == "שלום · مرحبا · 日本"

    def test_many_steps_run_in_one_call(self, probes):
        items = [{"tool": "ppt_make", "params": {"name": f"S{i}"}} for i in range(200)]
        report, _ = steps(probes, items)
        assert report["ran"] == 200
        assert report["steps"][-1]["result"]["shape_name"] == "S199"


# ---------------------------------------------------------------------------
# The helpers on their own
# ---------------------------------------------------------------------------
class TestHelpers:
    @pytest.mark.parametrize("value, mode", [
        ("catalog", "catalog"), ("", "catalog"), ("  CATALOG ", "catalog"),
        ("full", "full"), ("FULL", "full"), ("all", "full"),
    ])
    def test_the_mode_setting(self, value, mode):
        assert _server()._tool_mode(value) == mode

    def test_a_mistyped_mode_is_catalog_with_a_warning(self, caplog):
        with caplog.at_level(logging.WARNING):
            assert _server()._tool_mode("fll") == "catalog"
        assert "PPT_TOOL_MODE='fll' is not a setting" in caplog.text

    @pytest.mark.parametrize("description, summary", [
        ("One. Two.", "One."),
        ("Stop! Then go.", "Stop!"),
        ("Why? Because.", "Why?"),
        ("No full stop", "No full stop"),
        ("  Spread\n   over\nlines. Rest.", "Spread over lines."),
        ("", ""),
        (None, ""),
        ("v1.2 is fine. Next.", "v1.2 is fine."),
    ])
    def test_the_summary(self, description, summary):
        assert _server()._summary(description) == summary

    def test_a_long_first_sentence_is_cut_at_120(self):
        assert len(_server()._summary("x" * 300)) == 120

    def test_content_from_a_tuple_is_its_first_part(self):
        assert _server()._content_of((["a"], {"b": 1})) == ["a"]

    def test_content_from_a_dict_is_its_json(self):
        assert _server()._content_of({"a": 1}) == ['{"a": 1}']

    def test_content_from_a_result_that_failed_raises(self):
        class Block:
            type, text = "text", "it broke"

        class Result:
            content, isError = [Block()], True

        with pytest.raises(Exception, match="it broke"):
            _server()._content_of(Result())

    def test_content_from_a_result_that_worked_is_its_blocks(self):
        class Block:
            type, text = "text", "fine"

        class Result:
            content, isError = [Block()], False

        assert _server()._content_of(Result())[0].text == "fine"

    def test_nothing_is_no_content(self):
        assert _server()._content_of(None) == []


# ---------------------------------------------------------------------------
# On the real server
# ---------------------------------------------------------------------------
class TestOnTheServer:
    def test_the_three_are_registered_in_either_mode(self):
        names = {t.name for t in _server().mcp._tool_manager.list_tools()}
        assert {"ppt_describe", "ppt_run", "ppt_run_steps"} <= names

    def test_full_mode_lists_every_tool_but_the_three(self, monkeypatch):
        server = _server()
        monkeypatch.setattr(server, "CATALOG_MODE", False)
        listed = {t.name for t in anyio.run(server.mcp.list_tools)}
        registered = {t.name for t in server.mcp._tool_manager.list_tools()}
        assert listed == registered - {"ppt_describe", "ppt_run", "ppt_run_steps"}

    def test_catalog_mode_lists_the_three(self, monkeypatch):
        server = _server()
        monkeypatch.setattr(server, "CATALOG_MODE", True)
        listed = sorted(t.name for t in anyio.run(server.mcp.list_tools))
        assert listed == ["ppt_describe", "ppt_run", "ppt_run_steps"]

    def test_the_three_describe_how_to_use_them(self):
        tools = {t.name: t for t in _server().mcp._tool_manager.list_tools()}
        assert "ppt_run" in tools["ppt_describe"].description
        assert "params" in tools["ppt_run"].description
        assert "$prev" in tools["ppt_run_steps"].description

    def test_the_real_list_names_every_registered_tool(self):
        server = _server()
        lines = server._catalog_lines().splitlines()
        registered = {t.name for t in server.mcp._tool_manager.list_tools()}
        assert {line.split(":")[0] for line in lines} == registered - server._CATALOG_TOOLS

    def test_describe_answers_a_real_tool_without_powerpoint(self):
        found = json.loads(anyio.run(_server().ppt_describe, ["ppt_add_textbox"]))
        params = found["tools"]["ppt_add_textbox"]["params"]
        assert {"slide_index", "left", "top", "width", "height"} <= set(params["properties"])

    def test_run_refuses_a_bad_argument_before_reaching_powerpoint(self):
        with pytest.raises(Exception, match="slide_index"):
            anyio.run(lambda: _server().ppt_run("ppt_add_textbox", {
                "slide_index": 0, "left": 0, "top": 0, "width": 1, "height": 1}))
