"""A failure answers isError true, with its JSON text unchanged (#249).

Every tool answers a failure as json.dumps({"error": ...}). That used to go
out as a success, so a client that checks isError counted every failure as
one. _PowerPointServer.call_tool now marks it, in one place for every tool.

These go through a real client connected in-process, so the result is the one
a client receives, under whichever SDK major is installed. The server is a
stand-in with tools that need no PowerPoint and no lifespan (the real one
starts the PowerPoint worker), built from the same _PowerPointServer class,
with the catalog tools registered on it as the real server registers them.
"""

import json

import anyio
import pytest
from pydantic import BaseModel, ConfigDict

FAILURE = {"error": "ppt_merge_shapes is not available on macOS",
           "reason": "PowerPoint for Mac has no Merge Shapes in its dictionary.",
           "platform": "macOS"}


def _server():
    import server

    return server


class _Slide(BaseModel):
    model_config = ConfigDict(extra="forbid")
    slide_index: int = 1


@pytest.fixture
def srv(monkeypatch):
    server = _server()
    srv = server._PowerPointServer("is-error-probes")

    @srv.tool(name="ppt_refuse")
    async def refuse(params: _Slide) -> str:
        """Answer a failure the way every tool does."""
        return json.dumps(FAILURE)

    @srv.tool(name="ppt_work")
    async def work(params: _Slide) -> str:
        """Succeed."""
        return json.dumps({"success": True, "slide_index": params.slide_index})

    @srv.tool(name="ppt_nested")
    async def nested(params: _Slide) -> str:
        """Succeed with a failure inside, as batch formatting does per shape."""
        return json.dumps({"success": True, "results": [{"status": "error", "error": "x"}]})

    @srv.tool(name="ppt_words")
    async def words(params: _Slide) -> str:
        """Answer plain text that mentions an error."""
        return 'not json, though it says {"error": 1}'

    srv.tool(name="ppt_run")(server.ppt_run)
    srv.tool(name="ppt_run_steps")(server.ppt_run_steps)
    server._inline_schema_defs(srv)
    # ppt_run dispatches through the module's `mcp`.
    monkeypatch.setattr(server, "mcp", srv)
    return srv


@pytest.fixture
def flagged():
    """Skip where the SDK cannot send the flagged result (before mcp 1.19)."""
    if not _server()._SDK_PASSES_CALL_TOOL_RESULT:
        pytest.skip("this mcp release cannot send a CallToolResult from call_tool")


def call(srv, name, arguments):
    """What a client connected to `srv` receives: (isError, text, structured)."""
    async def go():
        try:
            from mcp.client import Client  # mcp >= 2.0
        except ImportError:
            from mcp.shared.memory import create_connected_server_and_client_session

            # The low-level server, which older 1.x releases require here.
            async with create_connected_server_and_client_session(srv._mcp_server) as session:
                return await session.call_tool(name, arguments)
        async with Client(srv) as client:
            return await client.call_tool(name, arguments)

    result = anyio.run(go)
    is_error = getattr(result, "isError", None)
    if is_error is None:
        is_error = result.is_error
    structured = getattr(result, "structuredContent", getattr(result, "structured_content", None))
    text = "".join(getattr(b, "text", "") or "" for b in result.content)
    return bool(is_error), text, structured


def test_a_failure_is_an_error_with_its_json_unchanged(srv, flagged):
    is_error, text, structured = call(srv, "ppt_refuse", {"params": {}})
    assert is_error
    assert text == json.dumps(FAILURE), "no prefix, nothing added"
    assert structured is None


def test_a_success_is_not_an_error_and_keeps_its_structured_content(srv):
    is_error, text, structured = call(srv, "ppt_work", {"params": {"slide_index": 2}})
    assert not is_error
    assert json.loads(text) == {"success": True, "slide_index": 2}
    assert structured == {"result": text}


@pytest.mark.parametrize("tool", ["ppt_nested", "ppt_words"])
def test_only_a_top_level_error_in_a_json_object_counts(srv, tool):
    is_error, _, _ = call(srv, tool, {"params": {}})
    assert not is_error


def test_ppt_run_of_a_failing_tool_is_an_error_with_the_same_json(srv, flagged):
    is_error, text, _ = call(srv, "ppt_run", {"tool": "ppt_refuse", "params": {}})
    assert is_error
    assert text == json.dumps(FAILURE)


def test_ppt_run_of_a_working_tool_is_not_an_error(srv):
    is_error, text, _ = call(srv, "ppt_run", {"tool": "ppt_work", "params": {}})
    assert not is_error
    assert json.loads(text)["success"] is True


def test_ppt_run_steps_with_a_failing_step_is_an_error_and_stops_there(srv, flagged):
    is_error, text, _ = call(srv, "ppt_run_steps", {"steps": [
        {"tool": "ppt_work", "params": {}},
        {"tool": "ppt_refuse", "params": {}},
        {"tool": "ppt_work", "params": {}},
    ]})
    assert is_error
    summary = json.loads(text)
    assert summary["ran"] == 2 and summary["of"] == 3
    assert summary["steps"][1]["result"] == FAILURE, "the tool's answer as it was"
    assert summary["error"] == f"Step 2 (ppt_refuse) failed: {FAILURE['error']}"


def test_ppt_run_steps_without_stopping_runs_on_and_counts_the_failures(srv, flagged):
    is_error, text, _ = call(srv, "ppt_run_steps", {"stop_on_error": False, "steps": [
        {"tool": "ppt_refuse", "params": {}},
        {"tool": "ppt_nope"},
        {"tool": "ppt_work", "params": {}},
    ]})
    assert is_error
    summary = json.loads(text)
    assert summary["ran"] == 3
    assert summary["error"].startswith("Step 1 (ppt_refuse) failed: ")
    assert summary["error"].endswith("(2 steps failed)")


def test_a_step_cannot_refer_to_a_failed_step(srv):
    _, text, _ = call(srv, "ppt_run_steps", {"stop_on_error": False, "steps": [
        {"tool": "ppt_refuse", "params": {}},
        {"tool": "ppt_work", "params": {"slide_index": "$prev.platform"}},
    ]})
    assert "refers to a step that has no result" in json.loads(text)["steps"][1]["error"]


def test_ppt_run_steps_that_all_work_is_not_an_error(srv):
    is_error, text, _ = call(srv, "ppt_run_steps", {"steps": [
        {"tool": "ppt_work", "params": {}}, {"tool": "ppt_nested", "params": {}}]})
    assert not is_error
    assert "error" not in json.loads(text)


@pytest.mark.parametrize("arguments, said", [
    ({"params": {}, "slide": 1}, "ppt_work: unknown argument 'slide'"),
    ({"params": {"wrong": 1}}, "ppt_work: unknown argument 'wrong' in params."),
])
def test_an_unknown_argument_is_an_error(srv, arguments, said):
    is_error, text, _ = call(srv, "ppt_work", arguments)
    assert is_error
    assert text.startswith(said)
