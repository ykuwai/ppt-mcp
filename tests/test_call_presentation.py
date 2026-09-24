"""Tests for the per-call `presentation` argument.

The session target set by ppt_activate_presentation is one value for the whole
server process, and one process can serve several conversations: Claude
Desktop starts the server once and routes every chat through it. When a second
conversation activates its deck, the first one's next call used to land there.
A tool call can now name its presentation, and that name wins over the session
target for that call only.

Covered here without COM and without PowerPoint:

* pick_presentation: how a name is matched against the open decks;
* the Windows wrapper: the per-call target reaches the worker through
  execute(), wins in _find_target_pres_impl (and so in _get_pres_impl and
  _get_target_window_impl), never falls back, and leaves the session target
  alone;
* the server: call_tool takes the argument out before validation, holds it in
  call_presentation while the tool runs, and every presentation-bound tool
  advertises it;
* the macOS wrapper, on macOS only.
"""

from __future__ import annotations

import sys
import threading
from pathlib import Path
from unittest.mock import MagicMock

import anyio
import pytest
from pydantic import BaseModel

_src_dir = str(Path(__file__).resolve().parents[1] / "src")
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)

from utils.com_wrapper import (  # noqa: E402
    PowerPointCOMWrapper,
    bind_call_presentation,
    call_presentation,
    pick_presentation,
)

macos_only = pytest.mark.skipif(
    sys.platform != "darwin", reason="the Apple Event backend needs appscript"
)


# ---------------------------------------------------------------------------
# pick_presentation
# ---------------------------------------------------------------------------
_A = ("C:\\Decks\\Quarterly.pptx", "Quarterly.pptx", "a")
_B = ("C:\\Decks\\Roadmap.pptx", "Roadmap.pptx", "b")
_A2 = ("D:\\Archive\\Quarterly.pptx", "Quarterly.pptx", "a2")
_UNSAVED = ("Presentation1", "Presentation1", "u")


@pytest.mark.parametrize("wanted, expected", [
    ("C:\\Decks\\Roadmap.pptx", "b"),
    ("c:\\decks\\ROADMAP.pptx", "b"),
    ("Roadmap.pptx", "b"),
    ("roadmap", "b"),
    ("  Roadmap.pptx  ", "b"),
    ("Presentation1", "u"),
])
def test_pick_matches_full_name_name_and_stem(wanted, expected):
    assert pick_presentation([_A, _B, _UNSAVED], wanted) == expected


def test_pick_full_name_tells_same_named_files_apart():
    assert pick_presentation([_A, _A2], "D:\\Archive\\Quarterly.pptx") == "a2"


def test_pick_refuses_an_ambiguous_name():
    with pytest.raises(ValueError, match="several open files") as caught:
        pick_presentation([_A, _A2], "Quarterly.pptx")
    assert "C:\\Decks\\Quarterly.pptx" in str(caught.value)
    assert "D:\\Archive\\Quarterly.pptx" in str(caught.value)


def test_pick_refuses_a_deck_that_is_not_open():
    with pytest.raises(ValueError, match="is not open") as caught:
        pick_presentation([_A, _B], "Budget.pptx")
    assert "Roadmap.pptx" in str(caught.value), "the open decks are listed"


# ---------------------------------------------------------------------------
# Windows wrapper
# ---------------------------------------------------------------------------
def _make_pres(full_name, window_count=1):
    pres = MagicMock()
    pres.FullName = full_name
    pres.Name = full_name.replace("/", "\\").rsplit("\\", 1)[-1]
    pres.Windows.Count = window_count
    windows = {i: MagicMock(name=f"{full_name}-win{i}") for i in range(1, window_count + 1)}
    pres.Windows.side_effect = lambda i: windows[i]
    return pres


def _make_app(presentations, active=None):
    app = MagicMock()
    app.Presentations.Count = len(presentations)
    app.Presentations.side_effect = lambda i: presentations[i - 1]
    app.Windows.Count = 1
    app.ActivePresentation = active if active is not None else presentations[0]
    return app


def _wrapper_with(app):
    w = PowerPointCOMWrapper()
    w._app = app
    w._get_app_impl = lambda allow_launch=False: app
    return w


def _as_call(w, wanted, func):
    """Run func the way the worker runs an operation of a call that named
    `wanted`."""
    return bind_call_presentation(w._call_local, wanted, func)()


def test_the_named_deck_wins_over_the_session_target():
    a, b = _make_pres("C:/a.pptx"), _make_pres("C:/b.pptx")
    app = _make_app([a, b])
    w = _wrapper_with(app)
    w._target_pres_full_name = "C:/a.pptx"

    assert _as_call(w, "b.pptx", w._get_pres_impl) is b
    assert w._target_pres_full_name == "C:/a.pptx", "the session target is left alone"
    assert w._get_pres_impl() is a, "outside the call the session target applies again"


def test_the_named_deck_is_found_without_any_session_target():
    a, b = _make_pres("C:/a.pptx"), _make_pres("C:/b.pptx")
    w = _wrapper_with(_make_app([a, b], active=a))
    assert _as_call(w, "C:/b.pptx", w._get_pres_impl) is b


def test_a_named_deck_that_is_not_open_never_falls_back():
    a = _make_pres("C:/a.pptx")
    w = _wrapper_with(_make_app([a], active=a))
    w._target_pres_full_name = "C:/a.pptx"
    with pytest.raises(ValueError, match="is not open"):
        _as_call(w, "gone.pptx", w._get_pres_impl)
    assert w._target_pres_full_name == "C:/a.pptx"


def test_the_named_deck_drives_its_own_window():
    a, b = _make_pres("C:/a.pptx"), _make_pres("C:/b.pptx")
    app = _make_app([a, b], active=a)
    w = _wrapper_with(app)
    w._target_pres_full_name = "C:/a.pptx"
    assert _as_call(w, "b.pptx", w._get_target_window_impl) is b.Windows(1)


def test_a_named_deck_without_a_window_is_not_swapped_for_the_active_one():
    a, headless = _make_pres("C:/a.pptx"), _make_pres("C:/h.pptx", window_count=0)
    w = _wrapper_with(_make_app([a, headless], active=a))
    assert _as_call(w, "h.pptx", w._get_target_window_impl) is None


def test_the_binding_is_undone_after_the_operation():
    w = PowerPointCOMWrapper()
    seen = []

    def op():
        seen.append(w._call_local.presentation)
        raise RuntimeError("boom")

    with pytest.raises(RuntimeError):
        bind_call_presentation(w._call_local, "b.pptx", op)()
    assert seen == ["b.pptx"]
    assert getattr(w._call_local, "presentation", None) is None


def test_the_binding_is_per_thread():
    """An abandoned worker and its replacement can overlap (issue #199)."""
    w = PowerPointCOMWrapper()
    inside = threading.Event()
    release = threading.Event()
    seen_elsewhere = []

    def slow():
        inside.set()
        release.wait(5)

    t = threading.Thread(target=bind_call_presentation(w._call_local, "b.pptx", slow))
    t.start()
    inside.wait(5)
    seen_elsewhere.append(getattr(w._call_local, "presentation", None))
    release.set()
    t.join(5)
    assert seen_elsewhere == [None]


def test_execute_hands_the_call_target_to_the_worker():
    """execute() runs on the calling thread and the operation on the worker, so
    the context variable has to travel with the queued item."""
    w = PowerPointCOMWrapper()
    w._running = True
    captured = []
    w.start = lambda: None

    def fake_put(item):
        func, args, kwargs, future, _idempotent, dequeued = item
        dequeued.set()
        future.set_result(func(*args, **kwargs))

    w._queue = MagicMock()
    w._queue.put.side_effect = fake_put

    def op():
        captured.append(getattr(w._call_local, "presentation", None))
        return "done"

    token = call_presentation.set("b.pptx")
    try:
        assert w.execute(op) == "done"
    finally:
        call_presentation.reset(token)
    assert w.execute(op) == "done"
    assert captured == ["b.pptx", None]


# ---------------------------------------------------------------------------
# Server
# ---------------------------------------------------------------------------
def _server_module():
    import server

    return server


def _text_of(result):
    """Text of a call_tool result on either SDK major."""
    content = getattr(result, "content", result)
    if isinstance(content, tuple):  # 1.x may return (content, structured)
        content = content[0]
    return "".join(getattr(block, "text", "") for block in content)


class _ProbeInput(BaseModel):
    value: int = 0


def _probe_server():
    server = _server_module()
    srv = server._PowerPointServer("probe")

    @srv.tool(name="ppt_probe")
    async def probe(params: _ProbeInput) -> str:
        return f"{call_presentation.get()}|{params.value}"

    @srv.tool(name="ppt_list_presentations")
    async def unbound(params: _ProbeInput) -> str:
        return f"{call_presentation.get()}"

    return srv


def test_call_tool_takes_the_argument_out_before_validation():
    srv = _probe_server()
    result = anyio.run(srv.call_tool, "ppt_probe",
                       {"params": {"value": 3}, "presentation": " b.pptx "})
    assert _text_of(result) == "b.pptx|3"


def test_call_tool_without_the_argument_sets_nothing():
    srv = _probe_server()
    token = call_presentation.set("leaked.pptx")
    try:
        result = anyio.run(srv.call_tool, "ppt_probe", {"params": {"value": 1}})
    finally:
        call_presentation.reset(token)
    assert _text_of(result) == "None|1"


def test_call_tool_resets_the_argument_afterwards():
    srv = _probe_server()
    anyio.run(srv.call_tool, "ppt_probe", {"params": {}, "presentation": "b.pptx"})
    assert call_presentation.get() is None


def test_call_tool_drops_the_argument_where_it_means_nothing():
    srv = _probe_server()
    result = anyio.run(srv.call_tool, "ppt_list_presentations",
                       {"params": {}, "presentation": "b.pptx"})
    assert _text_of(result) == "None"


def test_call_tool_refuses_a_non_string():
    srv = _probe_server()
    with pytest.raises(ValueError, match="must be a string"):
        anyio.run(srv.call_tool, "ppt_probe", {"params": {}, "presentation": 2})


def test_every_presentation_bound_tool_advertises_the_argument():
    server = _server_module()
    tools = server.mcp._tool_manager.list_tools()
    without = set(server._TOOLS_WITHOUT_PRESENTATION_ARG)
    names = {t.name for t in tools}
    assert without <= names, f"stale entries: {sorted(without - names)}"
    for tool in tools:
        props = tool.parameters.get("properties", {})
        required = tool.parameters.get("required", [])
        if tool.name in without:
            assert server.PRESENTATION_ARG not in props, tool.name
        else:
            # Type only: a description would be repeated in every schema.
            assert props.get(server.PRESENTATION_ARG) == {"type": "string"}, tool.name
            assert server.PRESENTATION_ARG not in required, tool.name
    assert server._presentation_arg_count == len(names - without)


def test_the_instructions_mention_the_argument():
    server = _server_module()
    assert "as `presentation` on every further call" in server.mcp.instructions


# ---------------------------------------------------------------------------
# macOS wrapper
# ---------------------------------------------------------------------------
@macos_only
def test_mac_the_named_deck_wins_over_the_session_target():
    from backend.mac_ae import PowerPointAppleEventWrapper

    def deck(full_name):
        pres = MagicMock()
        pres.full_name.return_value = full_name
        pres.name.return_value = full_name.rsplit("/", 1)[-1]
        return pres

    a, b = deck("/Users/x/a.pptx"), deck("/Users/x/b.pptx")
    w = PowerPointAppleEventWrapper()
    w._get_app_impl = lambda allow_launch=False: MagicMock()
    w._presentations = lambda app_ref: [a, b]
    w._target_pres_full_name = "/Users/x/a.pptx"

    assert bind_call_presentation(w._call_local, "b.pptx", w._get_pres_impl)() is b
    assert w._target_pres_full_name == "/Users/x/a.pptx"
    with pytest.raises(ValueError, match="is not open"):
        bind_call_presentation(w._call_local, "gone.pptx", w._get_pres_impl)()
