"""Tests for presentations that other people edit at the same time.

A deck opened from SharePoint or OneDrive takes in what co-authors have saved
only when it is saved. utils/coauthoring.py merges those changes before a tool
call when nothing here is unsaved, refuses a writing call whose slides the
merge moved, reports the state under "coauthoring", and hands a conflict to a
registered resolver or leaves it to the user.

Covered here without COM and without PowerPoint, against a small stand-in for
a deck whose Save() behaves the way PowerPoint was measured to behave:

* the state of a deck and what before_call / after_save make of it;
* the conflict resolver hook;
* bind_call_coauthoring: merge, operation and report in one go;
* the COM wrapper hands the call's co-authoring to the worker;
* the server marks presentation-bound calls, with their read/write nature and
  their own target fields.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from unittest.mock import MagicMock

import anyio
import pytest
from pydantic import BaseModel

_src_dir = str(Path(__file__).resolve().parents[1] / "src")
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)

from utils import coauthoring  # noqa: E402
from utils.coauthoring import (  # noqa: E402
    CallCoauthoring,
    MovedByMerge,
    bind_call_coauthoring,
    call_coauthoring,
)
from utils.com_wrapper import PowerPointCOMWrapper  # noqa: E402

URL = "https://contoso.sharepoint.com/sites/team/Shared Documents/Deck.pptx"


# ---------------------------------------------------------------------------
# A deck on a server
# ---------------------------------------------------------------------------
class _TextRange:
    def __init__(self, text):
        self.Text = text


class _TextFrame:
    def __init__(self, text):
        self.TextRange = _TextRange(text)

    @property
    def HasText(self):
        return bool(self.TextRange.Text)


class _Shape:
    def __init__(self, name, text="", left=10.0):
        self.Name, self.Type, self.HasTextFrame = name, 14, True
        self.Left, self.Top, self.Width, self.Height = left, 10.0, 100.0, 50.0
        self.TextFrame = _TextFrame(text)


class _Shapes(list):
    @property
    def HasTitle(self):
        return bool(self)

    @property
    def Title(self):
        return self[0]


class _Slide:
    def __init__(self, slide_id, *shapes):
        self.SlideID, self.SlideIndex = slide_id, 0
        self.Shapes = _Shapes(shapes)


class _Deck:
    """Save() merges what `theirs` brings while updates are pending. With
    `conflict`, PendingUpdates stays true and the dialog is open until
    answer() applies the other authors' version."""

    def __init__(self, full_name=URL, pending=False, saved=True, theirs=None, conflict=None):
        self.FullName = full_name
        self.Name = full_name.replace("\\", "/").rsplit("/", 1)[-1]
        self.Saved = -1 if saved else 0
        self.Coauthoring = MagicMock(PendingUpdates=pending, CoauthorCount=1)
        self.slides = [
            _Slide(256 + i, _Shape("Title", t), _Shape("Body", f"Body {t}"))
            for i, t in enumerate(("A", "B"))
        ]
        self.theirs, self.conflict = theirs, conflict
        self.dialog_open = False
        self.saves = 0

    @property
    def Slides(self):
        for i, slide in enumerate(self.slides, 1):
            slide.SlideIndex = i
        return list(self.slides)

    def Save(self):
        self.saves += 1
        self.Saved = -1
        if not self.Coauthoring.PendingUpdates:
            return
        if self.conflict:
            self.dialog_open = True
            return
        if self.theirs:
            self.theirs(self)
            self.theirs = None
        self.Coauthoring.PendingUpdates = False

    def answer(self):
        self.conflict(self)
        self.conflict, self.dialog_open = None, False
        return True

    def text(self, slide, shape):
        return self.slides[slide].Shapes[shape].TextFrame.TextRange.Text


def _their_body(deck):
    deck.slides[1].Shapes[1].TextFrame.TextRange.Text = "theirs"


def _their_title_and_move(deck):
    deck.slides[0].Shapes[0].TextFrame.TextRange.Text = "Title by others"
    deck.slides[1].Shapes[1].Left = 99.0


@pytest.fixture(autouse=True)
def _clean_state():
    coauthoring.set_conflict_resolver(None)
    coauthoring._waiting_for_user.clear()
    yield
    coauthoring.set_conflict_resolver(None)
    coauthoring._waiting_for_user.clear()


# ---------------------------------------------------------------------------
# State, before_call, after_save
# ---------------------------------------------------------------------------
def test_a_local_file_has_no_state():
    deck = _Deck(r"C:\Decks\Deck.pptx", pending=True)
    assert coauthoring.status(deck) is None
    assert coauthoring.before_call(deck, writes=False) is None
    assert deck.saves == 0


def test_the_state_of_a_deck_on_a_server():
    assert coauthoring.status(_Deck(), with_count=True) == {
        "pending_updates": False, "coauthor_count": 1}
    state = coauthoring.status(_Deck(pending=True))
    assert state["pending_updates"] is True
    assert "ppt_save_presentation" in state["note"]


def test_pending_changes_are_merged_before_a_read():
    deck = _Deck(pending=True, theirs=_their_body)
    state = coauthoring.before_call(deck, writes=False)
    assert deck.saves == 1 and deck.text(1, 1) == "theirs"
    assert state["merged_updates"] is True and state["pending_updates"] is False


def test_nothing_is_merged_while_something_here_is_unsaved():
    deck = _Deck(pending=True, saved=False, theirs=_their_body)
    state = coauthoring.before_call(deck, writes=True)
    assert deck.saves == 0
    assert state["pending_updates"] is True and "merged_updates" not in state


def test_a_write_proceeds_when_the_merge_moved_nothing():
    deck = _Deck(pending=True, theirs=_their_body)
    state = coauthoring.before_call(deck, writes=True)
    assert state["merged_updates"] is True and "moved" not in state


def test_a_slide_added_at_the_end_moves_nothing():
    deck = _Deck(pending=True, theirs=lambda d: d.slides.append(_Slide(300, _Shape("New"))))
    assert "moved" not in coauthoring.before_call(deck, writes=True)


@pytest.mark.parametrize("theirs", [
    lambda d: d.slides.insert(0, _Slide(300, _Shape("New"))),
    lambda d: d.slides.pop(0),
    lambda d: d.slides[0].Shapes.insert(0, _Shape("Picture 3")),
], ids=["slide inserted before", "slide deleted", "shape inserted before"])
def test_a_write_is_refused_when_the_merge_moved_slides_or_shapes(theirs):
    deck = _Deck(pending=True, theirs=theirs)
    with pytest.raises(MovedByMerge) as refused:
        coauthoring.before_call(deck, writes=True)
    assert "Nothing was changed" in str(refused.value)
    assert refused.value.state["merged_updates"] is True


def test_a_read_is_not_refused_when_the_merge_moved_slides():
    deck = _Deck(pending=True, theirs=lambda d: d.slides.insert(0, _Slide(300)))
    assert coauthoring.before_call(deck, writes=False)["merged_updates"] is True


def test_a_conflict_without_a_resolver_is_left_to_the_user():
    deck = _Deck(pending=True, conflict=_their_title_and_move)
    state = coauthoring.before_call(deck, writes=False)
    assert state["conflict"] is True and "Resolve conflict" in state["note"]
    assert "overwritten" not in state and deck.dialog_open

    # No further save while the dialog waits for the user.
    state = coauthoring.before_call(deck, writes=False)
    assert deck.saves == 1 and state["conflict"] is True

    # Answered by the user: back to normal.
    deck.Coauthoring.PendingUpdates, deck.dialog_open = False, False
    assert coauthoring.before_call(deck, writes=False) == {"pending_updates": False}
    assert not coauthoring._waiting_for_user


def test_a_registered_resolver_answers_a_conflict():
    coauthoring.set_conflict_resolver(lambda pres: pres.answer())
    deck = _Deck(pending=True, conflict=_their_title_and_move)
    deck.slides[0].Shapes[0].TextFrame.TextRange.Text = "Title from here"

    state = coauthoring.before_call(deck, writes=False)

    assert deck.saves == 2 and deck.text(0, 0) == "Title by others"
    assert state["conflict"] is True and state["merged_updates"] is True
    assert state["pending_updates"] is False
    by_id = {entry["slide_id"]: entry for entry in state["overwritten"]}
    assert set(by_id) == {256, 257}
    assert by_id[256]["changed_shapes"] == ["Title"]
    assert by_id[256]["text_here"] == {"Title": "Title from here"}
    assert by_id[257]["changed_shapes"] == ["Body"] and by_id[257]["slide_index"] == 2


def test_a_slide_the_others_deleted_is_reported():
    coauthoring.set_conflict_resolver(lambda pres: pres.answer())
    deck = _Deck(pending=True, conflict=lambda d: d.slides.pop(1))
    entry, = coauthoring.before_call(deck, writes=False)["overwritten"]
    assert entry["removed_by_others"] is True and entry["slide_index"] is None


@pytest.mark.parametrize("resolver", [
    lambda pres: False,
    lambda pres: 1 / 0,
], ids=["declines", "fails"])
def test_a_resolver_that_does_not_answer_leaves_the_conflict_to_the_user(resolver):
    coauthoring.set_conflict_resolver(resolver)
    deck = _Deck(pending=True, conflict=_their_title_and_move)
    state = coauthoring.before_call(deck, writes=False)
    assert state["conflict"] is True and "overwritten" not in state
    assert coauthoring._waiting_for_user


def test_after_save_reports_a_merge_and_a_plain_save():
    deck = _Deck(pending=True, theirs=_their_body)
    deck.Save()
    assert coauthoring.after_save(deck, pending_before=True)["merged_updates"] is True
    deck = _Deck()
    deck.Save()
    assert coauthoring.after_save(deck, pending_before=False) == {"pending_updates": False}


def test_after_save_reports_updates_the_upload_only_now_met():
    deck = _Deck()
    deck.Save()
    deck.Coauthoring.PendingUpdates = True  # the upload met the others' changes
    state = coauthoring.after_save(deck, pending_before=False)
    assert state["pending_updates"] is True and "conflict" not in state


# ---------------------------------------------------------------------------
# bind_call_coauthoring
# ---------------------------------------------------------------------------
def _wrapper_for(*decks, target=None):
    app = MagicMock()
    app.Presentations.Count = len(decks)
    app.Presentations.side_effect = lambda i: decks[i - 1]
    wrapper = MagicMock()
    wrapper._get_app_impl.return_value = app
    wrapper._get_pres_impl.return_value = target if target is not None else decks[0]
    return wrapper


def test_a_read_sees_the_merged_state_and_says_so():
    deck = _Deck(pending=True, theirs=_their_body)
    seen = []

    def op():
        seen.append(deck.text(1, 1))
        return {"slides": 2}

    result = bind_call_coauthoring(_wrapper_for(deck), CallCoauthoring(writes=False), op)()
    assert seen == ["theirs"]
    assert result["coauthoring"]["merged_updates"] is True


def test_nothing_is_added_when_nothing_happened():
    result = bind_call_coauthoring(
        _wrapper_for(_Deck()), CallCoauthoring(writes=True), lambda: {"ok": True})()
    assert result == {"ok": True}


def test_pending_changes_are_reported_after_a_write():
    deck = _Deck(saved=False)

    def op():
        deck.Coauthoring.PendingUpdates = True  # met while uploading meanwhile
        return {"success": True}

    result = bind_call_coauthoring(_wrapper_for(deck), CallCoauthoring(writes=True), op)()
    assert result["coauthoring"]["pending_updates"] is True


def test_a_refused_write_does_not_run():
    deck = _Deck(pending=True, theirs=lambda d: d.slides.insert(0, _Slide(300)))
    op = MagicMock(return_value={})
    with pytest.raises(MovedByMerge):
        bind_call_coauthoring(_wrapper_for(deck), CallCoauthoring(writes=True), op)()
    op.assert_not_called()


def test_a_result_that_is_not_a_dict_is_left_alone():
    deck = _Deck(pending=True, theirs=_their_body)
    result = bind_call_coauthoring(
        _wrapper_for(deck), CallCoauthoring(writes=False), lambda: "# Slide 1")()
    assert result == "# Slide 1" and deck.saves == 1


def test_the_tool_s_own_target_field_names_the_deck():
    local, shared = _Deck(r"C:\Decks\Local.pptx"), _Deck(pending=True, theirs=_their_body)
    call = CallCoauthoring(writes=False, presentation_name="Deck.pptx")
    result = bind_call_coauthoring(_wrapper_for(local, shared, target=local), call,
                                   lambda: {})()
    assert shared.saves == 1 and result["coauthoring"]["merged_updates"] is True


def test_a_merge_is_reported_by_every_operation_of_the_call():
    deck = _Deck(pending=True, theirs=_their_body)
    call = CallCoauthoring(writes=True)
    wrapper = _wrapper_for(deck)
    bind_call_coauthoring(wrapper, call, lambda: {})()
    second = bind_call_coauthoring(wrapper, call, lambda: {})()
    assert second["coauthoring"]["merged_updates"] is True and deck.saves == 1


def test_a_failing_lookup_never_fails_the_operation():
    wrapper = MagicMock()
    wrapper._get_app_impl.side_effect = RuntimeError("PowerPoint is gone")
    result = bind_call_coauthoring(wrapper, CallCoauthoring(writes=True), lambda: {"ok": 1})()
    assert result == {"ok": 1}


# ---------------------------------------------------------------------------
# COM wrapper
# ---------------------------------------------------------------------------
def test_execute_hands_the_call_s_coauthoring_to_the_worker():
    w = PowerPointCOMWrapper()
    w._running = True
    w.start = lambda: None
    deck = _Deck(pending=True, theirs=_their_body)
    w._get_app_impl = lambda allow_launch=False: _wrapper_for(deck)._get_app_impl()
    w._get_pres_impl = lambda: deck

    def fake_put(item):
        func, args, kwargs, future, _idempotent, dequeued = item
        dequeued.set()
        future.set_result(func(*args, **kwargs))

    w._queue = MagicMock()
    w._queue.put.side_effect = fake_put

    token = call_coauthoring.set(CallCoauthoring(writes=False))
    try:
        result = w.execute(lambda: {})
    finally:
        call_coauthoring.reset(token)
    assert result["coauthoring"]["merged_updates"] is True
    assert w.execute(lambda: {}) == {}


# ---------------------------------------------------------------------------
# Server
# ---------------------------------------------------------------------------
class _ProbeInput(BaseModel):
    value: int = 0
    presentation_name: str | None = None


def _probe_server():
    import server

    srv = server._PowerPointServer("probe")
    seen = {}

    def probe(name, read_only):
        @srv.tool(name=name, annotations={"readOnlyHint": read_only})
        async def tool(params: _ProbeInput) -> str:
            seen[name] = call_coauthoring.get()
            return "ok"

    probe("ppt_get_probe", True)
    probe("ppt_set_probe", False)
    probe("ppt_save_presentation", False)
    probe("ppt_list_presentations", True)
    return srv, seen


def _call(srv, name, params=None):
    async def go():
        return await srv.call_tool(name, {"params": params or {}})

    return anyio.run(go)


def test_the_server_marks_reads_and_writes():
    srv, seen = _probe_server()
    _call(srv, "ppt_get_probe")
    _call(srv, "ppt_set_probe")
    assert seen["ppt_get_probe"].writes is False
    assert seen["ppt_set_probe"].writes is True
    assert call_coauthoring.get() is None


def test_the_server_leaves_out_save_and_unbound_tools():
    srv, seen = _probe_server()
    _call(srv, "ppt_save_presentation")
    _call(srv, "ppt_list_presentations")
    assert seen["ppt_save_presentation"] is None
    assert seen["ppt_list_presentations"] is None


def test_the_server_passes_the_tool_s_own_target_field():
    srv, seen = _probe_server()
    _call(srv, "ppt_get_probe", {"presentation_name": "Deck.pptx"})
    assert seen["ppt_get_probe"].presentation_name == "Deck.pptx"


def test_the_instructions_explain_shared_files():
    import server

    text = server.mcp._lowlevel_server.instructions if hasattr(
        server.mcp, "_lowlevel_server") else server.mcp.instructions
    assert "## Shared files" in text and "coauthoring" in text
