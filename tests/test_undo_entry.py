"""One undo step per tool call (#298).

PowerPoint gathers every change made through COM since the last
Application.StartNewUndoEntry into one undo entry. Nothing started one, so a
single Undo took back everything the tools had done, slides included. Each tool
call that changes the deck now starts one before its first COM operation.
Runs anywhere, against a stand-in for the COM wrapper and PowerPoint.
"""

import sys

sys.path.insert(0, "src")

from utils.com_wrapper import UndoEntry, bind_undo_entry  # noqa: E402


class App:
    def __init__(self, fails=False):
        self.entries, self.fails = 0, fails

    def StartNewUndoEntry(self):
        if self.fails:
            raise RuntimeError("PowerPoint is busy")
        self.entries += 1


class Wrapper:
    def __init__(self, app):
        self._app = app

    def _get_app_impl(self):
        return self._app


def test_the_first_operation_of_a_call_starts_one_entry_and_the_rest_do_not():
    app = App()
    entry = UndoEntry()
    first = bind_undo_entry(Wrapper(app), entry, lambda: "added")
    second = bind_undo_entry(Wrapper(app), entry, lambda: "checked")
    assert first() == "added"
    assert second() == "checked"
    assert app.entries == 1


def test_each_call_is_its_own_step():
    app = App()
    for _ in range(3):
        bind_undo_entry(Wrapper(app), UndoEntry(), lambda: None)()
    assert app.entries == 3


def test_no_powerpoint_running_means_nothing_to_start():
    entry = UndoEntry()
    assert bind_undo_entry(Wrapper(None), entry, lambda: "created")() == "created"
    assert entry.started


def test_failing_to_start_one_never_changes_the_result():
    assert bind_undo_entry(Wrapper(App(fails=True)), UndoEntry(), lambda: 42)() == 42


def test_an_error_from_the_operation_still_propagates():
    def fails():
        raise ValueError("Shape 'Rectangle 9' not found")

    try:
        bind_undo_entry(Wrapper(App()), UndoEntry(), fails)()
    except ValueError as e:
        assert "Rectangle 9" in str(e)
    else:
        raise AssertionError("the operation's error was swallowed")


class TestWhichCallsGetOne:
    """server._undo_entry_for, against the tools the server registers."""

    @staticmethod
    def tools():
        import server

        return {t.name: t for t in server.mcp._tool_manager.list_tools()}, server

    def test_a_tool_that_changes_the_deck_gets_one(self):
        tools, server = self.tools()
        for name in ("ppt_add_slide", "ppt_add_shape", "ppt_update_shape", "ppt_set_text"):
            assert isinstance(server._undo_entry_for(name, tools[name]), UndoEntry), name

    def test_undo_redo_and_reading_do_not(self):
        tools, server = self.tools()
        for name in ("ppt_undo", "ppt_redo", "ppt_start_undo_entry"):
            assert server._undo_entry_for(name, tools[name]) is None, name
        read_only = [n for n, t in tools.items() if server._is_read_only(t)]
        assert "ppt_list_shapes" in read_only
        for name in read_only:
            assert server._undo_entry_for(name, tools[name]) is None, name
