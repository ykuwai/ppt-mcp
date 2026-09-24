"""Undo, Redo, ExecuteMso and the selection follow the target deck (Windows).

CommandBars.ExecuteMso acts on the active window, and ppt_get_selection read
app.ActiveWindow, whichever presentation either belonged to. With two decks
open, ppt_undo could undo the last edit of the deck the target is not.

The ExecuteMso tools now bring the target's window to the front first and
refuse when it does not get there; Undo and Redo then hand the front back to
the window that had it, since they leave nothing to look at. ppt_get_selection
reads the target's own window without activating anything (issue #183). With
no target set, behaviour is unchanged.

No COM and no PowerPoint required.
"""

from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

_src_dir = str(Path(__file__).resolve().parents[1] / "src")
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)

from utils.com_wrapper import PowerPointCOMWrapper, bind_call_presentation  # noqa: E402


def _make_app(*full_names, active=0, activatable=True):
    """Mock Application with one window per deck; Activate() moves
    app.ActiveWindow unless activatable is False."""
    app = MagicMock()
    decks = []
    for full_name in full_names:
        pres = MagicMock()
        pres.FullName = full_name
        pres.Name = full_name.rsplit("/", 1)[-1]
        window = MagicMock(name=f"{full_name}-win")
        window.Presentation = pres
        if activatable:
            window.Activate.side_effect = lambda w=window: setattr(app, "ActiveWindow", w)
        pres.Windows.Count = 1
        pres.Windows.side_effect = lambda i, w=window: w
        decks.append(pres)
    app.Presentations.Count = len(decks)
    app.Presentations.side_effect = lambda i: decks[i - 1]
    app.Windows.Count = len(decks)
    app.ActiveWindow = decks[active].Windows(1)
    app.ActivePresentation = decks[active]
    return app, decks


def _wrapper_with(app):
    w = PowerPointCOMWrapper()
    w._app = app
    w._get_app_impl = lambda allow_launch=False: app
    return w


# ---------------------------------------------------------------------------
# _activate_target_window_for_command_impl
# ---------------------------------------------------------------------------
def test_without_a_target_nothing_is_activated():
    app, (a, b) = _make_app("C:/a.pptx", "C:/b.pptx", active=1)
    w = _wrapper_with(app)
    w._activate_target_window_for_command_impl()
    assert app.ActiveWindow is b.Windows(1)
    a.Windows(1).Activate.assert_not_called()


def test_the_session_target_is_brought_to_the_front():
    app, (a, b) = _make_app("C:/a.pptx", "C:/b.pptx", active=1)
    w = _wrapper_with(app)
    w._target_pres_full_name = "C:/a.pptx"
    w._activate_target_window_for_command_impl()
    assert app.ActiveWindow is a.Windows(1)


def test_the_per_call_target_is_brought_to_the_front():
    app, (a, b) = _make_app("C:/a.pptx", "C:/b.pptx", active=0)
    w = _wrapper_with(app)
    w._target_pres_full_name = "C:/a.pptx"
    bind_call_presentation(w._call_local, "b.pptx",
                           w._activate_target_window_for_command_impl)()
    assert app.ActiveWindow is b.Windows(1)


def test_a_target_that_does_not_come_to_the_front_is_refused():
    app, (a, b) = _make_app("C:/a.pptx", "C:/b.pptx", active=1, activatable=False)
    w = _wrapper_with(app)
    w._target_pres_full_name = "C:/a.pptx"
    with pytest.raises(RuntimeError, match="Nothing was run"):
        w._activate_target_window_for_command_impl()


def test_a_target_without_a_window_is_refused():
    app, (a, b) = _make_app("C:/a.pptx", "C:/b.pptx", active=1)
    a.Windows.Count = 0
    w = _wrapper_with(app)
    w._target_pres_full_name = "C:/a.pptx"
    with pytest.raises(RuntimeError, match="has none"):
        w._activate_target_window_for_command_impl()


def test_the_window_that_gave_way_is_returned():
    app, (a, b) = _make_app("C:/a.pptx", "C:/b.pptx", active=1)
    w = _wrapper_with(app)
    w._target_pres_full_name = "C:/a.pptx"
    assert w._activate_target_window_for_command_impl() is b.Windows(1)


def test_nothing_is_returned_when_the_target_was_already_in_front():
    app, (a, b) = _make_app("C:/a.pptx", "C:/b.pptx", active=0)
    w = _wrapper_with(app)
    w._target_pres_full_name = "C:/a.pptx"
    assert w._activate_target_window_for_command_impl() is None


def test_nothing_is_returned_without_a_target():
    app, (a, b) = _make_app("C:/a.pptx", "C:/b.pptx", active=1)
    w = _wrapper_with(app)
    assert w._activate_target_window_for_command_impl() is None


# ---------------------------------------------------------------------------
# The tools
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("impl, args, command", [
    ("_undo_impl", (1,), "Undo"),
    ("_redo_impl", (1,), "Redo"),
    ("_execute_mso_impl", ("Bold", False), "Bold"),
])
def test_commandbars_run_in_the_target_window(impl, args, command):
    import ppt_com.edit_ops as edit_ops

    app, (a, b) = _make_app("C:/a.pptx", "C:/b.pptx", active=1)
    w = _wrapper_with(app)
    w._target_pres_full_name = "C:/a.pptx"
    window_at_command = []
    app.CommandBars.GetEnabledMso.return_value = True
    app.CommandBars.ExecuteMso.side_effect = (
        lambda name: window_at_command.append((name, app.ActiveWindow))
    )

    with patch.object(edit_ops, "ppt", w):
        getattr(edit_ops, impl)(*args)

    assert window_at_command == [(command, a.Windows(1))]


@pytest.mark.parametrize("impl", ["_undo_impl", "_redo_impl"])
def test_undo_and_redo_hand_the_front_back(impl):
    import ppt_com.edit_ops as edit_ops

    app, (a, b) = _make_app("C:/a.pptx", "C:/b.pptx", active=1)
    w = _wrapper_with(app)
    w._target_pres_full_name = "C:/a.pptx"
    app.CommandBars.GetEnabledMso.return_value = True

    with patch.object(edit_ops, "ppt", w):
        getattr(edit_ops, impl)(2)

    assert app.ActiveWindow is b.Windows(1)


def test_the_front_is_handed_back_when_the_command_fails():
    import ppt_com.edit_ops as edit_ops

    app, (a, b) = _make_app("C:/a.pptx", "C:/b.pptx", active=1)
    w = _wrapper_with(app)
    w._target_pres_full_name = "C:/a.pptx"
    app.CommandBars.GetEnabledMso.return_value = True
    app.CommandBars.ExecuteMso.side_effect = RuntimeError("boom")

    with patch.object(edit_ops, "ppt", w), pytest.raises(RuntimeError, match="boom"):
        edit_ops._undo_impl(1)
    assert app.ActiveWindow is b.Windows(1)


def test_undo_moves_nothing_when_the_target_was_already_in_front():
    import ppt_com.edit_ops as edit_ops

    app, (a, b) = _make_app("C:/a.pptx", "C:/b.pptx", active=0)
    w = _wrapper_with(app)
    w._target_pres_full_name = "C:/a.pptx"
    app.CommandBars.GetEnabledMso.return_value = True

    with patch.object(edit_ops, "ppt", w):
        edit_ops._undo_impl(1)

    assert app.ActiveWindow is a.Windows(1)
    b.Windows(1).Activate.assert_not_called()


def test_execute_mso_leaves_the_target_in_front():
    # A ribbon command can open a pane or a dialog on the target's window,
    # or start its slide show, so the front is not handed back there.
    import ppt_com.edit_ops as edit_ops

    app, (a, b) = _make_app("C:/a.pptx", "C:/b.pptx", active=1)
    w = _wrapper_with(app)
    w._target_pres_full_name = "C:/a.pptx"

    with patch.object(edit_ops, "ppt", w):
        edit_ops._execute_mso_impl("Bold", False)

    assert app.ActiveWindow is a.Windows(1)


def test_commandbars_are_not_run_when_the_target_cannot_be_reached():
    import ppt_com.edit_ops as edit_ops

    app, (a, b) = _make_app("C:/a.pptx", "C:/b.pptx", active=1, activatable=False)
    w = _wrapper_with(app)
    w._target_pres_full_name = "C:/a.pptx"
    app.CommandBars.GetEnabledMso.return_value = True

    with patch.object(edit_ops, "ppt", w), pytest.raises(RuntimeError):
        edit_ops._undo_impl(1)
    app.CommandBars.ExecuteMso.assert_not_called()


def test_the_selection_is_read_from_the_target_window():
    import ppt_com.advanced_ops as advanced_ops

    app, (a, b) = _make_app("C:/a.pptx", "C:/b.pptx", active=1)
    w = _wrapper_with(app)
    w._target_pres_full_name = "C:/a.pptx"
    a.Windows(1).Selection.Type = advanced_ops.ppSelectionText
    a.Windows(1).Selection.TextRange.Text = "from a"
    b.Windows(1).Selection.Type = advanced_ops.ppSelectionText
    b.Windows(1).Selection.TextRange.Text = "from b"

    with patch.object(advanced_ops, "ppt", w):
        result = advanced_ops._get_selection_impl()

    assert result["text"] == "from a"
    a.Windows(1).Activate.assert_not_called()
