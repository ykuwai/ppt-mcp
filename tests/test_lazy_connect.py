"""Tests for lazy-connect behavior (issue #148).

Verifies that the MCP server does not launch PowerPoint.exe at startup, and
that _connect_impl / _get_app_impl honor the allow_launch flag.
"""

from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import patch, MagicMock

import pytest
import pywintypes

_src_dir = str(Path(__file__).resolve().parents[1] / "src")
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)

from utils.com_wrapper import PowerPointCOMWrapper  # noqa: E402


def _make_not_running_error():
    """Construct a pywintypes.com_error indicating PowerPoint is not running.

    HRESULT -2147221021 is MK_E_UNAVAILABLE — PowerPoint is not in the Running
    Object Table. This is deliberately NOT one of the busy/retry HRESULTs in
    _BUSY_HRESULTS, so the wrapper treats it as a "no instance" signal rather
    than retrying.
    """
    err = pywintypes.com_error(-2147221021, "Operation unavailable", None, None)
    return err


def test_get_app_default_attach_only_raises_when_not_running():
    """_get_app_impl() with default allow_launch=False must NOT spawn PowerPoint."""
    w = PowerPointCOMWrapper()
    with patch("utils.com_wrapper.win32com.client.GetActiveObject",
               side_effect=_make_not_running_error()), \
         patch("utils.com_wrapper.win32com.client.Dispatch") as dispatch_mock:
        with pytest.raises(ConnectionError) as exc:
            w._get_app_impl()
        # Crucial: Dispatch (which spawns PowerPoint.exe) must never be called.
        dispatch_mock.assert_not_called()
        assert "PowerPoint is not running" in str(exc.value)


def test_get_app_with_allow_launch_calls_dispatch():
    """allow_launch=True must fall back to Dispatch when no instance running."""
    w = PowerPointCOMWrapper()
    fake_app = MagicMock()
    fake_app.Visible = False
    with patch("utils.com_wrapper.win32com.client.GetActiveObject",
               side_effect=_make_not_running_error()), \
         patch("utils.com_wrapper.win32com.client.Dispatch",
               return_value=fake_app) as dispatch_mock:
        app = w._get_app_impl(allow_launch=True)
        dispatch_mock.assert_called_once_with("PowerPoint.Application")
        assert app is fake_app
        # We launched it ourselves, so visibility should be forced True.
        assert fake_app.Visible is True


def test_attach_does_not_force_visibility():
    """Attaching to a hidden running instance must NOT yank it to foreground."""
    w = PowerPointCOMWrapper()
    fake_app = MagicMock()
    fake_app.Visible = False  # user had it hidden
    with patch("utils.com_wrapper.win32com.client.GetActiveObject",
               return_value=fake_app), \
         patch("utils.com_wrapper.win32com.client.Dispatch") as dispatch_mock:
        app = w._connect_impl()
        dispatch_mock.assert_not_called()
        assert app is fake_app
        # Critical: we attached, so we must leave Visible alone.
        assert fake_app.Visible is False


def test_explicit_visible_true_is_still_honored_on_attach():
    """When the caller explicitly passes visible=True, respect that on attach too."""
    w = PowerPointCOMWrapper()
    fake_app = MagicMock()
    fake_app.Visible = False
    with patch("utils.com_wrapper.win32com.client.GetActiveObject",
               return_value=fake_app):
        w._connect_impl(visible=True)
        assert fake_app.Visible is True


def test_stale_app_reconnects_with_attach_only_when_powerpoint_stopped():
    """If the cached _app is stale and PowerPoint actually stopped,
    _get_app_impl() with default allow_launch=False must surface a clear
    ConnectionError instead of silently spawning a new PowerPoint.
    """
    w = PowerPointCOMWrapper()
    stale_app = MagicMock()
    # Accessing .Name on a stale COM object raises a non-busy com_error.
    type(stale_app).Name = property(lambda self: (_ for _ in ()).throw(_make_not_running_error()))
    w._app = stale_app
    with patch("utils.com_wrapper.win32com.client.GetActiveObject",
               side_effect=_make_not_running_error()), \
         patch("utils.com_wrapper.win32com.client.Dispatch") as dispatch_mock:
        with pytest.raises(ConnectionError):
            w._get_app_impl()
        dispatch_mock.assert_not_called()


def test_open_presentation_forces_visible_when_powerpoint_running_hidden():
    """ppt_open_presentation must surface PowerPoint to the user, even if
    a previously running instance was hidden. Otherwise the user invokes
    the tool but sees nothing happen (issue #149 review feedback).
    """
    import os
    from ppt_com import presentation

    fake_app = MagicMock()
    fake_app.Visible = False  # running but hidden
    fake_pres = MagicMock()
    fake_pres.Name = "X.pptx"
    fake_pres.FullName = "C:\\X.pptx"
    fake_pres.Slides.Count = 1
    fake_pres.ReadOnly = 0
    fake_app.Presentations.Open.return_value = fake_pres
    fake_app.Presentations.Count = 1
    fake_app.Presentations.return_value = fake_pres  # for the index loop

    with patch("ppt_com.presentation.ppt._get_app_impl", return_value=fake_app), \
         patch.object(os.path, "exists", return_value=True):
        presentation._open_presentation_impl(
            file_path="C:\\X.pptx", read_only=False, with_window=True, activate=True
        )
    assert fake_app.Visible is True, (
        "ppt_open_presentation must make PowerPoint visible — "
        "otherwise the user invokes the tool but sees nothing."
    )


def test_open_presentation_preserves_hidden_when_with_window_false():
    """ppt_open_presentation with with_window=False is the headless workflow.
    The caller explicitly asked not to surface a window, so don't override
    Visible. (PR #149 review nit.)
    """
    import os
    from ppt_com import presentation

    fake_app = MagicMock()
    fake_app.Visible = False
    fake_pres = MagicMock()
    fake_pres.Name = "X.pptx"
    fake_pres.FullName = "C:\\X.pptx"
    fake_pres.Slides.Count = 1
    fake_pres.ReadOnly = 0
    fake_app.Presentations.Open.return_value = fake_pres
    fake_app.Presentations.Count = 1
    fake_app.Presentations.return_value = fake_pres

    with patch("ppt_com.presentation.ppt._get_app_impl", return_value=fake_app), \
         patch.object(os.path, "exists", return_value=True):
        presentation._open_presentation_impl(
            file_path="C:\\X.pptx", read_only=False, with_window=False, activate=True
        )
    assert fake_app.Visible is False, (
        "with_window=False is the headless workflow — Visible must not be forced."
    )


def test_create_presentation_forces_visible_when_powerpoint_running_hidden():
    """Same as above for ppt_create_presentation."""
    from ppt_com import presentation

    fake_app = MagicMock()
    fake_app.Visible = False
    fake_pres = MagicMock()
    fake_pres.Name = "Untitled.pptx"
    fake_pres.Slides.Count = 0
    fake_pres.PageSetup.SlideWidth = 960
    fake_pres.PageSetup.SlideHeight = 540
    fake_pres.TemplateName = ""
    fake_app.Presentations.Add.return_value = fake_pres
    fake_app.Presentations.Count = 1
    fake_app.Presentations.return_value = fake_pres

    with patch("ppt_com.presentation.ppt._get_app_impl", return_value=fake_app):
        presentation._create_presentation_impl(
            template_path=None, slide_width=None, slide_height=None, preset=None, activate=True
        )
    assert fake_app.Visible is True


def test_create_presentation_activate_false_does_not_set_target():
    """activate=False must not overwrite an existing session target.

    Issue #155 review: the Pydantic tests cover the schema, but the
    behavioral guarantee that opt-out preserves the target needs its own test.
    """
    from ppt_com import presentation

    fake_app = MagicMock()
    fake_app.Visible = True
    fake_pres = MagicMock()
    fake_pres.Name = "New.pptx"
    fake_pres.FullName = "C:\\New.pptx"
    fake_pres.Slides.Count = 0
    fake_pres.PageSetup.SlideWidth = 960
    fake_pres.PageSetup.SlideHeight = 540
    fake_pres.TemplateName = ""
    fake_app.Presentations.Add.return_value = fake_pres
    fake_app.Presentations.Count = 1
    fake_app.Presentations.return_value = fake_pres

    sentinel = "C:\\Existing.pptx"
    with patch("ppt_com.presentation.ppt._get_app_impl", return_value=fake_app), \
         patch.object(presentation.ppt, "_target_pres_full_name", sentinel):
        presentation._create_presentation_impl(
            template_path=None, slide_width=None, slide_height=None,
            preset=None, activate=False,
        )
        assert presentation.ppt._target_pres_full_name == sentinel, (
            "activate=False must leave the existing session target untouched"
        )


def test_open_presentation_activate_false_does_not_set_target():
    """Parallel test for ppt_open_presentation."""
    import os
    from ppt_com import presentation

    fake_app = MagicMock()
    fake_app.Visible = True
    fake_pres = MagicMock()
    fake_pres.Name = "X.pptx"
    fake_pres.FullName = "C:\\X.pptx"
    fake_pres.Slides.Count = 1
    fake_pres.ReadOnly = 0
    fake_app.Presentations.Open.return_value = fake_pres
    fake_app.Presentations.Count = 1
    fake_app.Presentations.return_value = fake_pres

    sentinel = "C:\\Existing.pptx"
    with patch("ppt_com.presentation.ppt._get_app_impl", return_value=fake_app), \
         patch.object(os.path, "exists", return_value=True), \
         patch.object(presentation.ppt, "_target_pres_full_name", sentinel):
        presentation._open_presentation_impl(
            file_path="C:\\X.pptx", read_only=False, with_window=True, activate=False
        )
        assert presentation.ppt._target_pres_full_name == sentinel, (
            "activate=False must leave the existing session target untouched"
        )


def test_create_presentation_activate_true_sets_target():
    """Confirms activate=True sets _target_pres_full_name to the new deck's FullName."""
    from ppt_com import presentation

    fake_app = MagicMock()
    fake_app.Visible = True
    fake_pres = MagicMock()
    fake_pres.Name = "New.pptx"
    fake_pres.FullName = "C:\\New.pptx"
    fake_pres.Slides.Count = 0
    fake_pres.PageSetup.SlideWidth = 960
    fake_pres.PageSetup.SlideHeight = 540
    fake_pres.TemplateName = ""
    fake_app.Presentations.Add.return_value = fake_pres
    fake_app.Presentations.Count = 1
    fake_app.Presentations.return_value = fake_pres

    with patch("ppt_com.presentation.ppt._get_app_impl", return_value=fake_app), \
         patch.object(presentation.ppt, "_target_pres_full_name", None):
        presentation._create_presentation_impl(
            template_path=None, slide_width=None, slide_height=None,
            preset=None, activate=True,
        )
        assert presentation.ppt._target_pres_full_name == "C:\\New.pptx"


def test_open_presentation_activate_true_sets_target():
    """Parallel positive-path test for ppt_open_presentation."""
    import os
    from ppt_com import presentation

    fake_app = MagicMock()
    fake_app.Visible = True
    fake_pres = MagicMock()
    fake_pres.Name = "Opened.pptx"
    fake_pres.FullName = "C:\\Opened.pptx"
    fake_pres.Slides.Count = 1
    fake_pres.ReadOnly = 0
    fake_app.Presentations.Open.return_value = fake_pres
    fake_app.Presentations.Count = 1
    fake_app.Presentations.return_value = fake_pres

    with patch("ppt_com.presentation.ppt._get_app_impl", return_value=fake_app), \
         patch.object(os.path, "exists", return_value=True), \
         patch.object(presentation.ppt, "_target_pres_full_name", None):
        presentation._open_presentation_impl(
            file_path="C:\\Opened.pptx", read_only=False, with_window=True, activate=True
        )
        assert presentation.ppt._target_pres_full_name == "C:\\Opened.pptx"


def test_server_lifespan_does_not_eager_connect():
    """Sanity check: server.app_lifespan must not call ppt.connect/_connect_impl."""
    from pathlib import Path
    server_src = (Path(__file__).resolve().parents[1] / "src" / "server.py").read_text(encoding="utf-8")
    # Inside app_lifespan, there must be no ppt.connect()/_connect_impl() call.
    lifespan_start = server_src.index("async def app_lifespan")
    lifespan_end = server_src.index("mcp = FastMCP(", lifespan_start)
    lifespan_block = server_src[lifespan_start:lifespan_end]
    assert "ppt.connect(" not in lifespan_block
    assert "_connect_impl(" not in lifespan_block
