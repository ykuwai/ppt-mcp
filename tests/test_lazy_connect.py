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


def _make_busy_error():
    """Construct a pywintypes.com_error mimicking a non-busy COM failure."""
    err = pywintypes.com_error(-2147221021, "Operation unavailable", None, None)
    return err


def test_get_app_default_attach_only_raises_when_not_running():
    """_get_app_impl() with default allow_launch=False must NOT spawn PowerPoint."""
    w = PowerPointCOMWrapper()
    with patch("utils.com_wrapper.win32com.client.GetActiveObject",
               side_effect=_make_busy_error()), \
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
               side_effect=_make_busy_error()), \
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
