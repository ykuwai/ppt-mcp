"""Tests for target-presentation / target-window resolution (issue #183).

These cover the pure branch logic of _find_target_pres_impl,
_get_target_window_impl and _activate_target_window_impl with a mocked
Application object — no COM and no PowerPoint required.

The behaviour under test is that resolving the target presentation must NOT
activate its window: doing so on every tool call pulled PowerPoint to the
foreground and stole focus from the user.  Only the explicitly UI-dependent
path (_activate_target_window_impl) may activate.
"""

from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import MagicMock

import pytest

_src_dir = str(Path(__file__).resolve().parents[1] / "src")
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)

from utils.com_wrapper import PowerPointCOMWrapper  # noqa: E402


def _make_pres(full_name, window_count=1):
    pres = MagicMock()
    pres.FullName = full_name
    pres.Windows.Count = window_count
    windows = {i: MagicMock(name=f"{full_name}-win{i}") for i in range(1, window_count + 1)}
    pres.Windows.side_effect = lambda i: windows[i]
    return pres


def _make_app(presentations, active=None):
    """Build a mock Application whose Presentations(i) is 1-based."""
    app = MagicMock()
    app.Presentations.Count = len(presentations)
    app.Presentations.side_effect = lambda i: presentations[i - 1]
    app.Windows.Count = 1
    app.ActivePresentation = active if active is not None else (
        presentations[0] if presentations else None
    )
    return app


def _wrapper_with(app):
    w = PowerPointCOMWrapper()
    w._app = app
    w._get_app_impl = lambda allow_launch=False: app  # bypass COM liveness check
    return w


# ---------------------------------------------------------------------------
# _find_target_pres_impl
# ---------------------------------------------------------------------------
def test_find_target_returns_none_when_no_target_set():
    app = _make_app([_make_pres("C:/a.pptx")])
    w = _wrapper_with(app)
    assert w._find_target_pres_impl(app) is None


def test_find_target_matches_on_full_name():
    a, b = _make_pres("C:/a.pptx"), _make_pres("C:/b.pptx")
    app = _make_app([a, b])
    w = _wrapper_with(app)
    w._target_pres_full_name = "C:/b.pptx"
    assert w._find_target_pres_impl(app) is b


def test_find_target_clears_stale_target_when_closed():
    app = _make_app([_make_pres("C:/a.pptx")])
    w = _wrapper_with(app)
    w._target_pres_full_name = "C:/gone.pptx"
    assert w._find_target_pres_impl(app) is None
    # The stale target must be forgotten so later calls fall back cleanly.
    assert w._target_pres_full_name is None


# ---------------------------------------------------------------------------
# _get_pres_impl — must resolve WITHOUT activating (the issue #183 fix)
# ---------------------------------------------------------------------------
def test_get_pres_does_not_activate_target_window():
    a, b = _make_pres("C:/a.pptx"), _make_pres("C:/b.pptx")
    app = _make_app([a, b], active=a)
    w = _wrapper_with(app)
    w._target_pres_full_name = "C:/b.pptx"

    assert w._get_pres_impl() is b
    b.Windows(1).Activate.assert_not_called()


def test_get_pres_falls_back_to_active_presentation():
    a = _make_pres("C:/a.pptx")
    app = _make_app([a], active=a)
    w = _wrapper_with(app)
    assert w._get_pres_impl() is a


# ---------------------------------------------------------------------------
# _get_target_window_impl
# ---------------------------------------------------------------------------
def test_target_window_is_the_targets_own_window():
    a, b = _make_pres("C:/a.pptx"), _make_pres("C:/b.pptx")
    app = _make_app([a, b], active=a)
    w = _wrapper_with(app)
    w._target_pres_full_name = "C:/b.pptx"

    window = w._get_target_window_impl()
    assert window is b.Windows(1)
    window.Activate.assert_not_called()


def test_target_window_is_none_when_target_has_no_window():
    """A deck opened with with_window=False must NOT fall back to ActiveWindow:
    that window belongs to a different presentation."""
    a = _make_pres("C:/a.pptx")
    headless = _make_pres("C:/headless.pptx", window_count=0)
    app = _make_app([a, headless], active=a)
    w = _wrapper_with(app)
    w._target_pres_full_name = "C:/headless.pptx"

    assert w._get_target_window_impl() is None


def test_target_window_falls_back_to_active_window_without_target():
    a = _make_pres("C:/a.pptx")
    app = _make_app([a], active=a)
    w = _wrapper_with(app)
    assert w._get_target_window_impl() is app.ActiveWindow


def test_target_window_is_none_when_no_windows_open():
    app = _make_app([])
    app.Windows.Count = 0
    w = _wrapper_with(app)
    assert w._get_target_window_impl() is None


# ---------------------------------------------------------------------------
# _activate_target_window_impl — the one path that may steal focus
# ---------------------------------------------------------------------------
def test_activate_target_window_activates():
    a, b = _make_pres("C:/a.pptx"), _make_pres("C:/b.pptx")
    app = _make_app([a, b], active=a)
    w = _wrapper_with(app)
    w._target_pres_full_name = "C:/b.pptx"

    window = w._activate_target_window_impl()
    assert window is b.Windows(1)
    window.Activate.assert_called_once()


def test_activate_target_window_raises_when_no_window():
    headless = _make_pres("C:/headless.pptx", window_count=0)
    app = _make_app([headless], active=headless)
    w = _wrapper_with(app)
    w._target_pres_full_name = "C:/headless.pptx"

    with pytest.raises(RuntimeError, match="active PowerPoint window"):
        w._activate_target_window_impl()


def test_activate_target_window_survives_activate_failure():
    """A failing Activate() must not abort the operation."""
    a = _make_pres("C:/a.pptx")
    app = _make_app([a], active=a)
    w = _wrapper_with(app)
    w._target_pres_full_name = "C:/a.pptx"
    a.Windows(1).Activate.side_effect = Exception("busy")

    assert w._activate_target_window_impl() is a.Windows(1)
