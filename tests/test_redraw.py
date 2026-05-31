"""Tests for src/utils/redraw.py.

Pure Python tests — no COM or PowerPoint dependency required. win32gui is
patched so the tests exercise the context-manager logic deterministically on
any platform, including CI without pywin32.
"""

import sys

sys.path.insert(0, "src")

import pytest

import utils.redraw as redraw
from utils.redraw import frozen_redraw


class _FakeWin32:
    """Records LockWindowUpdate calls; mimics the Win32 return convention."""

    def __init__(self, hwnd=12345, lock_succeeds=True):
        self._hwnd = hwnd
        self._lock_succeeds = lock_succeeds
        self.lock_calls = []

    def FindWindow(self, cls, title):
        return self._hwnd

    def LockWindowUpdate(self, hwnd):
        self.lock_calls.append(hwnd)
        # Win32: locking returns nonzero on success; unlocking (hwnd=0) is the
        # release call.
        if hwnd == 0:
            return 1
        return 1 if self._lock_succeeds else 0


@pytest.fixture
def fake_win32(monkeypatch):
    fake = _FakeWin32()
    monkeypatch.setattr(redraw, "_WIN32_AVAILABLE", True)
    monkeypatch.setattr(redraw, "win32gui", fake, raising=False)
    return fake


def test_locks_and_unlocks_around_block(fake_win32):
    with frozen_redraw():
        pass
    # First call locks the found window, last call (0) releases it.
    assert fake_win32.lock_calls == [12345, 0]


def test_unlocks_even_when_block_raises(fake_win32):
    with pytest.raises(ValueError):
        with frozen_redraw():
            raise ValueError("boom")
    # The lock must always be released so the editor never stays frozen.
    assert fake_win32.lock_calls == [12345, 0]


def test_does_not_suppress_exceptions(fake_win32):
    with pytest.raises(KeyError):
        with frozen_redraw():
            raise KeyError("propagate me")


def test_no_unlock_when_lock_fails(monkeypatch):
    fake = _FakeWin32(lock_succeeds=False)
    monkeypatch.setattr(redraw, "_WIN32_AVAILABLE", True)
    monkeypatch.setattr(redraw, "win32gui", fake, raising=False)
    with frozen_redraw():
        pass
    # Lock attempt was made, but since it failed we must NOT call unlock(0)
    # (that would release some other window's lock).
    assert fake.lock_calls == [12345]


def test_noop_when_window_not_found(monkeypatch):
    fake = _FakeWin32(hwnd=0)
    monkeypatch.setattr(redraw, "_WIN32_AVAILABLE", True)
    monkeypatch.setattr(redraw, "win32gui", fake, raising=False)
    with frozen_redraw():
        pass
    # hwnd 0 means "not found" -> never attempt to lock.
    assert fake.lock_calls == []


def test_noop_when_win32_unavailable(monkeypatch):
    monkeypatch.setattr(redraw, "_WIN32_AVAILABLE", False)
    # Must not raise even though win32gui may be absent.
    with frozen_redraw() as fr:
        assert fr.hwnd == 0
