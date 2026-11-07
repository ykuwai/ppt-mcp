"""execute() before the worker is started fails at once, and says why.

Only the server's lifespan starts the worker (``ppt.start()``). Code that
reaches the tools without running the server, a test or a script calling
``server.mcp.call_tool`` directly, used to queue onto nothing, wait out the
whole queue wait (60 s on Windows), and then report that "earlier operations
were still running" when there were none.

No PowerPoint is needed: nothing is queued, and nothing runs.
"""

import sys
import time

import pytest

sys.path.insert(0, "src")

NOT_STARTED = "has not been started"


def _fails_at_once(wrapper):
    ran = []
    start = time.monotonic()
    with pytest.raises(RuntimeError, match=NOT_STARTED) as info:
        wrapper.execute(lambda: ran.append("ran"))
    assert time.monotonic() - start < 1.0
    assert ran == []
    assert wrapper._queue.empty(), "nothing is left queued for a later worker"
    assert "ppt.start()" in str(info.value)


def test_the_com_wrapper_fails_at_once():
    from utils.com_wrapper import PowerPointCOMWrapper

    _fails_at_once(PowerPointCOMWrapper())


@pytest.mark.skipif(sys.platform != "darwin", reason="the Apple Event backend needs appscript")
def test_the_apple_event_wrapper_fails_at_once():
    from backend.mac_ae import PowerPointAppleEventWrapper

    _fails_at_once(PowerPointAppleEventWrapper())


@pytest.mark.skipif(sys.platform != "darwin", reason="the Apple Event backend needs appscript")
def test_the_apple_event_wrapper_runs_once_started():
    from backend.mac_ae import PowerPointAppleEventWrapper

    wrapper = PowerPointAppleEventWrapper()
    wrapper.start()
    try:
        assert wrapper.execute(lambda: "ran") == "ran"
    finally:
        wrapper.stop()


def test_the_com_wrapper_runs_once_started(monkeypatch):
    import utils.com_wrapper as com
    from unittest.mock import MagicMock

    # The worker initialises COM on its own thread; there is none to use here.
    monkeypatch.setattr(com, "pythoncom", MagicMock())
    wrapper = com.PowerPointCOMWrapper()
    wrapper.start()
    try:
        assert wrapper.execute(lambda: "ran") == "ran"
    finally:
        wrapper.stop()
