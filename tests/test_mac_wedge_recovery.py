"""Recovering from an Apple Event worker stuck inside one job (issue #245).

An Apple Event that never answers held the macOS worker for as long as it
took, and every call queued behind it waited too. Those callers were told
after two minutes that PowerPoint was still busy, which did not say that the
worker itself was stuck, and nothing after them could run either.

The Windows wrapper already notices this and replaces the worker (#199). This
is the same design on the macOS worker: a job that has used up its whole call
budget since the worker claimed it is a wedge, the generation is bumped, a
fresh worker takes over a fresh queue, and everything the stuck one had not
started is taken back and its callers told why. The stuck worker is left to
finish or not, and the generation keeps it off shared state when it does.

No PowerPoint required. The stuck job is a function waiting on an event.
"""

import sys
import threading
import time

import pytest

sys.path.insert(0, "src")

pytestmark = pytest.mark.skipif(
    sys.platform != "darwin", reason="the Apple Event backend needs appscript"
)

# Short enough to keep the suite quick, long enough that a healthy job in a
# loaded test run does not look stuck.
BUDGET = 0.3


@pytest.fixture
def ppt(monkeypatch):
    from backend import mac_ae

    monkeypatch.setattr(mac_ae, "_CALL_BUDGET", BUDGET)
    wrapper = mac_ae.PowerPointAppleEventWrapper()
    wrapper.start()
    yield wrapper
    wrapper.stop()


def _stuck_job():
    """A job that holds the worker until released, and says when it is in."""
    entered = threading.Event()
    release = threading.Event()

    def stuck():
        entered.set()
        release.wait(10)
        return "came free"

    return stuck, entered, release


def _in_background(ppt, func):
    """Run execute on another thread and keep what it raised or returned."""
    outcome = {}

    def run():
        start = time.monotonic()
        try:
            outcome["value"] = ppt.execute(func)
        except Exception as e:  # noqa: BLE001 - kept for the assertions
            outcome["error"] = e
        outcome["took"] = time.monotonic() - start

    thread = threading.Thread(target=run, daemon=True)
    thread.start()
    return thread, outcome


# --- noticing the wedge and replacing the worker ---------------------------

def test_the_next_call_is_served_after_a_wedge(ppt):
    """The whole point. One stuck job used to hold up everything after it."""
    from backend.mac_ae import AppleEventError

    stuck, entered, release = _stuck_job()
    try:
        with pytest.raises(AppleEventError) as caught:
            ppt.execute(stuck)

        assert entered.is_set(), "the worker really was inside the job"
        assert "fresh one has taken over" in str(caught.value)
        assert "check the deck" in str(caught.value)
        assert ppt._generation == 1
        assert ppt.execute(lambda: "served") == "served"
    finally:
        release.set()


def test_a_caller_queued_behind_the_wedge_is_told_so_and_its_work_never_runs(
        ppt, monkeypatch):
    """It used to wait out the whole queue budget and then hear that
    PowerPoint was busy. Now it hears at once that an earlier request got
    stuck, and its work is not run when the stuck one comes free."""
    from backend import mac_ae

    monkeypatch.setattr(mac_ae, "_QUEUE_WAIT", 30.0)
    stuck, entered, release = _stuck_job()
    ran = []
    try:
        first, _ = _in_background(ppt, stuck)
        assert entered.wait(5)
        behind, outcome = _in_background(ppt, lambda: ran.append("edited"))
        behind.join(5)

        assert "error" in outcome, "the caller behind the wedge never returned"
        message = str(outcome["error"])
        assert "got stuck" in message
        assert "Nothing was changed" in message
        assert "one after another" not in message, "it was told PowerPoint was busy"
        assert outcome["took"] < 5, "it waited out the queue budget"

        release.set()
        first.join(5)
        # The new worker is free, and the old one has come free too. Give
        # both every chance to run the job that was taken back.
        assert ppt.execute(lambda: "served") == "served"
        time.sleep(0.1)
        assert ran == []
    finally:
        release.set()


def test_a_job_that_finished_in_the_race_is_returned_not_treated_as_a_wedge(ppt):
    """A function that raises TimeoutError reaches execute as the same class
    the wait raises. It is the job's outcome, not a stuck worker."""

    def raises_timeout():
        raise TimeoutError("the job's own")

    with pytest.raises(TimeoutError, match="the job's own"):
        ppt.execute(raises_timeout)
    assert ppt._generation == 0


def test_a_slow_call_in_front_is_not_charged_to_the_one_behind(ppt, monkeypatch):
    """Time spent in the queue is not time inside the job. A caller that waits
    its turn behind a healthy job must not read that wait as a wedge."""
    from backend import mac_ae

    monkeypatch.setattr(mac_ae, "_QUEUE_WAIT", 5.0)
    release = threading.Event()
    holding = threading.Event()

    first, _ = _in_background(
        ppt, lambda: (holding.set(), release.wait(BUDGET * 0.6))
    )
    assert holding.wait(5)
    # Queued now, and served after more than one whole budget from here.
    second, outcome = _in_background(
        ppt, lambda: (time.sleep(BUDGET * 0.6), "landed")[1]
    )
    first.join(5)
    second.join(5)

    assert outcome.get("value") == "landed", outcome
    assert outcome["took"] > BUDGET
    assert ppt._generation == 0


# --- one replacement, however many callers notice --------------------------

def test_only_the_first_caller_of_a_generation_replaces_it(ppt):
    assert ppt._recover_from_wedge(0) is True
    assert ppt._recover_from_wedge(0) is False
    assert ppt._generation == 1


def test_nothing_is_replaced_once_stopped(ppt):
    ppt.stop()
    assert ppt._recover_from_wedge(0) is False
    assert ppt._generation == 0


def test_the_old_queue_is_left_with_a_sentinel(ppt):
    """A stuck worker that comes free goes back to its queue. Without the
    sentinel it would wait there for good."""
    old_queue = ppt._queue
    ppt._recover_from_wedge(0)
    assert old_queue.get_nowait() is None
    assert ppt._queue is not old_queue


# --- what the abandoned worker may and may not do --------------------------

def test_an_abandoned_worker_does_not_retry_or_clear_the_live_reference(ppt):
    """A job safe to repeat is normally retried after a lost connection, and
    the reference cleared. A worker that has been replaced must do neither:
    the retry would be a second Apple Event alongside the new worker's, and
    the reference is the new worker's now."""
    from backend.mac_ae import AE_CONNECTION_INVALID

    entered = threading.Event()
    release = threading.Event()
    calls = []

    def stuck_then_lost():
        calls.append(1)
        entered.set()
        release.wait(10)
        raise _command_error(AE_CONNECTION_INVALID)

    live = object()
    ppt._app = live
    thread, _ = _in_background(ppt, lambda: None)
    thread.join(5)

    old_thread = ppt._thread
    ppt._queue.put(_job(stuck_then_lost, idempotent=True))
    assert entered.wait(5)
    ppt._recover_from_wedge(0)

    release.set()
    old_thread.join(5)

    assert not old_thread.is_alive(), "the abandoned worker did not exit"
    assert calls == [1], "the abandoned worker sent the job again"
    assert ppt._app is live


def test_an_abandoned_worker_drops_a_job_that_arrives_late(ppt):
    """The old queue is emptied on replacement, but a job can reach the old
    worker in the gap. It must be taken back, not run."""
    from backend import mac_ae

    ppt.stop()
    ppt._running = True

    class ReplacedOnGet(mac_ae.Queue):
        def get(self, *args, **kwargs):
            item = super().get(*args, **kwargs)
            ppt._generation += 1  # replaced while this job sat in the queue
            return item

    queue = ReplacedOnGet()
    ran = []
    job = _job(lambda: ran.append("edited"))
    queue.put(job)

    ppt._worker(queue, 0)

    assert ran == []
    assert job.dropped and job.behind_a_wedge


def _job(func, idempotent=False):
    from backend.mac_ae import _Job

    return _Job(func, (), {}, idempotent)


def _command_error(number):
    """An appscript CommandError carrying a given number, as in
    test_mac_backend. `errornumber` is read-only on the real class."""
    from appscript.reference import CommandError

    class _Stub(CommandError):
        def __init__(self):
            Exception.__init__(self, f"stub error {number}")

        @property
        def errornumber(self):
            return number

        def __str__(self):
            return f"stub error {number}"

    return _Stub()


# --- deliberate shutdown and restart ---------------------------------------

def test_stop_after_a_wedge_stops_the_live_worker(ppt):
    ppt._recover_from_wedge(0)
    live = ppt._thread
    ppt.stop()
    live.join(5)
    assert not live.is_alive()


def test_a_restart_is_a_fresh_start(ppt):
    ppt._recover_from_wedge(0)
    ppt.stop()
    ppt.start()
    assert ppt._generation == 0
    assert ppt.execute(lambda: "served") == "served"
