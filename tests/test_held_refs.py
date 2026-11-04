"""Excel references are kept until their Excel has ended, then let go of (#299)."""

import sys
import threading
from concurrent.futures import Future
from pathlib import Path
from queue import Queue
from types import SimpleNamespace

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from utils.held_refs import HeldUntilEnded  # noqa: E402


class Reference:
    """Stands in for a COM reference, and says when it is let go of."""

    def __init__(self, gone):
        self.gone = gone

    def __del__(self):
        self.gone.append(threading.get_ident())


class Processes:
    def __init__(self):
        self.ended = set()

    def __call__(self, pid):
        return pid in self.ended


def test_references_wait_until_their_process_has_ended():
    processes, gone = Processes(), []
    held = HeldUntilEnded(processes)
    held.hold([Reference(gone), Reference(gone)], 1234)
    assert held.let_go_of_ended() == 0
    assert not gone and held.still_held() == 1
    processes.ended.add(1234)
    assert held.let_go_of_ended() == 1
    assert gone == [threading.get_ident()] * 2
    assert held.still_held() == 0


def test_with_no_process_known_they_go_at_once():
    gone = []
    held = HeldUntilEnded(Processes())
    held.hold([Reference(gone)], None)
    assert len(gone) == 1
    assert held.still_held() == 0


def test_nothing_to_hold_holds_nothing():
    held = HeldUntilEnded(Processes())
    held.hold([None, None], 1234)
    assert held.still_held() == 0


def test_each_set_goes_when_its_own_process_ends():
    processes, gone_a, gone_b = Processes(), [], []
    held = HeldUntilEnded(processes)
    held.hold([Reference(gone_a)], 1)
    held.hold([Reference(gone_b)], 2)
    processes.ended.add(2)
    assert held.let_go_of_ended() == 1
    assert not gone_a and len(gone_b) == 1
    assert held.still_held() == 1


def test_another_threads_references_are_left_alone():
    """A COM reference belongs to the thread that made it, so a worker never
    lets go of what an abandoned one holds."""
    processes, gone = Processes(), []
    held = HeldUntilEnded(processes)
    thread = threading.Thread(target=lambda: held.hold([Reference(gone)], 7))
    thread.start()
    thread.join()
    processes.ended.add(7)
    assert held.let_go_of_ended() == 0
    assert not gone and held.still_held() == 1


def test_a_failure_to_ask_is_asked_again_later():
    gone, answers = [], iter([OSError("no handle"), True])

    def ended(pid):
        answer = next(answers)
        if isinstance(answer, Exception):
            raise answer
        return answer

    held = HeldUntilEnded(ended)
    held.hold([Reference(gone)], 9)
    assert held.let_go_of_ended() == 0
    assert held.let_go_of_ended() == 1
    assert len(gone) == 1


def test_at_the_end_what_has_not_ended_stays_held():
    processes, gone_a, gone_b = Processes(), [], []
    held = HeldUntilEnded(processes)
    held.hold([Reference(gone_a)], 1)
    held.hold([Reference(gone_b)], 2)
    processes.ended.add(1)
    assert held.leave_behind() == 1
    assert len(gone_a) == 1 and not gone_b
    assert held.still_held() == 1


@pytest.mark.skipif(sys.platform != "win32", reason="asks Windows about a process")
def test_this_process_has_not_ended_and_a_made_up_one_has():
    import os

    from utils.held_refs import process_ended

    assert process_ended(os.getpid()) is False
    assert process_ended(0x7FFFFFF0) is True


@pytest.mark.skipif(sys.platform != "win32", reason="runs the STA COM worker")
def test_the_worker_lets_go_before_each_item_and_leaves_the_rest_at_its_end(monkeypatch):
    from utils import com_wrapper, held_refs

    calls = []
    fake = SimpleNamespace(
        let_go_of_ended=lambda: calls.append(("let go", threading.get_ident())),
        leave_behind=lambda: calls.append(("leave", threading.get_ident())),
    )
    monkeypatch.setattr(held_refs, "held", fake)

    w = com_wrapper.PowerPointCOMWrapper()
    w._running = True
    queue, future = Queue(), Future()

    def item():
        calls.append(("item", threading.get_ident()))
        return "done"

    queue.put((item, (), {}, future, False, threading.Event()))
    queue.put(None)
    thread = threading.Thread(target=w._com_worker, args=(queue, 0))
    thread.start()
    thread.join(10)
    assert future.result(1) == "done"
    worker = calls[0][1]
    assert calls == [("let go", worker), ("item", worker), ("leave", worker)]
    assert worker != threading.get_ident()
