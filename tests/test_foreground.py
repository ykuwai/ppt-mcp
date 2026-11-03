"""utils.foreground: a PowerPoint call that took the front hands it back.

Runs anywhere, against a stand-in for the Win32 foreground, since what it
decides is when to put the user's window back.
"""

import sys

sys.path.insert(0, "src")

from utils.foreground import run_keeping_foreground, shows_powerpoint  # noqa: E402


class Screen:
    """Stands in for the Win32 foreground: windows by handle, owned by an exe."""

    def __init__(self, front, owners):
        self.front, self.owners = front, dict(owners)
        self.input_tick = 100
        self.put_back_calls = []
        self.fail = set()

    def window(self):
        if "window" in self.fail:
            raise OSError("no foreground")
        return self.front

    def process_name(self, hwnd):
        return self.owners.get(hwnd, "")

    def last_input(self):
        return self.input_tick

    def is_window(self, hwnd):
        return hwnd in self.owners

    def put_back(self, hwnd):
        if "put_back" in self.fail:
            raise OSError("access denied")
        self.put_back_calls.append(hwnd)
        self.front = hwnd


NOTEPAD, POWERPOINT, EXCEL, BROWSER = 1, 2, 3, 4
OWNERS = {NOTEPAD: "notepad.exe", POWERPOINT: "powerpnt.exe", EXCEL: "excel.exe", BROWSER: "chrome.exe"}


def taking(screen, to, result="done", touch=False):
    def call():
        screen.front = to
        if touch:
            screen.input_tick += 1
        return result
    return call


def test_the_users_window_comes_back_when_powerpoint_took_the_front():
    screen = Screen(NOTEPAD, OWNERS)
    assert run_keeping_foreground(taking(screen, POWERPOINT), (), {}, screen) == "done"
    assert screen.put_back_calls == [NOTEPAD]
    assert screen.front == NOTEPAD


def test_excel_opened_for_a_charts_data_is_handed_back_too():
    screen = Screen(NOTEPAD, OWNERS)
    run_keeping_foreground(taking(screen, EXCEL), (), {}, screen)
    assert screen.put_back_calls == [NOTEPAD]


def test_nothing_moves_when_powerpoint_was_already_in_front():
    screen = Screen(POWERPOINT, OWNERS)
    run_keeping_foreground(taking(screen, POWERPOINT), (), {}, screen)
    screen = Screen(EXCEL, OWNERS)
    run_keeping_foreground(taking(screen, POWERPOINT), (), {}, screen)
    assert screen.put_back_calls == []


def test_nothing_moves_when_the_user_acted_during_the_call():
    screen = Screen(NOTEPAD, OWNERS)
    run_keeping_foreground(taking(screen, POWERPOINT, touch=True), (), {}, screen)
    assert screen.put_back_calls == []
    assert screen.front == POWERPOINT


def test_nothing_moves_when_something_other_than_office_took_the_front():
    screen = Screen(NOTEPAD, OWNERS)
    run_keeping_foreground(taking(screen, BROWSER), (), {}, screen)
    assert screen.put_back_calls == []


def test_a_window_that_closed_meanwhile_is_not_put_back():
    owners = dict(OWNERS)
    screen = Screen(NOTEPAD, owners)

    def call():
        screen.front = POWERPOINT
        del screen.owners[NOTEPAD]

    run_keeping_foreground(call, (), {}, screen)
    assert screen.put_back_calls == []


def test_a_call_meant_to_show_powerpoint_is_left_alone():
    screen = Screen(NOTEPAD, OWNERS)

    @shows_powerpoint
    def start_the_show():
        screen.front = POWERPOINT
        return "showing"

    assert run_keeping_foreground(start_the_show, (), {}, screen) == "showing"
    assert screen.put_back_calls == []


def test_failing_to_read_or_hand_back_never_changes_the_tools_result():
    screen = Screen(NOTEPAD, OWNERS)
    screen.fail.add("put_back")
    assert run_keeping_foreground(taking(screen, POWERPOINT, result=42), (), {}, screen) == 42
    screen = Screen(NOTEPAD, OWNERS)
    screen.fail.add("window")
    assert run_keeping_foreground(taking(screen, POWERPOINT, result=7), (), {}, screen) == 7


def test_an_error_from_the_call_still_propagates_and_the_front_comes_back():
    screen = Screen(NOTEPAD, OWNERS)

    def fails():
        screen.front = POWERPOINT
        raise ValueError("Shape 'Rectangle 9' not found")

    try:
        run_keeping_foreground(fails, (), {}, screen)
    except ValueError as e:
        assert "Rectangle 9" in str(e)
    else:
        raise AssertionError("the call's error was swallowed")
    assert screen.put_back_calls == [NOTEPAD]


def test_arguments_reach_the_call_and_no_reader_means_a_plain_call():
    seen = []

    def call(a, b=None):
        seen.append((a, b))
        return a

    assert run_keeping_foreground(call, (1,), {"b": 2}, None) == 1
    assert run_keeping_foreground(call, (3,), {"b": 4}, Screen(NOTEPAD, OWNERS)) == 3
    assert seen == [(1, 2), (3, 4)]
