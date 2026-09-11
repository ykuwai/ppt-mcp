"""Tests for the deck level Apple Event tools.

Covers ``ppt_mac/properties.py``, ``ppt_mac/sections.py``,
``ppt_mac/slideshow.py`` and ``ppt_mac/edit_ops.py``. Everything here needs
appscript, which only installs on macOS, so the whole file is skipped
elsewhere. None of it launches PowerPoint; the object graph is made of
stand-ins and the live behaviour is covered by MACOS_PORT.md and by running the
server.

The stand-ins are deliberately literal about the two habits the port is built
on. A collection answers only through positional indexing, and a write that is
meant to fail reports success and changes nothing, which is what the silent
no-op refusals exist to catch.
"""

import sys

import pytest

sys.path.insert(0, "src")

macos_only = pytest.mark.skipif(
    sys.platform != "darwin", reason="the Apple Event backend needs appscript"
)


# ---------------------------------------------------------------------------
# The swap itself
# ---------------------------------------------------------------------------
@macos_only
class TestImplementationsAreSwapped:
    """The whole port hangs on the ppt_com names pointing at the mac ones.

    A module scope import of the matching ``ppt_com`` module from a ``ppt_mac``
    one would run the swap block against a half built module and the swap would
    quietly not happen, leaving the COM code in place on a Mac. Nothing else in
    the suite would notice, so it is checked here.
    """

    @pytest.mark.parametrize(
        "module_name",
        ["properties", "sections", "slideshow", "edit_ops", "charts", "freeform", "groups"],
    )
    def test_every_impl_is_the_apple_event_one(self, module_name):
        import importlib

        com = importlib.import_module(f"ppt_com.{module_name}")
        mac = importlib.import_module(f"ppt_mac.{module_name}")

        swapped = [
            name for name in dir(mac)
            if name.startswith("_") and name.endswith("_impl")
        ]
        assert swapped, "the mac module defines no implementations at all"
        for name in swapped:
            assert getattr(com, name) is getattr(mac, name)


# ---------------------------------------------------------------------------
# properties.py
# ---------------------------------------------------------------------------
@macos_only
class TestDocumentProperties:
    """Built in properties, found by name and read back after every write."""

    def test_the_windows_spellings_come_back_whatever_macos_calls_them(self):
        """PowerPoint for Mac writes `Last author` and Windows `Last Author`."""
        from ppt_mac.properties import _get_properties_impl

        with _fake_deck(properties={"Title": "Deck", "Last author": "Rika"}):
            result = _get_properties_impl()

        assert result["properties"]["Title"] == "Deck"
        assert result["properties"]["Last Author"] == "Rika"

    def test_a_property_the_deck_lacks_reads_as_null(self):
        from ppt_mac.properties import _get_properties_impl

        with _fake_deck(properties={"Title": "Deck"}):
            result = _get_properties_impl()

        assert result["properties"]["Company"] is None

    def test_an_empty_property_reads_as_null_rather_than_empty_string(self):
        from ppt_mac.properties import _get_properties_impl

        with _fake_deck(properties={"Title": "Deck", "Company": ""}):
            result = _get_properties_impl()

        assert result["properties"]["Company"] is None

    def test_a_date_comes_back_in_the_shape_windows_gives_it(self):
        import datetime

        from ppt_mac.properties import _get_properties_impl

        stamp = datetime.datetime(2026, 9, 4, 13, 45, 0)
        with _fake_deck(properties={"Title": "Deck", "Creation date": stamp}):
            result = _get_properties_impl()

        assert result["properties"]["Creation Date"] == "2026-09-04 13:45:00"

    def test_writing_counts_only_what_read_back(self):
        from ppt_mac.properties import _set_properties_impl

        with _fake_deck(properties={"Title": "", "Author": ""}) as deck:
            result = _set_properties_impl(
                "New title", "Rika", None, None, None, None, None,
            )

        assert result["properties_set"] == 2
        assert result["set_names"] == ["Title", "Author"]
        assert deck.properties["Title"] == "New title"
        assert "warnings" not in result

    def test_a_write_that_does_not_stick_is_not_counted(self):
        """PowerPoint accepting a value and keeping the old one is the no-op."""
        from ppt_mac.properties import _set_properties_impl

        with _fake_deck(properties={"Title": "Old"}, refuse_writes={"Title"}):
            result = _set_properties_impl(
                "New title", None, None, None, None, None, None,
            )

        assert result["properties_set"] == 0
        assert result["set_names"] == []
        assert "still reads" in result["warnings"][0]

    def test_a_property_that_is_not_there_is_reported_not_invented(self):
        from ppt_mac.properties import _set_properties_impl

        with _fake_deck(properties={"Title": "Deck"}) as deck:
            result = _set_properties_impl(
                None, None, None, None, None, None, "Acme",
            )

        assert result["properties_set"] == 0
        assert "Company" in result["warnings"][0]
        assert "Company" not in deck.properties

    def test_an_unreadable_collection_refuses_rather_than_answering_nulls(self):
        from ppt_mac.properties import _get_properties_impl, _set_properties_impl

        with _fake_deck(properties={}):
            read = _get_properties_impl()
            written = _set_properties_impl(
                "Title", None, None, None, None, None, None,
            )

        for payload in (read, written):
            assert payload["platform"] == "macOS"
            assert "no document properties at all" in payload["reason"]

    def test_names_are_read_one_reference_at_a_time(self):
        """`document properties` is a mixed collection, so it is indexed.

        The fake refuses ``document properties.name`` outright, which is the
        bulk read that a collection answering by subclass cannot be trusted
        with. Asking for it is what this test would catch.
        """
        from ppt_mac.properties import _property_index

        deck = _FakeDeck(properties={"Title": "Deck", "Author": "Rika"})
        index = _property_index(deck.presentation)

        assert sorted(index) == ["author", "title"]


# ---------------------------------------------------------------------------
# sections.py
# ---------------------------------------------------------------------------
@macos_only
class TestSections:
    """Sections are commands here, and every command is checked afterwards."""

    def test_listing_reports_the_keys_windows_reports(self):
        from ppt_mac.sections import _list_sections_impl

        with _fake_deck(sections=[("Intro", 1, 2), ("Body", 3, 4)]):
            result = _list_sections_impl()

        assert result["sections_count"] == 2
        assert result["sections"][0] == {
            "index": 1, "name": "Intro", "first_slide": 1, "slides_count": 2,
        }
        assert result["sections"][1]["name"] == "Body"

    def test_a_deck_with_no_sections_lists_none_rather_than_failing(self):
        from ppt_mac.sections import _list_sections_impl

        with _fake_deck(sections=[]):
            result = _list_sections_impl()

        assert result == {"success": True, "sections_count": 0, "sections": []}

    def test_a_count_that_will_not_answer_becomes_a_readable_refusal(self):
        """MACOS_PORT section 9 leaves -1708 open, so it is named not leaked."""
        from ppt_mac.sections import _list_sections_impl

        with _fake_deck(sections=[], count_error=-1708):
            result = _list_sections_impl()

        assert result["error"] == "ppt_list_sections is not available on macOS"
        assert "-1708" in result["reason"]
        assert result["platform"] == "macOS"

    def test_adding_a_section_reads_its_position_back(self):
        from ppt_mac.sections import _add_section_impl

        with _fake_deck(slides=6, sections=[("Intro", 1, 2)]) as deck:
            result = _add_section_impl("Body", 3)

        assert result["success"] is True
        assert result["section_index"] == 2
        assert result["name"] == "Body"
        assert result["slide_index"] == 3
        # `before slide` is the parameter that means what Windows means.
        assert deck.inserted == [{"before_slide": 3, "titled": "Body"}]

    def test_a_section_that_never_appeared_refuses(self):
        from ppt_mac.sections import _add_section_impl

        with _fake_deck(slides=6, sections=[("Intro", 1, 6)], deaf=True):
            result = _add_section_impl("Body", 3)

        assert "silent no-op" in result["reason"]
        assert result["error"] == "ppt_add_section is not available on macOS"

    def test_a_slide_index_off_the_end_still_raises(self):
        from ppt_mac.sections import _add_section_impl

        with _fake_deck(slides=3, sections=[]):
            with pytest.raises(ValueError, match="out of range"):
                _add_section_impl("Body", 9)

    def test_renaming_reads_the_new_name_back(self):
        from ppt_mac.sections import _manage_section_impl

        with _fake_deck(sections=[("Intro", 1, 2)]) as deck:
            result = _manage_section_impl(1, "rename", "Opening", None)

        assert result == {
            "success": True, "action": "rename",
            "section_index": 1, "new_name": "Opening",
        }
        assert deck.sections[0].name == "Opening"

    def test_a_rename_that_did_not_take_names_the_rename_not_the_tool(self):
        from ppt_mac.sections import _manage_section_impl

        with _fake_deck(sections=[("Intro", 1, 2)], deaf=True):
            result = _manage_section_impl(1, "rename", "Opening", None)

        assert result["error"] == (
            "ppt_manage_section could not rename the section on macOS"
        )
        assert "is not available" not in result["error"]

    def test_a_move_is_checked_against_the_id_not_the_name(self):
        """Two sections may share a name, and then a name check proves nothing."""
        from ppt_mac.sections import _manage_section_impl

        with _fake_deck(sections=[("Part", 1, 2), ("Part", 3, 2)]) as deck:
            result = _manage_section_impl(2, "move", None, 1)

        assert result["moved_to"] == 1
        assert [s.identifier for s in deck.sections] == ["id2", "id1"]

    def test_a_move_that_did_not_happen_refuses(self):
        from ppt_mac.sections import _manage_section_impl

        with _fake_deck(sections=[("Intro", 1, 2), ("Body", 3, 2)], deaf=True):
            result = _manage_section_impl(2, "move", None, 1)

        assert result["error"] == (
            "ppt_manage_section could not move the section on macOS"
        )
        assert "silent no-op" in result["reason"]

    def test_deleting_keeps_the_slides_and_reports_the_old_name(self):
        from ppt_mac.sections import _manage_section_impl

        with _fake_deck(sections=[("Intro", 1, 2), ("Body", 3, 2)]) as deck:
            result = _manage_section_impl(1, "delete", None, None)

        assert result == {
            "success": True, "action": "delete", "deleted_section": "Intro",
        }
        assert deck.deleted_with_slides == [False]

    def test_a_delete_that_left_the_count_alone_refuses(self):
        from ppt_mac.sections import _manage_section_impl

        with _fake_deck(sections=[("Intro", 1, 2)], deaf=True):
            result = _manage_section_impl(1, "delete", None, None)

        assert result["error"] == (
            "ppt_manage_section could not delete the section on macOS"
        )

    @pytest.mark.parametrize(
        "args, message",
        [
            ((1, "explode", None, None), "Unknown action"),
            ((1, "rename", None, None), "new_name is required"),
            ((1, "move", None, None), "move_to_index is required"),
            ((7, "delete", None, None), "out of range"),
            ((1, "move", None, 7), "out of range"),
        ],
    )
    def test_bad_arguments_still_raise_the_way_windows_raises(self, args, message):
        """ppt_com turns a ValueError into the error JSON, so it stays a raise."""
        from ppt_mac.sections import _manage_section_impl

        with _fake_deck(sections=[("Intro", 1, 2)]):
            with pytest.raises(ValueError, match=message):
                _manage_section_impl(*args)

    def test_an_unknown_action_is_rejected_before_anything_is_asked(self):
        from ppt_mac.sections import _manage_section_impl

        with _no_powerpoint():
            with pytest.raises(ValueError, match="Unknown action"):
                _manage_section_impl(1, "explode", None, None)


# ---------------------------------------------------------------------------
# slideshow.py
# ---------------------------------------------------------------------------
@macos_only
class TestSlideShowEnums:
    """Show constants have to read the same number on both platforms."""

    def test_the_range_type_the_generator_missed_is_filled_in(self):
        """macOS calls it `slide show range` and Windows ppShowSlideRange."""
        from appscript import k

        from ppt_com.constants import ppShowAll, ppShowSlideRange
        from ppt_mac.slideshow import _RANGE_TYPES, _windows_constant

        assert _RANGE_TYPES[ppShowSlideRange] == k.slide_show_range
        assert _windows_constant(_RANGE_TYPES, k.slide_show_range) == 2
        assert _RANGE_TYPES[ppShowAll] == k.slide_show_range_show_all

    def test_show_types_translate_out_and_back(self):
        from appscript import k

        from backend.mac_enums import PpSlideShowType, to_keyword
        from ppt_com.constants import ppShowTypeKiosk
        from ppt_mac.slideshow import _windows_constant

        assert to_keyword(PpSlideShowType, ppShowTypeKiosk, "show type") == (
            k.slide_show_type_kiosk
        )
        assert _windows_constant(PpSlideShowType, k.slide_show_type_kiosk) == 3

    def test_a_macos_only_show_type_reads_as_unknown_not_as_a_near_miss(self):
        from appscript import k

        from backend.mac_enums import PpSlideShowType
        from ppt_com.constants import SHOW_TYPE_NAMES
        from ppt_mac.slideshow import _windows_constant

        presenter = _windows_constant(PpSlideShowType, k.slide_show_type_presenter)
        assert presenter is None
        assert SHOW_TYPE_NAMES.get(presenter, "unknown") == "unknown"

    def test_every_state_round_trips(self):
        from backend.mac_enums import PpSlideShowState, to_keyword
        from ppt_com.constants import SLIDESHOW_STATE_NAMES
        from ppt_mac.slideshow import _windows_constant

        for value in SLIDESHOW_STATE_NAMES:
            word = to_keyword(PpSlideShowState, value, "slide show state")
            assert _windows_constant(PpSlideShowState, word) == value

    def test_pointer_types_agree_on_the_ordinal(self):
        from appscript import k

        from ppt_mac.slideshow import _POINTER_TYPES, _windows_constant

        assert _POINTER_TYPES[2] == k.slide_show_pointer_pen
        assert _windows_constant(_POINTER_TYPES, k.slide_show_pointer_arrow) == 1


@macos_only
class TestSlideShow:
    """Starting, stopping and driving a show."""

    def test_starting_the_whole_deck_clears_any_earlier_range(self):
        from appscript import k

        from ppt_mac.slideshow import _slideshow_start_impl

        with _fake_deck(slides=5) as deck:
            result = _slideshow_start_impl(None, None, None, None)

        assert result["success"] is True
        assert result["show_type"] == "speaker"
        assert (result["start_slide"], result["end_slide"]) == (1, 5)
        assert result["total_slides"] == 5
        assert deck.settings["range_type"] == k.slide_show_range_show_all

    def test_a_range_sets_the_range_type_the_generator_missed(self):
        from appscript import k

        from ppt_mac.slideshow import _slideshow_start_impl

        with _fake_deck(slides=5) as deck:
            result = _slideshow_start_impl(2, 4, True, "kiosk")

        assert (result["start_slide"], result["end_slide"]) == (2, 4)
        assert deck.settings["range_type"] == k.slide_show_range
        assert deck.settings["starting_slide"] == 2
        assert deck.settings["ending_slide"] == 4
        assert deck.settings["show_type"] == k.slide_show_type_kiosk
        # A plain boolean, not the COM tri-state Windows sets.
        assert deck.settings["loop_until_stopped"] is True

    def test_a_show_that_never_opened_a_window_refuses(self):
        from ppt_mac.slideshow import _slideshow_start_impl

        with _fake_deck(slides=3, deaf=True):
            result = _slideshow_start_impl(None, None, None, None)

        assert result["error"] == "ppt_slideshow_start is not available on macOS"
        assert "silent no-op" in result["reason"]

    def test_the_window_the_command_answered_with_is_not_used(self):
        from ppt_mac.slideshow import _slideshow_start_impl

        with _fake_deck(slides=3) as deck:
            _slideshow_start_impl(None, None, None, None)

        assert deck.run_result_touched is False

    @pytest.mark.parametrize(
        "args, message",
        [
            ((None, None, None, "cinema"), "Unknown show_type"),
            ((9, None, None, None), "start_slide 9 out of range"),
            ((2, 9, None, None), "end_slide 9 out of range"),
            ((3, 2, None, None), "end_slide 2 out of range"),
        ],
    )
    def test_bad_arguments_still_raise(self, args, message):
        from ppt_mac.slideshow import _slideshow_start_impl

        with _fake_deck(slides=5):
            with pytest.raises(ValueError, match=message):
                _slideshow_start_impl(*args)

    def test_a_rejected_argument_leaves_the_settings_alone(self):
        """Validation runs before any write, so nothing is half applied."""
        from ppt_mac.slideshow import _slideshow_start_impl

        with _fake_deck(slides=5) as deck:
            with pytest.raises(ValueError):
                _slideshow_start_impl(None, None, None, "cinema")

        assert deck.settings == {}

    def test_stopping_with_nothing_running_is_not_an_error(self):
        from ppt_mac.slideshow import _slideshow_stop_impl

        with _fake_deck(slides=3):
            assert _slideshow_stop_impl() == {
                "success": True, "message": "No slide show was running.",
            }

    def test_stopping_ends_the_show(self):
        from ppt_mac.slideshow import _slideshow_start_impl, _slideshow_stop_impl

        with _fake_deck(slides=3) as deck:
            _slideshow_start_impl(None, None, None, None)
            result = _slideshow_stop_impl()

        assert result["message"] == "Slide show ended."
        assert deck.show is None

    def test_a_show_that_will_not_close_refuses(self):
        from ppt_mac.slideshow import _slideshow_stop_impl

        with _fake_deck(slides=3, running=1, deaf=True):
            result = _slideshow_stop_impl()

        assert result["error"] == "ppt_slideshow_stop is not available on macOS"
        assert "still open" in result["reason"]

    def test_next_and_previous_walk_the_show(self):
        from ppt_mac.slideshow import _slideshow_next_impl, _slideshow_previous_impl

        with _fake_deck(slides=4, running=2):
            forward = _slideshow_next_impl()
            back = _slideshow_previous_impl()

        assert forward["current_slide"] == 3
        assert forward["state"] == "running"
        assert back["current_slide"] == 2

    @pytest.mark.parametrize(
        "impl_name",
        ["_slideshow_next_impl", "_slideshow_previous_impl", "_slideshow_goto_impl"],
    )
    def test_driving_a_show_that_is_not_running_raises(self, impl_name):
        import ppt_mac.slideshow as slideshow

        impl = getattr(slideshow, impl_name)
        args = (2,) if impl_name == "_slideshow_goto_impl" else ()
        with _fake_deck(slides=4):
            with pytest.raises(RuntimeError, match="No slide show is running"):
                impl(*args)

    def test_going_to_a_slide_is_read_back(self):
        from ppt_mac.slideshow import _slideshow_goto_impl

        with _fake_deck(slides=4, running=1):
            result = _slideshow_goto_impl(3)

        assert result["current_slide"] == 3
        assert result["state"] == "running"

    def test_a_show_that_did_not_move_refuses_and_names_the_two_that_work(self):
        from ppt_mac.slideshow import _slideshow_goto_impl

        with _fake_deck(slides=4, running=1, deaf=True):
            result = _slideshow_goto_impl(3)

        assert "still on slide 1" in result["reason"]
        assert result["alternatives"] == [
            "ppt_slideshow_next", "ppt_slideshow_previous",
        ]

    def test_a_not_handled_go_to_slide_becomes_a_refusal_not_a_raw_error(self):
        """`go to slide` is declared on a document window's view, not a show's."""
        from ppt_mac.slideshow import _slideshow_goto_impl

        with _fake_deck(slides=4, running=1, goto_error=-1708):
            result = _slideshow_goto_impl(3)

        assert "-1708" in result["reason"]
        assert result["alternatives"] == [
            "ppt_slideshow_next", "ppt_slideshow_previous",
        ]

    def test_status_with_nothing_running(self):
        from ppt_mac.slideshow import _slideshow_get_status_impl

        with _fake_deck(slides=3):
            assert _slideshow_get_status_impl() == {"running": False}

    def test_status_reports_the_windows_numbers(self):
        from ppt_mac.slideshow import _slideshow_get_status_impl

        with _fake_deck(slides=3, running=2):
            result = _slideshow_get_status_impl()

        assert result["running"] is True
        assert result["current_slide"] == 2
        assert result["state"] == 1
        assert result["state_name"] == "running"
        assert result["pointer_type"] == 1


# ---------------------------------------------------------------------------
# edit_ops.py
# ---------------------------------------------------------------------------
@macos_only
class TestEditOps:
    """Undo, copy, format painting, and the two with no route."""

    def test_undo_takes_a_count_in_one_command(self):
        from ppt_mac.edit_ops import _undo_impl

        with _fake_deck(slides=2) as deck:
            result = _undo_impl(3)

        assert result["actions_undone"] == 3
        assert deck.undone == [3]
        # The warning used to say fewer than three might have happened. The
        # danger measured on a real deck runs the other way: one undo takes
        # back every edit this server made, whatever number is passed.
        warning = result["warnings"][0]
        assert "takes back everything this server has edited" in warning
        assert "not a measurement" in warning

    def test_redo_says_the_same_about_its_count(self):
        from ppt_mac.edit_ops import _redo_impl

        with _fake_deck(slides=2) as deck:
            result = _redo_impl(2)

        assert result["actions_redone"] == 2
        assert deck.redone == [2]
        assert "fewer than 2" in result["warnings"][0]

    def test_copying_a_shape_navigates_before_it_pastes(self):
        """`paste object` lands on whatever slide the view is showing."""
        from ppt_mac.edit_ops import _copy_shape_to_slide_impl

        with _fake_deck(slides=2, shapes={1: ["Title"], 2: []}) as deck:
            result = _copy_shape_to_slide_impl(1, "Title", 2)

        assert result["new_shape_name"] == "Title"
        assert result["destination_slide"] == 2
        assert deck.viewed_slide == 2
        # The selection is cleared between the navigation and the paste. A
        # paste leaves what it pasted selected, and `paste object` puts the
        # next one inside that selection when it is a chart, which changes no
        # shape count and so passes every check here while rewriting the chart.
        assert deck.order == ["copy", "goto", "unselect", "paste"]
        assert deck.selection_cleared

    def test_a_paste_is_never_attempted_from_the_wrong_slide(self):
        """A navigation that failed would drop the copy on the source slide."""
        from ppt_mac.edit_ops import _copy_shape_to_slide_impl

        with _fake_deck(
            slides=2, shapes={1: ["Title"], 2: []}, goto_error=-1728,
        ) as deck:
            result = _copy_shape_to_slide_impl(1, "Title", 2)

        assert "-1728" in result["reason"]
        assert "paste" not in deck.order
        assert deck.shapes[1] == ["Title"]
        assert deck.shapes[2] == []

    def test_a_paste_that_added_nothing_refuses(self):
        from ppt_mac.edit_ops import _copy_shape_to_slide_impl

        with _fake_deck(slides=2, shapes={1: ["Title"], 2: []}, deaf=True):
            result = _copy_shape_to_slide_impl(1, "Title", 2)

        assert result["error"] == (
            "ppt_copy_shape_to_slide is not available on macOS"
        )
        assert "silent no-op" in result["reason"]

    def test_a_shape_that_is_not_there_raises_before_anything_moves(self):
        from ppt_mac.edit_ops import _copy_shape_to_slide_impl

        with _fake_deck(slides=2, shapes={1: ["Title"], 2: []}) as deck:
            with pytest.raises(ValueError):
                _copy_shape_to_slide_impl(1, "Nowhere", 2)

        assert deck.order == []
        assert deck.viewed_slide is None

    def test_format_painting_needs_no_selection_and_refuses_nothing(self):
        from ppt_mac.edit_ops import _copy_formatting_impl

        with _fake_deck(slides=1, shapes={1: ["Title", "Body", "Note"]}) as deck:
            result = _copy_formatting_impl(1, "Title", ["Body", "Note"])

        assert result["source"] == "Title"
        assert result["applied_to"] == ["Body", "Note"]
        assert deck.picked_up == ["Title"]
        assert deck.applied == ["Body", "Note"]

    def test_one_bad_target_leaves_every_shape_unpainted(self):
        from ppt_mac.edit_ops import _copy_formatting_impl

        with _fake_deck(slides=1, shapes={1: ["Title", "Body"]}) as deck:
            with pytest.raises(ValueError):
                _copy_formatting_impl(1, "Title", ["Body", "Nowhere"])

        assert deck.picked_up == []
        assert deck.applied == []

    def test_an_undo_entry_names_the_tool_that_replaces_it(self):
        from ppt_mac.edit_ops import _start_undo_entry_impl

        with _no_powerpoint():
            result = _start_undo_entry_impl()

        assert result["error"] == "ppt_start_undo_entry is not available on macOS"
        assert result["platform"] == "macOS"
        assert "ppt_undo with times set to the number of edits" in (
            result["alternatives"]
        )

    def test_execute_mso_explains_that_an_idmso_names_nothing_here(self):
        from ppt_mac.edit_ops import _execute_mso_impl

        with _no_powerpoint():
            result = _execute_mso_impl("SelectAll", True)

        assert result["error"] == "ppt_execute_mso is not available on macOS"
        assert "command bar control" in result["reason"]
        assert "check_enabled" in result["reason"]
        assert result["alternatives"] == ["ppt_undo", "ppt_redo", "ppt_slideshow_start"]


# ---------------------------------------------------------------------------
# A deck made of stand-ins. Only the parts these four modules touch are
# modelled. Two of the behaviours are deliberate. A collection answers
# positionally and raises -1728 past its end, the way PowerPoint does. A
# `deaf` deck accepts every write, reports no error and changes nothing, which
# is the silent no-op the refusals exist to catch.
# ---------------------------------------------------------------------------
class _FakeCollection:
    """A collection that is counted and then indexed, one element at a time.

    ``names`` is given only where the real object model can be asked for every
    name in one event, which is true of a slide's shapes and is exactly what
    ``document properties`` cannot be trusted with, so that one is built
    without it and raises if anything asks.
    """

    def __init__(self, items, names=None):
        self._items = items
        self._names = names

    @property
    def name(self):
        if self._names is None:
            raise AssertionError(
                "this collection answers by subclass, so its names have to be "
                "read one reference at a time"
            )
        return _FakeCollection(list(self._names))

    def get(self):
        return list(self._items)

    def __getitem__(self, index):
        if index < 1 or index > len(self._items):
            raise _command_error(-1728)
        return self._items[index - 1]


class _FakeProperty:
    def __init__(self, value=None, on_set=None, keep=True):
        self._value = value
        self._on_set = on_set
        self._keep = keep

    def __call__(self):
        return self._value

    def get(self):
        return self._value

    def set(self, value):
        if self._on_set is not None:
            self._on_set(value)
        if self._keep:
            self._value = value


class _FakeDocumentProperty:
    def __init__(self, deck, name, value):
        self._deck = deck
        self._name = name
        # `keep` is off so that a refused write shows up as PowerPoint
        # accepting a value and still reading back the old one.
        self.value = _FakeProperty(value, on_set=self._write, keep=False)

    def name(self):
        return self._name

    def _write(self, value):
        if self._name in self._deck.refuse_writes:
            return
        self._deck.properties[self._name] = value
        self.value._value = value


class _FakeSection:
    def __init__(self, identifier, name, first_slide, slides_count):
        self.identifier = identifier
        self.name = name
        self.first_slide = first_slide
        self.slides_count = slides_count


class _FakeSectionProperties:
    """Sections, driven the way the dictionary drives them, through commands."""

    def __init__(self, deck):
        self._deck = deck

    def get_count_of_sections(self):
        if self._deck.count_error is not None:
            raise _command_error(self._deck.count_error)
        return len(self._deck.sections)

    def _at(self, at_position):
        return self._deck.sections[at_position - 1]

    def get_name_of_section(self, at_position=None):
        return self._at(at_position).name

    def get_id_of_section(self, at_position=None):
        return self._at(at_position).identifier

    def get_first_slide_of_section(self, at_position=None):
        return self._at(at_position).first_slide

    def get_slide_count_of_section(self, at_position=None):
        return self._at(at_position).slides_count

    def insert_section(self, before_slide=None, titled=None, before_section=None):
        self._deck.inserted.append(
            {"before_slide": before_slide, "titled": titled}
        )
        if self._deck.deaf:
            return 0
        identifier = f"id{len(self._deck.sections) + 1}"
        position = len(self._deck.sections)
        for i, section in enumerate(self._deck.sections):
            if section.first_slide > before_slide:
                position = i
                break
        self._deck.sections.insert(
            position, _FakeSection(identifier, titled, before_slide, 1)
        )
        return position + 1

    def rename_section(self, at_position=None, to=None):
        if self._deck.deaf:
            return
        self._at(at_position).name = to

    def move_section(self, at_position=None, to_position=None):
        if self._deck.deaf:
            return
        section = self._deck.sections.pop(at_position - 1)
        self._deck.sections.insert(to_position - 1, section)

    def delete_section(self, at_position=None, deleting_slides=None):
        self._deck.deleted_with_slides.append(deleting_slides)
        if self._deck.deaf:
            return
        self._deck.sections.pop(at_position - 1)


class _FakeShape:
    def __init__(self, deck, name):
        self._deck = deck
        self._name = name

    def name(self):
        return self._name

    def copy_shape(self):
        self._deck.order.append("copy")
        self._deck.clipboard = self._name

    def pick_up(self):
        self._deck.picked_up.append(self._name)

    def apply(self):
        self._deck.applied.append(self._name)


class _FakeSlideShowView:
    def __init__(self, deck):
        self._deck = deck

    def current_show_position(self):
        return self._deck.show

    def slide_state(self):
        from appscript import k

        return k.slide_show_state_running

    def pointer_type(self):
        from appscript import k

        return k.slide_show_pointer_arrow

    def go_to_next_slide(self):
        self._deck.show = min(self._deck.show + 1, self._deck.slide_count)

    def go_to_previous_slide(self):
        self._deck.show = max(self._deck.show - 1, 1)

    def go_to_slide(self, number=None):
        if self._deck.goto_error is not None:
            raise _command_error(self._deck.goto_error)
        if self._deck.deaf:
            return
        self._deck.show = number

    def exit_slide_show(self):
        if self._deck.deaf:
            return
        self._deck.show = None


class _FakeSlideShowWindow:
    def __init__(self, deck):
        self.slideshow_view = _FakeSlideShowView(deck)


class _FakeRunResult:
    """What `run slide show` hands back, which nothing is allowed to use."""

    def __init__(self, deck):
        self._deck = deck

    def __getattr__(self, name):
        self._deck.run_result_touched = True
        raise AssertionError(
            "the window `run slide show` answered with must not be used"
        )


class _FakeSettings:
    def __init__(self, deck):
        self._deck = deck
        for name in (
            "show_type", "range_type", "starting_slide", "ending_slide",
            "loop_until_stopped",
        ):
            setattr(self, name, _FakeProperty(
                on_set=lambda value, key=name: deck.settings.__setitem__(key, value)
            ))

    def run_slide_show(self):
        if not self._deck.deaf:
            self._deck.show = self._deck.settings.get("starting_slide", 1)
        return _FakeRunResult(self._deck)


class _FakeView:
    def __init__(self, deck):
        self._deck = deck

    def go_to_slide(self, number=None):
        if self._deck.goto_error is not None:
            raise _command_error(self._deck.goto_error)
        self._deck.order.append("goto")
        self._deck.viewed_slide = number

    def paste_object(self):
        self._deck.order.append("paste")
        if self._deck.deaf or self._deck.clipboard is None:
            return
        self._deck.shapes[self._deck.viewed_slide].append(self._deck.clipboard)


class _FakeSelection:
    def __init__(self, deck):
        self._deck = deck

    def unselect(self):
        self._deck.order.append("unselect")
        self._deck.selection_cleared = True


class _FakeDeck:
    def __init__(
        self, slides=1, shapes=None, properties=None, sections=None,
        refuse_writes=None, deaf=False, running=None, count_error=None,
        goto_error=None,
    ):
        self.slide_count = slides
        self.shapes = {i: list(shapes.get(i, [])) for i in range(1, slides + 1)} \
            if shapes else {i: [] for i in range(1, slides + 1)}
        self.properties = dict(properties or {})
        self.refuse_writes = set(refuse_writes or ())
        self.sections = [
            _FakeSection(f"id{i}", name, first, held)
            for i, (name, first, held) in enumerate(sections or [], start=1)
        ]
        self.deaf = deaf
        self.count_error = count_error
        self.goto_error = goto_error
        self.show = running
        self.settings = {}
        self.run_result_touched = False
        self.selection_cleared = False
        self.inserted = []
        self.deleted_with_slides = []
        self.undone = []
        self.redone = []
        self.order = []
        self.picked_up = []
        self.applied = []
        self.clipboard = None
        self.viewed_slide = None

    # -- the object graph the tools walk ----------------------------------
    @property
    def application(self):
        deck = self

        class _App:
            presentations = _FakeCollection([object()])

            @property
            def slide_show_windows(self):
                return _FakeCollection(
                    [_FakeSlideShowWindow(deck)] if deck.show else []
                )

            def count(self, each=None):
                return 1 if deck.show else 0

        return _App()

    @property
    def presentation(self):
        deck = self

        class _Pres:
            slides = _FakeCollection(
                [deck.slide(i) for i in range(1, deck.slide_count + 1)]
            )
            document_properties = _FakeCollection(
                [
                    _FakeDocumentProperty(deck, name, value)
                    for name, value in deck.properties.items()
                ]
            )
            section_properties = _FakeSectionProperties(deck)
            document_windows = _FakeCollection([deck.window])
            slide_show_settings = _FakeSettings(deck)

            def count(self, each=None):
                from appscript import k

                # What `target_window` asks before it hands the window over.
                # A deck can outlive its window, and every one of these tools
                # reaches the editor through it.
                if each == k.document_window:
                    return 1
                raise AssertionError(f"unexpected count of {each}")

            def undo(self, times=None):
                deck.undone.append(times)

            def redo(self, times=None):
                deck.redone.append(times)

        return _Pres()

    @property
    def window(self):
        return type("Window", (), {
            "view": _FakeView(self),
            # A paste leaves what it pasted selected, and the next one then
            # lands inside it when it is a chart. The tool clears the
            # selection first, so the fake has to have one to clear.
            "selection": _FakeSelection(self),
        })()

    def slide(self, index):
        deck = self

        class _Slide:
            @property
            def shapes(self):
                # Built fresh on every read, because a paste changes what the
                # slide holds and a snapshot would hide that.
                return _FakeCollection(
                    [_FakeShape(deck, name) for name in deck.shapes[index]],
                    names=list(deck.shapes[index]),
                )

        return _Slide()


class _fake_deck:  # noqa: N801 - reads as a context manager, not a class
    """Point the wrapper at a fake deck for the length of a `with` block."""

    def __init__(self, **kwargs):
        self._deck = _FakeDeck(**kwargs)

    def __enter__(self):
        from backend.mac_ae import ppt

        self._ppt = ppt
        self._app = ppt._get_app_impl
        self._pres = ppt._get_pres_impl
        ppt._get_app_impl = lambda *a, **kw: self._deck.application
        ppt._get_pres_impl = lambda *a, **kw: self._deck.presentation
        return self._deck

    def __exit__(self, *exc):
        self._ppt._get_app_impl = self._app
        self._ppt._get_pres_impl = self._pres
        return False


class _no_powerpoint:  # noqa: N801 - reads as a context manager, not a class
    """Make any call into PowerPoint an error, so a refusal has to be one.

    What it enforces is the rule that everything which would refuse or raise
    happens before the first Apple Event, so a call that was never going to
    work does not move the user's view first.
    """

    def __enter__(self):
        from backend.mac_ae import ppt

        def _explode(*args, **kwargs):
            raise AssertionError("a refusal must not touch PowerPoint")

        self._ppt = ppt
        self._app = ppt._get_app_impl
        self._pres = ppt._get_pres_impl
        ppt._get_app_impl = _explode
        ppt._get_pres_impl = _explode
        return self

    def __exit__(self, *exc):
        self._ppt._get_app_impl = self._app
        self._ppt._get_pres_impl = self._pres
        return False


def _command_error(number):
    """Build an appscript CommandError carrying a given OSError number.

    ``errornumber`` is a read-only property on the real class, so the stub
    subclasses it rather than assigning through.
    """
    from appscript.reference import CommandError

    class _Stub(CommandError):
        def __init__(self):
            Exception.__init__(self, f"stub error {number}")

        @property
        def errornumber(self):
            return number

        @property
        def errormessage(self):
            return f"stub error {number}"

        def __str__(self):
            return f"stub error {number}"

    return _Stub()
