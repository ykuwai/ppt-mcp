"""Tests for the macOS Apple Event backend.

The colour helpers are pure and run everywhere. Everything else needs
appscript, which only installs on macOS, so those are skipped elsewhere. None
of this launches PowerPoint; the live behaviour is covered by MACOS_PORT.md and
by running the server.
"""

import sys
import time

import pytest

sys.path.insert(0, "src")

from utils.color import (  # noqa: E402
    hex_to_int,
    hex_to_rgb_list,
    rgb_list_to_hex,
    rgb_list_to_int,
)

macos_only = pytest.mark.skipif(
    sys.platform != "darwin", reason="the Apple Event backend needs appscript"
)


class TestColourBridging:
    """Windows packs a colour into one integer, macOS takes [R, G, B]."""

    def test_hex_to_rgb_list(self):
        assert hex_to_rgb_list("#1F6FEB") == [31, 111, 235]

    def test_hex_to_rgb_list_accepts_shorthand(self):
        assert hex_to_rgb_list("#F00") == [255, 0, 0]

    def test_round_trip_through_the_mac_form(self):
        assert rgb_list_to_hex(hex_to_rgb_list("#1F6FEB")) == "#1F6FEB"

    def test_both_platforms_report_the_same_number(self):
        """A deck inspected on a Mac has to read the same as on Windows."""
        assert rgb_list_to_int([31, 111, 235]) == hex_to_int("#1F6FEB")

    @pytest.mark.parametrize("empty", [None, [], [1, 2]])
    def test_missing_colour_is_not_an_error(self, empty):
        """PowerPoint answers `missing value` for a shape that has no colour."""
        assert rgb_list_to_hex(empty) is None
        assert rgb_list_to_int(empty) is None


@macos_only
class TestEnumTable:
    """The generated Windows constant to macOS enumerator table."""

    def test_shared_office_constants_pair_up(self):
        from appscript import k

        from backend.mac_enums import MsoAutoShapeType, PpSlideLayout

        assert MsoAutoShapeType[1] == k.autoshape_rectangle
        assert MsoAutoShapeType[5] == k.autoshape_rounded_rectangle
        assert PpSlideLayout[12] == k.slide_layout_blank

    def test_pairing_is_by_name_not_by_number(self):
        """ppSaveAsPNG is 18 on Windows and `save as PNG` is 24 on macOS.

        Pairing on the number would put the wrong format behind that constant,
        which is the failure this table exists to prevent.
        """
        from appscript import k

        from backend.mac_enums import PpSaveAsFileType

        assert PpSaveAsFileType[18] == k.save_as_PNG
        assert PpSaveAsFileType[32] == k.save_as_PDF

    def test_different_naming_schemes_are_bridged(self):
        """Windows numbers theme colours, macOS names them as ordinals."""
        from appscript import k

        from backend.mac_enums import MsoAutoShapeType, MsoThemeColorIndex

        assert MsoThemeColorIndex[5] == k.first_accent_theme_color
        assert MsoAutoShapeType[92] == k.autoshape_five_point_star

    def test_a_constant_macos_lacks_says_so(self):
        from backend.mac_enums import PpSaveAsFileType, to_keyword

        with pytest.raises(ValueError, match="no slide format matching"):
            to_keyword(PpSaveAsFileType, 33, "slide format")  # ppSaveAsXPS


@macos_only
class TestQuirkAbsorption:
    """PowerPoint answers oddly in ways that would otherwise reach the tools."""

    def test_empty_collection_counts_as_zero(self):
        """PowerPoint raises -1728 for an empty collection, not an empty list."""
        from backend.mac_ae import AE_NO_SUCH_OBJECT, count, elements

        class Raises:
            def get(self):
                raise _command_error(AE_NO_SUCH_OBJECT)

        assert elements(Raises()) == []
        assert count(Raises()) == 0

    def test_a_real_failure_still_propagates(self):
        """Only the empty-collection codes are swallowed, nothing else."""
        from backend.mac_ae import elements

        class Raises:
            def get(self):
                raise _command_error(-1743)  # permission refused

        with pytest.raises(Exception):
            elements(Raises())

    def test_single_element_becomes_a_list(self):
        from backend.mac_ae import elements

        class One:
            def get(self):
                return "only one"

        assert elements(One()) == ["only one"]

    def test_missing_value_counts_as_empty(self):
        from appscript import k

        from backend.mac_ae import count

        class Missing:
            def get(self):
                return k.missing_value

        assert count(Missing()) == 0


@macos_only
class TestErrorTranslation:
    """The errors worth rewriting are the ones that misdirect."""

    def test_permission_refusal_says_where_to_grant_it(self):
        from backend.mac_ae import PowerPointAppleEventWrapper

        translated = PowerPointAppleEventWrapper._translate(_command_error(-1743))
        assert "Automation" in str(translated)
        assert translated.errornumber == -1743

    def test_lost_connection_says_powerpoint_may_have_quit(self):
        from backend.mac_ae import PowerPointAppleEventWrapper

        translated = PowerPointAppleEventWrapper._translate(_command_error(-609))
        assert "quit" in str(translated)

    def test_an_unremarkable_error_is_passed_through(self):
        from backend.mac_ae import PowerPointAppleEventWrapper

        translated = PowerPointAppleEventWrapper._translate(_command_error(-2700))
        assert translated.errornumber == -2700


@macos_only
class TestUnsupportedTools:
    """What a tool says when the platform genuinely cannot do it."""

    def test_it_names_the_platform_the_reason_and_a_way_forward(self):
        import json

        from backend import PLATFORM_NAME
        from backend.unsupported import unsupported

        payload = json.loads(
            unsupported(
                "ppt_add_chart",
                "PowerPoint for Mac exposes no chart object to Apple Events",
                ["ppt_add_shape", "ppt_add_table"],
            )
        )
        assert payload["error"] == f"ppt_add_chart is not available on {PLATFORM_NAME}"
        assert "chart object" in payload["reason"]
        assert payload["alternatives"] == ["ppt_add_shape", "ppt_add_table"]

    def test_alternatives_are_omitted_rather_than_padded(self):
        import json

        from backend.unsupported import unsupported

        payload = json.loads(unsupported("ppt_x", "because"))
        assert "alternatives" not in payload


@macos_only
class TestAnimationEnums:
    """Animation constants have to read the same number on both platforms."""

    def test_effects_translate_out_and_back(self):
        from appscript import k

        from backend.mac_enums import MsoAnimEffect, to_keyword
        from ppt_mac.animation import _windows_constant

        assert to_keyword(MsoAnimEffect, 10, "animation effect") == k.animation_type_fade
        assert _windows_constant(MsoAnimEffect, k.animation_type_fade) == 10

    def test_every_effect_the_tool_advertises_is_reachable(self):
        """It once was not. Thirty seven of the fifty four failed on macOS.

        The generator only read banner sections of constants.py, and this
        vocabulary lives in a friendly name map in animation.py with no named
        constants behind it, so nothing could tell it the words were missing.
        """
        from backend.mac_enums import MsoAnimEffect, to_keyword
        from ppt_com.animation import ANIMATION_EFFECT_MAP
        from ppt_mac.animation import _WHAT_EFFECT

        for name, value in ANIMATION_EFFECT_MAP.items():
            assert to_keyword(MsoAnimEffect, value, _WHAT_EFFECT), name

    def test_the_motion_paths_are_the_same_words_in_another_order(self):
        """Windows writes path_arc_down, macOS writes `arc down path`."""
        from appscript import k

        from backend.mac_enums import MsoAnimEffect, to_keyword
        from ppt_mac.animation import _WHAT_EFFECT

        assert to_keyword(MsoAnimEffect, 86, _WHAT_EFFECT) == (
            k.animation_type_circle_path
        )
        assert to_keyword(MsoAnimEffect, 122, _WHAT_EFFECT) == (
            k.animation_type_arc_down_path
        )

    def test_every_direction_round_trips(self):
        from backend.mac_enums import MsoAnimDirection, to_keyword
        from ppt_com.constants import ANIM_DIRECTION_MAP
        from ppt_mac.animation import _windows_constant

        for value in ANIM_DIRECTION_MAP.values():
            word = to_keyword(MsoAnimDirection, value, "animation direction")
            assert _windows_constant(MsoAnimDirection, word) == value

    def test_the_two_names_the_generator_missed_are_filled_in(self):
        """Windows says None where macOS says `no after effect` and `no levels`."""
        from appscript import k

        from ppt_mac.animation import _AFTER_EFFECTS, _BUILD_LEVELS, _windows_constant

        assert _windows_constant(_AFTER_EFFECTS, k.no_after_effect) == 0
        assert _windows_constant(_AFTER_EFFECTS, k.dim) == 1
        assert _windows_constant(_BUILD_LEVELS, k.text_by_no_levels) == 0
        assert _windows_constant(_BUILD_LEVELS, k.text_by_first_level) == 2

    def test_an_emphasis_effect_reads_back_as_its_windows_number(self):
        """A deck holding one has to report the same effect on both platforms."""
        from appscript import k

        from backend.mac_enums import MsoAnimEffect
        from ppt_mac.animation import _windows_constant

        assert _windows_constant(MsoAnimEffect, k.animation_type_teeter) == 80
        assert _windows_constant(MsoAnimEffect, k.animation_type_grow_shrink) == 59

    def test_a_transition_with_four_directions_and_no_plain_form(self):
        """These are not missing. They are four each, and choosing is a guess."""
        from backend.mac_enums import PpEntryEffect, to_keyword
        from ppt_mac.animation import _WHAT_TRANSITION

        assert to_keyword(PpEntryEffect, 3844, _WHAT_TRANSITION)  # fade
        with pytest.raises(ValueError, match="four directional variants"):
            to_keyword(PpEntryEffect, 3845, _WHAT_TRANSITION)  # ppEffectPush

    def test_every_shape_the_tool_advertises_is_reachable(self):
        """Fourteen of the sixty used to fail, for the same reason as above."""
        from backend.mac_enums import MsoAutoShapeType, to_keyword
        from ppt_com.shapes import SHAPE_NAME_MAP

        for name, value in SHAPE_NAME_MAP.items():
            assert to_keyword(MsoAutoShapeType, value, "shape type"), name


@macos_only
class TestAnimationRefusals:
    """What the animation tools say to an argument macOS cannot honour."""

    def test_sequence_index_names_the_argument_not_the_tool(self):
        """The tool works; only that one argument has to go, and it says so."""
        from ppt_mac.animation import _remove_animation_impl, _update_animation_impl

        for payload in (
            _remove_animation_impl(1, 1, 2),
            _update_animation_impl(
                1, 1, 2, "fade", None, None, None, None, None,
                None, None, None, None, None, None, None, None,
                None, None, None, None,
            ),
        ):
            assert "sequence_index" in payload["error"]
            assert "is not available" not in payload["error"]
            assert payload["platform"] == "macOS"
            assert "sequence count stays at zero" in payload["reason"]

    def test_a_shape_click_trigger_names_trigger_shape(self):
        from ppt_mac.animation import _add_animation_impl

        payload = _add_animation_impl(
            1, "Title", "fade", "on_shape_click", None, None, False,
            None, None, None, None, None, None,
            "Button", None, None,
            None, None, None, None,
        )
        assert "trigger_shape" in payload["error"]
        assert "-1708" in payload["reason"]
        assert "ppt_add_animation with trigger='on_click'" in payload["alternatives"]

    def test_a_refusal_never_reaches_powerpoint(self):
        """A refused call is answered before anything is asked of the app."""
        from backend.mac_ae import ppt
        from ppt_mac.animation import _remove_animation_impl

        def _explode(*args, **kwargs):
            raise AssertionError("a refusal must not touch PowerPoint")

        original = ppt._get_app_impl
        ppt._get_app_impl = _explode
        try:
            assert "error" in _remove_animation_impl(1, 1, 1)
        finally:
            ppt._get_app_impl = original


@macos_only
class TestAnimationRemoval:
    """There is no way to remove one animation here, so nothing is attempted."""

    def test_removal_is_refused_and_names_the_shape(self):
        from ppt_mac.animation import _remove_animation_impl

        with _fake_deck(["Title", "Body", "Footer"]) as deck:
            result = _remove_animation_impl(1, 2, None)

        assert "error" in result
        assert "Body" in result["reason"]
        assert result["alternatives"] == [
            "ppt_list_animations", "ppt_clear_animations", "ppt_add_animation",
        ]
        # The whole point: nothing was cleared, so nothing was flattened.
        assert deck.cleared == []

    def test_the_error_line_names_the_tool_limit_not_the_tool(self):
        from ppt_mac.animation import _remove_animation_impl

        with _fake_deck(["Body", "Title", "Body"]) as deck:
            result = _remove_animation_impl(1, 1, None)

        assert result["error"] == (
            "ppt_remove_animation cannot remove a single animation on macOS"
        )
        assert deck.cleared == []

    def test_an_index_past_the_end_says_the_range(self):
        from ppt_mac.animation import _remove_animation_impl

        with _fake_deck(["Title"]):
            with pytest.raises(ValueError, match=r"out of range \(1-1\)"):
                _remove_animation_impl(1, 4, None)


@macos_only
class TestSlideTransition:
    """Seven of the eleven transitions Windows names exist here."""

    def test_a_transition_macos_has_is_set_and_read_back(self):
        from appscript import k

        from ppt_mac.animation import _set_slide_transition_impl

        with _fake_deck([], shape_names=["Title"]) as deck:
            result = _set_slide_transition_impl(1, "dissolve", 1.25, True, False, None)

        assert result == {"success": True, "slide_index": 1, "effect": 1537}
        assert deck.transition.entry_effect() == k.entry_effect_dissolve
        assert deck.transition.transition_duration() == 1.25
        assert deck.transition.advance_on_click() is True

    def test_a_transition_macos_lacks_raises_rather_than_landing_nearby(self):
        from ppt_mac.animation import _set_slide_transition_impl

        with _fake_deck([], shape_names=["Title"]) as deck:
            with pytest.raises(ValueError, match="push, wipe, split and reveal"):
                _set_slide_transition_impl(1, "push", None, None, None, None)

        assert deck.transition.entry_effect() is None


@macos_only
class TestAnimationAdding:
    """Adding works; the parts of the Windows call that cannot follow say so."""

    def test_the_effect_is_refetched_by_index_rather_than_taken_from_the_add(self):
        from ppt_mac.animation import _add_animation_impl

        with _fake_deck([], shape_names=["Title"]) as deck:
            result = _add_animation_impl(
                1, "Title", "fade", "after_previous", 0.75, None, True,
                "left", 2, True, True, True, True,
                None, None, None,
                None, "by_word", True, False,
            )

        assert result["success"] is True
        assert result["animation_index"] == 1
        assert result["shape_name"] == "Title"
        assert result["effect"] == 10
        assert len(deck.effects) == 1
        assert deck.effects[0].timing.duration() == 0.75
        assert deck.effects[0].timing.repeat_count() == 2
        assert deck.effects[0].exit_animation() is True

    def test_the_trigger_goes_in_at_the_add_because_nothing_else_takes_one(self):
        from appscript import k

        from ppt_mac.animation import _add_animation_impl

        with _fake_deck([], shape_names=["Title"]) as deck:
            _add_animation_impl(
                1, "Title", "zoom", "after_previous", None, None, False,
                None, None, None, None, None, None,
                None, None, None,
                "first_level", None, None, None,
            )

        added = deck.added[0]
        assert added["fx"] == k.animation_type_zoom
        assert added["trigger"] == k.after_previous
        assert added["level"] == k.text_by_first_level

    def test_a_delay_is_reported_rather_than_dropped(self):
        from ppt_mac.animation import _add_animation_impl

        with _fake_deck([], shape_names=["Title"]):
            result = _add_animation_impl(
                1, "Title", "fade", "on_click", None, 0.5, False,
                None, None, None, None, None, None,
                None, None, None,
                None, None, None, None,
            )

        assert any("delay was not applied" in w for w in result["warnings"])

    def test_a_dim_colour_is_never_written(self):
        """Writing one wipes every animation on the slide, so it is not written."""
        from ppt_mac.animation import _add_animation_impl

        with _fake_deck([], shape_names=["Title"]) as deck:
            result = _add_animation_impl(
                1, "Title", "fade", "on_click", None, None, False,
                None, None, None, None, None, None,
                None, "dim", "#808080",
                None, None, None, None,
            )

        shape = deck.shape("Title")
        assert shape.animation_settings.dim_color() is None
        dim = [w for w in result["warnings"] if "dim_color" in w]
        assert len(dim) == 1
        assert "#808080" in dim[0]
        assert "wipes every animation" in dim[0]
        assert any("after_effect" in w and "-1708" in w for w in result["warnings"])

    def test_animate_in_reverse_is_refused_rather_than_attempted(self):
        from ppt_mac.animation import _add_animation_impl

        with _fake_deck([], shape_names=["Title"]) as deck:
            result = _add_animation_impl(
                1, "Title", "fade", "on_click", None, None, False,
                None, None, None, None, None, None,
                None, None, None,
                None, None, True, None,
            )

        shape = deck.shape("Title")
        assert shape.animation_settings.animate_text_in_reverse() is not True
        assert any(
            "animate_in_reverse was not applied" in w for w in result["warnings"]
        )

    def test_a_shape_wide_setting_says_it_is_shape_wide(self):
        from ppt_mac.animation import _add_animation_impl

        with _fake_deck([], shape_names=["Title"]):
            result = _add_animation_impl(
                1, "Title", "fade", "on_click", None, None, False,
                None, None, None, None, None, None,
                None, None, None,
                None, None, None, True,
            )

        assert any("belongs to the shape here" in w for w in result["warnings"])


@macos_only
class TestAnimationUpdating:
    """Updating an effect in place, and the three things that cannot be updated."""

    def test_what_cannot_be_changed_comes_back_as_warnings_not_a_failure(self):
        from ppt_mac.animation import _update_animation_impl

        with _fake_deck(["Title"]) as deck:
            result = _update_animation_impl(
                1, 1, None,
                "zoom", "with_previous", 1.5, 0.25, 3, True,
                "up", None, None, None, None, None,
                None, None,
                "first_level", None, None, None,
            )

        assert result["success"] is True
        assert result["animation_index"] == 1
        assert deck.effects[0].timing.duration() == 1.5
        assert deck.effects[0].exit_animation() is True
        joined = " ".join(result["warnings"])
        for phrase in ("trigger was not applied", "delay was not applied",
                       "move_to was not applied", "build_level was not applied"):
            assert phrase in joined

    def test_the_unreadable_fields_report_nothing_rather_than_a_guess(self):
        from ppt_mac.animation import _update_animation_impl

        with _fake_deck(["Title"]):
            result = _update_animation_impl(
                1, 1, None,
                None, None, 2.0, None, None, None,
                None, None, None, None, None, None,
                None, None,
                None, None, None, None,
            )

        assert result["trigger_type"] is None
        assert result["trigger_name"] is None
        assert result["delay"] is None


@macos_only
class TestAnimationListing:
    """Reading the timeline, without ever asking for the whole of it."""

    def test_the_effects_collection_is_never_materialised(self):
        """The fake raises if anything asks for it, which is what -609 costs."""
        from ppt_mac.animation import _clear_animations_impl, _list_animations_impl

        with _fake_deck(["Title", "Body"]):
            _list_animations_impl(1)
            _clear_animations_impl(1, False)

    def test_it_reports_the_keys_windows_reports(self):
        from ppt_mac.animation import _list_animations_impl

        with _fake_deck(["Title"]):
            result = _list_animations_impl(1)

        assert result["main_sequence_count"] == 1
        assert result["interactive_sequences"] == []
        assert result["interactive_count"] == 0
        animation = result["animations"][0]
        for key in (
            "index", "shape_name", "effect_type", "effect_name", "trigger_type",
            "trigger_name", "duration", "exit", "category", "direction",
            "direction_name", "after_effect", "after_effect_name",
            "build_level", "build_level_name", "text_unit_effect",
            "text_unit_effect_name", "animate_in_reverse", "animate_background",
        ):
            assert key in animation
        assert animation["index"] == 1
        assert animation["shape_name"] == "Title"

    def test_macos_keywords_come_back_as_the_windows_numbers(self):
        """A caller reading a deck on a Mac has to see the numbers Windows reports."""
        from ppt_mac.animation import _list_animations_impl

        with _fake_deck(["Title"]):
            animation = _list_animations_impl(1)["animations"][0]

        assert (animation["effect_type"], animation["effect_name"]) == (10, "fade")
        assert (animation["direction"], animation["direction_name"]) == (3, "down")
        assert (animation["after_effect"], animation["after_effect_name"]) == (1, "dim")
        assert (animation["build_level"], animation["build_level_name"]) == (0, "none")
        assert animation["text_unit_effect_name"] == "by_paragraph"
        assert animation["animate_background"] is True
        assert animation["duration"] == 0.5
        assert animation["category"] == "entrance"

    def test_a_trigger_is_reported_as_unknown_rather_than_guessed(self):
        """Nothing on a macOS effect carries a trigger, so None is the honest answer."""
        from ppt_mac.animation import _list_animations_impl

        with _fake_deck(["Title"]):
            animation = _list_animations_impl(1)["animations"][0]

        assert animation["trigger_type"] is None
        assert animation["trigger_name"] is None

    def test_clearing_counts_before_and_after_rather_than_reporting_intent(self):
        from ppt_mac.animation import _clear_animations_impl

        with _fake_deck(["Title", "Body", "Body"]) as deck:
            result = _clear_animations_impl(1, True)

        assert result["cleared_count"] == 3
        assert result["remaining_count"] == 0
        assert result["interactive_cleared"] == 0
        assert deck.transition_cleared is True


# ---------------------------------------------------------------------------
# A slide made of stand-ins, for the decisions that are worth testing without
# PowerPoint. Only the parts the animation tools touch are modelled, and the
# effects collection deliberately raises if anything asks for it whole, because
# that is the call that kills the application.
# ---------------------------------------------------------------------------
class _FakeCollection:
    def __init__(self, items):
        self._items = items

    def get(self):
        return list(self._items)

    def __getitem__(self, index):
        if index < 1 or index > len(self._items):
            raise _command_error(-1728)
        return self._items[index - 1]


class _FakeEffects(_FakeCollection):
    def get(self):
        raise AssertionError("the effects collection must never be materialised")

    def count(self, each=None):
        raise AssertionError("the effects collection must never be counted")


class _FakeProperty:
    def __init__(self, on_set, value=None):
        self._on_set = on_set
        self._value = value

    def __call__(self):
        return self._value

    def get(self):
        return self._value

    def set(self, value):
        self._value = value
        self._on_set(value)


class _FakeShape:
    def __init__(self, deck, name):
        self._deck = deck
        self._name = name
        self.animation_settings = _bag(
            animate_text_in_reverse=False,
            animate_background=False,
            dim_color=None,
        )
        self.animation_settings.animate = _FakeProperty(self._on_animate, True)

    def name(self):
        return self._name

    def _on_animate(self, value):
        if not value:
            self._deck.clear_shape(self._name)


def _bag(**properties):
    """A stand-in for one of PowerPoint's little property-only classes."""
    holder = type("Bag", (), {})()
    for name, value in properties.items():
        setattr(holder, name, _FakeProperty(lambda _v: None, value))
    return holder


class _FakeEffect:
    def __init__(self, shape):
        from appscript import k

        self.shape = shape
        self.animation_effect_type = _FakeProperty(
            lambda _v: None, k.animation_type_fade
        )
        self.exit_animation = _FakeProperty(lambda _v: None, False)
        self.timing = _bag(
            duration=0.5, repeat_count=0, autoreverse=False, rewind=False,
            smooth_start=True, smooth_end=True,
        )
        self.effect_parameters = _bag(direction=k.down)
        self.effect_information = _bag(
            after_effect_information=k.dim,
            build_by_level=k.text_by_no_levels,
            text_unit_effect_information=k.by_paragraph,
            animate_text_in_reverse_information=False,
            animate_background_information=True,
        )

    def get(self):
        return self


class _FakeSequence:
    def __init__(self, deck):
        self._deck = deck

    @property
    def effects(self):
        return _FakeEffects(self._deck.effects)

    def count(self, each=None):
        return len(self._deck.effects)

    def add_effect(self, **kwargs):
        self._deck.added.append(kwargs)
        self._deck.effects.append(_FakeEffect(kwargs["for_"]))

    def convert_to_text_unit_effect(self, Effect=None, unit=None):  # noqa: N803
        self._deck.converted.append(unit)


class _FakeDeck:
    """One slide, one shape per distinct name, one effect per name given."""

    def __init__(self, effect_shape_names, shape_names=None):
        self.cleared: list = []
        self.added: list = []
        self.converted: list = []
        self.transition_cleared = False
        names = list(dict.fromkeys(list(shape_names or []) + list(effect_shape_names)))
        self._shapes = {name: _FakeShape(self, name) for name in names}
        self.effects = [
            _FakeEffect(self._shapes[name]) for name in effect_shape_names
        ]

    def shape(self, name):
        return self._shapes[name]

    def clear_shape(self, name):
        self.cleared.append(name)
        self.effects = [e for e in self.effects if e.shape.name() != name]

    def clear_transition(self, _value):
        self.transition_cleared = True

    @property
    def slide(self):
        deck = self
        sequence = _FakeSequence(self)

        transition = _bag(
            transition_duration=None,
            advance_on_click=None,
            advance_on_time=None,
            advance_time=None,
        )
        transition.entry_effect = _FakeProperty(deck.clear_transition)

        class _Slide:
            shapes = _FakeCollection(list(deck._shapes.values()))
            timeline = type("Timeline", (), {"main_sequence": sequence})()
            slide_show_transition = transition

        deck.transition = transition
        return _Slide()

    @property
    def presentation(self):
        return type("Pres", (), {"slides": _FakeCollection([self.slide])})()


class _fake_deck:  # noqa: N801 - reads as a context manager, not a class
    """Point the wrapper at a fake slide for the length of a `with` block."""

    def __init__(self, effect_shape_names, shape_names=None):
        self._deck = _FakeDeck(effect_shape_names, shape_names)

    def __enter__(self):
        from backend.mac_ae import ppt

        self._ppt = ppt
        self._app = ppt._get_app_impl
        self._pres = ppt._get_pres_impl
        ppt._get_app_impl = lambda *a, **kw: object()
        ppt._get_pres_impl = lambda *a, **kw: self._deck.presentation
        return self._deck

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


@macos_only
class TestQueueingBehindAnotherCall:
    """Only one request at a time reaches PowerPoint, and the rest wait.

    A caller that stopped waiting used to leave its work in the queue, where the
    worker ran it later against a deck that had moved on. That is how a tool call
    reported failure with no message at all and put its picture on the slide
    regardless, and how retrying it put a second one there.
    """

    def _wrapper(self):
        from backend.mac_ae import PowerPointAppleEventWrapper

        wrapper = PowerPointAppleEventWrapper()
        wrapper.start()
        return wrapper

    def test_work_taken_back_after_the_wait_never_runs(self, monkeypatch):
        import threading

        from backend import mac_ae

        monkeypatch.setattr(mac_ae, "_QUEUE_WAIT", 0.2)
        release = threading.Event()
        ran = []

        wrapper = self._wrapper()
        try:
            blocker = threading.Thread(
                target=lambda: wrapper.execute(lambda: release.wait(5)), daemon=True
            )
            blocker.start()
            # Let the blocker reach the worker before the second call queues.
            time_waited = 0.0
            while wrapper._queue.unfinished_tasks == 0 and time_waited < 2:
                time.sleep(0.01)
                time_waited += 0.01

            with pytest.raises(mac_ae.AppleEventError) as caught:
                wrapper.execute(lambda: ran.append("second"))
            assert "one after another" in str(caught.value)

            release.set()
            blocker.join(5)
            # The worker is now free. Give it every chance to run the dropped
            # job, so this fails if the job was left in the queue.
            wrapper.execute(lambda: None)
            assert ran == []
        finally:
            release.set()
            wrapper.stop()

    def test_a_call_is_not_charged_for_the_time_it_spent_in_line(self, monkeypatch):
        """It waits its turn and then runs, rather than failing at the back."""
        import threading

        from backend import mac_ae

        monkeypatch.setattr(mac_ae, "_QUEUE_WAIT", 5)
        release = threading.Event()

        wrapper = self._wrapper()
        try:
            blocker = threading.Thread(
                target=lambda: wrapper.execute(lambda: release.wait(5)), daemon=True
            )
            blocker.start()
            time.sleep(0.1)
            releaser = threading.Timer(0.3, release.set)
            releaser.start()
            assert wrapper.execute(lambda: "landed") == "landed"
            blocker.join(5)
        finally:
            release.set()
            wrapper.stop()


@macos_only
class TestTheWaitsAreTheRightWayRound:
    """A call that would have recovered must not be taken back while it waits."""

    def test_a_whole_call_fits_inside_the_queue_wait(self):
        from backend import mac_ae

        # One call in front that times out and is retried spends the whole
        # budget. A shorter queue wait would take back everything behind it.
        assert mac_ae._QUEUE_WAIT >= mac_ae._CALL_BUDGET

    def test_the_call_budget_covers_every_retry_and_the_pauses_between(self):
        from backend import mac_ae

        attempts = mac_ae.DEFAULT_TIMEOUT * (mac_ae._RETRY_MAX + 1)
        pauses = mac_ae._RETRY_INTERVAL * mac_ae._RETRY_MAX
        assert mac_ae._CALL_BUDGET >= attempts + pauses


@macos_only
class TestWhatIsSafeToRunTwice:
    """One impl is many Apple Events, so a failure part-way through is not a
    failure to start. Re-running it repeats whatever already landed, which is
    how a caller once got two copies of the same picture. Windows reached the
    same conclusion in #200; this is the macOS half of it.
    """

    @staticmethod
    def _wrapper():
        from backend.mac_ae import PowerPointAppleEventWrapper

        ppt = PowerPointAppleEventWrapper()
        ppt.start()
        return ppt

    def test_an_editing_call_is_not_sent_again(self):
        from backend.mac_ae import AE_TIMED_OUT, AppleEventError

        calls = []

        def edits_the_deck():
            calls.append(1)
            raise _command_error(AE_TIMED_OUT)

        ppt = self._wrapper()
        with pytest.raises(AppleEventError) as caught:
            ppt.execute(edits_the_deck)

        assert calls == [1], "the deck was edited twice"
        assert "may already have been applied" in str(caught.value)

    def test_a_caller_that_says_it_is_safe_is_retried(self):
        from backend.mac_ae import AE_TIMED_OUT

        calls = []

        def connecting():
            calls.append(1)
            if len(calls) < 3:
                raise _command_error(AE_TIMED_OUT)
            return "connected"

        ppt = self._wrapper()
        assert ppt.execute(connecting, idempotent=True) == "connected"
        assert len(calls) == 3

    def test_idempotent_is_not_passed_to_the_work(self):
        """`ppt_connect` passes it, and every impl would choke on it.

        Develop added `idempotent=True` at that call site while this backend
        forwarded unknown keywords straight through, so the tool raised a
        TypeError on macOS until the keyword became one this wrapper owns.
        """
        seen = {}

        def takes_one_argument(visible):
            seen["visible"] = visible
            return "ok"

        ppt = self._wrapper()
        assert ppt.execute(takes_one_argument, True, idempotent=True) == "ok"
        assert seen == {"visible": True}


@macos_only
class TestWorkGoesAwayWithTheCallerThatQueuedIt:
    """macOS had half of #198 and #199 and not the other half.

    `utils.offload` frees the event loop and, when a caller goes away, cancels
    everything that caller had queued. It cancels through
    `pending_com_futures`, which this backend registered nothing with, so a
    cancelled request's queue still ran, later, against a deck that had moved
    on. That is the shape the two-stage wait was built for, arriving by a
    different door.

    Written the way `tests/test_offload_cancel.py` writes the Windows half,
    through `run_offloaded` itself, because the context only reaches the
    worker thread when anyio puts it there. A plain `threading.Thread` starts
    with an empty context and would pass this test while proving nothing.
    """

    @staticmethod
    def _wrapper():
        from backend.mac_ae import PowerPointAppleEventWrapper

        ppt = PowerPointAppleEventWrapper()
        ppt.start()
        return ppt

    def test_a_cancelled_request_takes_its_queue_with_it(self, monkeypatch):
        import threading

        import anyio

        from backend import mac_ae
        from utils.offload import run_offloaded

        monkeypatch.setattr(mac_ae, "_QUEUE_WAIT", 1.0)
        ppt = self._wrapper()
        holding = threading.Event()
        release = threading.Event()
        entered = threading.Event()
        ran = []

        # Fill the worker, so the job below waits in line rather than running.
        threading.Thread(
            target=lambda: ppt.execute(lambda: (holding.set(), release.wait(5))),
            daemon=True,
        ).start()
        holding.wait(5)

        def queues_an_edit():
            entered.set()
            try:
                ppt.execute(lambda: ran.append("edited"))
            except Exception:  # noqa: BLE001 - not running it is the point
                pass

        async def main():
            async with anyio.create_task_group() as tg:
                tg.start_soon(run_offloaded, queues_an_edit)
                await anyio.to_thread.run_sync(entered.wait, 5)
                await anyio.sleep(0.2)
                tg.cancel_scope.cancel()

        anyio.run(main)
        release.set()
        threading.Event().wait(0.5)

        assert ran == [], "the abandoned edit reached PowerPoint anyway"

    def test_nothing_watching_changes_nothing(self):
        """Internal callers have no request behind them, and still work."""
        ppt = self._wrapper()
        assert ppt.execute(lambda: "done") == "done"

    def test_the_cancelled_caller_is_freed_at_once(self, monkeypatch):
        """Discarding the work is only half of it.

        `drop` marked the job and woke nobody, so the thread that queued it sat
        in `started.wait()` for the entire queue budget, which is over two
        minutes. The Apple Event was correctly discarded and a burst of
        cancellations could still hold the shared thread pool closed behind it.
        """
        import threading
        import time

        from backend import mac_ae

        monkeypatch.setattr(mac_ae, "_QUEUE_WAIT", 30.0)
        ppt = self._wrapper()
        holding = threading.Event()
        release = threading.Event()
        took = []

        threading.Thread(
            target=lambda: ppt.execute(lambda: (holding.set(), release.wait(5))),
            daemon=True,
        ).start()
        holding.wait(5)

        job_seen = []
        real_put = ppt._queue.put

        def capture(job):
            job_seen.append(job)
            real_put(job)

        monkeypatch.setattr(ppt._queue, "put", capture)

        def waits():
            start = time.monotonic()
            try:
                ppt.execute(lambda: None)
            except Exception:  # noqa: BLE001 - being freed is the point
                pass
            took.append(time.monotonic() - start)

        waiter = threading.Thread(target=waits, daemon=True)
        waiter.start()
        while not job_seen:
            time.sleep(0.02)
        time.sleep(0.1)
        assert job_seen[0].cancel() is True

        waiter.join(5)
        release.set()
        assert took, "the caller never returned"
        assert took[0] < 2.0, f"freed only after {took[0]:.1f}s, budget was 30s"
