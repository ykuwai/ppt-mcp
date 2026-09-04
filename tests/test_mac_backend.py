"""Tests for the macOS Apple Event backend.

The colour helpers are pure and run everywhere. Everything else needs
appscript, which only installs on macOS, so those are skipped elsewhere. None
of this launches PowerPoint; the live behaviour is covered by MACOS_PORT.md and
by running the server.
"""

import sys

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
