"""Tests for src/utils/color.py and src/utils/units.py.

Pure Python tests — no COM or PowerPoint dependency required.
"""

import sys

sys.path.insert(0, "src")

import pytest

from utils.color import (
    THEME_COLOR_MAP,
    get_theme_color_index,
    hex_to_int,
    hex_to_rgb,
    int_to_hex,
    int_to_rgb,
    rgb_to_int,
)
from utils.units import (
    CM_PER_INCH,
    EMU_PER_CM,
    EMU_PER_INCH,
    EMU_PER_POINT,
    POINTS_PER_CM,
    POINTS_PER_INCH,
    SLIDE_HEIGHT_16_9,
    SLIDE_HEIGHT_4_3,
    SLIDE_WIDTH_16_9,
    SLIDE_WIDTH_4_3,
    cm_to_emu,
    cm_to_points,
    emu_to_cm,
    emu_to_inches,
    emu_to_points,
    inches_to_emu,
    inches_to_points,
    points_to_cm,
    points_to_emu,
    points_to_inches,
)


# ============================================================
# color.py — rgb_to_int
# ============================================================


class TestRgbToInt:
    """Tests for rgb_to_int(r, g, b) -> BGR integer."""

    def test_black(self):
        assert rgb_to_int(0, 0, 0) == 0

    def test_white(self):
        assert rgb_to_int(255, 255, 255) == 0xFFFFFF

    def test_pure_red(self):
        # Red channel sits in lowest byte
        assert rgb_to_int(255, 0, 0) == 255

    def test_pure_green(self):
        assert rgb_to_int(0, 255, 0) == 255 << 8

    def test_pure_blue(self):
        assert rgb_to_int(0, 0, 255) == 255 << 16

    def test_arbitrary_color(self):
        # Coral: (255, 127, 80)
        expected = 255 + (127 << 8) + (80 << 16)
        assert rgb_to_int(255, 127, 80) == expected

    def test_grey_128(self):
        expected = 128 + (128 << 8) + (128 << 16)
        assert rgb_to_int(128, 128, 128) == expected


# ============================================================
# color.py — int_to_rgb
# ============================================================


class TestIntToRgb:
    """Tests for int_to_rgb(color_int) -> (R, G, B)."""

    def test_black(self):
        assert int_to_rgb(0) == (0, 0, 0)

    def test_white(self):
        assert int_to_rgb(0xFFFFFF) == (255, 255, 255)

    def test_pure_red(self):
        assert int_to_rgb(255) == (255, 0, 0)

    def test_pure_green(self):
        assert int_to_rgb(255 << 8) == (0, 255, 0)

    def test_pure_blue(self):
        assert int_to_rgb(255 << 16) == (0, 0, 255)

    def test_arbitrary_color(self):
        bgr = 100 + (150 << 8) + (200 << 16)
        assert int_to_rgb(bgr) == (100, 150, 200)


# ============================================================
# color.py — round-trip rgb_to_int <-> int_to_rgb
# ============================================================


class TestRgbRoundTrip:
    """Verify rgb_to_int and int_to_rgb are true inverses."""

    @pytest.mark.parametrize(
        "r, g, b",
        [
            (0, 0, 0),
            (255, 255, 255),
            (255, 0, 0),
            (0, 255, 0),
            (0, 0, 255),
            (1, 2, 3),
            (128, 64, 32),
            (10, 200, 77),
        ],
    )
    def test_round_trip(self, r, g, b):
        assert int_to_rgb(rgb_to_int(r, g, b)) == (r, g, b)


# ============================================================
# color.py — hex_to_rgb
# ============================================================


class TestHexToRgb:
    """Tests for hex_to_rgb(hex_str) -> (R, G, B)."""

    def test_with_hash(self):
        assert hex_to_rgb("#FF0000") == (255, 0, 0)

    def test_without_hash(self):
        assert hex_to_rgb("00FF00") == (0, 255, 0)

    def test_lowercase(self):
        assert hex_to_rgb("#ff8040") == (255, 128, 64)

    def test_mixed_case(self):
        assert hex_to_rgb("#aAbBcC") == (170, 187, 204)

    def test_black(self):
        assert hex_to_rgb("#000000") == (0, 0, 0)

    def test_white(self):
        assert hex_to_rgb("#FFFFFF") == (255, 255, 255)

    def test_shorthand_3_char(self):
        # "#F00" -> "#FF0000"
        assert hex_to_rgb("#F00") == (255, 0, 0)

    def test_shorthand_3_char_mixed(self):
        # "#abc" -> "#aabbcc"
        assert hex_to_rgb("#abc") == (170, 187, 204)

    def test_invalid_length_raises(self):
        with pytest.raises(ValueError, match="Invalid hex color"):
            hex_to_rgb("#12345")

    def test_invalid_too_long_raises(self):
        with pytest.raises(ValueError, match="Invalid hex color"):
            hex_to_rgb("#1234567")

    def test_invalid_single_char_raises(self):
        with pytest.raises(ValueError, match="Invalid hex color"):
            hex_to_rgb("#A")


# ============================================================
# color.py — hex_to_int
# ============================================================


class TestHexToInt:
    """Tests for hex_to_int(hex_str) -> BGR integer."""

    def test_red(self):
        assert hex_to_int("#FF0000") == 255

    def test_blue(self):
        assert hex_to_int("#0000FF") == 255 << 16

    def test_coral(self):
        expected = rgb_to_int(255, 127, 80)
        assert hex_to_int("#FF7F50") == expected


# ============================================================
# color.py — int_to_hex
# ============================================================


class TestIntToHex:
    """Tests for int_to_hex(color_int) -> '#RRGGBB'."""

    def test_black(self):
        assert int_to_hex(0) == "#000000"

    def test_white(self):
        assert int_to_hex(0xFFFFFF) == "#FFFFFF"

    def test_pure_red(self):
        assert int_to_hex(255) == "#FF0000"

    def test_pure_blue(self):
        assert int_to_hex(255 << 16) == "#0000FF"

    def test_leading_zeros(self):
        # RGB (1, 2, 3) -> "#010203"
        bgr = rgb_to_int(1, 2, 3)
        assert int_to_hex(bgr) == "#010203"


# ============================================================
# color.py — hex round-trip
# ============================================================


class TestHexRoundTrip:
    """Verify hex_to_int and int_to_hex are true inverses."""

    @pytest.mark.parametrize(
        "hex_str",
        ["#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#1A2B3C"],
    )
    def test_round_trip(self, hex_str):
        assert int_to_hex(hex_to_int(hex_str)) == hex_str


# ============================================================
# color.py — get_theme_color_index
# ============================================================


class TestGetThemeColorIndex:
    """Tests for get_theme_color_index(name)."""

    def test_all_valid_names(self):
        for name, expected in THEME_COLOR_MAP.items():
            assert get_theme_color_index(name) == expected

    def test_case_insensitive(self):
        assert get_theme_color_index("Accent1") == 5
        assert get_theme_color_index("DARK1") == 1

    def test_hyphen_separator(self):
        assert get_theme_color_index("followed-hyperlink") == 12

    def test_space_separator(self):
        assert get_theme_color_index("followed hyperlink") == 12

    def test_unknown_name_raises(self):
        with pytest.raises(ValueError, match="Unknown theme color"):
            get_theme_color_index("nonexistent")


# ============================================================
# units.py — Constants
# ============================================================


class TestUnitConstants:
    """Verify unit constants have expected values."""

    def test_points_per_inch(self):
        assert POINTS_PER_INCH == 72.0

    def test_cm_per_inch(self):
        assert CM_PER_INCH == 2.54

    def test_points_per_cm(self):
        assert POINTS_PER_CM == pytest.approx(72.0 / 2.54)

    def test_emu_per_point(self):
        assert EMU_PER_POINT == 12700

    def test_emu_per_inch(self):
        assert EMU_PER_INCH == 914400

    def test_emu_per_cm(self):
        assert EMU_PER_CM == 360000

    def test_slide_16_9(self):
        assert SLIDE_WIDTH_16_9 == 960.0
        assert SLIDE_HEIGHT_16_9 == 540.0

    def test_slide_4_3(self):
        assert SLIDE_WIDTH_4_3 == 720.0
        assert SLIDE_HEIGHT_4_3 == 540.0


# ============================================================
# units.py — inches <-> points
# ============================================================


class TestInchesPoints:
    """Tests for inches_to_points and points_to_inches."""

    def test_zero_inches(self):
        assert inches_to_points(0) == 0.0

    def test_one_inch(self):
        assert inches_to_points(1) == 72.0

    def test_fractional_inch(self):
        assert inches_to_points(0.5) == pytest.approx(36.0)

    def test_zero_points(self):
        assert points_to_inches(0) == 0.0

    def test_72_points(self):
        assert points_to_inches(72) == 1.0

    def test_round_trip(self):
        for val in [0, 1, 2.5, 10, 0.125]:
            assert points_to_inches(inches_to_points(val)) == pytest.approx(val)


# ============================================================
# units.py — cm <-> points
# ============================================================


class TestCmPoints:
    """Tests for cm_to_points and points_to_cm."""

    def test_zero_cm(self):
        assert cm_to_points(0) == 0.0

    def test_one_cm(self):
        assert cm_to_points(1) == pytest.approx(POINTS_PER_CM)

    def test_2_54_cm_equals_1_inch(self):
        assert cm_to_points(2.54) == pytest.approx(72.0)

    def test_zero_points(self):
        assert points_to_cm(0) == 0.0

    def test_round_trip(self):
        for val in [0, 1, 2.54, 5, 10.5]:
            assert points_to_cm(cm_to_points(val)) == pytest.approx(val)


# ============================================================
# units.py — EMU <-> points
# ============================================================


class TestEmuPoints:
    """Tests for emu_to_points and points_to_emu."""

    def test_zero_emu(self):
        assert emu_to_points(0) == 0.0

    def test_one_point_in_emu(self):
        assert emu_to_points(12700) == pytest.approx(1.0)

    def test_zero_points_to_emu(self):
        assert points_to_emu(0) == 0

    def test_one_point_to_emu(self):
        assert points_to_emu(1) == 12700

    def test_points_to_emu_rounds(self):
        # 0.5 points -> 6350 EMU
        assert points_to_emu(0.5) == 6350

    def test_round_trip(self):
        for pt in [0, 1, 10, 72, 100]:
            assert emu_to_points(points_to_emu(pt)) == pytest.approx(pt)


# ============================================================
# units.py — inches <-> EMU
# ============================================================


class TestInchesEmu:
    """Tests for inches_to_emu and emu_to_inches."""

    def test_zero(self):
        assert inches_to_emu(0) == 0
        assert emu_to_inches(0) == 0.0

    def test_one_inch(self):
        assert inches_to_emu(1) == 914400

    def test_one_inch_reverse(self):
        assert emu_to_inches(914400) == pytest.approx(1.0)

    def test_round_trip(self):
        for val in [0, 1, 0.5, 2.5]:
            assert emu_to_inches(inches_to_emu(val)) == pytest.approx(val)


# ============================================================
# units.py — cm <-> EMU
# ============================================================


class TestCmEmu:
    """Tests for cm_to_emu and emu_to_cm."""

    def test_zero(self):
        assert cm_to_emu(0) == 0
        assert emu_to_cm(0) == 0.0

    def test_one_cm(self):
        assert cm_to_emu(1) == 360000

    def test_one_cm_reverse(self):
        assert emu_to_cm(360000) == pytest.approx(1.0)

    def test_round_trip(self):
        for val in [0, 1, 2.54, 5, 10]:
            assert emu_to_cm(cm_to_emu(val)) == pytest.approx(val)


# ============================================================
# units.py — cross-unit consistency
# ============================================================


class TestCrossUnitConsistency:
    """Verify that different conversion paths produce the same result."""

    def test_inch_via_points_vs_direct_emu(self):
        """1 inch converted to EMU via points should match direct conversion."""
        via_points = points_to_emu(inches_to_points(1))
        direct = inches_to_emu(1)
        assert via_points == direct

    def test_cm_via_points_vs_direct_emu(self):
        """1 cm converted to EMU via points should match direct conversion."""
        via_points = points_to_emu(cm_to_points(1))
        direct = cm_to_emu(1)
        # Allow small rounding difference because points_to_emu rounds
        assert via_points == pytest.approx(direct, abs=1)
