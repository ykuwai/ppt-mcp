"""Tests for the macOS theme and header footer tools.

Nothing here launches PowerPoint. The object graph is faked down to the few
properties these tools touch, which is enough to pin the two things that are
easy to get wrong: the colour conversion between the packed integer every tool
in this project speaks and the three element list Apple Events take, and the
route to the theme colour scheme, which is through the slide master and not
through `designs`.
"""

import sys
from contextlib import contextmanager
from unittest import mock

import pytest

sys.path.insert(0, "src")

macos_only = pytest.mark.skipif(
    sys.platform != "darwin", reason="the Apple Event backend needs appscript"
)


class _Property:
    """One readable and writable Apple Event property."""

    def __init__(self, value=None, raises=None):
        self.value = value
        self.raises = raises
        self.writes = []

    def __call__(self):
        if self.raises:
            raise self.raises
        return self.value

    def set(self, value):
        if self.raises:
            raise self.raises
        self.writes.append(value)
        self.value = value


class _ThemeColor:
    def __init__(self, rgb):
        self.RGB = _Property(list(rgb))


class _Indexed:
    """A one based element collection, reached only by index."""

    def __init__(self, items):
        self._items = items

    def __getitem__(self, i):
        return self._items[i - 1]


class _HeaderOrFooter:
    def __init__(self, raises=None):
        self.visible = _Property(False, raises)
        self.header_footer_text = _Property("", raises)
        self.date_format = _Property(None, raises)
        self.use_date_format = _Property(True, raises)


class _Slide:
    def __init__(self, raises=None):
        self.headers_and_footers = mock.Mock()
        self.headers_and_footers.footer = _HeaderOrFooter(raises)
        self.headers_and_footers.slide_number = _HeaderOrFooter(raises)
        self.headers_and_footers.date_and_time = _HeaderOrFooter(raises)


DEFAULT_COLOURS = [
    (0, 0, 0), (255, 255, 255), (14, 40, 65), (232, 232, 232),
    (21, 96, 130), (233, 113, 50), (25, 107, 36), (15, 158, 213),
    (160, 43, 43), (77, 77, 77), (70, 120, 134), (150, 97, 137),
]


class _Deck:
    def __init__(self, slide_count=2, slide_raises=None):
        self.theme_colors = _Indexed([_ThemeColor(c) for c in DEFAULT_COLOURS])
        scheme = mock.Mock()
        scheme.theme_colors = self.theme_colors
        theme = mock.Mock()
        theme.theme_color_scheme = scheme
        self.slide_master = mock.Mock()
        self.slide_master.theme = theme
        self.slides = _Indexed([_Slide(slide_raises) for _ in range(slide_count)])
        self.designs = object()   # empty on a real deck, see the module docstring
        self._slide_count = slide_count


@contextmanager
def _fake_deck(slide_count=2, slide_raises=None):
    from ppt_mac import themes as mac_themes

    deck = _Deck(slide_count, slide_raises)
    with mock.patch.object(mac_themes.ppt, "_get_pres_impl", return_value=deck), \
            mock.patch.object(mac_themes.ppt, "_get_app_impl", return_value=mock.Mock()), \
            mock.patch.object(mac_themes, "count", lambda ref: slide_count
                              if ref is deck.slides else 0):
        yield deck


@macos_only
class TestThemeColours:
    """Twelve colours, read and written through the slide master."""

    def test_all_twelve_come_back_as_hex(self):
        from ppt_mac.themes import _get_theme_colors_impl

        with _fake_deck():
            result = _get_theme_colors_impl()

        assert result["success"] is True
        assert len(result["colors"]) == 12
        assert result["colors"][0] == {
            "index": 1, "name": "dark1", "color_hex": "#000000",
        }
        assert result["colors"][4]["name"] == "accent1"
        assert result["colors"][11]["name"] == "followed_hyperlink"

    def test_a_deck_with_no_designs_still_reports_one(self):
        """`designs` counts zero here, and a zero would read as no theme."""
        from ppt_mac.themes import _get_theme_colors_impl

        with _fake_deck():
            assert _get_theme_colors_impl()["design_count"] == 1

    def test_a_windows_integer_lands_as_the_right_three_numbers(self):
        from utils.color import hex_to_int

        from ppt_mac.themes import _set_theme_colors_impl

        with _fake_deck() as deck:
            result = _set_theme_colors_impl({5: hex_to_int("#2B579A")})

        assert deck.theme_colors[5].RGB.writes == [[43, 87, 154]]
        assert result["changed"] == [{"name": "accent1", "color_hex": "#2B579A"}]
        assert result["changed_count"] == 1
        assert "warnings" not in result

    def test_a_colour_powerpoint_did_not_keep_is_reported(self):
        """Read back, never echoed, or a rejected colour looks like a success."""
        from utils.color import hex_to_int

        from ppt_mac.themes import _set_theme_colors_impl

        with _fake_deck() as deck:
            deck.theme_colors[5].RGB.set = lambda value: None
            result = _set_theme_colors_impl({5: hex_to_int("#2B579A")})

        assert result["changed_count"] == 0
        assert "accent1" in result["warnings"][0]
        assert "#156082" in result["warnings"][0]


@macos_only
class TestHeadersAndFooters:
    """Per slide, because there is no deck wide object here."""

    def test_the_footer_is_made_visible_before_its_text_is_written(self):
        from ppt_mac.themes import _set_headers_footers_impl

        with _fake_deck(slide_count=1) as deck:
            _set_headers_footers_impl("Draft", None, None, None, None, None)

        footer = deck.slides[1].headers_and_footers.footer
        assert footer.visible.writes == [True]
        assert footer.header_footer_text.writes == ["Draft"]

    def test_a_fixed_date_turns_the_format_off_first(self):
        from ppt_mac.themes import _set_headers_footers_impl

        with _fake_deck(slide_count=1) as deck:
            _set_headers_footers_impl(None, None, None, None, None, "2026-09-04")

        date = deck.slides[1].headers_and_footers.date_and_time
        assert date.use_date_format.writes == [False]
        assert date.header_footer_text.writes == ["2026-09-04"]

    def test_a_layout_with_no_footer_placeholder_is_counted_not_swallowed(self):
        from appscript.reference import CommandError

        from ppt_mac.themes import _set_headers_footers_impl

        refuses = CommandError(None, None, RuntimeError("-1728"), None)
        with _fake_deck(slide_count=3, slide_raises=refuses):
            result = _set_headers_footers_impl("Draft", None, None, None, None, None)

        assert result["slides_updated"] == 0
        assert "3 of 3 slides" in result["warnings"][0]

    def test_an_unknown_date_format_names_the_argument(self):
        from ppt_mac.themes import _date_format_keyword

        with pytest.raises(ValueError, match="date format"):
            _date_format_keyword(9999)


@macos_only
class TestBatchCarriesWhatAnOperationSaid:
    """A step that worked and warned has to still say it warned."""

    def test_warnings_and_partial_travel_onto_the_step(self):
        from ppt_mac.batch_apply import _carried_over

        assert _carried_over({
            "success": True,
            "warnings": ["transparency was not applied"],
            "partial": True,
            "unsupported": ["first_line_indent"],
            "note": "cleared by shape",
        }) == {
            "warnings": ["transparency was not applied"],
            "partial": True,
            "unsupported": ["first_line_indent"],
            "note": "cleared by shape",
        }

    def test_a_plain_success_carries_nothing(self):
        from ppt_mac.batch_apply import _carried_over

        assert _carried_over({"success": True, "shape_name": "Title"}) == {}
        assert _carried_over({"success": True, "warnings": []}) == {}
        assert _carried_over("not a dict") == {}
