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


@contextmanager
def _refusing_deck(module):
    """A PowerPoint that answers nothing, so only `make` can be observed.

    The two refusals below are the only thing between a caller and losing
    every open deck, so what they are tested for is that `make` is never
    reached, not that the code stops at any particular earlier line.
    """
    app = mock.Mock()
    app.make.side_effect = AssertionError("make must not be reached")
    shape = mock.Mock()
    shape.has_table.return_value = True
    with mock.patch.object(module.ppt, "_get_app_impl", return_value=app), \
            mock.patch.object(module.ppt, "_get_pres_impl", return_value=mock.MagicMock()), \
            mock.patch.object(module, "goto_slide", lambda *a, **kw: None), \
            mock.patch.object(module, "_get_table_shape", return_value=shape), \
            mock.patch.object(module, "_dimensions", return_value=(3, 3)):
        yield app


@macos_only
class TestTheTwoRefusalsThatProtectTheDeck:
    """Inserting a table column at a position kills PowerPoint. Nothing else does.

    These two are the only thing between a caller and losing every open deck,
    so they are pinned the way the effects collection is, with a fake that
    raises if the code reaches for `make` at all.
    """

    def test_a_row_at_a_position_refuses_and_never_reaches_make(self):
        from ppt_mac import tables

        with _refusing_deck(tables) as app:
            result = tables._add_table_row_impl(1, 1, 2, None)

        assert result["error"] == (
            "ppt_add_table_row cannot insert at a position on macOS"
        )
        assert "-1708" in result["reason"]
        app.make.assert_not_called()

    def test_a_column_at_a_position_refuses_and_never_reaches_make(self):
        from ppt_mac import tables

        with _refusing_deck(tables) as app:
            result = tables._add_table_column_impl(1, 1, 2, None)

        assert result["error"] == (
            "ppt_add_table_column cannot insert at a position on macOS"
        )
        assert "kills PowerPoint" in result["reason"]
        app.make.assert_not_called()

    def test_appending_is_not_refused(self):
        """The refusal is about `position`, not about the tool."""
        import inspect

        from ppt_mac import tables

        for impl in (tables._add_table_row_impl, tables._add_table_column_impl):
            source = inspect.getsource(impl)
            assert "if position is not None:" in source
            assert source.index("if position is not None:") < source.index("make")

    def test_neither_refusal_moves_the_view_first(self):
        """Refused for its arguments, so the user's view stays where it was."""
        import inspect

        from ppt_mac import tables

        for impl in (tables._add_table_row_impl, tables._add_table_column_impl):
            source = inspect.getsource(impl)
            assert source.index("if position is not None:") < source.index(
                "goto_slide("
            ), impl.__name__


@macos_only
class TestSections:
    """Two things about sections that only running them showed."""

    def test_adding_one_checks_the_name_rather_than_the_count(self):
        """A section inserted mid deck brings a second one with it.

        The slides in front of it need a section too, so a deck with none
        goes straight to two. Counting one more would refuse a success.
        """
        import inspect

        from ppt_mac import sections

        source = inspect.getsource(sections._add_section_impl)
        assert "if after != before + 1" not in source
        assert "_locate_section" in source

    def test_deleting_a_section_that_has_another_after_it_is_refused(self):
        from appscript.reference import CommandError

        from ppt_mac import sections

        deck = mock.MagicMock()
        sp = deck.section_properties
        sp.delete_section.side_effect = CommandError(
            None, None, RuntimeError("-50"), None
        )
        with mock.patch.object(sections.ppt, "_get_app_impl", return_value=mock.Mock()), \
                mock.patch.object(sections.ppt, "_get_pres_impl", return_value=deck), \
                mock.patch.object(sections, "_count_sections", return_value=3), \
                mock.patch.object(sections, "_name_of", return_value="はじめに"), \
                mock.patch.object(sections, "error_number", return_value=-50), \
                mock.patch.object(sections, "goto_slide", lambda *a, **kw: None):
            result = sections._manage_section_impl(1, "delete", None, None)

        assert result["error"] == (
            "ppt_manage_section cannot delete a section that has another "
            "after it on macOS"
        )
        assert "はじめに" in result["reason"]
        assert "section 2" in result["reason"]

    def test_deleting_the_last_section_is_not_refused(self):
        """Deleting from the back works, and so does deleting the only one."""
        from appscript.reference import CommandError

        from ppt_mac import sections

        deck = mock.MagicMock()
        deck.section_properties.delete_section.side_effect = CommandError(
            None, None, RuntimeError("-50"), None
        )
        with mock.patch.object(sections.ppt, "_get_app_impl", return_value=mock.Mock()), \
                mock.patch.object(sections.ppt, "_get_pres_impl", return_value=deck), \
                mock.patch.object(sections, "_count_sections", return_value=3), \
                mock.patch.object(sections, "_name_of", return_value="第二部"), \
                mock.patch.object(sections, "error_number", return_value=-50), \
                mock.patch.object(sections, "goto_slide", lambda *a, **kw: None):
            with pytest.raises(CommandError):
                sections._manage_section_impl(3, "delete", None, None)


@macos_only
class TestTheDocumentedRefusalsMatchTheCode:
    """MACOS_PORT lists every tool that refuses. A test keeps the list true.

    The list is what a user reads before deciding whether this is usable on a
    Mac, so it going stale would be worse than not having it.
    """

    @staticmethod
    def _always_refusing():
        import ast
        import pathlib

        found = set()
        for path in sorted(pathlib.Path("src/ppt_mac").glob("*.py")):
            text = path.read_text()
            for node in ast.parse(text).body:
                if not isinstance(node, ast.FunctionDef):
                    continue
                if not node.name.endswith("_impl"):
                    continue
                returns = [
                    n for n in ast.walk(node)
                    if isinstance(n, ast.Return) and n.value is not None
                ]
                if not returns:
                    continue
                sources = [ast.get_source_segment(text, r) or "" for r in returns]
                # `_refuse_for_shape` and `_refuse_for_node_tool` build the
                # same payload for a whole module's worth of tools, so a
                # refusal is any return that goes through one of the three.
                if all(
                    "_refusal" in s or "_refuse_" in s or "unsupported" in s
                    for s in sources
                ):
                    found.add("ppt_" + node.name[1:-5])
        return found

    def test_the_list_names_exactly_the_tools_that_always_refuse(self):
        import pathlib
        import re

        doc = pathlib.Path("MACOS_PORT.md").read_text()
        section = doc[doc.index("### 6.1 Every tool that refuses"):]
        section = section[:section.index("A further")]
        documented = set(re.findall(r"`(ppt_[a-z_]+)`", section))

        assert documented == self._always_refusing()

    def test_the_headline_count_is_the_real_one(self):
        import ast
        import pathlib
        import re

        doc = pathlib.Path("MACOS_PORT.md").read_text()
        stated = re.search(
            r"\*\*(\d+) tools\. (\d+) do the job\. (\d+) always refuse\.\*\*", doc
        )
        assert stated, "the headline count is not in MACOS_PORT any more"
        total = sum(
            1
            for path in pathlib.Path("src/ppt_mac").glob("*.py")
            for node in ast.parse(path.read_text()).body
            if isinstance(node, ast.FunctionDef) and node.name.endswith("_impl")
        )
        refusing = len(self._always_refusing())
        assert int(stated.group(1)) == total
        assert int(stated.group(3)) == refusing
        assert int(stated.group(2)) == total - refusing


@macos_only
class TestTheAccentsPresentationInfoReports:
    """The server's own instructions tell a caller to build a deck from the
    accents `ppt_get_presentation_info` reports. Every one of them came back
    null on decks whose palette reads perfectly well elsewhere, because the
    colours were read out of the materialised collection instead of being asked
    for one at a time."""

    def test_the_accents_are_asked_for_by_position(self):
        import inspect

        from ppt_mac import presentation

        source = inspect.getsource(presentation._get_presentation_info_impl)
        accents = source[source.index("accent_colors = {"):]
        assert "scheme.theme_colors[position]" in accents
        assert "elements(scheme.theme_colors)" not in accents

    def test_all_six_accents_are_reported_even_when_none_can_be_read(self):
        """A deck with no palette answers null six times, not an empty dict."""
        import inspect

        from ppt_mac import presentation

        source = inspect.getsource(presentation._get_presentation_info_impl)
        for name in ("accent1", "accent2", "accent3", "accent4", "accent5", "accent6"):
            assert f'"{name}"' in source
        assert "accent_colors[key] = None" in source


@macos_only
class TestTheFontAdviceSuitsThisPlatform:
    """Naming a font PowerPoint cannot find does not fail, it substitutes.

    So font advice that is wrong for the platform does not produce an error a
    caller can act on. It produces a deck set in something nobody chose, and
    the only sign is the look of it. The advice has to be right where it is
    first read, which is why it is no longer corrected further down.
    """

    @staticmethod
    def _instructions():
        import src.server as server

        return server.mcp.instructions

    def test_the_windows_fonts_are_not_recommended_here(self):
        advice = self._instructions()
        headline = advice[advice.index("Preferred fonts:"):]
        headline = headline[:headline.index("\n")]
        # The first sentence is the recommendation. What follows it only
        # says which Windows fonts are absent, so Segoe UI belongs there.
        recommendation = headline.split(". ")[0]
        assert "Hiragino Sans" in recommendation
        assert "Segoe UI" not in recommendation
        assert "BIZ UDP" not in recommendation

    def test_the_advice_is_not_contradicted_later(self):
        advice = self._instructions()
        assert advice.count("Preferred fonts:") == 1


class TestEveryMacTestSaysItIsAMacTest:
    """Not a macOS test itself. It runs everywhere, and that is the point.

    CI runs the suite on Windows, where appscript is not installed, so a test
    class that reaches into `ppt_mac` without `@macos_only` fails there and
    nowhere else. That has now happened three times, each time as a surprise
    from a change that looked unrelated: `TestShapesThatWalkOffTheSlide` when
    it was written, and `TestGroupItems` when `ppt_mac/groups.py` stopped
    refusing and started importing appscript to do the work.

    A class is exempt when it never mentions the macOS side at all, which is
    how the handful of cross-platform checks living in these files stay
    running on both.
    """

    @staticmethod
    def _offenders():
        import ast
        import pathlib

        found = []
        for path in sorted(pathlib.Path("tests").glob("test_mac_*.py")):
            text = path.read_text()
            for node in ast.parse(text).body:
                if not isinstance(node, ast.ClassDef):
                    continue
                if not node.name.startswith("Test"):
                    continue
                if node.name == "TestEveryMacTestSaysItIsAMacTest":
                    # This class names the thing it looks for, in prose, and
                    # would otherwise be its own only finding.
                    continue
                if any(
                    isinstance(d, ast.Name) and d.id == "macos_only"
                    for d in node.decorator_list
                ):
                    continue
                body = ast.get_source_segment(text, node) or ""
                if "ppt_mac" in body or "mac_ae" in body:
                    found.append(f"{path}:{node.lineno} {node.name}")
        return found

    def test_no_class_reaches_into_ppt_mac_unguarded(self):
        offenders = self._offenders()
        assert not offenders, (
            "these will fail on Windows, where appscript is not installed. "
            "Add @macos_only:\n  " + "\n  ".join(offenders)
        )


class TestBothReadmesCountTheSame:
    """Runs everywhere. The number drifts, and only in one language.

    MACOS_PORT's headline is pinned against the code. The READMEs repeat it in
    prose, in English and Japanese, and the Japanese one was left saying 21
    when the English one had moved to 12 — through three rounds of tools
    leaving the refusal list. Nobody reads both.
    """

    @staticmethod
    def _documented_refusals():
        import pathlib
        import re

        doc = pathlib.Path("MACOS_PORT.md").read_text()
        stated = re.search(
            r"\*\*(\d+) tools\. (\d+) do the job\. (\d+) always refuse\.\*\*", doc
        )
        assert stated, "the headline count is not in MACOS_PORT any more"
        return int(stated.group(3))

    def test_the_english_readme_agrees(self):
        import pathlib
        import re

        refusing = self._documented_refusals()
        text = pathlib.Path("README.md").read_text()
        said = re.search(r"\*\*(\d+) tools refuse on macOS", text)
        assert said, "README.md no longer says how many refuse"
        assert int(said.group(1)) == refusing

    def test_the_japanese_readme_agrees(self):
        import pathlib
        import re

        refusing = self._documented_refusals()
        text = pathlib.Path("README_ja.md").read_text()
        said = re.search(r"\*\*macOS で断るツールは (\d+) です", text)
        assert said, "README_ja.md no longer says how many refuse"
        assert int(said.group(1)) == refusing
