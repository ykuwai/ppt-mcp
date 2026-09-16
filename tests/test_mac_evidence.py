"""Tests for the two ways a macOS tool can report something that did not happen.

Pure unit tests over a stand-in object graph. Nothing here launches PowerPoint
or connects to it; the live behaviour is covered by MACOS_PORT.md and by
running the server. appscript only installs on macOS, so the whole file is
skipped elsewhere.

The modules gathered here have nothing else in common. What they share is the
rule MACOS_PORT section 5 leaves behind, that no mutation is trusted because it
did not raise, and the two ways of breaking it that a review of every tool
turned up.

The first is a tool that writes part of what it was asked for and then raises
on an argument it had not looked at yet, which leaves the caller with an error
and a half edited deck. So every name a caller can misspell is checked before
the first write, and the tests for it make any approach to PowerPoint an
assertion failure.

The second is a tool that reports what it was handed rather than what
PowerPoint is holding. Where a read route exists it is used and the answer is
the measurement; where none exists the answer says so in ``warnings`` rather
than letting an echo pass for evidence.
"""

import os
import sys

import pytest

sys.path.insert(0, "src")

macos_only = pytest.mark.skipif(
    sys.platform != "darwin", reason="the Apple Event backend needs appscript"
)


# ---------------------------------------------------------------------------
# Family one. Every name is checked before PowerPoint is touched at all.
# ---------------------------------------------------------------------------
@macos_only
class TestArgumentsAreCheckedFirst:
    """A misspelled argument costs nothing, not a stray shape and not the view."""

    def test_a_bad_vertical_anchor_never_makes_a_text_box(self):
        from ppt_mac.shapes import _add_textbox_impl

        with _no_powerpoint():
            with pytest.raises(ValueError, match="vertical_anchor"):
                _add_textbox_impl(
                    1, 10, 10, 100, 50, "Hello",
                    None, None, None, None, None, None, "centre",
                )

    def test_a_bad_alignment_never_makes_a_text_box(self):
        from ppt_mac.shapes import _add_textbox_impl

        with _no_powerpoint():
            with pytest.raises(ValueError, match="align"):
                _add_textbox_impl(
                    1, 10, 10, 100, 50, "Hello",
                    None, None, None, None, None, "middle", None,
                )

    def test_a_bad_fill_type_never_makes_a_shape(self):
        from ppt_mac.shapes import _add_shape_impl

        with _no_powerpoint():
            with pytest.raises(ValueError, match="fill_type"):
                _add_shape_impl(
                    1, 1, 10, 10, 100, 50, "Hello",
                    None, None, None, None, None, None,
                    None, "checkerboard", None, None, None,
                    None, None, None, None, None,
                )

    def test_a_bad_text_frame_argument_writes_no_margins(self):
        from ppt_mac.text import _set_textframe_impl

        for kwargs, message in (
            ({"orientation": "sideways"}, "orientation"),
            ({"auto_size": "grow"}, "auto_size"),
            ({"vertical_anchor": "centre"}, "vertical_anchor"),
        ):
            arguments = {
                "auto_size": None, "word_wrap": True,
                "margin_left": 5, "margin_right": 5,
                "margin_top": 5, "margin_bottom": 5,
                "orientation": None, "vertical_anchor": None,
            }
            arguments.update(kwargs)
            with _no_powerpoint():
                with pytest.raises(ValueError, match=message):
                    _set_textframe_impl(1, "Box", **arguments)

    def test_a_bad_cell_alignment_writes_no_text(self):
        from ppt_mac.tables import _set_table_cell_impl

        with _no_powerpoint():
            with pytest.raises(ValueError, match="alignment"):
                _set_table_cell_impl(
                    1, "Table", 1, 1, "Hello",
                    None, None, None, None, None, None,
                    None, "centred", None,
                )

    def test_a_bad_cell_vertical_alignment_writes_no_text(self):
        from ppt_mac.tables import _set_table_cell_impl

        with _no_powerpoint():
            with pytest.raises(ValueError, match="vertical_alignment"):
                _set_table_cell_impl(
                    1, "Table", 1, 1, "Hello",
                    None, None, None, None, None, None,
                    None, None, "centre",
                )

    def test_a_bad_dash_style_never_moves_the_view(self):
        from ppt_mac.formatting import _set_line_impl

        with _no_powerpoint():
            with pytest.raises(ValueError, match="dash_style"):
                _set_line_impl(1, "Box", "#FF0000", 2.0, "wiggly", None, None)

    def test_a_bad_fill_type_never_moves_the_view(self):
        from ppt_mac.formatting import _set_fill_impl

        with _no_powerpoint():
            with pytest.raises(ValueError, match="fill_type"):
                _set_fill_impl(1, "Box", "checkerboard",
                               "#FF0000", None, None, None, None)


# ---------------------------------------------------------------------------
# Family two. The answer is the measurement, or it says it is not one.
# ---------------------------------------------------------------------------
@macos_only
class TestFillIsMeasured:
    """The busiest styling tool in the server used to echo its own argument."""

    def test_the_fill_type_comes_from_the_shape(self):
        from appscript import k

        from ppt_mac.formatting import _set_fill_impl

        with _fake_deck(["Box"]) as deck:
            result = _set_fill_impl(1, "Box", "solid", "#2B579A",
                                    None, None, None, None)

        assert deck.shape("Box").fill_format.fill_format_type() == k.fill_solid
        assert result["fill_type"] == "solid"
        assert result["color_hex"] == "#2B579A"
        assert "warnings" not in result

    def test_a_fill_powerpoint_declined_is_reported_rather_than_claimed(self):
        from appscript import k

        from ppt_mac.formatting import _set_fill_impl

        with _fake_deck(["Box"]) as deck:
            # PowerPoint took the command, said nothing, and left the shape
            # wearing the gradient it already had.
            deck.shape("Box").fill_format.fill_format_type.clamp_to = k.fill_gradient
            result = _set_fill_impl(1, "Box", "solid", "#2B579A",
                                    None, None, None, None)

        assert result["fill_type"] == "gradient"
        assert "did not take it" in result["warnings"][0]

    def test_no_fill_is_read_off_the_visibility_rather_than_the_type(self):
        from ppt_mac.formatting import _set_fill_impl

        with _fake_deck(["Box"]) as deck:
            result = _set_fill_impl(1, "Box", "none", None,
                                    None, None, None, None)

        assert deck.shape("Box").fill_format.visible() is False
        assert result["fill_type"] == "none"
        assert result["fill_visible"] is False

    def test_a_solid_fill_after_no_fill_makes_the_shape_visible_again(self):
        """"none" is a visibility here, so a later fill has to turn it back on."""
        from ppt_mac.formatting import _set_fill_impl

        with _fake_deck(["Box"]) as deck:
            _set_fill_impl(1, "Box", "none", None, None, None, None, None)
            result = _set_fill_impl(1, "Box", "solid", "#2B579A",
                                    None, None, None, None)

        assert deck.shape("Box").fill_format.visible() is True
        assert result["fill_type"] == "solid"
        assert "warnings" not in result

    def test_a_fill_that_will_not_answer_says_the_type_is_not_a_measurement(self):
        from ppt_mac.formatting import _set_fill_impl

        with _fake_deck(["Box"]) as deck:
            deck.shape("Box").fill_format.visible.read_raises = True
            deck.shape("Box").fill_format.fill_format_type.read_raises = True
            result = _set_fill_impl(1, "Box", "solid", None,
                                    None, None, None, None)

        assert result["fill_type"] == "solid"
        assert "rather than what was measured" in result["warnings"][0]


@macos_only
class TestLineVisibility:
    """Hiding a border is a stand-in here, and both tools now say so."""

    def test_hiding_a_border_says_what_it_actually_did(self):
        from ppt_mac.formatting import _set_line_impl

        with _fake_deck(["Box"]) as deck:
            result = _set_line_impl(1, "Box", None, None, None, False, None)

        assert deck.shape("Box").line_format.line_weight() == 0.0
        assert deck.shape("Box").line_format.transparency() == 1.0
        assert result["status"] == "success"
        assert "no visible property" in result["warnings"][0]
        # The part worth the words: the flag PowerPoint reads is untouched, so
        # setting a colour later undoes this without the caller asking.
        assert "brings the border back" in result["warnings"][0]
        # Named as the caller passes it, not as the plumbing calls it.
        assert "line_visible" in result["warnings"][0]

    def test_it_says_which_way_round_the_substitution_went(self):
        """One sentence for hiding and another for showing.

        A reader was left working out whether transparency 0.0 meant opaque or
        see through, because both directions got the same sentence.
        """
        from ppt_mac.shapes import _LINE_VISIBILITY_WARNING

        hiding = _LINE_VISIBILITY_WARNING[False]
        showing = _LINE_VISIBILITY_WARNING[True]

        assert "weight 0 and full transparency" in hiding
        assert "The border is drawn." in showing
        assert hiding != showing

    def test_a_shape_with_no_border_is_not_a_success(self):
        """A picture answers with an error, and that used to be swallowed."""
        from ppt_mac.formatting import _set_line_impl

        with _fake_deck(["Picture"]) as deck:
            deck.shape("Picture").line_format.line_weight.raises = True
            deck.shape("Picture").line_format.transparency.raises = True
            result = _set_line_impl(1, "Picture", None, None, None, False, None)

        assert "status" not in result
        assert result["error"] == "ppt_set_line could not change this shape's border"
        assert "no line format" in result["reason"]

    def test_a_border_that_could_not_be_hidden_still_reports_the_colour(self):
        from ppt_mac.formatting import _set_line_impl

        with _fake_deck(["Picture"]) as deck:
            line = deck.shape("Picture").line_format
            line.transparency.raises = True
            result = _set_line_impl(1, "Picture", "#FF0000", None, None, True, None)

        assert result["status"] == "success"
        assert result["color_hex"] == "#FF0000"
        assert any("no line format" in warning for warning in result["warnings"])

    def test_a_clamped_weight_is_reported_as_powerpoint_holds_it(self):
        from ppt_mac.formatting import _set_line_impl

        with _fake_deck(["Box"]) as deck:
            deck.shape("Box").line_format.line_weight.clamp_to = 0.25
            result = _set_line_impl(1, "Box", None, 0.1, None, None, None)

        assert result["weight"] == 0.25


@macos_only
class TestShadow:
    """Nothing reads a shadow back, so nothing here pretends to have measured one."""

    def test_the_answer_says_it_is_not_a_measurement(self):
        from ppt_mac.formatting import _set_shadow_impl

        with _fake_deck(["Box"]):
            result = _set_shadow_impl(1, "Box", True, 5, 3, 3, "#000000", 0.5)

        assert result["shadow_visible"] is True
        assert "rather than a measurement" in result["warnings"][0]

    def test_styling_a_hidden_shadow_says_the_styling_went_nowhere(self):
        from ppt_mac.formatting import _set_shadow_impl

        with _fake_deck(["Box"]):
            result = _set_shadow_impl(1, "Box", False, 5, None, None, None, None)

        assert any("blur" in warning for warning in result["warnings"])


@macos_only
class TestTextAtCreation:
    """The creation path holds to the same standard ppt_set_text does."""

    def test_a_shape_whose_text_did_not_land_is_not_a_success(self):
        from ppt_mac.shapes import _add_shape_impl

        with _fake_deck([]) as deck:
            deck.text_writes = False
            with pytest.raises(RuntimeError, match="came back empty"):
                _add_shape_impl(
                    1, 1, 10, 10, 100, 50, "Hello",
                    None, None, None, None, None, None,
                    None, None, None, None, None,
                    None, None, None, None, None,
                )

    def test_a_text_box_whose_text_did_not_land_is_not_a_success(self):
        from ppt_mac.shapes import _add_textbox_impl

        with _fake_deck([]) as deck:
            deck.text_writes = False
            deck.make_is_a_text_box = True
            with pytest.raises(RuntimeError, match="came back empty"):
                _add_textbox_impl(
                    1, 10, 10, 100, 50, "Hello",
                    None, None, None, None, None, None, None,
                )

    def test_text_that_lands_is_not_complained_about(self):
        from ppt_mac.shapes import _add_shape_impl

        with _fake_deck([]) as deck:
            result = _add_shape_impl(
                1, 1, 10, 10, 100, 50, "Hello",
                None, None, None, None, None, None,
                None, None, None, None, None,
                None, None, None, None, None,
            )

        assert result["success"] is True
        assert deck.slide_shapes[-1].text_frame.text_range.content() == "Hello"


# ---------------------------------------------------------------------------
# Tables
# ---------------------------------------------------------------------------
@macos_only
class TestTableData:
    """Data that does not fit used to be dropped without a word."""

    def test_rows_past_the_end_of_the_table_are_reported(self):
        from ppt_mac.tables import _set_table_data_impl

        with _fake_deck(["Table"], table=(2, 2)):
            result = _set_table_data_impl(
                1, "Table", [["a", "b"], ["c", "d"], ["e", "f"]], 1, 1, False
            )

        assert result["cells_set"] == 4
        assert result["rows_written"] == 2
        assert "1 row(s) of data had nowhere to go" in result["warnings"][0]

    def test_columns_past_the_end_of_the_table_are_reported(self):
        from ppt_mac.tables import _set_table_data_impl

        with _fake_deck(["Table"], table=(2, 2)):
            result = _set_table_data_impl(
                1, "Table", [["a", "b", "c"], ["d", "e", "f"]], 1, 1, False
            )

        assert result["cells_set"] == 4
        assert any("fell past" in warning for warning in result["warnings"])

    def test_data_that_fits_says_nothing(self):
        from ppt_mac.tables import _set_table_data_impl

        with _fake_deck(["Table"], table=(2, 2)) as deck:
            result = _set_table_data_impl(
                1, "Table", [["a", "b"], ["c", "d"]], 1, 1, False
            )

        assert "warnings" not in result
        assert deck.table_text(2, 2) == "d"


@macos_only
class TestTableBorders:
    """The count used to rise for every cell the loop walked past."""

    def test_nothing_asked_for_is_nothing_updated(self):
        from ppt_mac.tables import _set_table_borders_impl

        with _fake_deck(["Table"], table=(2, 2)):
            result = _set_table_borders_impl(
                1, "Table", 1, 1, None, None, ["top"], None, None, None, None
            )

        assert result["cells_updated"] == 0
        assert result["borders_written"] == 0
        assert "No border property was given" in result["warnings"][0]

    def test_the_count_is_of_borders_that_were_written(self):
        from ppt_mac.tables import _set_table_borders_impl

        with _fake_deck(["Table"], table=(2, 2)) as deck:
            result = _set_table_borders_impl(
                1, "Table", 1, 1, None, None, ["top", "bottom"],
                None, "#FF0000", None, None,
            )

        assert result["cells_updated"] == 4
        assert result["borders_written"] == 8
        assert len(deck.border_writes) == 8


@macos_only
class TestSplitAndMerge:
    """One can be measured off the table and the other cannot."""

    def test_a_split_reports_the_table_it_left_behind(self):
        from ppt_mac.tables import _split_table_cells_impl

        with _fake_deck(["Table"], table=(3, 3)) as deck:
            deck.split_adds_rows = 1
            result = _split_table_cells_impl(1, "Table", 1, 1, 2, 1)

        assert result["table_rows"] == 4
        assert result["table_columns"] == 3
        assert "warnings" not in result

    def test_a_split_that_did_not_land_says_what_was_expected(self):
        """One cell into n rows adds n minus 1 to the table. Measured live."""
        from ppt_mac.tables import _split_table_cells_impl

        with _fake_deck(["Table"], table=(3, 3)):
            result = _split_table_cells_impl(1, "Table", 1, 1, 2, 1)

        assert result["table_rows"] == 3
        assert "measures 3 by 3" in result["warnings"][0]
        assert "should have left it 4 by 3" in result["warnings"][0]

    def test_a_merge_says_it_could_not_be_measured(self):
        from ppt_mac.tables import _merge_table_cells_impl

        with _fake_deck(["Table"], table=(3, 3)) as deck:
            result = _merge_table_cells_impl(1, "Table", 1, 1, 1, 2)

        assert deck.merges == [((1, 1), (1, 2))]
        assert "rather than a measurement" in result["warnings"][0]


# ---------------------------------------------------------------------------
# Layout
# ---------------------------------------------------------------------------
@macos_only
class TestAlignAndDistribute:
    """The count used to be the length of the list the caller passed in."""

    def test_the_count_is_of_shapes_that_moved(self):
        from ppt_mac.layout import _align_shapes_impl

        with _fake_deck(["A", "B"]) as deck:
            deck.shape("B").left_position.set(200.0)
            result = _align_shapes_impl(1, ["A", "B"], "left", False)

        assert result["aligned_count"] == 2
        assert result["success"] is True
        assert "warnings" not in result

    def test_a_shape_that_would_not_move_is_named(self):
        from ppt_mac.layout import _align_shapes_impl

        with _fake_deck(["A", "B"]) as deck:
            deck.shape("B").left_position.set(200.0)
            # A placeholder driven by its layout takes the write and stays put.
            deck.shape("B").left_position.clamp_to = 200.0
            result = _align_shapes_impl(1, ["A", "B"], "left", False)

        assert result["aligned_count"] == 1
        assert "B did not move" in result["warnings"][0]

    def test_distributing_counts_the_same_way(self):
        from ppt_mac.layout import _distribute_shapes_impl

        with _fake_deck(["A", "B", "C"]) as deck:
            deck.shape("B").left_position.set(100.0)
            deck.shape("C").left_position.set(300.0)
            # Distributing leaves the outermost two where they are, so the
            # middle one is the only one with anywhere to go.
            deck.shape("B").left_position.clamp_to = 100.0
            result = _distribute_shapes_impl(1, ["A", "B", "C"], "horizontal", False)

        assert result["distributed_count"] == 2
        assert "B did not move" in result["warnings"][0]


@macos_only
class TestSlideBackground:
    """One call can name twenty slides, and it used to report all twenty."""

    def test_only_the_slides_that_took_it_are_counted(self):
        from ppt_mac.formatting import _set_fill_impl  # noqa: F401 - shared fakes
        from ppt_mac.layout import _set_slide_background_impl

        with _fake_deck([], slides=3) as deck:
            deck.background_refuses = {2}
            result = _set_slide_background_impl(
                1, "solid", "#2B579A", None, None, None, None, None,
                slide_indices=[1, 2, 3],
            )

        assert result["slide_indices"] == [1, 3]
        assert result["success"] is True
        assert "slide 2" in result["warnings"][0]

    def test_a_call_that_painted_nothing_is_not_a_success(self):
        from ppt_mac.layout import _set_slide_background_impl

        with _fake_deck([], slides=1) as deck:
            deck.background_refuses = {1}
            result = _set_slide_background_impl(
                1, "solid", "#2B579A", None, None, None, None, None,
            )

        assert result["success"] is False
        assert result["slide_indices"] == []
        # The back compatible key still names the slide that was asked for.
        assert result["slide_index"] == 1

    def test_a_clamped_transparency_is_not_a_background_that_did_not_take(self):
        from ppt_mac.layout import _set_slide_background_impl

        with _fake_deck([], slides=1) as deck:
            deck.slides[0].background.fill_format.transparency.clamp_to = 0.0
            result = _set_slide_background_impl(
                1, "solid", "#2B579A", None, None, None, None, 0.5,
            )

        assert result["success"] is True
        assert result["slide_indices"] == [1]
        assert "keeps the value inside its own range" in result["warnings"][0]

    def test_following_the_master_is_read_back_too(self):
        from ppt_mac.layout import _set_slide_background_impl

        with _fake_deck([], slides=1) as deck:
            result = _set_slide_background_impl(
                1, "master", None, None, None, None, None, None,
            )

        assert deck.slides[0].follow_master_background() is True
        assert result["slide_indices"] == [1]


# ---------------------------------------------------------------------------
# Window and selection
# ---------------------------------------------------------------------------
@macos_only
class TestWindowSelection:
    """An empty list and an unanswerable question are different answers."""

    def test_the_selected_shapes_are_reached_positionally(self):
        from ppt_mac.app import _get_active_window_info_impl

        with _fake_window(["Box", "Circle"]) as window:
            result = _get_active_window_info_impl()

        assert window.selection.shape_range.materialised == 0
        assert result["selection_type"] == "shapes"
        assert [s["name"] for s in result["selected_shapes"]] == ["Box", "Circle"]

    def test_a_selection_powerpoint_will_not_describe_is_not_an_empty_one(self):
        from ppt_mac.app import _get_active_window_info_impl

        with _fake_window(["Box"], shapes_refuse=True):
            result = _get_active_window_info_impl()

        assert result["selection_type"] == "shapes"
        assert result["selected_shapes"] == []
        assert "would not say which ones" in result["warnings"][0]

    def test_the_banned_route_is_not_used(self):
        """`elements(...shapes)` is the addressing bug, so it is not in here."""
        import inspect

        from ppt_mac import app as mac_app

        source = inspect.getsource(mac_app._get_active_window_info_impl)
        assert "elements(selection" not in source
        assert "shapes_of(selection.shape_range)" in source


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------
@macos_only
class TestExportEvidence:
    """A file that exists and is empty is the failure this catches."""

    def test_an_empty_pdf_is_refused(self, tmp_path, monkeypatch):
        from ppt_mac import export as mac_export

        destination = tmp_path / "deck.pdf"

        def _move(source, target):
            os.remove(source)
            open(target, "wb").close()

        with _fake_deck([], slides=2):
            monkeypatch.setattr(
                mac_export, "_staging_path",
                lambda suffix: str(tmp_path / f"staged{suffix}"),
            )
            monkeypatch.setattr(
                mac_export, "_save_pdf",
                lambda pres, path: open(path, "wb").write(b"%PDF-1.4"),
            )
            monkeypatch.setattr(mac_export.shutil, "move", _move)
            with pytest.raises(RuntimeError, match="wrote no PDF"):
                mac_export._export_pdf_impl(str(destination), None, None)

    def test_an_empty_slide_image_is_refused(self, tmp_path, monkeypatch):
        from ppt_mac import export as mac_export

        staged_png = tmp_path / "staged.png"
        staged_png.write_bytes(b"")

        with _fake_deck([], slides=1):
            monkeypatch.setattr(
                mac_export, "_staging_path",
                lambda suffix: str(tmp_path / f"staged_pdf{suffix}"),
            )
            monkeypatch.setattr(
                mac_export, "_save_pdf",
                lambda pres, path: open(path, "wb").write(b"%PDF-1.4"),
            )
            monkeypatch.setattr(
                mac_export, "_render_pdf_pages",
                lambda pdf, pages, width=None, height=None,
                uti="public.png", suffix=".png": [
                    (1, str(staged_png), 100, 100)
                ],
            )
            with pytest.raises(RuntimeError, match="arrived"):
                mac_export._export_images_impl(
                    str(tmp_path / "out"), "png", 1, None, None, None,
                    None, None, None,
                )


# ---------------------------------------------------------------------------
# A deck made of stand-ins, so none of this needs PowerPoint. Only the parts
# the tools touch are modelled, and the collections behave the way PowerPoint's
# do, including answering -1728 for an empty one.
# ---------------------------------------------------------------------------
class _Prop:
    """A property that remembers what was written to it.

    ``clamp_to`` is PowerPoint keeping a value inside its own range or
    declining the write outright, which it does without raising. ``raises`` is
    a property the shape does not really have, which answers with an error to
    both a read and a write, and ``read_raises`` is one that takes a write and
    will not say what it holds.
    """

    def __init__(self, value=None):
        self.value = value
        self.writes: list = []
        self.clamp_to = None
        self.raises = False
        self.read_raises = False

    def __call__(self):
        if self.raises or self.read_raises:
            raise _command_error(-1728)
        return self.value

    def get(self):
        return self()

    def set(self, value):
        if self.raises:
            raise _command_error(-1728)
        self.writes.append(value)
        self.value = self.clamp_to if self.clamp_to is not None else value


def _bag(**properties):
    """A stand-in for one of PowerPoint's little property-only classes."""
    holder = type("Bag", (), {})()
    for name, value in properties.items():
        setattr(holder, name, _Prop(value))
    return holder


class _FakeList:
    def __init__(self, items):
        self._items = list(items)

    def get(self):
        if not self._items:
            # PowerPoint raises rather than answering an empty list.
            raise _command_error(-1728)
        return list(self._items)

    def __getitem__(self, index):
        if index < 1 or index > len(self._items):
            raise _command_error(-1728)
        return self._items[index - 1]


class _FakeShapes(_FakeList):
    """A shapes collection, which can also be asked for every name at once."""

    @property
    def name(self):
        return _FakeList([shape.name() for shape in self._items])


class _FakeFont:
    def __init__(self):
        self.font_name = _Prop("Calibri")
        self.east_asian_name = _Prop("Calibri")
        self.font_size = _Prop(None)
        self.bold = _Prop(None)
        self.italic = _Prop(None)
        self.font_color = _Prop(None)


class _FakeTextRange:
    def __init__(self, deck):
        self._deck = deck
        self.content = _Prop(None)
        self.font = _FakeFont()
        self.paragraph_format = _bag(alignment=None)
        self.indent_level = _Prop(1)


class _FakeTextFrame:
    def __init__(self, deck):
        self.text_range = _FakeTextRange(deck)
        self.vertical_anchor = _Prop(None)
        self.word_wrap = _Prop(True)
        self.margin_left = _Prop(0.0)
        self.margin_right = _Prop(0.0)
        self.margin_top = _Prop(0.0)
        self.margin_bottom = _Prop(0.0)
        self.auto_size = _Prop(None)
        self.text_orientation = _Prop(None)


class _FakeFill:
    """A fill format, which answers what it was last told to be."""

    def __init__(self):
        from appscript import k

        self.visible = _Prop(True)
        self.fore_color = _Prop([0, 0, 0])
        self.back_color = _Prop([0, 0, 0])
        self.transparency = _Prop(0.0)
        self.fill_format_type = _Prop(k.fill_solid)
        self.solid_calls = 0
        self.gradient_calls = 0

    def solid(self):
        from appscript import k

        self.solid_calls += 1
        self.fill_format_type.set(k.fill_solid)

    def two_color_gradient(self, style=None, variant=None):
        from appscript import k

        self.gradient_calls += 1
        self.fill_format_type.set(k.fill_gradient)

    def user_picture(self, picture_file=None):
        from appscript import k

        self.fill_format_type.set(k.fill_picture)


class _FakeShape:
    def __init__(self, deck, name, shape_type=None, has_table=False):
        from appscript import k

        self._deck = deck
        self._name = _Prop(name)
        self._type = shape_type or k.shape_type_auto
        self._has_table = has_table
        self.width = _Prop(120.0)
        self.height = _Prop(60.0)
        self.left_position = _Prop(10.0)
        self.top = _Prop(20.0)
        self.auto_shape_type = _Prop(k.autoshape_rectangle)
        self.fill_format = _FakeFill()
        self.line_format = _bag(
            fore_color=[0, 0, 0], line_weight=1.0, transparency=0.0,
        )
        self.shadow_format = _bag(
            visible=False, blur=0.0, offset_X=0.0, offset_Y=0.0,
            fore_color=[0, 0, 0], transparency=0.0,
        )
        self.text_frame = _FakeTextFrame(deck)
        self.adjustments = _FakeList([_bag(adjustment_value=0.0)])
        self.table_object = _FakeTable(deck) if has_table else None

    def name(self):
        return self._name()

    def shape_type(self):
        return self._type

    def has_text_frame(self):
        return True

    def has_table(self):
        return self._has_table

    def number_of_rows(self):
        return self._deck.table_rows

    def number_of_columns(self):
        return self._deck.table_cols

    def z_order_position(self):
        return self._deck.slide_shapes.index(self) + 1


class _FakeCell:
    def __init__(self, deck, row, col):
        self._deck = deck
        self.row = row
        self.col = col
        self.shape = _FakeShape(deck, f"Cell{row}_{col}")

    def merge(self, merge_with=None):
        self._deck.merges.append(
            ((self.row, self.col), (merge_with.row, merge_with.col))
        )

    def split(self, number_of_rows=None, number_of_columns=None):
        self._deck.table_rows += self._deck.split_adds_rows
        self._deck.splits.append((self.row, self.col,
                                  number_of_rows, number_of_columns))

    def get_border(self, edge=None):
        border = _bag(transparency=0.0, fore_color=[0, 0, 0], line_weight=1.0)
        self._deck.border_writes.append((self.row, self.col, edge))
        return border


class _FakeRow:
    def __init__(self, deck, row):
        self._deck = deck
        self._row = row
        self.height = _Prop(20.0)

    @property
    def cells(self):
        return _FakeList([
            self._deck.cell(self._row, c)
            for c in range(1, self._deck.table_cols + 1)
        ])


class _FakeTable:
    def __init__(self, deck):
        self._deck = deck

    @property
    def rows(self):
        return _FakeList([
            _FakeRow(self._deck, r) for r in range(1, self._deck.table_rows + 1)
        ])

    @property
    def columns(self):
        return _FakeList([
            _bag(width=60.0) for _ in range(1, self._deck.table_cols + 1)
        ])


class _FakeSlide:
    def __init__(self, deck, shapes, index=1):
        self._deck = deck
        self._index = index
        self.slide_shapes = shapes
        self.follow_master_background = _Prop(False)
        self._background = type("Background", (), {})()
        self._background.fill_format = _FakeFill()
        self.end = object()

    @property
    def background(self):
        """A slide that takes the write, raises nothing and changes nothing.

        Read on the way through rather than rigged up front, because a test
        names the refusing slides inside the `with` block.
        """
        from appscript import k

        if self._index in self._deck.background_refuses:
            self._background.fill_format.fill_format_type.clamp_to = (
                k.fill_background
            )
            self._background.fill_format.visible.clamp_to = True
        return self._background

    @property
    def shapes(self):
        return _FakeShapes(self.slide_shapes)


class _FakeDeck:
    """One or more slides, each holding the shapes it was given."""

    def __init__(self, shape_names, slides=1, table=None):
        self.table_rows, self.table_cols = table or (0, 0)
        self.split_adds_rows = 0
        self.text_writes = True
        self.make_is_a_text_box = False
        self.background_refuses: set = set()
        self.merges: list = []
        self.splits: list = []
        self.border_writes: list = []
        self._cells: dict = {}

        self.slide_shapes = [
            _FakeShape(self, name, has_table=(table is not None))
            for name in shape_names
        ]
        self.slides = [_FakeSlide(self, self.slide_shapes, index=1)]
        for index in range(2, slides + 1):
            self.slides.append(_FakeSlide(self, [], index=index))
        # The deck stands in for the application as well, which is what the
        # export tools ask for the open presentations.
        self.presentations = _FakeList([object()])

    # -- the parts the tools reach for -------------------------------------
    def shape(self, name):
        for shape in self.slide_shapes:
            if shape.name() == name:
                return shape
        raise KeyError(name)

    def cell(self, row, col):
        if (row, col) not in self._cells:
            self._cells[(row, col)] = _FakeCell(self, row, col)
        return self._cells[(row, col)]

    def table_text(self, row, col):
        return self.cell(row, col).shape.text_frame.text_range.content()

    def make(self, new=None, at=None, with_properties=None):
        from appscript import k

        shape_type = k.shape_type_text_box if self.make_is_a_text_box else None
        if new == k.text_box:
            shape_type = k.shape_type_text_box
        shape = _FakeShape(self, f"Shape {len(self.slide_shapes) + 1}",
                           shape_type=shape_type)
        if with_properties:
            for key, value in (
                (k.width, "width"), (k.height, "height"),
                (k.left_position, "left_position"), (k.top, "top"),
            ):
                if key in with_properties:
                    getattr(shape, value).set(with_properties[key])
        if not self.text_writes:
            # PowerPoint takes the content write, answers nothing and leaves
            # the frame empty, which is the silent no-op this all exists for.
            shape.text_frame.text_range.content.clamp_to = ""
        self.slide_shapes.append(shape)
        return shape

    @property
    def presentation(self):
        deck = self

        class _Pres:
            slides = _FakeList(deck.slides)
            page_setup = _bag(slide_width=960.0)
            slide_master = _bag(height=540.0)
            document_windows = _FakeList([_bag(caption="deck")])
            presentations = _FakeList([object()])

        return _Pres()


class _fake_deck:  # noqa: N801 - reads as a context manager, not a class
    """Point the wrapper at a fake deck for the length of a `with` block."""

    def __init__(self, shape_names, slides=1, table=None):
        self._deck = _FakeDeck(shape_names, slides=slides, table=table)

    def __enter__(self):
        from backend.mac_ae import ppt

        self._ppt = ppt
        self._app = ppt._get_app_impl
        self._pres = ppt._get_pres_impl
        ppt._get_app_impl = lambda *a, **kw: self._deck
        ppt._get_pres_impl = lambda *a, **kw: self._deck.presentation

        return self._deck

    def __exit__(self, *exc):
        self._ppt._get_app_impl = self._app
        self._ppt._get_pres_impl = self._pres
        return False


class _FakeShapeRange:
    """A selection's shape range, which must never be materialised."""

    def __init__(self, deck, names, refuse=False):
        self._shapes = [_FakeShape(deck, name) for name in names]
        deck.slide_shapes.extend(self._shapes)
        self.materialised = 0
        self._refuse = refuse

    def get(self):
        self.materialised += 1
        raise AssertionError("a selection range must not be materialised")

    @property
    def shapes(self):
        if self._refuse:
            # The subclass addressing problem, which answers -1728 rather than
            # answering with nothing.
            raise _command_error(-1728)
        return _FakeShapes(self._shapes)


class _fake_window:  # noqa: N801 - reads as a context manager, not a class
    """A window with a shape selection in it."""

    def __init__(self, names, shapes_refuse=False):
        from appscript import k

        self._deck = _FakeDeck([])
        self.selection = type("Selection", (), {})()
        self.selection.selection_type = lambda: k.selection_type_shapes
        self.selection.shape_range = _FakeShapeRange(
            self._deck, names, refuse=shapes_refuse
        )
        window = self

        class _Window:
            caption = staticmethod(lambda: "deck")
            view_type = staticmethod(lambda: k.slide_view)
            view = _bag(slide=None)
            selection = window.selection

        self._window = _Window()

    def __enter__(self):
        from backend.mac_ae import ppt

        deck = self._deck
        window = self._window

        class _App:
            document_windows = _FakeList([window])
            active_window = window
            presentations = _FakeList([object()])

        self._ppt = ppt
        self._app = ppt._get_app_impl
        self._pres = ppt._get_pres_impl
        ppt._get_app_impl = lambda *a, **kw: _App()
        ppt._get_pres_impl = lambda *a, **kw: deck.presentation
        return self._window

    def __exit__(self, *exc):
        self._ppt._get_app_impl = self._app
        self._ppt._get_pres_impl = self._pres
        return False


class _no_powerpoint:  # noqa: N801 - reads as a context manager, not a class
    """Make any approach to PowerPoint fail, so a check has to come first."""

    def __enter__(self):
        from backend.mac_ae import ppt

        def _explode(*args, **kwargs):
            raise AssertionError("the argument must be checked before this")

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


@macos_only
class TestTheCopyOutsideTheContainer:
    """PowerPoint can only hold a deck inside its container, so the caller's own
    file is a copy. It used to fall behind on every ordinary save without saying
    so, while the advice was to save at every break."""

    def test_a_save_refreshes_the_copy_it_was_given(self, tmp_path):
        from ppt_mac import presentation

        inside = tmp_path / "container" / "deck.pptx"
        inside.parent.mkdir()
        inside.write_bytes(b"second version")
        outside = tmp_path / "Desktop" / "deck.pptx"
        outside.parent.mkdir()
        outside.write_bytes(b"first version")

        presentation._EXTERNAL_COPIES[str(inside)] = str(outside)
        try:
            where = presentation._refresh_external_copy(str(inside))
        finally:
            presentation._EXTERNAL_COPIES.pop(str(inside), None)

        assert where == str(outside)
        assert outside.read_bytes() == b"second version"

    def test_a_deck_with_no_copy_outside_is_left_alone(self, tmp_path):
        from ppt_mac import presentation

        inside = tmp_path / "deck.pptx"
        inside.write_bytes(b"only version")

        assert presentation._refresh_external_copy(str(inside)) is None
        assert presentation._refresh_external_copy(None) is None

    def test_a_copy_that_cannot_be_written_does_not_fail_the_save(self, tmp_path):
        """The save landed. Saying where it did not reach beats raising."""
        from ppt_mac import presentation

        inside = tmp_path / "deck.pptx"
        inside.write_bytes(b"data")
        unwritable = tmp_path / "gone" / "deck.pptx"

        presentation._EXTERNAL_COPIES[str(inside)] = str(unwritable)
        try:
            assert presentation._refresh_external_copy(str(inside)) is None
        finally:
            presentation._EXTERNAL_COPIES.pop(str(inside), None)


class TestBatchCallsMatchTheImplementations:
    """Every operation ppt_batch_apply_formatting offers has to actually run.

    `format_text` was passing ten arguments to an implementation that takes
    eleven, so the operation failed with a Python TypeError on every call, on
    both platforms. Nothing caught it because the dispatch calls positionally
    and nothing counted. This counts.
    """

    def _dispatch_source(self):
        import inspect

        from ppt_com import batch_apply

        return inspect.getsource(batch_apply._dispatch_op)

    def test_every_operation_is_dispatched(self):
        from ppt_com.batch_apply import SUPPORTED_OPERATIONS

        source = self._dispatch_source()
        for name in SUPPORTED_OPERATIONS:
            assert f'"{name}"' in source, f"{name} has no branch in _dispatch_op"

    def test_each_call_passes_every_argument_its_impl_takes(self):
        import ast
        import inspect
        import textwrap

        from ppt_com import batch_apply, effects, formatting, text

        modules = {
            "_formatting": formatting,
            "_effects": effects,
            "_text": text,
        }
        tree = ast.parse(textwrap.dedent(self._dispatch_source()))
        checked = 0
        for node in ast.walk(tree):
            if not isinstance(node, ast.Call):
                continue
            func = node.func
            if not (
                isinstance(func, ast.Attribute)
                and isinstance(func.value, ast.Name)
                and func.value.id in modules
                and func.attr.endswith("_impl")
            ):
                continue
            impl = getattr(modules[func.value.id], func.attr)
            wanted = len(inspect.signature(impl).parameters)
            assert len(node.args) == wanted, (
                f"{func.value.id}.{func.attr} takes {wanted} arguments and "
                f"the batch dispatch passes {len(node.args)}"
            )
            checked += 1
        assert checked == len(batch_apply.SUPPORTED_OPERATIONS)


@macos_only
class TestWhatTheCodexReviewFound:
    """Three defects found by review on #210, each with the evidence for it."""

    def test_a_jpg_export_holds_jpeg_bytes(self, tmp_path):
        """The format was accepted, named in the file, and then ignored.

        `uti` was worked out from the caller's `format` and never reached
        Quartz, which had "public.png" written into it, so a file called
        Slide1.jpg held PNG bytes. Anything decoding by the format it asked
        for refuses that.
        """
        import Quartz

        from ppt_mac.export import _render_pdf_pages

        pdf = str(tmp_path / "probe.pdf")
        box = Quartz.CGRectMake(0, 0, 200, 100)
        url = Quartz.CFURLCreateFromFileSystemRepresentation(
            None, pdf.encode(), len(pdf.encode()), False
        )
        ctx = Quartz.CGPDFContextCreateWithURL(url, box, None)
        Quartz.CGContextBeginPage(ctx, box)
        Quartz.CGContextSetRGBFillColor(ctx, 1, 0, 0, 1)
        Quartz.CGContextFillRect(ctx, Quartz.CGRectMake(10, 10, 50, 50))
        Quartz.CGContextEndPage(ctx)
        Quartz.CGPDFContextClose(ctx)

        (_, png, _, _) = _render_pdf_pages(pdf, [1])[0]
        assert open(png, "rb").read(4) == b"\x89PNG"

        (_, jpg, _, _) = _render_pdf_pages(
            pdf, [1], uti="public.jpeg", suffix=".jpg"
        )[0]
        assert open(jpg, "rb").read(2) == b"\xff\xd8", "a .jpg holding PNG bytes"

    def test_two_decks_of_one_name_get_two_staged_paths(self, monkeypatch, tmp_path):
        """A basename is not unique, and both halves of the bug followed.

        Staging `~/a/report.pptx` and `~/b/report.pptx` wrote one over the
        other in the container, and keyed both external copies the same, so a
        save on either refreshed whichever registered last.
        """
        from ppt_mac import presentation as mac_pres

        staging = tmp_path / "container"
        staging.mkdir()
        monkeypatch.setattr(mac_pres, "EXPORT_STAGING_DIR", str(staging))
        monkeypatch.setattr(mac_pres, "_full_names", lambda app: [])
        monkeypatch.setattr(mac_pres, "_EXTERNAL_COPIES", {})

        first = mac_pres._free_staged_path(None, str(tmp_path / "a" / "report.pptx"))
        mac_pres._EXTERNAL_COPIES[os.path.abspath(first)] = os.path.abspath(
            str(tmp_path / "a" / "report.pptx")
        )
        second = mac_pres._free_staged_path(None, str(tmp_path / "b" / "report.pptx"))

        assert first != second
        assert os.path.basename(first) == "report.pptx", "the first keeps the name"
        assert os.path.basename(second) == "report-2.pptx"

    def test_saving_the_same_deck_again_reuses_its_own_staged_file(
        self, monkeypatch, tmp_path
    ):
        from ppt_mac import presentation as mac_pres

        staging = tmp_path / "container"
        staging.mkdir()
        monkeypatch.setattr(mac_pres, "EXPORT_STAGING_DIR", str(staging))
        monkeypatch.setattr(mac_pres, "_full_names", lambda app: [])
        monkeypatch.setattr(mac_pres, "_EXTERNAL_COPIES", {})

        target = str(tmp_path / "a" / "report.pptx")
        first = mac_pres._free_staged_path(None, target)
        mac_pres._EXTERNAL_COPIES[os.path.abspath(first)] = os.path.abspath(target)

        assert mac_pres._free_staged_path(None, target) == first

    def test_closing_with_save_carries_the_last_edits_out(self, monkeypatch, tmp_path):
        """The save on the way out is the one most worth copying.

        `ppt_close_presentation(save_changes=True)` wrote the container file
        and closed, so the caller's own copy kept everything except the edits
        they had just asked to keep, and the deck was gone before anyone could
        notice.
        """
        from ppt_mac import presentation as mac_pres

        inside = tmp_path / "container" / "deck.pptx"
        inside.parent.mkdir()
        outside = tmp_path / "Desktop" / "deck.pptx"
        outside.parent.mkdir()
        outside.write_bytes(b"stale")

        class _Saved:
            def set(self, value):
                pass

        class _Pres:
            saved = _Saved()

            def name(self):
                return "deck.pptx"

            def full_name(self):
                return str(inside)

            def path(self):
                return str(inside.parent)

            def save(self):
                inside.write_bytes(b"the edits made just before closing")

            def close(self):
                pass

        monkeypatch.setattr(mac_pres.ppt, "_get_app_impl", lambda: object())
        monkeypatch.setattr(mac_pres, "_resolve_presentation", lambda *a, **k: _Pres())
        monkeypatch.setattr(mac_pres, "_full_names", lambda app: [])
        monkeypatch.setattr(
            mac_pres, "_EXTERNAL_COPIES",
            {os.path.abspath(str(inside)): os.path.abspath(str(outside))},
        )

        result = mac_pres._close_presentation_impl(True, None, None)

        assert outside.read_bytes() == b"the edits made just before closing"
        assert result["also_copied_to"] == os.path.abspath(str(outside))
        # And the deck is gone, so nothing should still be pointing at it.
        assert mac_pres._EXTERNAL_COPIES == {}


@macos_only
class TestADocumentThatOutlivedItsWindow:
    """PowerPoint closes a window and sometimes does not free the document.

    Its own diagnostic log shows the close completing without the matching
    `Document::~Document` and `RemoveFromGlobalList`. The deck then sits in
    `presentations` answering questions, holding all nine of its slides, and
    invisible to the person at the machine. Everything here reaches the editor
    through `document_windows[1]`, which in that state raises -1728 and
    explains nothing (#191).
    """

    def test_a_deck_with_no_window_is_named_as_the_problem(self, monkeypatch):
        from backend import mac_ae

        monkeypatch.setattr(mac_ae, "count_of", lambda container, each: 0)
        with pytest.raises(mac_ae.AppleEventError) as caught:
            mac_ae.target_window(object())

        said = str(caught.value)
        assert "no window" in said
        assert "ppt_close_presentation" in said, "no way out was offered"
        assert "intact" in said, "a caller would think the slides were lost"

    def test_a_deck_with_a_window_hands_it_over(self, monkeypatch):
        from backend import mac_ae

        class _Pres:
            document_windows = {1: "the window"}

        monkeypatch.setattr(mac_ae, "count_of", lambda container, each: 1)
        assert mac_ae.target_window(_Pres()) == "the window"

    def test_nothing_reaches_the_window_around_the_guard(self):
        """The explanation is only worth writing if every call site uses it.

        Two places address `document_windows[1]` directly and both are meant
        to. `target_window` is the guard itself, and `presentation.py` has a
        deck that was just created or just opened, where a caller who passed
        `with_window=False` asked for no window on purpose and would be told
        to close the deck and open it again for no reason.
        """
        import pathlib

        reaching = {
            str(path)
            for path in sorted(pathlib.Path("src").rglob("*.py"))
            for line in path.read_text(encoding="utf-8").splitlines()
            # Comments about the trap are the point of the routing, not a
            # breach of it.
            if "document_windows[1]" in line and not line.strip().startswith("#")
        }

        assert reaching == {
            "src/backend/mac_ae.py",
            "src/ppt_mac/presentation.py",
        }

    def test_targeting_a_windowless_deck_is_refused_rather_than_done(self):
        """It used to activate the window and log whatever came back.

        So a deck that had outlived its window became the session target
        anyway, and every tool after it failed with a bare -1728 instead.
        """
        from backend import mac_ae
        from backend.mac_ae import ppt

        class _Pres:
            def count(self, each=None):
                from appscript import k

                assert each == k.document_window
                return 0

            def name(self):
                return "第二部.pptx"

            def full_name(self):
                return "/deck/第二部.pptx"

        was = ppt._target_pres_full_name
        original = ppt._get_app_impl
        ppt._get_app_impl = lambda *a, **kw: object()
        original_list = mac_ae.PowerPointAppleEventWrapper._presentations
        mac_ae.PowerPointAppleEventWrapper._presentations = (
            lambda self, app_ref: [_Pres()]
        )
        try:
            with pytest.raises(mac_ae.AppleEventError) as caught:
                ppt._set_target_pres_impl(1)
            # The session still points where it did, so the next call is not
            # quietly working on a deck nobody can see.
            assert ppt._target_pres_full_name == was
        finally:
            ppt._get_app_impl = original
            mac_ae.PowerPointAppleEventWrapper._presentations = original_list

        assert "no window" in str(caught.value)


@macos_only
class TestAnEmptyDeckAndADeadReference:
    """`elements` folds -1728 into an empty list, so the two arrive alike.

    A caller was told "The presentation has 0 slides" while PowerPoint was in
    the act of dying underneath it, and went looking for the missing slides
    rather than the missing application.
    """

    class _Slides:
        def get(self):
            return []

    def test_a_reference_that_cannot_answer_says_so(self):
        from appscript.reference import CommandError

        from backend import mac_ae

        class _Dead:
            slides = TestAnEmptyDeckAndADeadReference._Slides()

            def name(self):
                raise _command_error(-1728)

        with pytest.raises(mac_ae.AppleEventError) as caught:
            mac_ae.slide_at(_Dead(), 1)

        said = str(caught.value)
        assert "not an empty deck" in said
        assert "ppt_list_presentations" in said

    def test_a_genuinely_empty_deck_still_reads_as_out_of_range(self):
        from backend import mac_ae

        class _Empty:
            slides = TestAnEmptyDeckAndADeadReference._Slides()

            def name(self):
                return "empty.pptx"

        with pytest.raises(ValueError, match="out of range"):
            mac_ae.slide_at(_Empty(), 1)


class TestABatchOperationSaysWhatItCannotTake:
    """Runs on both platforms: the dispatch is shared, and so was the bug.

    Pydantic drops unknown keys, so a batch `format_text` given `font_color`
    applied nothing and answered `"status": "success"`. `font_color` is what
    `ppt_add_shape` and `ppt_add_textbox` call it; this one calls it `color`.
    An agent building a deck hit exactly that, and only caught it by looking at
    the slide afterwards.
    """

    def test_a_sibling_tools_name_is_refused_and_translated(self):
        from ppt_com.batch_apply import _dispatch_op

        with pytest.raises(ValueError) as caught:
            _dispatch_op(1, "Card", "format_text", {"font_color": "#FFFFFF"})

        said = str(caught.value)
        assert "does not take 'font_color'" in said
        assert "did you mean 'color'" in said, "the near miss guessed wrong"
        assert "Nothing was applied" in said

    def test_a_plain_typo_still_gets_a_guess(self):
        from ppt_com.batch_apply import _dispatch_op

        with pytest.raises(ValueError) as caught:
            _dispatch_op(1, "Card", "format_text", {"colour": "red"})
        assert "did you mean 'color'" in str(caught.value)

    def test_the_arguments_it_does_take_are_listed(self):
        from ppt_com.batch_apply import _dispatch_op

        with pytest.raises(ValueError) as caught:
            _dispatch_op(1, "Card", "set_line", {"line_visible": False})
        said = str(caught.value)
        assert "did you mean 'visible'" in said
        assert "It takes: color, dash_style, transparency, visible, weight." in said
