"""Tests for the macOS advanced operation and batch formatting tools.

Pure unit tests against a fake object graph. None of this launches PowerPoint,
and none of it may; the live behaviour is covered by MACOS_PORT.md and by
running the server. appscript only installs on macOS, so everything that
imports it is skipped elsewhere.

What is worth testing here is what the two modules had to decide. Which Windows
constant becomes which macOS enumerator and back again, what each tool says when
macOS cannot do it at all, that a refusal is answered before PowerPoint is
touched, that everything naming a path to PowerPoint goes through its container
first, and that a refusal coming back from a batched operation is recorded as a
failure rather than as a success.
"""

import os
import sys

import pytest

sys.path.insert(0, "src")

macos_only = pytest.mark.skipif(
    sys.platform != "darwin", reason="the Apple Event backend needs appscript"
)


# ---------------------------------------------------------------------------
# Enumerator tables
# ---------------------------------------------------------------------------
@macos_only
class TestViewTypes:
    """The project's own view table and the generated one disagree twice."""

    def test_every_view_carries_a_windows_number_and_a_name(self):
        from ppt_mac.advanced_ops import _VIEW_TYPES

        numbers = sorted(number for number, _ in _VIEW_TYPES.values())
        assert numbers == list(range(1, 12))
        assert all(name for _, name in _VIEW_TYPES.values())

    def test_the_generated_table_agrees_wherever_it_has_an_entry(self):
        """The local table adds 2, 3 and 11; it must not contradict the rest."""
        from backend.mac_enums import PpViewType
        from ppt_mac.advanced_ops import _VIEW_KEYWORD_BY_WINDOWS

        for number, word in PpViewType.items():
            assert _VIEW_KEYWORD_BY_WINDOWS[number] == word

    def test_every_name_the_tool_accepts_reaches_a_macos_view(self):
        from ppt_com.constants import VIEW_TYPE_MAP
        from ppt_mac.advanced_ops import _VIEW_KEYWORD_BY_WINDOWS

        for name, number in VIEW_TYPE_MAP.items():
            assert number in _VIEW_KEYWORD_BY_WINDOWS, name

    def test_the_two_places_the_project_map_is_mislabelled(self):
        """Both platforms put ppViewSlide at 1 and print preview at 10."""
        from appscript import k

        from ppt_com.constants import VIEW_TYPE_MAP
        from ppt_mac.advanced_ops import _VIEW_KEYWORD_BY_WINDOWS

        assert _VIEW_KEYWORD_BY_WINDOWS[VIEW_TYPE_MAP["normal"]] == k.slide_view
        assert _VIEW_KEYWORD_BY_WINDOWS[VIEW_TYPE_MAP["reading"]] == k.print_preview
        # And normal view, which the project map has no name for, is 9.
        assert _VIEW_KEYWORD_BY_WINDOWS[9] == k.normal_view


@macos_only
class TestOtherEnums:
    """Selection type, picture colour type and the export formats."""

    def test_selection_types_translate_both_ways(self):
        from appscript import k

        from backend.mac_enums import PpSelectionType
        from ppt_mac.advanced_ops import _WIN_SELECTION_TYPE

        assert PpSelectionType[2] == k.selection_type_shapes
        assert _WIN_SELECTION_TYPE[k.selection_type_shapes] == 2
        assert _WIN_SELECTION_TYPE[k.selection_type_none] == 0
        assert _WIN_SELECTION_TYPE[k.selection_type_text] == 3

    def test_picture_colour_types_translate_both_ways(self):
        from appscript import k

        from ppt_com.constants import PICTURE_COLOR_TYPE_MAP, PICTURE_COLOR_TYPE_NAMES
        from backend.mac_enums import MsoPictureColorType, to_keyword
        from ppt_mac.advanced_ops import _WIN_PICTURE_COLOR_TYPE

        for name, number in PICTURE_COLOR_TYPE_MAP.items():
            word = to_keyword(MsoPictureColorType, number, "picture colour type")
            assert _WIN_PICTURE_COLOR_TYPE[word] == number
            assert PICTURE_COLOR_TYPE_NAMES[number] == name
        assert MsoPictureColorType[2] == k.picture_color_gray_scale

    def test_the_four_export_formats_macos_can_write(self):
        from appscript import k

        from ppt_com.constants import SHAPE_FORMAT_MAP
        from ppt_mac.advanced_ops import (
            _SHAPE_FORMATS,
            _SHAPE_FORMAT_NAMES_UNSUPPORTED,
        )

        assert _SHAPE_FORMATS[SHAPE_FORMAT_MAP["png"]] == (k.save_as_PNG_file, ".png")
        assert _SHAPE_FORMATS[SHAPE_FORMAT_MAP["jpg"]] == (k.save_as_JPG_file, ".jpg")
        assert _SHAPE_FORMATS[SHAPE_FORMAT_MAP["gif"]] == (k.save_as_GIF_file, ".gif")
        assert _SHAPE_FORMATS[SHAPE_FORMAT_MAP["bmp"]] == (k.save_as_BMP_file, ".bmp")
        # Every Windows format is accounted for, as a keyword or as a refusal.
        covered = set(_SHAPE_FORMATS) | set(_SHAPE_FORMAT_NAMES_UNSUPPORTED)
        assert covered == set(SHAPE_FORMAT_MAP.values())


# ---------------------------------------------------------------------------
# Refusals
# ---------------------------------------------------------------------------
@macos_only
class TestTagRefusals:
    """There is no Tags collection anywhere in the dictionary."""

    def test_setting_a_tag_refuses(self):
        from ppt_mac.advanced_ops import _set_tag_impl

        result = _set_tag_impl(1, "Box", "owner", "rika", "shape")

        assert result["error"] == "ppt_set_tag is not available on macOS"
        assert result["platform"] == "macOS"
        assert "no Tags collection" in result["reason"]
        assert "command bar control" in result["reason"]
        assert result["alternatives"]

    def test_reading_tags_refuses(self):
        from ppt_mac.advanced_ops import _get_tags_impl

        result = _get_tags_impl(1, "Box", "shape")

        assert result["error"] == "ppt_get_tags is not available on macOS"
        assert "nothing to read" in result["reason"]

    def test_neither_touches_powerpoint(self):
        from ppt_mac.advanced_ops import _get_tags_impl, _set_tag_impl

        with _no_powerpoint():
            assert "error" in _set_tag_impl(1, "Box", "k", "v", "shape")
            assert "error" in _get_tags_impl(1, "Box", "shape")


@macos_only
class TestSelectionRefusal:
    """Selecting needs a shape range, and there is no way to build one."""

    def test_selecting_refuses(self):
        from ppt_mac.advanced_ops import _select_shapes_impl

        result = _select_shapes_impl(1, ["Box", "Circle"])

        assert result["error"] == "ppt_select_shapes is not available on macOS"
        assert result["platform"] == "macOS"
        assert "no `select` command" in result["reason"]

    def test_it_says_the_same_thing_layout_says(self):
        """Three modules refusing for one reason should not word it three ways."""
        from ppt_mac.advanced_ops import _select_shapes_impl
        from ppt_mac.layout import _NO_SHAPE_RANGE

        assert _select_shapes_impl(1, ["Box"])["reason"] == _NO_SHAPE_RANGE

    def test_the_refusal_never_reaches_powerpoint(self):
        from ppt_mac.advanced_ops import _select_shapes_impl

        with _no_powerpoint():
            assert "error" in _select_shapes_impl(1, ["Box"])


@macos_only
class TestCopyAnimationRefusal:
    """The only route left is the write that rewrites the whole slide."""

    def test_copying_animation_refuses(self):
        from ppt_mac.advanced_ops import _copy_animation_impl

        result = _copy_animation_impl(1, "Box", "Circle")

        assert result["error"] == "ppt_copy_animation is not available on macOS"
        assert "pick up" in result["reason"]
        assert "animation settings" in result["reason"]
        assert "ppt_add_animation" in result["alternatives"][0]

    def test_the_refusal_never_reaches_powerpoint(self):
        """A slide's animations must not be risked to find out it cannot work."""
        from ppt_mac.advanced_ops import _copy_animation_impl

        with _no_powerpoint():
            assert "error" in _copy_animation_impl(1, "Box", "Circle")


@macos_only
class TestExportFormatRefusal:
    """wmf and emf have no counterpart, and only those two are refused."""

    @pytest.mark.parametrize("number,name", [(4, "wmf"), (5, "emf")])
    def test_the_vector_formats_refuse_by_name(self, number, name):
        from ppt_mac.advanced_ops import _export_shape_impl

        with _no_powerpoint():
            result = _export_shape_impl(1, "Box", "/tmp/out", number, None, None)

        assert result["error"] == f"ppt_export_shape cannot write {name} on macOS"
        assert "is not available" not in result["error"]
        assert "png" in result["alternatives"][0]

    def test_an_unknown_format_number_raises(self):
        from ppt_mac.advanced_ops import _export_shape_impl

        with _no_powerpoint():
            with pytest.raises(ValueError, match="Unknown format"):
                _export_shape_impl(1, "Box", "/tmp/out", 99, None, None)

    def test_an_unknown_format_name_raises(self):
        from ppt_mac.advanced_ops import _export_shape_impl

        with _no_powerpoint():
            with pytest.raises(ValueError, match="Unknown format 'tiff'"):
                _export_shape_impl(1, "Box", "/tmp/out", "tiff", None, None)


# ---------------------------------------------------------------------------
# Paths and the sandbox
# ---------------------------------------------------------------------------
@macos_only
class TestStaging:
    """Nothing hands PowerPoint a path outside its own container."""

    def test_a_staged_file_lands_in_the_container(self, tmp_path, monkeypatch):
        from ppt_mac import advanced_ops

        monkeypatch.setattr(advanced_ops, "EXPORT_STAGING_DIR", str(tmp_path))
        path = advanced_ops._staged_file(".svg")

        assert os.path.dirname(path) == str(tmp_path)
        assert path.endswith(".svg")
        assert os.path.exists(path)

    def test_an_export_is_written_in_the_container_and_moved_out(
        self, tmp_path, monkeypatch
    ):
        from ppt_mac import export
        from ppt_mac.advanced_ops import _export_shape_impl

        staging = tmp_path / "container"
        monkeypatch.setattr(export, "EXPORT_STAGING_DIR", str(staging))
        destination = tmp_path / "out" / "box.png"

        with _fake_deck(["Box"]) as deck:
            result = _export_shape_impl(1, "Box", str(destination), 2, None, None)

        (asked_for,) = deck.shape("Box").exports
        assert os.path.dirname(asked_for) == str(staging)
        assert result["file_path"] == str(destination)
        assert destination.read_bytes() == b"PNG"
        # Nothing is left behind in the container.
        assert os.listdir(staging) == []

    def test_an_export_that_writes_nothing_is_reported_as_the_no_op_it_is(
        self, tmp_path, monkeypatch
    ):
        from ppt_mac import export
        from ppt_mac.advanced_ops import _export_shape_impl

        monkeypatch.setattr(export, "EXPORT_STAGING_DIR", str(tmp_path / "c"))
        destination = tmp_path / "box.png"

        with _fake_deck(["Box"]) as deck:
            deck.shape("Box").export_writes = False
            result = _export_shape_impl(1, "Box", str(destination), 2, None, None)

        assert result["error"] == "ppt_export_shape wrote no file"
        # The tool works; this call did not land, and the headline says which.
        assert "is not available" not in result["error"]
        assert "wrote no file" in result["reason"]
        assert not destination.exists()

    def test_a_size_is_named_in_a_warning_rather_than_dropped(
        self, tmp_path, monkeypatch
    ):
        from ppt_mac import export
        from ppt_mac.advanced_ops import _export_shape_impl

        monkeypatch.setattr(export, "EXPORT_STAGING_DIR", str(tmp_path / "c"))

        with _fake_deck(["Box"]):
            result = _export_shape_impl(
                1, "Box", str(tmp_path / "box.png"), 2, 800, 600
            )

        assert result["success"] is True
        assert "takes no size" in result["warnings"][0]

    def test_a_download_is_staged_in_the_container_and_cleaned_up(
        self, tmp_path, monkeypatch
    ):
        from ppt_mac import advanced_ops

        staging = tmp_path / "container"
        monkeypatch.setattr(advanced_ops, "EXPORT_STAGING_DIR", str(staging))
        monkeypatch.setattr(
            advanced_ops.urllib.request, "urlopen",
            lambda url: _FakeResponse(b"\x89PNG", "image/png"),
        )

        with _fake_deck([]) as deck:
            result = advanced_ops._add_picture_from_url_impl(
                1, "https://example.test/a.png", 10, 20, None, None, None, False
            )

        (given,) = deck.picture_paths
        assert os.path.dirname(given) == str(staging)
        assert result["source_url"] == "https://example.test/a.png"
        assert os.listdir(staging) == []

    def test_an_svg_icon_is_coloured_then_rendered_to_a_png_before_it_is_handed_over(
        self, tmp_path, monkeypatch
    ):
        """PowerPoint for Mac cannot read an SVG, so it is never handed one."""
        from ppt_mac import advanced_ops

        staging = tmp_path / "container"
        monkeypatch.setattr(advanced_ops, "EXPORT_STAGING_DIR", str(staging))
        monkeypatch.setattr(
            advanced_ops.urllib.request, "urlopen",
            lambda url: _FakeResponse(
                b'<svg viewBox="0 0 24 24"><path fill="currentColor"/></svg>',
                "image/svg+xml",
            ),
        )
        monkeypatch.setattr(advanced_ops, "_sips_renders_svg", lambda: True)

        rendered = []

        def fake_render(svg_path, png_path, pixels):
            rendered.append((open(svg_path, encoding="utf-8").read(), pixels))
            with open(png_path, "wb") as handle:
                handle.write(b"\x89PNG\r\n\x1a\n")
            return True

        monkeypatch.setattr(advanced_ops, "_sips_to_png", fake_render)

        with _fake_deck([]) as deck:
            result = advanced_ops._add_svg_icon_impl(
                1, "bolt", 0, 0, 72, 72, "#1F6FEB", "outlined", False
            )

        (svg_text, pixels) = rendered[0]
        assert "currentColor" not in svg_text
        assert "#1F6FEB" in svg_text
        assert pixels == 72 * advanced_ops._SVG_RENDER_SCALE

        handed_over = deck.picture_paths[0]
        assert handed_over.endswith(".png")
        assert os.path.dirname(handed_over) == str(staging)
        assert "PNG" in deck.picture_contents[0]
        assert result["icon_name"] == "bolt"
        assert result["source_url"].endswith("/outlined/bolt.svg")
        assert os.listdir(staging) == []

    def test_an_svg_icon_refuses_where_sips_cannot_read_an_svg(self, monkeypatch):
        """Rather than leave the empty box PowerPoint puts there instead."""
        from ppt_mac import advanced_ops

        monkeypatch.setattr(advanced_ops, "_sips_renders_svg", lambda: False)
        result = advanced_ops._add_svg_icon_impl(
            1, "bolt", 0, 0, 72, 72, "#1F6FEB", "outlined", False
        )

        assert "error" in result
        assert "macOS 13" in result["reason"]

    def test_sips_really_renders_an_svg_on_this_machine(self):
        """The one measurement the whole icon route rests on."""
        from ppt_mac import advanced_ops

        assert advanced_ops._sips_renders_svg() is True

    def test_a_picture_powerpoint_would_not_read_is_not_reported_as_added(
        self, tmp_path, monkeypatch
    ):
        """An unreadable file leaves a plain autoshape behind and no error."""
        from appscript import k

        from ppt_mac import advanced_ops

        monkeypatch.setattr(advanced_ops, "EXPORT_STAGING_DIR", str(tmp_path))
        monkeypatch.setattr(
            advanced_ops.urllib.request, "urlopen",
            lambda url: _FakeResponse(b"not an image", "image/png"),
        )

        with _fake_deck([]) as deck:
            deck.make_type = k.shape_type_auto
            with pytest.raises(RuntimeError, match="plain autoshape"):
                advanced_ops._add_picture_from_url_impl(
                    1, "https://example.test/a.png", 0, 0, None, None, None, False
                )

        # And the empty box is taken off the slide again. Leaving it there made
        # every retry add another one for someone to find by hand afterwards.
        assert len(deck.deleted_shapes) == 1


# ---------------------------------------------------------------------------
# Argument validation
# ---------------------------------------------------------------------------
@macos_only
class TestValidation:
    """A bad argument is refused before anything on the slide is touched."""

    def test_an_unknown_crop_shape_name_is_named(self):
        from ppt_mac.advanced_ops import _crop_picture_impl

        with _no_powerpoint():
            with pytest.raises(ValueError, match="Unknown crop_shape 'sausage'"):
                _crop_picture_impl(
                    1, "Pic", None, None, None, None,
                    "sausage", None, None, None,
                )

    def test_an_unknown_crop_fit_is_named(self):
        from ppt_mac.advanced_ops import _crop_picture_impl

        with _no_powerpoint():
            with pytest.raises(ValueError, match="Unknown crop_fit 'round'"):
                _crop_picture_impl(
                    1, "Pic", None, None, None, None, None, "round", None, None,
                )

    def test_cropping_something_that_is_not_a_picture_says_so(self):
        from ppt_mac.advanced_ops import _crop_picture_impl

        with _fake_deck(["Box"]):
            with pytest.raises(ValueError, match="is not a picture"):
                _crop_picture_impl(
                    1, "Box", 5, None, None, None, None, None, None, None,
                )

    def test_default_fonts_needs_at_least_one_font(self):
        from ppt_mac.advanced_ops import _set_default_fonts_impl

        with _no_powerpoint():
            with pytest.raises(ValueError, match="At least one of"):
                _set_default_fonts_impl(None, None, False)

    def test_an_unknown_view_type_lists_the_ones_that_work(self):
        from ppt_mac.advanced_ops import _set_view_impl

        with _no_powerpoint():
            with pytest.raises(ValueError, match="Unknown view_type 'zoomed'"):
                _set_view_impl("zoomed", None)

    def test_an_unknown_colour_name_lists_the_theme_names(self):
        from ppt_mac.advanced_ops import _resolve_color

        with pytest.raises(ValueError, match="Unknown color 'puce'"):
            _resolve_color(None, "puce")

    def test_a_hex_colour_never_asks_powerpoint_for_the_palette(self):
        from ppt_mac.advanced_ops import _resolve_color

        assert _resolve_color(None, "#123456") == "#123456"


# ---------------------------------------------------------------------------
# Writes, and reading them back
# ---------------------------------------------------------------------------
@macos_only
class TestWritesAreVerified:
    """Nothing is trusted because it did not raise."""

    def test_hiding_a_slide_writes_a_real_boolean(self):
        from ppt_mac.advanced_ops import _set_slide_hidden_impl

        with _fake_deck([]) as deck:
            result = _set_slide_hidden_impl(1, True)

        assert deck.slide_object.slide_show_transition.hidden() is True
        assert result == {"success": True, "slide_index": 1, "hidden": True}

    def test_a_hidden_flag_that_does_not_land_is_reported_as_a_no_op(self):
        from ppt_mac.advanced_ops import _set_slide_hidden_impl

        with _fake_deck([]) as deck:
            deck.slide_object.slide_show_transition.hidden.clamp_to = False
            result = _set_slide_hidden_impl(1, True)

        assert result["error"] == "ppt_set_slide_hidden did not change the slide"
        assert "is not available" not in result["error"]
        assert "silent no-op" in result["reason"]

    def test_locking_the_aspect_ratio_reads_back(self):
        from ppt_mac.advanced_ops import _lock_aspect_ratio_impl

        with _fake_deck(["Box"]) as deck:
            result = _lock_aspect_ratio_impl(1, "Box", True)

        assert deck.shape("Box").lock_aspect_ratio() is True
        assert result == {"success": True, "shape_name": "Box", "locked": True}

    def test_a_lock_that_does_not_land_is_reported_as_a_no_op(self):
        from ppt_mac.advanced_ops import _lock_aspect_ratio_impl

        with _fake_deck(["Box"]) as deck:
            deck.shape("Box").lock_aspect_ratio.clamp_to = False
            result = _lock_aspect_ratio_impl(1, "Box", True)

        assert result["error"] == "ppt_lock_aspect_ratio did not change the shape"
        assert "is not available" not in result["error"]
        assert "silent no-op" in result["reason"]


@macos_only
class TestView:
    """The view type is written on the window and the zoom on its view."""

    def test_a_view_and_a_zoom_both_land(self):
        from appscript import k

        from ppt_mac.advanced_ops import _set_view_impl

        with _fake_deck([]) as deck:
            result = _set_view_impl("slide_sorter", 75)

        assert deck.window.view_type() == k.slide_sorter_view
        assert deck.window.view.zoom() == 75
        assert result["view_type"] == "slide_sorter"
        assert result["view_type_id"] == 7
        assert result["zoom"] == 75
        assert "warnings" not in result

    def test_a_view_powerpoint_declines_is_reported_rather_than_claimed(self):
        from ppt_mac.advanced_ops import _set_view_impl

        with _fake_deck([]) as deck:
            deck.window.view_type.clamp_to = _slide_view()
            result = _set_view_impl("outline", None)

        assert result["success"] is True
        assert result["view_type"] == "normal"
        assert "did not switch" in result["warnings"][0]

    def test_a_zoom_powerpoint_kept_inside_its_own_range_is_reported(self):
        from ppt_mac.advanced_ops import _set_view_impl

        with _fake_deck([]) as deck:
            deck.window.view.zoom.clamp_to = 100
            result = _set_view_impl(None, 400)

        assert result["zoom"] == 100
        assert "keeps a view inside its own range" in result["warnings"][0]


@macos_only
class TestSelectionReading:
    """Refusing to select is not a reason to refuse to read the selection."""

    def test_a_shape_selection_is_counted_through_its_container(self):
        from ppt_mac.advanced_ops import _get_selection_impl

        with _fake_deck([]) as deck:
            deck.select_shapes(["Box", "Circle"])
            result = _get_selection_impl()

        assert result["type"] == 2
        assert result["type_name"] == "shapes"
        assert result["shape_names"] == ["Box", "Circle"]
        assert result["count"] == 2
        # Never `.get()` and never `.count()` on the range itself.
        assert deck.selection.shape_range.materialised == 0

    def test_a_slide_selection_reports_its_indices(self):
        from ppt_mac.advanced_ops import _get_selection_impl

        with _fake_deck([]) as deck:
            deck.select_slides([1, 3])
            result = _get_selection_impl()

        assert result["type"] == 1
        assert result["type_name"] == "slides"
        assert result["slide_indices"] == [1, 3]

    def test_a_text_selection_reports_its_content(self):
        from ppt_mac.advanced_ops import _get_selection_impl

        with _fake_deck([]) as deck:
            deck.select_text("hello")
            result = _get_selection_impl()

        assert result["type"] == 3
        assert result["type_name"] == "text"
        assert result["text"] == "hello"

    def test_an_empty_selection_reports_none(self):
        from ppt_mac.advanced_ops import _get_selection_impl

        with _fake_deck([]):
            result = _get_selection_impl()

        assert result["type"] == 0
        assert result["type_name"] == "none"


@macos_only
class TestFonts:
    """No replace command, so the deck is walked and the gaps are named."""

    def test_replacing_a_font_walks_the_shapes_and_says_what_it_missed(self):
        from ppt_mac.advanced_ops import _replace_font_impl

        with _fake_deck(["A", "B"]) as deck:
            deck.shape("A").set_fonts("Arial", "Arial")
            deck.shape("B").set_fonts("Meiryo", "Meiryo")
            result = _replace_font_impl("Arial", "Inter")

        assert deck.shape("A").text_frame.text_range.font.font_name() == "Inter"
        assert deck.shape("B").text_frame.text_range.font.font_name() == "Meiryo"
        assert result["success"] is True
        assert result["original_font"] == "Arial"
        assert result["replacement_font"] == "Inter"
        assert result["shapes_updated"] == 1
        assert any("Masters and layouts" in w for w in result["warnings"])

    def test_a_shape_mixing_fonts_is_skipped_and_counted(self):
        from ppt_mac.advanced_ops import _replace_font_impl

        with _fake_deck(["A"]) as deck:
            deck.shape("A").set_fonts(None, None)
            result = _replace_font_impl("Arial", "Inter")

        assert result["shapes_updated"] == 0
        assert any("mix more than one font" in w for w in result["warnings"])

    def test_listing_fonts_walks_the_theme_and_the_shapes(self):
        """`presentation.fonts` counts and then will not resolve, so it is not used.

        Live, a deck using one font answers 2 for its font count, nothing for
        the collection's contents, and -1728 for `fonts[1]`. The tool walks the
        theme font scheme and every shape with text instead, which is the
        question it was asked.
        """
        from ppt_mac.advanced_ops import _list_fonts_impl

        with _fake_deck(["A", "B"]) as deck:
            deck.font_scheme.major[1].name.set("Inter")
            deck.font_scheme.minor[3].name.set("Meiryo")
            result = _list_fonts_impl()

        assert result["success"] is True
        assert "Inter" in result["fonts"]
        assert "Meiryo" in result["fonts"]
        assert result["fonts"] == sorted(result["fonts"])
        assert result["fonts_count"] == len(result["fonts"])
        assert any("will not enumerate" in w for w in result["warnings"])

    def test_the_font_collection_is_never_touched(self):
        """Reaching for it is the mistake this tool exists not to make."""
        import inspect

        from ppt_mac import advanced_ops

        source = inspect.getsource(advanced_ops._list_fonts_impl)
        assert "pres.fonts" not in source
        assert "k.font" not in source

    def test_theme_fonts_are_written_at_the_latin_and_east_asian_positions(self):
        from ppt_mac.advanced_ops import _set_default_fonts_impl

        with _fake_deck([]) as deck:
            result = _set_default_fonts_impl("Inter", "Meiryo", False)

        scheme = deck.font_scheme
        assert scheme.major[1].name() == "Inter"
        assert scheme.major[3].name() == "Meiryo"
        assert scheme.minor[1].name() == "Inter"
        assert scheme.minor[3].name() == "Meiryo"
        assert result["theme_updated"] is True
        assert result["latin"] == "Inter"
        assert result["east_asian"] == "Meiryo"
        assert "slides_processed" not in result

    def test_applying_to_existing_text_counts_the_shapes_and_warns_on_groups(self):
        from ppt_mac.advanced_ops import _set_default_fonts_impl

        with _fake_deck(["A", "B"]) as deck:
            result = _set_default_fonts_impl("Inter", None, True)

        assert deck.shape("A").text_frame.text_range.font.font_name() == "Inter"
        assert result["slides_processed"] == 1
        assert result["shapes_updated"] == 2
        assert any("grouped shapes" in w for w in result["warnings"])

    def test_a_theme_write_that_does_not_land_says_so(self):
        from ppt_mac.advanced_ops import _set_default_fonts_impl

        with _fake_deck([]) as deck:
            deck.font_scheme.minor[1].name.clamp_to = "Calibri"
            result = _set_default_fonts_impl("Inter", None, False)

        assert result["theme_updated"] is False
        assert "reads back as" in result["warnings"][0]

    def test_every_slot_written_is_read_back_not_just_the_first(self):
        """The Latin face landing says nothing about the East Asian one."""
        from ppt_mac.advanced_ops import _set_default_fonts_impl

        with _fake_deck([]) as deck:
            deck.font_scheme.minor[3].name.clamp_to = "Calibri"
            result = _set_default_fonts_impl("Inter", "Meiryo", False)

        assert deck.font_scheme.minor[1].name() == "Inter"
        assert result["theme_updated"] is False
        assert "'Meiryo'" in result["warnings"][0]

    def test_the_major_family_is_probed_as_well_as_the_minor(self):
        """Writing both families and reading one leaves half the write unchecked."""
        from ppt_mac.advanced_ops import _set_default_fonts_impl

        with _fake_deck([]) as deck:
            deck.font_scheme.major[1].name.clamp_to = "Calibri"
            result = _set_default_fonts_impl("Inter", None, False)

        assert deck.font_scheme.minor[1].name() == "Inter"
        assert result["theme_updated"] is False
        assert "major theme font" in result["warnings"][0]

    def test_a_font_that_reads_back_unchanged_is_not_counted(self):
        """A write that did not raise is not the same as a font that changed."""
        from ppt_mac.advanced_ops import _replace_font_impl

        with _fake_deck(["A"]) as deck:
            deck.shape("A").set_fonts("Arial", "Arial")
            font = deck.shape("A").text_frame.text_range.font
            font.font_name.clamp_to = "Arial"
            font.east_asian_name.clamp_to = "Arial"
            result = _replace_font_impl("Arial", "Inter")

        assert result["shapes_updated"] == 0
        assert any("still read back" in w for w in result["warnings"])


@macos_only
class TestPictures:
    """Cropping and picture adjustments, which macOS has in full."""

    def test_manual_crops_are_written_and_read_back(self):
        from ppt_mac.advanced_ops import _crop_picture_impl

        with _fake_picture_deck("Pic") as deck:
            result = _crop_picture_impl(
                1, "Pic", 5, 6, 7, 8, None, None, None, None,
            )

        crops = deck.shape("Pic").picture_format
        assert (crops.crop_left(), crops.crop_right()) == (5, 6)
        assert (crops.crop_top(), crops.crop_bottom()) == (7, 8)
        assert result["shape_name"] == "Pic"
        assert result["crop_left"] == 5

    def test_a_square_crop_measures_the_image_and_centres_the_cut(self):
        """Scale to the original size is the only way to read the true ratio."""
        from ppt_mac.advanced_ops import _crop_picture_impl

        with _fake_picture_deck("Pic") as deck:
            pic = deck.shape("Pic")
            pic.width.set(200.0)
            pic.height.set(100.0)
            pic.original = (400.0, 200.0)
            _crop_picture_impl(
                1, "Pic", None, None, None, None, None, "square", None, None,
            )

        crops = pic.picture_format
        # 400 by 200 loses 200 from the width, half off each side by default.
        assert crops.crop_left() == 100.0
        assert crops.crop_right() == 100.0
        assert crops.crop_top() == 0
        assert (pic.width(), pic.height()) == (100.0, 100.0)
        assert pic.lock_aspect_ratio() is False
        assert [entry[0] for entry in pic.scaled] == ["width", "height"]

    def test_a_square_crop_honours_the_anchor(self):
        from ppt_mac.advanced_ops import _crop_picture_impl

        with _fake_picture_deck("Pic") as deck:
            pic = deck.shape("Pic")
            pic.width.set(100.0)
            pic.height.set(200.0)
            pic.original = (200.0, 400.0)
            _crop_picture_impl(
                1, "Pic", None, None, None, None, None, "1:1", 0.0, None,
            )

        crops = pic.picture_format
        assert crops.crop_top() == 0.0
        assert crops.crop_bottom() == 200.0

    def test_picture_format_speaks_hex_where_apple_events_speak_lists(self):
        from appscript import k

        from ppt_mac.advanced_ops import _set_picture_format_impl

        with _fake_picture_deck("Pic") as deck:
            result = _set_picture_format_impl(
                1, "Pic", 0.7, 0.3, "grayscale", "#FF0000", None,
            )

        pf = deck.shape("Pic").picture_format
        assert pf.color_type() == k.picture_color_gray_scale
        assert pf.transparency_color() == [255, 0, 0]
        assert pf.transparent_background() is True
        assert result["color_type"] == 2
        assert result["color_type_name"] == "grayscale"
        assert result["transparent_color_hex"] == "#FF0000"
        assert result["brightness"] == 0.7

    def test_a_shape_that_is_not_a_picture_is_refused_by_both_tools(self):
        from ppt_mac.advanced_ops import _set_picture_format_impl

        with _fake_deck(["Box"]):
            with pytest.raises(ValueError, match="ppt_set_picture_format"):
                _set_picture_format_impl(
                    1, "Box", 0.5, None, None, None, None,
                )


@macos_only
class TestDefaultShapeStyle:
    """A throwaway template shape, which has to be gone afterwards."""

    def test_a_source_shape_is_captured_and_the_answer_is_encoded(self):
        import json

        from ppt_mac.advanced_ops import _set_default_shape_style_from_shape_impl

        with _fake_deck(["Box"]) as deck:
            answer = _set_default_shape_style_from_shape_impl(1, "Box")

        assert deck.shape("Box").defaults_set == 1
        payload = json.loads(answer)
        assert payload["success"] is True
        assert payload["source_shape"] == "Box"
        # Nothing reads a default style back, so the answer has to say that
        # rather than let the command standing in for evidence.
        assert "rather than a measurement" in payload["warnings"][0]

    def test_the_template_shape_is_deleted_even_though_it_was_styled(self):
        import json

        from ppt_mac.advanced_ops import _set_default_shape_style_impl

        with _fake_deck([]) as deck:
            answer = _set_default_shape_style_impl(
                "solid", "#101010", False, "#202020", 0.75,
                "Inter", 18, True, False, "#FFFFFF",
            )

        payload = json.loads(answer)
        assert payload["success"] is True
        assert "rather than a measurement" in payload["warnings"][0]
        assert deck.slide_shapes == []
        (template,) = deck.deleted_shapes
        assert template.defaults_set == 1
        assert template.fill_format.solid_calls == 1
        assert template.fill_format.fore_color() == [16, 16, 16]
        assert template.line_format.fore_color() == [32, 32, 32]
        assert template.text_frame.text_range.font.font_name() == "Inter"

    def test_a_deck_with_no_slides_says_so_before_anything_is_made(self):
        from ppt_mac.advanced_ops import _set_default_shape_style_impl

        with _fake_deck([]) as deck:
            deck.no_slides = True
            with pytest.raises(ValueError, match="no slides"):
                _set_default_shape_style_impl(
                    None, None, None, None, None, None, None, None, None, None,
                )


# ---------------------------------------------------------------------------
# Batch formatting
# ---------------------------------------------------------------------------
@macos_only
class TestBatchRecording:
    """A refusal is a dict, not a raise, so a batch has to look at it."""

    def test_a_refusal_dict_is_recorded_as_an_error(self):
        from ppt_mac.batch_apply import _describe_failure

        failure = _describe_failure({
            "error": "ppt_set_glow is not available on macOS",
            "reason": "there is no glow here",
            "platform": "macOS",
        })

        assert failure == (
            "ppt_set_glow is not available on macOS. there is no glow here"
        )

    def test_a_refusal_with_no_reason_still_reads(self):
        from ppt_mac.batch_apply import _describe_failure

        assert _describe_failure({"error": "nope"}) == "nope"

    def test_a_successful_dict_is_not_mistaken_for_a_failure(self):
        from ppt_mac.batch_apply import _describe_failure

        assert _describe_failure({"success": True, "shape_name": "Box"}) is None

    def test_anything_that_is_not_a_dict_counts_as_success(self):
        from ppt_mac.batch_apply import _describe_failure

        assert _describe_failure(None) is None
        assert _describe_failure("{}") is None

    def test_the_batch_records_a_refusal_as_an_error(self, monkeypatch):
        from ppt_com import batch_apply as com_batch
        from ppt_mac.batch_apply import _batch_apply_impl

        monkeypatch.setattr(
            com_batch, "_dispatch_op",
            lambda *a: {
                "error": "ppt_set_glow is not available on macOS",
                "reason": "no glow format here",
            },
        )

        with _fake_deck(["Box"]):
            result = _batch_apply_impl(1, ["Box"], [{"tool": "set_glow"}])

        (step,) = result["results"][0]["operations"]
        assert step["status"] == "error"
        assert "no glow format here" in step["error"]

    def test_the_batch_still_records_a_real_success(self, monkeypatch):
        from ppt_com import batch_apply as com_batch
        from ppt_mac.batch_apply import _batch_apply_impl

        monkeypatch.setattr(
            com_batch, "_dispatch_op", lambda *a: {"success": True},
        )

        with _fake_deck(["Box"]):
            result = _batch_apply_impl(1, ["Box"], [{"tool": "set_fill"}])

        assert result["results"][0]["operations"] == [
            {"tool": "set_fill", "status": "success"}
        ]

    def test_a_raise_is_still_recorded_per_operation(self, monkeypatch):
        from ppt_com import batch_apply as com_batch
        from ppt_mac.batch_apply import _batch_apply_impl

        def _boom(*args):
            raise ValueError("bad radius")

        monkeypatch.setattr(com_batch, "_dispatch_op", _boom)

        with _fake_deck(["Box"]):
            result = _batch_apply_impl(1, ["Box"], [{"tool": "set_glow"}])

        (step,) = result["results"][0]["operations"]
        assert step == {
            "tool": "set_glow", "status": "error", "error": "bad radius",
        }

    def test_a_missing_shape_is_recorded_and_the_rest_carries_on(self, monkeypatch):
        from ppt_com import batch_apply as com_batch
        from ppt_mac.batch_apply import _batch_apply_impl

        monkeypatch.setattr(
            com_batch, "_dispatch_op", lambda *a: {"success": True},
        )

        with _fake_deck(["Box"]):
            result = _batch_apply_impl(
                1, ["Ghost", "Box"], [{"tool": "set_fill"}]
            )

        assert result["results"][0]["shape"] == "Ghost"
        assert "not found" in result["results"][0]["error"]
        assert result["results"][0]["operations"] == []
        assert result["results"][1]["operations"][0]["status"] == "success"


# ---------------------------------------------------------------------------
# A deck made of stand-ins, so none of this needs PowerPoint. Only the parts
# the two modules touch are modelled, and the collections behave the way
# PowerPoint's do, including answering -1728 for an empty one.
# ---------------------------------------------------------------------------
def _slide_view():
    from appscript import k

    return k.slide_view


class _Prop:
    """A property that remembers what was written to it."""

    def __init__(self, value=None):
        self.value = value
        self.writes: list = []
        self.clamp_to = None

    def __call__(self):
        return self.value

    def get(self):
        return self.value

    def set(self, value):
        self.writes.append(value)
        self.value = self.clamp_to if self.clamp_to is not None else value


def _count_call(holder, counter):
    """A no-argument command on a stand-in, which only records that it ran."""

    def _call():
        setattr(holder, counter, getattr(holder, counter) + 1)

    return _call


def _bag(**properties):
    """A stand-in for one of PowerPoint's little property-only classes."""
    holder = type("Bag", (), {})()
    for name, value in properties.items():
        setattr(holder, name, _Prop(value))
    return holder


class _FakeResponse:
    """What urlopen hands back, cut down to what the two tools read."""

    def __init__(self, body, content_type):
        self._body = body
        self.headers = {"Content-Type": content_type}

    def read(self):
        return self._body


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


class _FakeRange:
    """A selection's shape or slide range, which must never be materialised."""

    def __init__(self, items):
        self._items = list(items)
        self.materialised = 0

    def count(self, each=None):
        return len(self._items)

    def get(self):
        self.materialised += 1
        raise AssertionError("a selection range must not be materialised")

    @property
    def shapes(self):
        return _FakeList(self._items)

    @property
    def slides(self):
        return _FakeList(self._items)


class _FakeFont:
    def __init__(self, latin="Calibri", east_asian="Calibri"):
        self.font_name = _Prop(latin)
        self.east_asian_name = _Prop(east_asian)
        self.font_size = _Prop(None)
        self.bold = _Prop(None)
        self.italic = _Prop(None)
        self.font_color = _Prop(None)


class _FakeTextFrame:
    def __init__(self):
        self.text_range = type("TR", (), {})()
        self.text_range.font = _FakeFont()


class _FakeShape:
    def __init__(self, deck, name, shape_type=None):
        from appscript import k

        self._deck = deck
        self._name = _Prop(name)
        self._type = shape_type or k.shape_type_auto
        self.children: list = []
        self.lock_aspect_ratio = _Prop(False)
        self.width = _Prop(120.0)
        self.height = _Prop(60.0)
        self.left_position = _Prop(10.0)
        self.top = _Prop(20.0)
        self.auto_shape_type = _Prop(k.autoshape_rectangle)
        self.picture_format = _bag(
            brightness=0.5, contrast=0.5, color_type=k.picture_color_automatic,
            crop_left=0.0, crop_right=0.0, crop_top=0.0, crop_bottom=0.0,
            transparency_color=[0, 0, 0], transparent_background=False,
        )
        self.fill_format = _bag(visible=True, fore_color=None)
        self.fill_format.solid_calls = 0
        self.fill_format.solid = _count_call(self.fill_format, "solid_calls")
        self.line_format = _bag(fore_color=None, line_weight=1.0, transparency=0.0)
        self.text_frame = _FakeTextFrame()
        self.adjustments = _FakeList([_bag(adjustment_value=0.0)])
        self.exports: list = []
        self.export_writes = True
        # What the image measures at 100 per cent, which is what a scale
        # relative to the original size puts the shape back to.
        self.original = (240.0, 120.0)
        self.defaults_set = 0
        self.deleted = 0
        self.scaled: list = []

    def name(self):
        return self._name()

    def shape_type(self):
        return self._type

    def has_text_frame(self):
        return True

    def z_order_position(self):
        return self._deck.slide_shapes.index(self) + 1

    def set_fonts(self, latin, east_asian):
        self.text_frame.text_range.font.font_name = _Prop(latin)
        self.text_frame.text_range.font.east_asian_name = _Prop(east_asian)

    def save_as_picture(self, picture_type=None, file_name=None):
        self.exports.append(file_name)
        if self.export_writes:
            with open(file_name, "wb") as handle:
                handle.write(b"PNG")

    def scale_width(self, factor=None, relative_to_original_size=None, scale=None):
        self.scaled.append(("width", factor, relative_to_original_size))
        if relative_to_original_size:
            self.width.set(self.original[0] * factor)

    def scale_height(self, factor=None, relative_to_original_size=None, scale=None):
        self.scaled.append(("height", factor, relative_to_original_size))
        if relative_to_original_size:
            self.height.set(self.original[1] * factor)

    def set_shapes_default_properties(self):
        self.defaults_set += 1

    def delete(self):
        self.deleted += 1
        self._deck.slide_shapes.remove(self)
        self._deck.deleted_shapes.append(self)

    @property
    def shapes(self):
        return _FakeShapes(self.children)


class _FakeThemeFontScheme:
    """Major and minor families, four positions each, as macOS orders them."""

    def __init__(self):
        self.major = {i: _bag(name="Calibri") for i in (1, 2, 3, 4)}
        self.minor = {i: _bag(name="Calibri") for i in (1, 2, 3, 4)}

    @property
    def major_theme_fonts(self):
        return _FakeList([self.major[i] for i in (1, 2, 3, 4)])

    @property
    def minor_theme_fonts(self):
        return _FakeList([self.minor[i] for i in (1, 2, 3, 4)])


class _FakeDeck:
    """One slide holding one shape per name given, plus what `make` adds."""

    def __init__(self, shape_names):
        from appscript import k

        self.slide_shapes = [_FakeShape(self, name) for name in shape_names]
        self.font_names: list = []
        self.font_scheme = _FakeThemeFontScheme()
        self.make_type = k.shape_type_picture
        self.picture_paths: list = []
        self.picture_contents: list = []
        self.deleted_shapes: list = []
        self.no_slides = False
        self.selection = _bag(selection_type=k.selection_type_none)
        self.selection.shape_range = _FakeRange([])
        self.selection.slide_range = _FakeRange([])
        self.selection.text_range = _bag(content=None)
        self.window = _bag(view_type=k.slide_view)
        self.window.view = _bag(zoom=100)
        self.window.selection = self.selection
        self._slide = _FakeSlide(self)

    # -- what the tests reach for ------------------------------------------

    @property
    def slide_object(self):
        return self._slide

    def shape(self, name):
        for shape in self.slide_shapes:
            if shape.name() == name:
                return shape
        raise KeyError(name)

    def select_shapes(self, names):
        from appscript import k

        self.selection.selection_type = _Prop(k.selection_type_shapes)
        self.selection.shape_range = _FakeRange(
            [_FakeShape(self, name) for name in names]
        )

    def select_slides(self, indices):
        from appscript import k

        self.selection.selection_type = _Prop(k.selection_type_slides)
        self.selection.slide_range = _FakeRange(
            [_bag(slide_index=index) for index in indices]
        )

    def select_text(self, content):
        from appscript import k

        self.selection.selection_type = _Prop(k.selection_type_text)
        self.selection.text_range = _bag(content=content)

    # -- the object model --------------------------------------------------

    def make(self, new=None, at=None, with_properties=None):
        properties = with_properties or {}
        from appscript import k

        shape = _FakeShape(self, f"Shape_{len(self.slide_shapes) + 1}")
        if new == k.picture:
            shape._type = self.make_type
            path = properties[k.file_name]
            self.picture_paths.append(path)
            with open(path, "rb") as handle:
                self.picture_contents.append(handle.read().decode("utf-8", "replace"))
        self.slide_shapes.append(shape)
        return object()

    def count(self, each=None):
        from appscript import k

        if each == k.font:
            return len(self.font_names)
        if each == k.document_window:
            # What `target_window` asks before it hands the window over.
            return 0 if self.window is None else 1
        raise _command_error(-1708)

    @property
    def presentation(self):
        return self

    # `presentation` and the application are the same stand-in here, because
    # nothing in these two modules needs them to be different objects.
    @property
    def slides(self):
        return _FakeList([] if self.no_slides else [self._slide])

    @property
    def fonts(self):
        return _FakeList([_bag(font_name=name) for name in self.font_names])

    def __getattr__(self, name):
        # `pres.fonts[i]` and `pres.document_windows[1]` are elements rather
        # than properties, so they are answered here to keep the class short.
        if name == "document_windows":
            return _FakeList([self.window])
        raise AttributeError(name)

    @property
    def slide_master(self):
        scheme = self.font_scheme
        return type(
            "Master", (), {
                "theme": type(
                    "Theme", (), {"theme_font_scheme": scheme}
                )()
            }
        )()


class _FakeSlide:
    def __init__(self, deck):
        self._deck = deck
        self.end = object()
        self.slide_show_transition = _bag(hidden=False)

    @property
    def shapes(self):
        return _FakeShapes(self._deck.slide_shapes)


class _fake_deck:  # noqa: N801 - reads as a context manager, not a class
    """Point the wrapper at a fake deck for the length of a `with` block."""

    def __init__(self, shape_names):
        self._deck = _FakeDeck(shape_names)

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


class _fake_picture_deck(_fake_deck):  # noqa: N801 - a context manager
    """The same deck, with its one shape answering as a picture."""

    def __init__(self, name):
        super().__init__([name])
        from appscript import k

        self._deck.shape(name)._type = k.shape_type_picture


class _no_powerpoint:  # noqa: N801 - reads as a context manager, not a class
    """Make any approach to PowerPoint fail, so a refusal has to come first."""

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


# ---------------------------------------------------------------------------
# The icon search and the icon package have to agree
# ---------------------------------------------------------------------------
class TestIconSearchOffersOnlyWhatCanBeInserted:
    """Searching reads Google Fonts and inserting reads a pinned npm package.

    They do not hold the same names. `auto_awesome` is on the site and has never
    shipped in the package, so a search that offered it sent the caller to a 404.
    """

    def _stub(self, monkeypatch, available):
        from ppt_com import advanced_ops

        monkeypatch.setattr(advanced_ops, "_icon_on_cdn_cache", {})
        monkeypatch.setattr(
            advanced_ops, "_icon_is_on_cdn", lambda name: name in available
        )
        return advanced_ops

    def test_a_name_the_package_does_not_serve_is_dropped(self, monkeypatch):
        advanced_ops = self._stub(monkeypatch, {"star"})
        results = [{"name": "auto_awesome"}, {"name": "star"}]

        kept = advanced_ops._drop_what_cannot_be_inserted(results, 5)

        assert [icon["name"] for icon in kept] == ["star"]

    def test_the_next_match_takes_the_dropped_one_s_place(self, monkeypatch):
        advanced_ops = self._stub(monkeypatch, {"b", "c"})
        results = [{"name": "a"}, {"name": "b"}, {"name": "c"}]

        kept = advanced_ops._drop_what_cannot_be_inserted(results, 2)

        assert [icon["name"] for icon in kept] == ["b", "c"]

    def test_a_cdn_that_cannot_be_reached_filters_nothing(self, monkeypatch):
        """Not knowing is not the same answer as no."""
        import urllib.error

        from ppt_com import advanced_ops

        monkeypatch.setattr(advanced_ops, "_icon_on_cdn_cache", {})

        def unreachable(request, timeout=None):
            raise urllib.error.URLError("no network")

        monkeypatch.setattr(advanced_ops.urllib.request, "urlopen", unreachable)

        assert advanced_ops._icon_is_on_cdn("anything") is True
        assert advanced_ops._icon_on_cdn_cache == {}


@macos_only
class TestTheVisibilityStandInIsExplainedOnce:
    """A deck is built a shape at a time and the long form buried the rest."""

    def test_the_reason_is_given_once_and_then_only_the_fact(self, monkeypatch):
        from ppt_mac import shapes

        monkeypatch.setattr(shapes, "_line_visibility_explained", False)

        first = shapes._line_visibility_warning(False)
        second = shapes._line_visibility_warning(False)

        assert "no visible property" in first
        assert "no visible property" not in second
        assert "visible=False" in second


# Guarded because the checks below reach into `ppt_mac.text`, which imports
# appscript, and appscript only installs on macOS. CI runs on Windows, where an
# unguarded class fails at collection rather than skipping. The two classes in
# this file that are deliberately not guarded test `ppt_com` code that runs on
# both platforms.
@macos_only
class TestShapesThatWalkOffTheSlide:
    """A box set to grow with its text never overflows. It grows past the
    slide edge instead, and nothing used to say so."""

    class _Shape:
        def __init__(self, left, top, width, height):
            self._box = (left, top, width, height)

        def left_position(self):
            return self._box[0]

        def top(self):
            return self._box[1]

        def width(self):
            return self._box[2]

        def height(self):
            return self._box[3]

    def _edges(self, left, top, width, height):
        from ppt_mac.text import _off_slide_edges

        return _off_slide_edges(self._Shape(left, top, width, height), 960, 540)

    def test_a_shape_inside_the_slide_is_not_reported(self):
        assert self._edges(100, 100, 200, 100) == []

    def test_a_shape_flush_against_an_edge_is_not_reported(self):
        assert self._edges(0, 0, 960, 540) == []

    def test_the_edge_a_shape_hangs_over_is_named(self):
        assert self._edges(60, 60, 150, 530) == ["bottom"]
        assert self._edges(900, 100, 150, 100) == ["right"]
        assert self._edges(-20, -10, 100, 100) == ["left", "top"]

    def test_a_fraction_of_a_point_past_the_edge_is_rounding(self):
        assert self._edges(0, 0, 960.2, 540.2) == []

    def test_a_shape_that_will_not_answer_is_skipped(self):
        from ppt_mac.text import _off_slide_edges

        class Mute:
            def left_position(self):
                raise RuntimeError("no")

        assert _off_slide_edges(Mute(), 960, 540) == []

    def test_a_deck_that_will_not_say_its_size_reports_nothing(self):
        from ppt_mac.text import _off_slide_edges

        assert _off_slide_edges(self._Shape(0, 0, 9999, 9999), None, None) == []
