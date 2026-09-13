"""Tests for the macOS hyperlink, comment and media tools.

Pure unit tests over a stand-in object graph. Nothing here launches PowerPoint
or connects to it; the live behaviour is covered by MACOS_PORT.md and by
running the server. appscript only installs on macOS, so the whole file is
skipped elsewhere.

The three modules are together because they fail together. Each one leans on
something PowerPoint's dictionary does not really support, a command result
that has to be trusted, a class with no properties, and a media object that is
an element of nothing, so what is worth testing in all three is the refusal and
the read back rather than the happy path alone.
"""

import sys

import pytest

sys.path.insert(0, "src")

macos_only = pytest.mark.skipif(
    sys.platform != "darwin", reason="the Apple Event backend needs appscript"
)


# ---------------------------------------------------------------------------
# Hyperlinks
# ---------------------------------------------------------------------------
@macos_only
class TestHyperlinkEnums:
    """The two translations a hyperlink needs, and the table nobody generated."""

    def test_the_click_and_mouseover_words_come_from_the_generated_table(self):
        from appscript import k

        from backend.mac_enums import PpMouseActivation, to_keyword
        from ppt_com.hyperlinks import ACTION_ON_MAP

        assert to_keyword(
            PpMouseActivation, ACTION_ON_MAP["click"], "mouse activation"
        ) == k.mouse_activation_mouse_click
        assert to_keyword(
            PpMouseActivation, ACTION_ON_MAP["mouseover"], "mouse activation"
        ) == k.mouse_activation_mouse_over

    def test_hyperlink_types_translate_out_and_back(self):
        """constants.py has no msoHyperlink group, so this table lives in the port."""
        from appscript import k

        from ppt_mac.hyperlinks import MsoHyperlinkType, _windows_constant

        assert MsoHyperlinkType[0] == k.hyperlink_type_text_range
        assert MsoHyperlinkType[1] == k.hyperlink_type_shape
        assert MsoHyperlinkType[2] == k.hyperlink_type_inline_shape
        for value, word in MsoHyperlinkType.items():
            assert _windows_constant(MsoHyperlinkType, word) == value

    def test_the_raw_code_answers_when_the_name_does_not(self):
        """`hyperlink type` is declared mHyT while the enumeration is mHlT.

        The name route works anyway, because appscript names an enumerator
        from one table covering the whole dictionary and these three codes
        appear nowhere else in it. The code route is there in case a future
        build breaks that, and the low byte is the Windows number.
        """
        from ppt_mac.hyperlinks import _hyperlink_type

        class _Unnamed:
            code = b"\x00\x96\x00\x02"

        assert _hyperlink_type(_Unnamed()) == 2

    def test_a_word_the_table_does_not_know_reports_nothing(self):
        """`hyperlink type` is declared mHyT while the enumeration is mHlT.

        So an answer that maps to nothing is a real possibility here rather
        than a defensive branch, and a near miss would be worse than a null.
        """
        from appscript import k

        from ppt_mac.hyperlinks import (
            MsoHyperlinkType,
            _hyperlink_type,
            _windows_constant,
        )

        assert _windows_constant(MsoHyperlinkType, k.shape_type_auto) is None
        assert _hyperlink_type(k.shape_type_auto) is None

    def test_the_action_words_come_from_the_generated_table(self):
        from appscript import k

        from backend.mac_enums import PpActionType, to_keyword
        from ppt_com.constants import ppActionHyperlink, ppActionNone

        assert to_keyword(PpActionType, ppActionHyperlink, "action type") == (
            k.action_type_hyperlink_action
        )
        assert to_keyword(PpActionType, ppActionNone, "action type") == (
            k.action_type_none
        )


@macos_only
class TestHyperlinkValidation:
    """What the tools check before PowerPoint is asked for anything."""

    def test_an_unknown_trigger_names_the_ones_that_exist(self):
        from ppt_mac.hyperlinks import _add_hyperlink_impl, _remove_hyperlink_impl

        with pytest.raises(ValueError, match="click, mouseover"):
            _add_hyperlink_impl(1, "Title", "https://x", None, None, "hover")
        with pytest.raises(ValueError, match="click, mouseover"):
            _remove_hyperlink_impl(1, "Title", "hover")

    def test_a_trigger_is_matched_however_it_is_spelled(self):
        with _fake_deck() as deck:
            from ppt_mac.hyperlinks import _add_hyperlink_impl

            result = _add_hyperlink_impl(1, "Title", "https://x", None, None, " Click ")

        assert result["action_on"] == "click"
        assert deck.shape("Title").events == [1]   # position 1 is the click

    def test_a_shape_index_past_the_end_says_the_range(self):
        from ppt_mac.hyperlinks import _add_hyperlink_impl

        with _fake_deck():
            with pytest.raises(ValueError, match=r"out of range \(1-3\)"):
                _add_hyperlink_impl(1, 7, "https://x", None, None, "click")

    def test_a_shape_name_that_is_not_there_says_so(self):
        from ppt_mac.hyperlinks import _remove_hyperlink_impl

        with _fake_deck():
            with pytest.raises(ValueError, match="'Missing' not found"):
                _remove_hyperlink_impl(1, "Missing", "click")


@macos_only
class TestHyperlinkRefusals:
    """The two things that go wrong, and what each of them says."""

    def test_a_screen_tip_names_the_argument_not_the_tool(self):
        """The tool works; only that one argument has to go, and it says so."""
        from ppt_mac.hyperlinks import _add_hyperlink_impl

        payload = _add_hyperlink_impl(1, "Title", "https://x", None, "Open it", "click")

        assert payload["error"] == "ppt_add_hyperlink cannot set screen_tip on macOS"
        assert "is not available" not in payload["error"]
        assert payload["platform"] == "macOS"
        assert "no screen tip" in payload["reason"]
        assert payload["alternatives"] == ["ppt_add_hyperlink without screen_tip"]

    def test_a_refused_screen_tip_never_reaches_powerpoint(self):
        """A refused call is answered before the view is moved."""
        from backend.mac_ae import ppt
        from ppt_mac.hyperlinks import _add_hyperlink_impl

        def _explode(*args, **kwargs):
            raise AssertionError("a refusal must not touch PowerPoint")

        original = ppt._get_app_impl
        ppt._get_app_impl = _explode
        try:
            payload = _add_hyperlink_impl(1, "Title", "https://x", None, "tip", "click")
        finally:
            ppt._get_app_impl = original
        assert "screen_tip" in payload["error"]

    def test_an_action_setting_that_does_not_resolve_says_which_position(self):
        """The route is built here, and it can still fail on an odd shape."""
        from ppt_mac.hyperlinks import _add_hyperlink_impl, _remove_hyperlink_impl

        with _fake_deck(action_setting_error=-1728):
            add = _add_hyperlink_impl(1, "Title", "https://x", None, None, "click")
            remove = _remove_hyperlink_impl(1, "Title", "click")

        for payload in (add, remove):
            assert payload["platform"] == "macOS"
            assert "`action settings` position 1" in payload["reason"]
            assert "built rather than asked for" in payload["reason"]
            assert "-1728" in payload["reason"]
        assert add["error"] == "ppt_add_hyperlink is not available on macOS"
        assert remove["error"] == "ppt_remove_hyperlink is not available on macOS"

    def test_an_address_that_did_not_land_is_reported_rather_than_claimed(self):
        from ppt_mac.hyperlinks import _add_hyperlink_impl

        with _fake_deck(address_is_frozen=True):
            payload = _add_hyperlink_impl(1, "Title", "https://x", None, None, "click")

        assert "silent no-op" in payload["reason"]
        assert "success" not in payload

    def test_an_action_that_stays_a_hyperlink_is_reported(self):
        from ppt_mac.hyperlinks import _remove_hyperlink_impl

        with _fake_deck(action_is_frozen=True):
            payload = _remove_hyperlink_impl(1, "Title", "click")

        assert "silent no-op" in payload["reason"]
        assert "success" not in payload


@macos_only
class TestHyperlinkWrites:
    """Adding and removing, and the read back that proves each one landed."""

    def test_the_action_is_set_before_the_address(self):
        """The address does not hold unless the action already says hyperlink."""
        from ppt_mac.hyperlinks import _add_hyperlink_impl

        with _fake_deck() as deck:
            _add_hyperlink_impl(1, "Title", "https://x", "3,,", None, "click")

        assert deck.journal == [
            ("action", "k.action_type_hyperlink_action"),
            ("hyperlink address", "https://x"),
            ("hyperlink sub address", "3,,"),
        ]

    def test_the_answer_is_read_back_rather_than_echoed(self):
        """PowerPoint completes an address it thinks is partial."""
        from ppt_mac.hyperlinks import _add_hyperlink_impl

        with _fake_deck(address_completed_to="https://www.x.com/"):
            result = _add_hyperlink_impl(1, "Title", "www.x.com", None, None, "click")

        assert result == {
            "success": True,
            "shape_name": "Title",
            "address": "https://www.x.com/",
            "sub_address": None,
            "action_on": "click",
        }

    def test_removing_accepts_the_word_an_untouched_shape_reads_back_as(self):
        """A shape that never carried an action says `unset`, not `none`."""
        from ppt_mac.hyperlinks import _remove_hyperlink_impl

        with _fake_deck(action_reads_back_unset=True) as deck:
            result = _remove_hyperlink_impl(1, "Body", "mouseover")

        assert result == {
            "success": True,
            "shape_name": "Body",
            "action_on": "mouseover",
        }
        assert deck.shape("Body").events == [2]   # position 2 is the mouse over

    def test_removing_clears_the_action_it_was_asked_for(self):
        from ppt_mac.hyperlinks import _remove_hyperlink_impl

        with _fake_deck() as deck:
            result = _remove_hyperlink_impl(1, "Title", "click")

        assert result["success"] is True
        assert deck.journal == [("action", "k.action_type_none")]


@macos_only
class TestHyperlinkListing:
    """A slide really does hold hyperlinks, so this one is a plain walk."""

    def test_every_link_is_reported_with_its_windows_type_number(self):
        from ppt_mac.hyperlinks import _get_hyperlinks_impl

        with _fake_deck():
            result = _get_hyperlinks_impl(1)

        assert result["success"] is True
        assert result["slide_index"] == 1
        assert result["hyperlinks_count"] == 2
        assert result["hyperlinks"][0] == {
            "index": 1,
            "address": "https://example.com",
            "sub_address": None,
            "type": 1,
        }
        assert result["hyperlinks"][1] == {
            "index": 2,
            "address": None,
            "sub_address": "Slide 3",
            "type": 0,
        }

    def test_a_slide_with_no_links_answers_with_none_rather_than_failing(self):
        """PowerPoint raises -1728 for an empty collection instead of an empty list."""
        from ppt_mac.hyperlinks import _get_hyperlinks_impl

        with _fake_deck(links=[]):
            result = _get_hyperlinks_impl(1)

        assert result["hyperlinks_count"] == 0
        assert result["hyperlinks"] == []


# ---------------------------------------------------------------------------
# Comments
# ---------------------------------------------------------------------------
@macos_only
class TestCommentListing:
    """A comment is a shape, so the comments are found among the shapes."""

    def test_only_comment_shapes_are_listed(self):
        from ppt_mac.comments import _list_comments_impl

        with _fake_deck():
            result = _list_comments_impl(1)

        assert result["comments_count"] == 1
        assert [c["text"] for c in result["comments"]] == ["Check this number"]

    def test_what_macos_cannot_read_is_empty_rather_than_guessed(self):
        from ppt_mac.comments import _list_comments_impl

        with _fake_deck():
            comment = _list_comments_impl(1)["comments"][0]

        assert comment["author"] is None
        assert comment["author_initials"] is None
        # The Windows side answers with an empty string when the date will not
        # read, so this matches it rather than inventing a null.
        assert comment["datetime"] == ""
        assert comment["left"] == 40.0
        assert comment["top"] == 12.5

    def test_the_answer_says_why_those_three_are_empty(self):
        from ppt_mac.comments import _list_comments_impl

        with _fake_deck():
            result = _list_comments_impl(1)

        assert "no author" in result["note"]

    def test_a_slide_with_no_comments_is_not_an_error(self):
        from ppt_mac.comments import _list_comments_impl

        with _fake_deck(with_comment=False):
            result = _list_comments_impl(1)

        assert result["comments_count"] == 0
        assert result["comments"] == []


@macos_only
class TestCommentAdding:
    """`make new comment` has no declared place to go, so it is verified."""

    def test_a_comment_that_arrives_is_written_and_read_back(self):
        from ppt_mac.comments import _add_comment_impl

        with _fake_deck(make="comment") as deck:
            result = _add_comment_impl(1, "Fix the total", "Ada", "AL", 20.0, 30.0)

        assert result["success"] is True
        assert result["text"] == "Fix the total"
        assert deck.slide.shapes_list[-1].left_position() == 20.0

    def test_author_and_initials_are_reported_as_dropped_not_echoed(self):
        from ppt_mac.comments import _add_comment_impl

        with _fake_deck(make="comment"):
            result = _add_comment_impl(1, "Fix the total", "Ada", "AL", 0, 0)

        assert result["author"] is None
        assert "author and author_initials were ignored" in result["warnings"][0]
        assert "no initials" in result["warnings"][0]

    def test_a_make_that_does_nothing_is_not_reported_as_success(self):
        from ppt_mac.comments import _add_comment_impl

        with _fake_deck(make="nothing") as deck:
            result = _add_comment_impl(1, "Fix the total", "Ada", "AL", 0, 0)

        assert "silent no-op" in result["reason"]
        assert "nothing was left behind" in result["reason"]
        assert len(deck.slide.shapes_list) == 3

    def test_a_stray_shape_is_deleted_again_rather_than_left_on_the_slide(self):
        """A `make` that falls through leaves an empty autoshape behind."""
        from ppt_mac.comments import _add_comment_impl

        with _fake_deck(make="stray") as deck:
            result = _add_comment_impl(1, "Fix the total", "Ada", "AL", 0, 0)

        assert "made a shape rather than a comment" in result["reason"]
        assert "deleted again" in result["reason"]
        assert len(deck.slide.shapes_list) == 3

    def test_a_stray_shape_that_will_not_go_names_it_for_the_user(self):
        from ppt_mac.comments import _add_comment_impl

        with _fake_deck(make="stray", delete_is_frozen=True) as deck:
            result = _add_comment_impl(1, "Fix the total", "Ada", "AL", 0, 0)

        assert "'Rectangle 9'" in result["reason"]
        assert "by hand" in result["reason"]
        assert len(deck.slide.shapes_list) == 4

    def test_a_make_powerpoint_refuses_says_what_the_dictionary_says(self):
        from ppt_mac.comments import _add_comment_impl

        with _fake_deck(make="error"):
            result = _add_comment_impl(1, "Fix the total", "Ada", "AL", 0, 0)

        assert result["error"] == "ppt_add_comment is not available on macOS"
        assert "-1708" in result["reason"]
        assert "element of `shape` and of nothing else" in result["reason"]
        assert result["alternatives"] == ["ppt_set_slide_notes", "ppt_add_textbox"]


@macos_only
class TestCommentDeleting:
    """The index is a position among the comment shapes, in z order."""

    def test_deleting_counts_before_and_after(self):
        from ppt_mac.comments import _delete_comment_impl

        with _fake_deck() as deck:
            result = _delete_comment_impl(1, 1)

        assert result == {"success": True}
        assert len(deck.slide.shapes_list) == 2

    def test_an_index_past_the_end_says_how_many_there_are(self):
        from ppt_mac.comments import _delete_comment_impl

        with _fake_deck():
            with pytest.raises(ValueError, match=r"out of range \(1-1\)"):
                _delete_comment_impl(1, 4)

    def test_a_delete_that_did_nothing_is_reported(self):
        from ppt_mac.comments import _delete_comment_impl

        with _fake_deck(delete_is_frozen=True) as deck:
            result = _delete_comment_impl(1, 1)

        assert "still holds the same number of shapes" in result["reason"]
        assert len(deck.slide.shapes_list) == 3


# ---------------------------------------------------------------------------
# Media
# ---------------------------------------------------------------------------
@macos_only
class TestMediaInsertion:
    """`media2 object` is an element of nothing and `make` accepts it anyway.

    Run on this machine before any of this was written: an `.aiff` and an
    `.mp4` both arrive on the slide as `shape type media`, and both are inside
    the saved `.pptx` under `ppt/media/` at the file's own byte size. The
    module used to refuse all three media tools on the grounds that nothing in
    the dictionary could put a file on a slide, which was a reading of the
    dictionary rather than a measurement.
    """

    def test_a_movie_arrives_as_a_media_shape(self, monkeypatch):
        from ppt_mac.media import _add_video_impl

        with _fake_media_deck(monkeypatch) as deck:
            clip = deck.a_file("intro.mp4")
            result = _add_video_impl(1, clip, 60, 40, None, None, False)

        assert result["success"] is True
        assert result["shape_name"] == "intro"
        # The caller's own path, not the container copy it was read from.
        assert result["file_path"] == clip
        assert deck.made == [("media2 object", 60, 40)]

    def test_the_view_is_sent_to_the_slide_before_the_insert(self, monkeypatch):
        """Every write here shows the slide it is editing, and this is a write."""
        from ppt_mac.media import _add_audio_impl

        with _fake_media_deck(monkeypatch) as deck:
            _add_audio_impl(2, deck.a_file("bell.aiff"), 0, 0, None, None, False)

        assert deck.navigated_to == [2]

    def test_the_file_is_read_from_inside_powerpoints_container(self, monkeypatch):
        """A path outside it raises the Grant Access sheet.

        In #191 that sheet closed every open document and took PowerPoint with
        it, so nothing is ever named where the caller left it.
        """
        from ppt_mac.media import _add_video_impl

        with _fake_media_deck(monkeypatch) as deck:
            clip = deck.a_file("intro.mp4")
            _add_video_impl(1, clip, 0, 0, None, None, False)

        handed_over = deck.files_named[0]
        assert handed_over != clip
        assert handed_over.startswith(str(deck.container))

    def test_the_staged_copy_is_removed_once_the_shape_exists(self, monkeypatch):
        """The embed happens at insert time, not at save time.

        Checked rather than assumed, because a copy deleted too early would
        have produced a deck that lost its audio when it was saved. The two
        staged files were deleted, the deck was saved afterwards, and both
        media parts were in the archive whole.
        """
        import os

        from ppt_mac.media import _add_audio_impl

        with _fake_media_deck(monkeypatch) as deck:
            source = deck.a_file("bell.aiff")
            _add_audio_impl(1, source, 0, 0, None, None, False)
            staged = deck.files_named[0]
            # It existed while PowerPoint was reading it, and is gone by the
            # time the call returns. The caller's own file is untouched.
            assert deck.existed_when_named == [True]
            assert not os.path.exists(staged)
            assert os.path.exists(source)

    def test_a_make_that_added_nothing_is_refused(self, monkeypatch):
        """MACOS_PORT section 5. Success and an unchanged slide arrive alike."""
        from ppt_mac.media import _add_video_impl

        with _fake_media_deck(monkeypatch, make="nothing") as deck:
            result = _add_video_impl(1, deck.a_file("intro.mp4"), 0, 0, None, None, False)

        assert "error" in result
        assert "still holds" in result["reason"]

    def test_a_make_that_left_an_autoshape_is_refused(self, monkeypatch):
        """A declined `make` comes back as an empty autoshape and says nothing."""
        from ppt_mac.media import _add_video_impl

        with _fake_media_deck(monkeypatch, make="stray") as deck:
            with pytest.raises(RuntimeError, match="rather than a"):
                _add_video_impl(1, deck.a_file("intro.mp4"), 0, 0, None, None, False)

    def test_the_default_25_point_height_is_reported(self, monkeypatch):
        """PowerPoint sizes media at 25 points tall, not at its native size.

        A 320 by 240 movie arrives 33.3 by 25 and a 640 by 360 one 44.4 by 25,
        so the aspect ratio is the file's and the height is always 25. Windows
        would have used the native size, so the difference is said out loud.
        """
        from ppt_mac.media import _add_video_impl

        with _fake_media_deck(monkeypatch, landed_size=(44.44, 25.0)) as deck:
            result = _add_video_impl(1, deck.a_file("intro.mp4"), 0, 0, None, None, False)

        assert any("44.44 by 25.0" in w for w in result["warnings"])
        assert any("native size" in w for w in result["warnings"])

    def test_a_size_that_was_asked_for_is_read_back_not_echoed(self, monkeypatch):
        from ppt_mac.media import _add_video_impl

        with _fake_media_deck(monkeypatch) as deck:
            result = _add_video_impl(1, deck.a_file("intro.mp4"), 0, 0, 320, 180, False)

        assert deck.shape.width() == 320
        assert deck.shape.height() == 180
        assert deck.shape.lock_aspect_ratio() is False
        assert "warnings" not in result

    def test_one_dimension_keeps_the_aspect_ratio(self, monkeypatch):
        from ppt_mac.media import _add_video_impl

        with _fake_media_deck(monkeypatch) as deck:
            _add_video_impl(1, deck.a_file("intro.mp4"), 0, 0, 320, None, False)

        assert deck.shape.lock_aspect_ratio() is True
        assert deck.shape.width() == 320

    def test_a_size_that_did_not_land_is_reported(self, monkeypatch):
        from ppt_mac.media import _add_video_impl

        with _fake_media_deck(monkeypatch, size_is_frozen=True) as deck:
            result = _add_video_impl(1, deck.a_file("intro.mp4"), 0, 0, 320, 180, False)

        assert any("320 was asked for" in w for w in result["warnings"])
        assert any("180 was asked for" in w for w in result["warnings"])

    def test_a_second_clip_of_the_same_name_is_warned_about(self, monkeypatch):
        """PowerPoint names a media shape after its file and does not number it.

        Two clips called `intro.mp4` from two folders both arrive as `intro`,
        and every tool that addresses a shape by name then acts on the first.
        """
        from ppt_mac.media import _add_video_impl

        with _fake_media_deck(monkeypatch, existing=["intro"]) as deck:
            result = _add_video_impl(1, deck.a_file("intro.mp4"), 0, 0, None, None, False)

        assert any("already had a shape called 'intro'" in w
                   for w in result["warnings"])

    def test_link_to_file_is_refused_by_name_before_anything_is_copied(self, monkeypatch):
        """`link to file` reads back `missing value`, so embedding is the only mode.

        Named as an argument rather than as the tool, so dropping it and
        calling again works, and refused before the file is staged so that a
        refused call leaves the container alone.
        """
        from ppt_mac.media import _add_audio_impl, _add_video_impl

        with _fake_media_deck(monkeypatch) as deck:
            clip = deck.a_file("intro.mp4")
            video = _add_video_impl(1, clip, 0, 0, None, None, True)
            audio = _add_audio_impl(1, clip, 0, 0, None, None, True)

        assert video["error"] == "ppt_add_video cannot link to a file on macOS"
        assert audio["error"] == "ppt_add_audio cannot link to a file on macOS"
        assert deck.files_named == []
        assert deck.made == []

    def test_a_file_that_is_not_there_is_named_before_powerpoint_is_touched(self, monkeypatch):
        from ppt_mac.media import _add_audio_impl

        with _fake_media_deck(monkeypatch) as deck:
            with pytest.raises(FileNotFoundError, match="Audio file not found"):
                _add_audio_impl(
                    1, str(deck.container / "nothing.aiff"), 0, 0, None, None, False,
                )
            assert deck.navigated_to == []


@macos_only
class TestMediaPlaybackSettings:
    """Two of the eight settings exist here, and only one of them is safe.

    Volume, mute, trim and fade have no words anywhere in the dictionary.
    `loop until stopped` and `hide while not playing` both read and write, and
    the difference between them was measured on a slide holding a fly in, a
    bounce and an exit fade: three writes of loop left all three effects
    exactly as they were, and one write of hide dropped the exit outright and
    turned the bounce into a plain appear.
    """

    def test_volume_and_the_rest_are_refused_by_name_and_nothing_is_written(self, monkeypatch):
        from ppt_mac.media import _set_media_settings_impl

        with _fake_media_deck(monkeypatch, existing=["clip"]) as deck:
            result = _set_media_settings_impl(
                1, "clip", 0.5, True, None, None, None, 200, True, None,
            )

        assert result["error"] == (
            "ppt_set_media_settings cannot set volume, muted, fade_out on macOS"
        )
        assert "no volume, mute, trim or fade" in result["reason"]
        # Refused whole. A call that set loop and dropped the volume would be
        # the silent half success MACOS_PORT section 5 is about.
        assert deck.journal == []

    def test_loop_is_written_and_read_back(self, monkeypatch):
        from ppt_mac.media import _set_media_settings_impl

        with _fake_media_deck(monkeypatch, existing=["clip"]) as deck:
            result = _set_media_settings_impl(
                1, "clip", None, None, None, None, None, None, True, None,
            )

        assert result == {"success": True, "shape_name": "clip"}
        assert deck.play.loop_until_stopped() is True

    def test_a_loop_that_did_not_land_is_reported(self, monkeypatch):
        from ppt_mac.media import _set_media_settings_impl

        with _fake_media_deck(monkeypatch, existing=["clip"], playback_is_frozen=True):
            result = _set_media_settings_impl(
                1, "clip", None, None, None, None, None, None, True, None,
            )

        assert result["error"] == "ppt_set_media_settings could not set loop"
        assert "silent no-op" in result["reason"]

    def test_hiding_the_frame_lands_on_a_slide_with_no_animations(self, monkeypatch):
        from ppt_mac.media import _set_media_settings_impl

        with _fake_media_deck(monkeypatch, existing=["clip"], effects=0) as deck:
            result = _set_media_settings_impl(
                1, "clip", None, None, None, None, None, None, None, True,
            )

        assert result["success"] is True
        assert deck.play.hide_while_not_playing() is True

    def test_hiding_the_frame_is_refused_on_a_slide_that_has_animations(self, monkeypatch):
        """One write of it dropped an exit effect and flattened a bounce.

        Measured on a slide holding three effects, and reproduced: writing
        `False` over a value that was already `False` did it too. That is the
        flattening of MACOS_PORT section 5.2, so the argument is refused by
        name rather than paid for with the slide.
        """
        from ppt_mac.media import _set_media_settings_impl

        with _fake_media_deck(monkeypatch, existing=["clip"], effects=3) as deck:
            result = _set_media_settings_impl(
                1, "clip", None, None, None, None, None, None, True, True,
            )

        assert result["error"] == (
            "ppt_set_media_settings cannot set hide_while_not_playing on "
            "slide 1 because it has animations"
        )
        assert "holds 3 animation effects" in result["reason"]
        assert "5.2" in result["reason"]
        # Refused before the loop write, so the whole call is one decision.
        assert deck.journal == []

    def test_the_timelines_own_sequences_are_never_counted(self, monkeypatch):
        """Counting those looked free and is not, so the gate must not use them.

        A media shape creates one `sequence` on the timeline for its own
        playback. The count answered 0 on a fresh slide, 0 after a text box, 0
        with one effect in the main sequence, and 1 as soon as a movie was
        inserted, so a gate that added it in would refuse the write on every
        slide this tool is ever called about. The stand-in below raises if it
        is ever asked.
        """
        from ppt_mac.media import _set_media_settings_impl

        with _fake_media_deck(monkeypatch, existing=["clip"], effects=0) as deck:
            result = _set_media_settings_impl(
                1, "clip", None, None, None, None, None, None, None, True,
            )

        assert result["success"] is True
        assert deck.play.hide_while_not_playing() is True

    def test_the_hide_that_did_land_says_what_the_count_cannot_see(self, monkeypatch):
        """Degrade honestly. The guard has an edge and the caller is told."""
        from ppt_mac.media import _set_media_settings_impl

        with _fake_media_deck(monkeypatch, existing=["clip"], effects=0):
            result = _set_media_settings_impl(
                1, "clip", None, None, None, None, None, None, None, True,
            )

        assert any("triggered by clicking a shape" in w
                   for w in result["warnings"])

    def test_effects_that_cannot_be_counted_are_treated_as_animations(self, monkeypatch):
        """The count guards a write that destroys what it cannot see."""
        from ppt_mac.media import _set_media_settings_impl

        with _fake_media_deck(monkeypatch, existing=["clip"], effects=None) as deck:
            result = _set_media_settings_impl(
                1, "clip", None, None, None, None, None, None, None, True,
            )

        assert "could not be counted" in result["reason"]
        assert deck.journal == []

    def test_loop_is_written_on_an_animated_slide(self, monkeypatch):
        """Three writes of it left a slide of three effects untouched."""
        from ppt_mac.media import _set_media_settings_impl

        with _fake_media_deck(monkeypatch, existing=["clip"], effects=3) as deck:
            result = _set_media_settings_impl(
                1, "clip", None, None, None, None, None, None, False, None,
            )

        assert result["success"] is True
        assert deck.play.loop_until_stopped() is False

    def test_a_shape_that_is_not_media_is_named_before_the_first_write(self, monkeypatch):
        """`play settings` exists on every shape, so a picture would take it.

        It would report success and change nothing anybody could ever see.
        """
        from ppt_mac.media import _set_media_settings_impl

        with _fake_media_deck(monkeypatch, existing=["clip"]) as deck:
            with pytest.raises(ValueError, match="is not a media shape"):
                _set_media_settings_impl(
                    1, "Title", None, None, None, None, None, None, True, None,
                )
            assert deck.journal == []

    def test_a_call_that_asked_for_nothing_says_so(self, monkeypatch):
        from ppt_mac.media import _set_media_settings_impl

        with _fake_media_deck(monkeypatch, existing=["clip"]):
            result = _set_media_settings_impl(
                1, "clip", None, None, None, None, None, None, None, None,
            )

        assert "nothing was changed" in result["warnings"][0]


@macos_only
class TestSwapIsWiredUp:
    """A module that defines nothing here keeps its COM version silently."""

    def test_every_impl_the_windows_module_has_is_replaced(self):
        import ppt_com.comments as com_comments
        import ppt_com.hyperlinks as com_hyperlinks
        import ppt_com.media as com_media
        import ppt_mac.comments as mac_comments
        import ppt_mac.hyperlinks as mac_hyperlinks
        import ppt_mac.media as mac_media

        for com, mac in (
            (com_hyperlinks, mac_hyperlinks),
            (com_comments, mac_comments),
            (com_media, mac_media),
        ):
            names = [n for n in dir(com) if n.startswith("_") and n.endswith("_impl")]
            assert names
            for name in names:
                assert getattr(com, name) is getattr(mac, name), name


# ---------------------------------------------------------------------------
# A slide made of stand-ins, for the decisions worth testing without
# PowerPoint. Only the parts these three modules touch are modelled, and every
# property records what it was asked to store so that a write that would have
# been silently lost is visible from the outside.
# ---------------------------------------------------------------------------
class _FakeCollection:
    def __init__(self, items):
        self._items = items

    def get(self):
        if not self._items:
            # PowerPoint raises for an empty collection rather than answering
            # with an empty list, which is the quirk `elements` absorbs.
            raise _command_error(-1728)
        return list(self._items)

    def __getitem__(self, index):
        if index < 1 or index > len(self._items):
            raise _command_error(-1728)
        return self._items[index - 1]


class _FakeProperty:
    def __init__(self, value=None, journal=None, label=None, frozen=False, completed_to=None):
        self._value = value
        self._journal = journal
        self._label = label
        self._frozen = frozen
        self._completed_to = completed_to

    def __call__(self):
        return self._value

    def get(self):
        return self._value

    def set(self, value):
        if self._journal is not None:
            self._journal.append((self._label, str(value)))
        if self._frozen:
            return
        # PowerPoint rewrites an address it considers partial, which is why the
        # tools report the read back rather than what they were given.
        self._value = self._completed_to if self._completed_to is not None else value


class _FakeHyperlink:
    def __init__(self, address, sub_address, type_word):
        from appscript import k

        self.hyperlink_address = _FakeProperty(
            address if address is not None else k.missing_value
        )
        self.hyperlink_sub_address = _FakeProperty(
            sub_address if sub_address is not None else k.missing_value
        )
        self._type = type_word

    def hyperlink_type(self):
        return self._type


class _FakeActionSetting:
    def __init__(self, deck):
        from appscript import k

        # A shape that already carries a link is what a frozen action stands
        # for, since the point of that case is a clear that does not clear.
        if deck.action_is_frozen:
            starts_as = k.action_type_hyperlink_action
        elif deck.action_reads_back_unset:
            starts_as = k.action_type_unset
        else:
            starts_as = k.action_type_none
        self.action = _FakeProperty(
            starts_as,
            journal=deck.journal,
            label="action",
            frozen=deck.action_is_frozen,
        )
        hyperlink = _FakeHyperlink(None, None, k.hyperlink_type_shape)
        hyperlink.hyperlink_address = _FakeProperty(
            "", journal=deck.journal, label="hyperlink address",
            frozen=deck.address_is_frozen, completed_to=deck.address_completed_to,
        )
        hyperlink.hyperlink_sub_address = _FakeProperty(
            "", journal=deck.journal, label="hyperlink sub address",
        )
        self.hyperlink = hyperlink


class _FakeTextRange:
    def __init__(self, text):
        self.content = _FakeProperty(text)


class _FakeShape:
    def __init__(self, deck, name, shape_type, text="", left=0.0, top=0.0):
        self._deck = deck
        self._name = name
        self._shape_type = shape_type
        self._text = text
        self.text_frame = type("TextFrame", (), {})()
        self.text_frame.text_range = _FakeTextRange(text)
        self.left_position = _FakeProperty(left)
        self.top = _FakeProperty(top)
        self.events: list = []
        self._settings: dict = {}

    def name(self):
        return self._name

    def shape_type(self):
        return self._shape_type

    def has_text_frame(self):
        return True

    @property
    def action_settings(self):
        """Reached by position, never through `get action setting for`.

        The command answers with a reference that will not resolve, so the
        port builds this one itself. Position 1 is the click, position 2 is
        the mouse over.
        """
        shape = self

        class _Settings:
            def __getitem__(self, position):
                if shape._deck.action_setting_error is not None:
                    raise _command_error(shape._deck.action_setting_error)
                shape.events.append(position)
                return shape._setting(position)

        return _Settings()

    def _setting(self, position):
        if position not in self._settings:
            self._settings[position] = _FakeActionSetting(self._deck)
        return self._settings[position]

    def delete(self):
        if self._deck.delete_is_frozen:
            return
        self._deck.slide.shapes_list.remove(self)


class _FakeSlide:
    def __init__(self, shapes, links):
        self.shapes_list = shapes
        self.links = links
        self.end = "the end of the slide"

    @property
    def shapes(self):
        return _FakeCollection(self.shapes_list)

    @property
    def hyperlinks(self):
        return _FakeCollection(self.links)


class _FakeApp:
    def __init__(self, deck):
        self._deck = deck

    def make(self, new=None, at=None, with_properties=None):
        from appscript import k

        behaviour = self._deck.make
        if behaviour == "error":
            raise _command_error(-1708)
        if behaviour == "comment":
            self._deck.slide.shapes_list.append(
                _FakeShape(self._deck, "Comment 9", k.shape_type_comment)
            )
        elif behaviour == "stray":
            # The fallen through `make` from MACOS_PORT section 5, an empty
            # autoshape where the caller asked for something else.
            self._deck.slide.shapes_list.append(
                _FakeShape(self._deck, "Rectangle 9", k.shape_type_auto)
            )


class _FakeDeck:
    """One slide holding a title, a body and, by default, one comment."""

    def __init__(
        self, links=None, with_comment=True, make="nothing",
        action_setting_error=None, action_is_frozen=False,
        address_is_frozen=False, address_completed_to=None,
        action_reads_back_unset=False, delete_is_frozen=False,
    ):
        from appscript import k

        self.journal: list = []
        self.make = make
        self.action_setting_error = action_setting_error
        self.action_is_frozen = action_is_frozen
        self.address_is_frozen = address_is_frozen
        self.address_completed_to = address_completed_to
        self.action_reads_back_unset = action_reads_back_unset
        self.delete_is_frozen = delete_is_frozen

        shapes = [
            _FakeShape(self, "Title", k.shape_type_text_box),
            _FakeShape(self, "Body", k.shape_type_auto),
        ]
        if with_comment:
            shapes.append(
                _FakeShape(
                    self, "Comment 1", k.shape_type_comment,
                    text="Check this number", left=40.0, top=12.5,
                )
            )
        if links is None:
            links = [
                _FakeHyperlink(
                    "https://example.com", None, k.hyperlink_type_shape
                ),
                _FakeHyperlink("", "Slide 3", k.hyperlink_type_text_range),
            ]
        self.slide = _FakeSlide(shapes, links)
        self.app = _FakeApp(self)

    def shape(self, name):
        for shape in self.slide.shapes_list:
            if shape.name() == name:
                return shape
        raise KeyError(name)

    @property
    def presentation(self):
        return type("Pres", (), {"slides": _FakeCollection([self.slide])})()


class _fake_deck:  # noqa: N801 - reads as a context manager, not a class
    """Point the wrapper at a fake slide for the length of a `with` block."""

    def __init__(self, **kwargs):
        self._deck = _FakeDeck(**kwargs)

    def __enter__(self):
        from backend.mac_ae import ppt

        self._ppt = ppt
        self._app = ppt._get_app_impl
        self._pres = ppt._get_pres_impl
        ppt._get_app_impl = lambda *a, **kw: self._deck.app
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


# ---------------------------------------------------------------------------
# A slide that holds media, for the tools that put a file on one. Separate from
# the deck above because these three need things the other two modules never
# touch: a container to stage into, a shape that appears only when `make` is
# answered, and an animation timeline whose count decides whether one of the
# writes is allowed at all.
# ---------------------------------------------------------------------------
class _FakePlaySettings:
    """`play settings`, reached only through `shape animation settings`."""

    def __init__(self, journal, frozen):
        self.loop_until_stopped = _FakeProperty(
            False, journal=journal, label="loop until stopped", frozen=frozen,
        )
        self.hide_while_not_playing = _FakeProperty(
            False, journal=journal, label="hide while not playing", frozen=frozen,
        )


class _FakeMediaShape:
    def __init__(self, deck, name, shape_type, size=(44.44, 25.0), frozen=False):
        self._name = name
        self._shape_type = shape_type
        self.width = _FakeProperty(size[0], frozen=frozen)
        self.height = _FakeProperty(size[1], frozen=frozen)
        self.lock_aspect_ratio = _FakeProperty(False)
        self.animation_settings = type(
            "AnimationSettings", (),
            {"animation_play_settings": _FakePlaySettings(
                deck.journal, deck.playback_is_frozen,
            )},
        )()

    def name(self):
        return self._name

    def shape_type(self):
        return self._shape_type


class _FakeShapesCollection(_FakeCollection):
    """A slide's shapes, which also answer a bulk read of their names.

    `elements(slide.shapes.name)` is one Apple Event for every name, which is
    how the port reads them, so the stand-in has to offer the same route.
    """

    @property
    def name(self):
        return _FakeCollection([shape.name() for shape in self._items])


class _FakeTimeline:
    """A slide's animation timeline, counted the one way that is safe and right.

    `effects.count()` and `effects.get()` both kill PowerPoint with -609.
    Asking the sequence how many `effect` elements it holds does not, so that
    is the only question this answers. Asking the timeline how many `sequence`
    elements it holds is safe too and is the wrong question, because a media
    shape brings one of those with it, so it raises here rather than answering.
    """

    def __init__(self, effects, sequences=0):
        self._effects = effects
        self._sequences = sequences

    def count(self, each=None):
        raise AssertionError(
            "the timeline's sequences are not what decides; a media shape "
            "creates one of them for its own playback"
        )

    @property
    def main_sequence(self):
        sequence = self

        class _Sequence:
            def count(self, each=None):
                from appscript import k

                assert each == k.effect, "the effects collection is never asked"
                if sequence._effects is None:
                    raise _command_error(-1728)
                return sequence._effects

        return _Sequence()


class _FakeMediaSlide:
    def __init__(self, deck, shapes, effects, sequences=0):
        self._deck = deck
        self.shapes_list = shapes
        self.end = "the end of the slide"
        self.timeline = _FakeTimeline(effects, sequences)

    @property
    def shapes(self):
        return _FakeShapesCollection(self.shapes_list)


class _FakeMediaDeck:
    """One slide, a container to stage into, and a `make` that can decline."""

    def __init__(self, tmp_path, make="media", existing=(),
                 landed_size=(44.44, 25.0), size_is_frozen=False,
                 playback_is_frozen=False, effects=0, sequences=0):
        from appscript import k

        self.journal: list = []
        self.playback_is_frozen = playback_is_frozen
        self._make = make
        self._landed_size = landed_size
        self._size_is_frozen = size_is_frozen
        self.container = tmp_path / "container"
        self.container.mkdir()
        self.source = tmp_path / "elsewhere"
        self.source.mkdir()
        self.made: list = []
        self.files_named: list = []
        self.existed_when_named: list = []
        self.navigated_to: list = []
        self.shape = None
        self.play = None

        shapes = [_FakeMediaShape(self, "Title", k.shape_type_text_box)]
        for name in existing:
            shape = _FakeMediaShape(self, name, k.shape_type_media)
            shapes.append(shape)
            self.shape = shape
            self.play = shape.animation_settings.animation_play_settings
        self.slide = _FakeMediaSlide(self, shapes, effects, sequences)
        self.app = self

    def a_file(self, name):
        """A real file, outside the container, the way a caller would name one."""
        path = self.source / name
        path.write_bytes(b"not really a movie, but really a file")
        return str(path)

    # -- the application ---------------------------------------------------
    def make(self, new=None, at=None, with_properties=None):
        import os

        from appscript import k

        path = with_properties[k.file_name]
        self.files_named.append(path)
        self.existed_when_named.append(os.path.exists(path))
        if self._make == "nothing":
            return object()
        shape_type = (
            k.shape_type_auto if self._make == "stray" else k.shape_type_media
        )
        shape = _FakeMediaShape(
            self, os.path.splitext(os.path.basename(path))[0], shape_type,
            size=self._landed_size, frozen=self._size_is_frozen,
        )
        self.slide.shapes_list.append(shape)
        self.shape = shape
        self.play = shape.animation_settings.animation_play_settings
        self.made.append(("media2 object", with_properties[k.left_position],
                          with_properties[k.top]))
        return object()

    # -- the presentation --------------------------------------------------
    @property
    def presentation(self):
        deck = self

        class _Pres:
            slides = _FakeCollection([deck.slide, deck.slide])

            def count(self, each=None):
                from appscript import k

                # What `target_window` asks before handing the window over.
                assert each == k.document_window
                return 1

            @property
            def document_windows(self):
                return _FakeCollection([deck._window()])

        return _Pres()

    def _window(self):
        deck = self

        class _View:
            def go_to_slide(self, number=None):
                deck.navigated_to.append(number)

        return type("Window", (), {"view": _View()})()


class _fake_media_deck:  # noqa: N801 - reads as a context manager, not a class
    """Point the wrapper and the staging directory at a fake deck."""

    def __init__(self, monkeypatch, **kwargs):
        import pathlib
        import tempfile

        self._monkeypatch = monkeypatch
        self._tmp = tempfile.TemporaryDirectory()
        self._deck = _FakeMediaDeck(pathlib.Path(self._tmp.name), **kwargs)

    def __enter__(self):
        from backend import mac_ae
        from backend.mac_ae import ppt

        self._ppt = ppt
        self._app = ppt._get_app_impl
        self._pres = ppt._get_pres_impl
        ppt._get_app_impl = lambda *a, **kw: self._deck.app
        ppt._get_pres_impl = lambda *a, **kw: self._deck.presentation
        # The staging directory is the real container everywhere else, and a
        # test has no business writing into it.
        self._monkeypatch.setattr(
            mac_ae, "EXPORT_STAGING_DIR", str(self._deck.container),
        )
        return self._deck

    def __exit__(self, *exc):
        self._ppt._get_app_impl = self._app
        self._ppt._get_pres_impl = self._pres
        self._tmp.cleanup()
        return False
