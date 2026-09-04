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
class TestMediaRefusals:
    """All three refuse, and the three reasons are not the same reason."""

    def test_video_and_audio_name_the_missing_insert_route(self):
        from ppt_mac.media import _add_audio_impl, _add_video_impl

        for payload in (
            _add_video_impl(1, "/tmp/clip.mp4", 0, 0, None, None, False),
            _add_audio_impl(1, "/tmp/clip.m4a", 0, 0, None, None, False),
        ):
            assert payload["platform"] == "macOS"
            assert "elements of no container" in payload["reason"]
            assert "insert from file" in payload["reason"]
            assert "sandboxed" in payload["reason"]
        assert _add_video_impl(1, "x", 0, 0, None, None, False)["error"] == (
            "ppt_add_video is not available on macOS"
        )

    def test_audio_says_what_import_sound_file_actually_does(self):
        """It is close enough to look like the answer, and it is not."""
        from ppt_mac.media import _add_audio_impl

        payload = _add_audio_impl(1, "/tmp/clip.m4a", 0, 0, None, None, False)

        assert "import sound file" in payload["reason"]
        assert any("transition" in item for item in payload["alternatives"])

    def test_playback_says_loop_and_hide_exist_and_are_left_alone(self):
        """Refusing because it is unsafe reads differently from refusing an absence."""
        from ppt_mac.media import _set_media_settings_impl

        payload = _set_media_settings_impl(
            1, "Movie 1", 0.5, None, None, None, None, None, True, True
        )

        assert payload["error"] == "ppt_set_media_settings is not available on macOS"
        assert "no volume, mute, trim or fade" in payload["reason"]
        assert "do exist" in payload["reason"]
        assert "5.2" in payload["reason"]
        assert payload["alternatives"] == [
            "ppt_get_shape_info, which reports a media shape's type",
            "ppt_update_shape, which moves and resizes it",
        ]

    def test_no_media_tool_ever_reaches_powerpoint(self):
        """Nothing here opens a file or edits a slide, so nothing connects."""
        from backend.mac_ae import ppt
        from ppt_mac.media import (
            _add_audio_impl,
            _add_video_impl,
            _set_media_settings_impl,
        )

        def _explode(*args, **kwargs):
            raise AssertionError("a refusal must not touch PowerPoint")

        original = ppt._get_app_impl
        ppt._get_app_impl = _explode
        try:
            assert "error" in _add_video_impl(1, "x", 0, 0, None, None, False)
            assert "error" in _add_audio_impl(1, "x", 0, 0, None, None, False)
            assert "error" in _set_media_settings_impl(
                1, 1, None, None, None, None, None, None, None, None
            )
        finally:
            ppt._get_app_impl = original


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
