"""ppt_set_slide_transition on several slides in one call (#194, item 1).

`ppt_set_slide_background` already took `slide_indices`, so a caller reached
for the same here and found one call per slide. The transition tool takes the
same pair of arguments now, checks every slide before changing the first, and
answers a single slide exactly as it did.

Nothing here reaches PowerPoint; both implementations run against fakes.
"""

import sys

import pytest
from pydantic import ValidationError

sys.path.insert(0, "src")

from ppt_com.animation import SetSlideTransitionInput  # noqa: E402

macos_only = pytest.mark.skipif(
    sys.platform != "darwin", reason="the Apple Event backend needs appscript"
)
# On macOS the ppt_com modules carry the Mac implementations.
windows_only = pytest.mark.skipif(
    sys.platform != "win32", reason="drives the COM implementation"
)


class TestArguments:
    def test_one_slide_as_before(self):
        inp = SetSlideTransitionInput(slide_index=2, effect="fade")
        assert (inp.slide_index, inp.slide_indices) == (2, None)

    def test_several_slides(self):
        inp = SetSlideTransitionInput(slide_indices=[1, 2, 3], effect="fade")
        assert (inp.slide_index, inp.slide_indices) == (None, [1, 2, 3])

    def test_neither_is_refused(self):
        with pytest.raises(ValidationError, match="slide_index or slide_indices"):
            SetSlideTransitionInput(effect="fade")

    def test_an_empty_list_is_refused(self):
        with pytest.raises(ValidationError, match="must not be empty"):
            SetSlideTransitionInput(slide_indices=[], effect="fade")

    def test_an_index_below_one_is_refused(self):
        with pytest.raises(ValidationError, match=">= 1, got 0"):
            SetSlideTransitionInput(slide_indices=[1, 0], effect="fade")

    def test_the_description_says_so(self):
        assert "slide_indices" in SetSlideTransitionInput.model_fields["slide_index"].description


# ---------------------------------------------------------------------------
# Windows
# ---------------------------------------------------------------------------
class _ComTransition:
    def __init__(self):
        self.EntryEffect = None
        self.Duration = None
        self.AdvanceOnClick = None
        self.AdvanceOnTime = None
        self.AdvanceTime = None


class _ComSlide:
    def __init__(self):
        self.SlideShowTransition = _ComTransition()


class _ComDeck:
    def __init__(self, count):
        self.slides = [_ComSlide() for _ in range(count)]

    def Slides(self, index):  # noqa: N802 - the COM name
        if not 1 <= index <= len(self.slides):
            raise ValueError(f"Slide index {index} out of range")
        return self.slides[index - 1]


@pytest.fixture
def com_deck(monkeypatch):
    import ppt_com.animation as animation

    deck = _ComDeck(4)
    monkeypatch.setattr(animation.ppt, "_get_app_impl", lambda *a, **kw: object())
    monkeypatch.setattr(animation.ppt, "_get_pres_impl", lambda *a, **kw: deck)
    monkeypatch.setattr(animation, "goto_slide", lambda app, index: None)
    return deck


@windows_only
class TestOnWindows:
    def test_every_slide_listed_gets_the_transition(self, com_deck):
        from ppt_com.animation import TRANSITION_EFFECT_MAP, _set_slide_transition_impl

        result = _set_slide_transition_impl(None, "fade", 0.7, True, None, None, [1, 3, 4])
        fade = TRANSITION_EFFECT_MAP["fade"]
        assert result == {"success": True, "slide_indices": [1, 3, 4], "effect": fade}
        effects = [s.SlideShowTransition.EntryEffect for s in com_deck.slides]
        assert effects == [fade, None, fade, fade]
        assert com_deck.slides[2].SlideShowTransition.Duration == 0.7

    def test_one_slide_answers_as_it_did(self, com_deck):
        from ppt_com.animation import TRANSITION_EFFECT_MAP, _set_slide_transition_impl

        result = _set_slide_transition_impl(2, "fade", None, None, None, None)
        assert result == {"success": True, "slide_index": 2,
                          "effect": TRANSITION_EFFECT_MAP["fade"]}

    def test_an_index_past_the_end_changes_no_slide(self, com_deck):
        from ppt_com.animation import _set_slide_transition_impl

        with pytest.raises(ValueError, match="out of range"):
            _set_slide_transition_impl(None, "fade", None, None, None, None, [1, 9])
        assert all(s.SlideShowTransition.EntryEffect is None for s in com_deck.slides)

    def test_the_tool_passes_the_list_through(self, com_deck, monkeypatch):
        import json

        import ppt_com.animation as animation
        from ppt_com.animation import set_slide_transition

        # The handler hands the call to the COM worker; run it here, on the fake.
        monkeypatch.setattr(animation.ppt, "execute", lambda fn, *args, **kw: fn(*args, **kw))

        out = json.loads(set_slide_transition(
            SetSlideTransitionInput(slide_indices=[2, 4], effect="fade")))
        assert out["slide_indices"] == [2, 4]


# ---------------------------------------------------------------------------
# macOS
# ---------------------------------------------------------------------------
class _MacProperty:
    def __init__(self):
        self.value = None

    def __call__(self):
        return self.value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class _MacTransition:
    def __init__(self):
        for name in ("entry_effect", "transition_duration", "advance_on_click",
                     "advance_on_time", "advance_time"):
            setattr(self, name, _MacProperty())


class _MacSlide:
    def __init__(self):
        self.slide_show_transition = _MacTransition()


@pytest.fixture
def mac_deck(monkeypatch):
    pytest.importorskip("appscript")
    import ppt_mac.animation as animation

    slides = [_MacSlide() for _ in range(4)]

    def slide_at(pres, index):
        if not 1 <= index <= len(slides):
            raise ValueError(f"Slide index {index} is out of range (1-{len(slides)})")
        return slides[index - 1]

    monkeypatch.setattr(animation.ppt, "_get_app_impl", lambda *a, **kw: object())
    monkeypatch.setattr(animation.ppt, "_get_pres_impl", lambda *a, **kw: object())
    monkeypatch.setattr(animation, "goto_slide", lambda app, index: None)
    monkeypatch.setattr(animation, "_slide", slide_at)
    return slides


@macos_only
class TestOnTheMac:
    def test_every_slide_listed_gets_the_transition(self, mac_deck):
        from appscript import k

        from ppt_mac.animation import _set_slide_transition_impl

        result = _set_slide_transition_impl(None, "dissolve", 1.25, None, None, None, [2, 4])
        assert result == {"success": True, "slide_indices": [2, 4], "effect": 1537}
        effects = [s.slide_show_transition.entry_effect() for s in mac_deck]
        assert effects == [None, k.entry_effect_dissolve, None, k.entry_effect_dissolve]
        assert mac_deck[3].slide_show_transition.transition_duration() == 1.25

    def test_one_slide_answers_as_it_did(self, mac_deck):
        from ppt_mac.animation import _set_slide_transition_impl

        result = _set_slide_transition_impl(1, "dissolve", None, None, None, None)
        assert result == {"success": True, "slide_index": 1, "effect": 1537}

    def test_an_index_past_the_end_changes_no_slide(self, mac_deck):
        from ppt_mac.animation import _set_slide_transition_impl

        with pytest.raises(ValueError, match="out of range"):
            _set_slide_transition_impl(None, "dissolve", None, None, None, None, [1, 9])
        assert all(s.slide_show_transition.entry_effect() is None for s in mac_deck)

    def test_a_transition_the_mac_lacks_is_refused_before_any_slide(self, mac_deck):
        from ppt_mac.animation import _set_slide_transition_impl

        # 3845 is circle on Windows, and not in the Mac's table.
        with pytest.raises(ValueError, match="matching the Windows constant 3845"):
            _set_slide_transition_impl(None, 3845, None, None, None, None, [1, 2])
        assert all(s.slide_show_transition.entry_effect() is None for s in mac_deck)
