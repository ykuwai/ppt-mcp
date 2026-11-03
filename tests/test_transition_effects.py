"""Slide transitions by PowerPoint's own numbers.

`fade`, `push`, `wipe`, `split` and `reveal` used to be sent as 3844 to 3848.
PowerPoint for Windows took 3844 as no transition and saved a cut, and 3845 to
3848 as circle, diamond and comb; macOS knew none of them, so it refused push,
wipe, split and reveal. The numbers are PowerPoint's now, checked on Windows
by what each saves (`<p:fade/>`, `<p:push dir="u"/>` ...) and on the Mac
against its dictionary, whose codes end in the same numbers.
"""

import sys

import pytest
from pydantic import ValidationError

sys.path.insert(0, "src")

from ppt_com.animation import TRANSITION_EFFECT_MAP, SetSlideTransitionInput  # noqa: E402

macos_only = pytest.mark.skipif(
    sys.platform != "darwin", reason="the Apple Event backend needs appscript"
)


class TestTheNumbers:
    def test_the_plain_names_are_powerpoints_defaults(self):
        # Fade; Push from the bottom; Wipe from the right; Split vertically
        # out; Reveal smoothly from the right.
        assert {name: TRANSITION_EFFECT_MAP[name]
                for name in ("fade", "push", "wipe", "split", "reveal")} == {
            "fade": 3849, "push": 3855, "wipe": 2817, "split": 3587, "reveal": 3894,
        }

    def test_no_name_sends_the_numbers_that_were_wrong(self):
        assert not {3844, 3845, 3846, 3847, 3848} & set(TRANSITION_EFFECT_MAP.values())

    def test_each_direction_is_its_own_number(self):
        for kind in ("push", "wipe"):
            numbers = {TRANSITION_EFFECT_MAP[f"{kind}_{way}"]
                       for way in ("up", "down", "left", "right")}
            assert len(numbers) == 4, kind
        assert len({TRANSITION_EFFECT_MAP[f"split_{o}_{d}"]
                    for o in ("vertical", "horizontal") for d in ("in", "out")}) == 4

    def test_fade_through_black_is_the_other_fade(self):
        assert TRANSITION_EFFECT_MAP["fade_through_black"] == 1793

    @macos_only
    def test_every_name_has_a_mac_word(self):
        from backend.mac_enums import PpEntryEffect

        missing = {name for name, number in TRANSITION_EFFECT_MAP.items()
                   if number not in PpEntryEffect}
        assert not missing


class TestTheName:
    def test_the_default_is_fade(self):
        assert SetSlideTransitionInput(slide_index=1).effect == "fade"

    @pytest.mark.parametrize("written", ["Push Down", "push-down", "PUSH_DOWN"])
    def test_a_name_is_read_however_it_is_spelled(self, written):
        assert SetSlideTransitionInput(slide_index=1, effect=written).effect == "push_down"

    def test_an_unknown_name_is_refused_with_the_close_one(self):
        with pytest.raises(ValidationError, match="Did you mean 'push'"):
            SetSlideTransitionInput(slide_index=1, effect="pussh")

    def test_an_unknown_name_lists_the_known_ones(self):
        with pytest.raises(ValidationError, match="reveal_black_right"):
            SetSlideTransitionInput(slide_index=1, effect="sparkle")

    def test_a_number_is_passed_as_it_is(self):
        assert SetSlideTransitionInput(slide_index=1, effect=3854).effect == 3854

    def test_the_description_names_the_directions(self):
        description = SetSlideTransitionInput.model_fields["effect"].description
        assert "push_down" in description and "fade_through_black" in description
