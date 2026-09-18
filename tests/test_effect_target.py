"""Tests for choosing between a shape effect and the one on its text.

Pure Python over a stand in shape. The two are different effects with the same
name, and the shape one draws nothing on a text box with no fill and no line,
which is the failure this covers.
"""

import sys

sys.path.insert(0, "src")

import pytest
from pydantic import ValidationError

from ppt_com.effects import (
    SetGlowInput,
    SetReflectionInput,
    effect_of,
    nothing_drawn_warning,
    will_not_draw,
)
from ppt_com.formatting import SetShadowInput
from ppt_com.shapes import _glow_of


class Visibility:
    def __init__(self, visible):
        self.Visible = visible


class Effect:
    def __init__(self, where):
        self.where = where


class Font:
    def __init__(self):
        self.Glow = Effect("text")
        self.Shadow = Effect("text")
        self.Reflection = Effect("text")


class TextFrame2:
    def __init__(self):
        self.TextRange = type("R", (), {"Font": Font()})()


class FakeShape:
    def __init__(self, name="TextBox 10", has_text_frame=True,
                 fill=False, line=False):
        self.Name = name
        self.HasTextFrame = has_text_frame
        self.Fill = Visibility(-1 if fill else 0)
        self.Line = Visibility(-1 if line else 0)
        self.Glow = Effect("shape")
        self.Shadow = Effect("shape")
        self.Reflection = Effect("shape")
        self.TextFrame2 = TextFrame2()


class TestWhichEffectIsReached:
    @pytest.mark.parametrize("name", ["Glow", "Shadow", "Reflection"])
    def test_shape_reaches_the_shape_one(self, name):
        assert effect_of(FakeShape(), "shape", name).where == "shape"

    @pytest.mark.parametrize("name", ["Glow", "Shadow", "Reflection"])
    def test_text_reaches_the_one_on_the_font(self, name):
        assert effect_of(FakeShape(), "text", name).where == "text"

    def test_no_target_at_all_is_the_shape(self):
        # The default, so the old behaviour is untouched.
        assert effect_of(FakeShape(), None, "Glow").where == "shape"

    def test_text_on_a_shape_with_no_text_frame_is_refused(self):
        shape = FakeShape(name="Straight Connector 2", has_text_frame=False)
        with pytest.raises(ValueError, match="has no text frame"):
            effect_of(shape, "text", "Glow")

    def test_and_the_message_names_the_shape_and_the_way_out(self):
        shape = FakeShape(name="Straight Connector 2", has_text_frame=False)
        with pytest.raises(ValueError) as caught:
            effect_of(shape, "text", "Glow")
        assert "Straight Connector 2" in str(caught.value)
        assert "target='shape'" in str(caught.value)

    def test_a_shape_that_will_not_say_whether_it_has_text_is_refused(self):
        class Awkward(FakeShape):
            @property
            def HasTextFrame(self):
                raise RuntimeError("COM said no")

            @HasTextFrame.setter
            def HasTextFrame(self, value):
                pass

        with pytest.raises(ValueError, match="has no text frame"):
            effect_of(Awkward(), "text", "Glow")


class TestWhenAShapeEffectWillDrawNothing:
    def test_a_default_text_box_has_nothing_to_draw_around(self):
        assert will_not_draw(FakeShape()) is True

    def test_a_filled_shape_has(self):
        assert will_not_draw(FakeShape(fill=True)) is False

    def test_an_outlined_shape_has(self):
        assert will_not_draw(FakeShape(line=True)) is False

    def test_a_shape_with_no_text_is_not_warned_about(self):
        # A line or a picture with no fill is perfectly ordinary and its
        # effect draws on what is there.
        assert will_not_draw(FakeShape(has_text_frame=False)) is False

    def test_a_shape_that_will_not_answer_is_not_warned_about(self):
        class Awkward(FakeShape):
            @property
            def Fill(self):
                raise RuntimeError("COM said no")

            @Fill.setter
            def Fill(self, value):
                pass

        assert will_not_draw(Awkward()) is False

    def test_the_warning_names_the_shape_and_the_way_out(self):
        message = nothing_drawn_warning(FakeShape(), "glow")
        assert "TextBox 10" in message
        assert "will not change" in message
        assert "target='text'" in message


class Glow:
    """A Glow object as COM answers it, undefined properties and all."""

    UNDEFINED = -2147483648

    def __init__(self, radius, color=0x000000, transparency=UNDEFINED):
        self.Radius = radius
        self.Color = type("C", (), {"RGB": color})()
        self.Transparency = transparency


class TestReadingAGlowBack:
    """COM answers -2147483648 for a property of an effect that is not set.
    Reporting that number is worse than reporting nothing, because it reads
    like a measurement.
    """

    def test_a_glow_that_is_set_comes_back_whole(self):
        glow = _glow_of(Glow(19.0, color=0xFFFFFF, transparency=0.0))
        assert glow == {"radius": 19.0, "color_hex": "#FFFFFF", "transparency": 0.0}

    def test_no_radius_means_no_glow_and_nothing_else_is_reported(self):
        assert _glow_of(Glow(0.0)) == {
            "radius": 0.0, "color_hex": None, "transparency": None}

    def test_an_undefined_transparency_is_not_a_number(self):
        glow = _glow_of(Glow(19.0, color=0xFFFFFF, transparency=Glow.UNDEFINED))
        assert glow["transparency"] is None
        assert glow["radius"] == 19.0

    def test_an_undefined_radius_is_not_a_glow_either(self):
        assert _glow_of(Glow(Glow.UNDEFINED))["radius"] is None

    def test_something_that_is_not_a_glow_at_all_is_None(self):
        assert _glow_of(object()) is None


def models():
    return [
        (SetGlowInput, dict(slide_index=1, shape_name_or_index="T", radius=6)),
        (SetReflectionInput, dict(slide_index=1, shape_name_or_index="T")),
        (SetShadowInput, dict(slide_index=1, shape_name_or_index="T", visible=True)),
    ]


class TestWhatTheThreeToolsAccept:
    @pytest.mark.parametrize("model,args", models())
    def test_the_default_is_the_shape(self, model, args):
        assert model(**args).target == "shape"

    @pytest.mark.parametrize("model,args", models())
    def test_text_is_accepted(self, model, args):
        assert model(**args, target="text").target == "text"

    @pytest.mark.parametrize("model,args", models())
    def test_anything_else_is_refused(self, model, args):
        with pytest.raises(ValidationError):
            model(**args, target="font")
