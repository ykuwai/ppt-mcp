"""Theme colour names wherever a deck is built and styled (#194, item 3).

The server's instructions tell callers to use the deck's accent colours, and
an agent building a deck wrote the same hex twenty-five times for want of a
way to name one. A name is now taken by the tools that make and style shapes,
text and table cells, and is set as the theme colour itself, so the shape
follows the theme when it changes; a hex value read from the theme would not.

Nothing here reaches PowerPoint: the arguments are checked by the models, and
the two setters are driven against fakes. The live behaviour was checked on
PowerPoint for Mac and Windows.
"""

import sys

import pytest
from pydantic import ValidationError

sys.path.insert(0, "src")

from utils.color import (  # noqa: E402
    THEME_COLOR_MAP,
    check_color,
    hex_to_rgb,
    set_com_color,
    theme_color_name,
)

macos_only = pytest.mark.skipif(
    sys.platform != "darwin", reason="the Apple Event backend needs appscript"
)

SHAPE = {"slide_index": 1, "shape_name_or_index": "Box"}


def _models():
    """Each tool that takes a name, the model it is checked by, the minimal
    arguments the model needs, and its colour arguments."""
    from ppt_com.effects import SetGlowInput
    from ppt_com.formatting import SetFillInput, SetLineInput, SetShadowInput
    from ppt_com.shapes import AddShapeInput, AddTextboxInput
    from ppt_com.tables import SetTableCellInput
    from ppt_com.text import FormatTextInput, FormatTextRangeInput, RunSpec

    box = {"slide_index": 1, "left": 0, "top": 0, "width": 10, "height": 10}
    return [
        ("ppt_add_shape", AddShapeInput, dict(box, shape_type="rectangle"),
         ["font_color", "fill_color", "fill_color2", "line_color"]),
        ("ppt_add_textbox", AddTextboxInput, box, ["font_color"]),
        ("ppt_set_fill", SetFillInput, dict(SHAPE, fill_type="solid"),
         ["color", "gradient_color1", "gradient_color2"]),
        ("ppt_set_line", SetLineInput, SHAPE, ["color"]),
        ("ppt_set_shadow", SetShadowInput, dict(SHAPE, visible=True), ["color"]),
        ("ppt_set_glow", SetGlowInput, dict(SHAPE, radius=5), ["color"]),
        ("ppt_set_table_cell", SetTableCellInput, dict(SHAPE, row=1, col=1),
         ["color", "fill_color"]),
        ("ppt_format_text", FormatTextInput, SHAPE, ["color"]),
        ("ppt_format_text_range", FormatTextRangeInput, dict(SHAPE, search_text="x"),
         ["color"]),
        ("ppt_set_text runs", RunSpec, {"text": "x"}, ["color"]),
    ]


def _cases():
    return [(tool, model, base, field)
            for tool, model, base, fields in _models() for field in fields]


def _id(case):
    return f"{case[0]}.{case[3]}"


class TestNames:
    @pytest.mark.parametrize("written", [
        "accent1", "Accent1", "ACCENT1", "Accent 1", "accent-1", "accent_1", " accent1 ",
    ])
    def test_a_name_is_read_however_it_is_spelled(self, written):
        assert theme_color_name(written) == "accent1"

    def test_every_name_in_the_map_is_read(self):
        for name in THEME_COLOR_MAP:
            assert theme_color_name(name) == name
        assert theme_color_name("Followed Hyperlink") == "followed_hyperlink"

    @pytest.mark.parametrize("value", ["#156082", "156082", "accent7", "blue", "", None, 5])
    def test_anything_else_is_not_a_name(self, value):
        assert theme_color_name(value) is None


class TestCheckColor:
    def test_a_name_comes_back_spelled_as_the_map_spells_it(self):
        assert check_color("Accent 2") == "accent2"

    def test_a_hex_comes_back_untouched(self):
        assert check_color("#156082") == "#156082"
        assert check_color("FFF") == "FFF"

    def test_none_stays_none(self):
        assert check_color(None) is None

    def test_nonsense_is_refused(self):
        with pytest.raises(ValueError, match="Invalid hex color"):
            check_color("accent7")

    def test_no_colour_says_what_to_pass_instead(self):
        with pytest.raises(ValueError, match="fill_type='none'"):
            check_color("transparent", "For no fill, pass fill_type='none'.")


class TestEveryToolTakesAName:
    @pytest.mark.parametrize("case", _cases(), ids=_id)
    def test_a_name_is_taken_and_spelled_as_the_map_spells_it(self, case):
        _, model, base, field = case
        assert getattr(model(**base, **{field: "Accent 3"}), field) == "accent3"

    @pytest.mark.parametrize("case", _cases(), ids=_id)
    def test_a_hex_is_taken_as_before(self, case):
        _, model, base, field = case
        assert getattr(model(**base, **{field: "#1F2937"}), field) == "#1F2937"

    @pytest.mark.parametrize("case", _cases(), ids=_id)
    def test_nonsense_is_refused_before_anything_is_drawn(self, case):
        _, model, base, field = case
        with pytest.raises(ValidationError, match="Invalid hex color"):
            model(**base, **{field: "accent9"})

    @pytest.mark.parametrize("case", _cases(), ids=_id)
    def test_the_description_says_a_name_is_taken(self, case):
        _, model, _, field = case
        assert "'accent1'" in model.model_fields[field].description

    def test_the_message_for_a_hex_only_argument_names_only_tools_that_take_one(self):
        # hex_to_rgb is what the arguments that still take only hex reach; its
        # message lists where a name works, so that list has to stay true.
        with pytest.raises(ValueError) as caught:
            hex_to_rgb("accent1")
        said = str(caught.value)
        tools = {tool for tool, *_ in _models()}
        for tool in tools - {"ppt_set_text runs"}:
            assert tool in said, tool

    def test_a_batch_of_text_formats_checks_a_name_before_writing(self):
        from ppt_com.text import _check_format_spec

        _check_format_spec({"color": "accent1"})
        with pytest.raises(ValueError, match="Invalid hex color"):
            _check_format_spec({"color": "accent9"})


class _ColorFormat:
    """A COM ColorFormat: which of its two properties was written."""

    def __init__(self):
        self.written = {}

    def __setattr__(self, name, value):
        if name == "written":
            object.__setattr__(self, name, value)
        else:
            self.written[name] = value


class TestOnWindows:
    def test_a_name_sets_the_theme_colour_and_not_rgb(self):
        fmt = _ColorFormat()
        set_com_color(fmt, "accent1")
        assert fmt.written == {"ObjectThemeColor": 5}  # msoThemeColorAccent1

    def test_dark_and_light_are_their_own_constants(self):
        for name, value in (("dark1", 1), ("light1", 2), ("dark2", 3), ("light2", 4)):
            fmt = _ColorFormat()
            set_com_color(fmt, name)
            assert fmt.written == {"ObjectThemeColor": value}

    def test_a_hex_sets_rgb_in_powerpoints_order(self):
        fmt = _ColorFormat()
        set_com_color(fmt, "#156082")
        assert fmt.written == {"RGB": 0x15 + (0x60 << 8) + (0x82 << 16)}


class _Property:
    def __init__(self, log, name):
        self.log, self.name = log, name

    def set(self, value):
        self.log.append((self.name, value))


class _MacObject:
    """An appscript reference: any property, recording what is set on it."""

    def __init__(self):
        self.log = []

    def __getattr__(self, name):
        return _Property(self.log, name)


@macos_only
class TestOnTheMac:
    def test_a_name_sets_the_theme_index_beside_the_colour(self):
        from appscript import k

        from backend.mac_enums import set_mac_color

        fill = _MacObject()
        set_mac_color(fill, "fore_color", "accent1")
        assert fill.log == [("fore_color_theme_index", k.first_accent_theme_color)]

    def test_each_colour_finds_its_own_theme_index(self):
        from appscript import k

        from backend.mac_enums import set_mac_color

        for prop in ("fore_color", "back_color", "font_color", "color"):
            obj = _MacObject()
            set_mac_color(obj, prop, "dark2")
            assert obj.log == [(f"{prop}_theme_index", k.second_dark_theme_color)]

    def test_a_hex_sets_the_colour_itself(self):
        from backend.mac_enums import set_mac_color

        font = _MacObject()
        set_mac_color(font, "font_color", "#156082")
        assert font.log == [("font_color", [0x15, 0x60, 0x82])]
