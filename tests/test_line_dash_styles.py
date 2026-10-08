"""Dash styles, connector transparency and the line block of get_shape_info.

Issue #242. ``round_dot`` used to send 2, which is msoLineSquareDot, so round
dots were out of reach, and only five of the twelve styles COM accepts could be
asked for at all. The three line tools now share one table of all twelve.

Pure Python over stand in objects; nothing here needs PowerPoint. The macOS
side is covered in test_mac_effects and test_mac_evidence, which only run on a
Mac, plus a read of the generated enumerator table below, which runs anywhere.
"""

import ast
import pathlib
import sys
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

sys.path.insert(0, "src")

import pytest
from pydantic import ValidationError

from ppt_com.constants import (
    DASH_STYLE_ALIASES,
    DASH_STYLE_MAP,
    DASH_STYLE_NAMES,
    dash_style_value,
)
from ppt_com.connectors import FormatConnectorInput
from ppt_com.formatting import SetLineInput
from ppt_com.tables import SetTableBordersInput

SRC = pathlib.Path(__file__).resolve().parents[1] / "src"

ALL_TWELVE = {
    "solid": 1,
    "square_dot": 2,
    "round_dot": 3,
    "dash": 4,
    "dash_dot": 5,
    "dash_dot_dot": 6,
    "long_dash": 7,
    "long_dash_dot": 8,
    "long_dash_dot_dot": 9,
    "sys_dash": 10,
    "sys_dot": 11,
    "sys_dash_dot": 12,
}


# ---------------------------------------------------------------------------
# The table
# ---------------------------------------------------------------------------
class TestTheTable:
    def test_every_name_sends_the_documented_number(self):
        """MsoLineDashStyle 1 to 9, plus the three sys styles COM accepts."""
        assert DASH_STYLE_MAP == ALL_TWELVE

    def test_round_dot_is_round_and_square_dot_is_square(self):
        assert DASH_STYLE_MAP["round_dot"] == 3
        assert DASH_STYLE_MAP["square_dot"] == 2

    def test_every_number_reads_back_as_its_name(self):
        for name, number in ALL_TWELVE.items():
            assert DASH_STYLE_NAMES[number] == name

    def test_dot_is_kept_and_still_means_round(self):
        """ppt_set_table_borders took 'dot', which sent 3; 3 is round dot."""
        assert DASH_STYLE_ALIASES == {"dot": 3}
        assert dash_style_value("dot") == 3

    def test_case_and_space_are_ignored(self):
        assert dash_style_value("  Sys_Dash ") == 10

    def test_an_unknown_name_is_refused_with_the_valid_ones(self):
        with pytest.raises(ValueError, match="Unknown dash_style 'wiggly'") as err:
            dash_style_value("wiggly")
        assert "sys_dash_dot" in str(err.value)


# ---------------------------------------------------------------------------
# The three tools accept all twelve and nothing else
# ---------------------------------------------------------------------------
def _set_line(dash_style):
    return SetLineInput(slide_index=1, shape_name_or_index="Line", dash_style=dash_style)


def _format_connector(dash_style):
    return FormatConnectorInput(
        slide_index=1, shape_name_or_index="Connector", dash_style=dash_style
    )


def _table_borders(dash_style):
    return SetTableBordersInput(
        slide_index=1, shape_name_or_index="Table", sides=["top"],
        dash_style=dash_style,
    )


BUILDERS = [_set_line, _format_connector, _table_borders]


class TestTheToolsAcceptEveryName:
    @pytest.mark.parametrize("build", BUILDERS)
    @pytest.mark.parametrize("name", list(ALL_TWELVE))
    def test_each_name_is_accepted(self, build, name):
        assert build(name).dash_style == name

    @pytest.mark.parametrize("build", BUILDERS)
    def test_an_unknown_name_is_rejected_by_the_model(self, build):
        with pytest.raises(ValidationError, match="Unknown dash_style"):
            build("wiggly")

    @pytest.mark.parametrize("build", BUILDERS)
    def test_a_name_is_normalised(self, build):
        assert build("Round_Dot").dash_style == "round_dot"

    def test_the_table_alias_still_validates(self):
        assert _table_borders("dot").dash_style == "dot"

    @pytest.mark.parametrize(
        "model", [SetLineInput, FormatConnectorInput, SetTableBordersInput]
    )
    def test_the_description_lists_every_name(self, model):
        description = model.model_fields["dash_style"].description
        for name in ALL_TWELVE:
            assert f"'{name}'" in description


class TestSetLineDescription:
    def test_it_says_line_cap_cannot_be_set_and_what_to_use(self):
        import ppt_com.formatting as formatting

        captured = {}

        class _Recorder:
            def tool(self, name, **kwargs):
                def register(func):
                    captured[name] = func.__doc__
                    return func
                return register

        formatting.register_tools(_Recorder())
        doc = captured["ppt_set_line"]
        assert "Line cap" in doc
        assert "ppt_copy_formatting" in doc


# ---------------------------------------------------------------------------
# What reaches COM on Windows
# ---------------------------------------------------------------------------
def _line():
    return SimpleNamespace(
        Visible=None, Weight=None, DashStyle=None, Transparency=None,
        ForeColor=SimpleNamespace(RGB=None),
        BeginArrowheadStyle=None, EndArrowheadStyle=None,
    )


class TestSetLineReachesCom:
    @pytest.mark.parametrize("name,number", list(ALL_TWELVE.items()))
    def test_each_name_is_written_as_its_number(self, name, number):
        from ppt_com import formatting

        shape = SimpleNamespace(Name="Line 1", Line=_line())
        with patch.object(formatting, "ppt", MagicMock()), \
                patch.object(formatting, "goto_slide"), \
                patch.object(formatting, "_get_shape", return_value=shape):
            formatting._set_line_impl(1, "Line 1", None, None, name, None, None)
        assert shape.Line.DashStyle == number

    def test_an_unknown_name_moves_nothing(self):
        from ppt_com import formatting

        goto = MagicMock()
        with patch.object(formatting, "ppt", MagicMock()), \
                patch.object(formatting, "goto_slide", goto):
            with pytest.raises(ValueError, match="Unknown dash_style"):
                formatting._set_line_impl(1, "Line 1", "#FF0000", 2, "wiggly",
                                          None, None)
        goto.assert_not_called()


class TestFormatConnectorTransparency:
    def test_transparency_reaches_line_transparency(self):
        from ppt_com import connectors

        shape = SimpleNamespace(Name="Connector 1", Line=_line())
        with patch.object(connectors, "ppt", MagicMock()), \
                patch.object(connectors, "goto_slide"), \
                patch.object(connectors, "_get_shape", return_value=shape):
            result = connectors._format_connector_impl(
                1, "Connector 1", None, None, "sys_dash",
                None, None, None, None, None, None,
                None, None, None, None, 0.45,
            )
        assert result["success"] is True
        assert shape.Line.Transparency == 0.45
        assert shape.Line.DashStyle == 10

    def test_the_wrapper_passes_transparency_through(self):
        from ppt_com import connectors

        fake_ppt = MagicMock()
        fake_ppt.execute.return_value = {"success": True}
        params = FormatConnectorInput(
            slide_index=1, shape_name_or_index="Connector 1", transparency=0.3,
        )
        with patch.object(connectors, "ppt", fake_ppt):
            connectors.format_connector(params)
        args = fake_ppt.execute.call_args.args
        assert args[0] is connectors._format_connector_impl
        assert args[-1] == 0.3

    @pytest.mark.parametrize("value", [-0.1, 1.5])
    def test_transparency_out_of_range_is_rejected(self, value):
        with pytest.raises(ValidationError):
            FormatConnectorInput(
                slide_index=1, shape_name_or_index="Connector 1",
                transparency=value,
            )


class TestTableBordersReachCom:
    @pytest.mark.parametrize("name,number", [("dot", 3), ("round_dot", 3),
                                             ("square_dot", 2), ("sys_dot", 11)])
    def test_the_name_is_written_to_each_border(self, name, number):
        from ppt_com import tables

        border = SimpleNamespace(DashStyle=None)
        cell = SimpleNamespace(Borders=SimpleNamespace(Item=lambda side: border))
        table = SimpleNamespace(
            Rows=SimpleNamespace(Count=1), Columns=SimpleNamespace(Count=1),
            Cell=lambda r, c: cell,
        )
        shape = SimpleNamespace(Name="Table 1", Table=table)
        with patch.object(tables, "ppt", MagicMock()), \
                patch.object(tables, "goto_slide"), \
                patch.object(tables, "_get_table_shape", return_value=shape):
            tables._set_table_borders_impl(
                1, "Table 1", 1, 1, None, None, ["top"], None, None, None, name,
            )
        assert border.DashStyle == number


# ---------------------------------------------------------------------------
# get_shape_info reports a line in words ppt_set_line takes
# ---------------------------------------------------------------------------
def _shape_info(dash_style, transparency=0.45):
    from ppt_com import shapes

    shape = MagicMock()
    shape.Name = "Straight Connector 1"
    shape.Id = 7
    shape.Type = 9
    shape.Left = shape.Top = shape.Width = shape.Height = 10.0
    shape.Rotation = 0.0
    shape.ZOrderPosition = 1
    shape.HasTextFrame = False
    shape.Fill.Type = 1
    shape.Fill.Visible = 0
    shape.Fill.ForeColor.RGB = 0
    shape.Fill.Transparency = 0.0
    shape.Line.Visible = -1
    shape.Line.Weight = 6.0
    shape.Line.ForeColor.RGB = 0xD6C4FF
    shape.Line.DashStyle = dash_style
    shape.Line.Transparency = transparency
    with patch.object(shapes, "ppt", MagicMock()), \
            patch.object(shapes, "_get_shape", return_value=shape):
        return shapes._get_shape_info_impl(1, "Straight Connector 1", None)


class TestShapeInfoLine:
    def test_dash_style_is_the_name_and_transparency_is_there(self):
        line = _shape_info(10)["line"]
        assert line["dash_style"] == "sys_dash"
        assert line["transparency"] == 0.45
        assert line["weight"] == 6.0

    def test_round_dot_reads_back_as_round_dot(self):
        assert _shape_info(3)["line"]["dash_style"] == "round_dot"

    def test_a_number_with_no_name_stays_a_number(self):
        """msoLineDashStyleMixed on a group has no name ppt_set_line takes."""
        assert _shape_info(-2)["line"]["dash_style"] == -2


# ---------------------------------------------------------------------------
# The macOS enumerator table, read as text so it runs anywhere
# ---------------------------------------------------------------------------
def _mac_dash_table():
    tree = ast.parse((SRC / "backend" / "mac_enums.py").read_text(encoding="utf-8"))
    for node in tree.body:
        if (isinstance(node, ast.AnnAssign)
                and getattr(node.target, "id", None) == "MsoLineDashStyle"):
            return {
                key.value: value.attr
                for key, value in zip(node.value.keys, node.value.values)
            }
    raise AssertionError("mac_enums.py has no MsoLineDashStyle table")


class TestMacEnumeratorTable:
    def test_two_is_square_and_three_is_round(self):
        table = _mac_dash_table()
        assert table[2] == "line_dash_style_square_dot"
        assert table[3] == "line_dash_style_round_dot"

    def test_nothing_is_guessed_for_the_styles_never_matched(self):
        """9 to 12 were never paired against the Mac dictionary, so absent."""
        table = _mac_dash_table()
        assert not set(table) & {9, 10, 11, 12}

    def test_each_mapped_number_has_the_name_windows_gives_it(self):
        for number, enumerator in _mac_dash_table().items():
            assert enumerator == "line_dash_style_" + DASH_STYLE_NAMES[number]
