"""Tests for Pydantic model validators in ppt_com modules.

Covers all model_validator decorated methods in:
- freeform.py: NodeSpec, BuildFreeformInput, InsertNodeInput
- tables.py: MergeTableCellsInput, SetTableBordersInput, SetTableDataInput
- advanced_ops.py: SetDefaultShapeStyleInput, CropPictureInput
- shapes.py: AddShapeInput
- layout.py: SetSlideBackgroundInput
- text.py: GetAllTextInput
- connectors.py: FormatConnectorInput

These are pure Python tests — no COM or PowerPoint required.
"""

import sys
from unittest.mock import patch, MagicMock

sys.path.insert(0, "src")

import pytest
from pydantic import ValidationError

from ppt_com.freeform import (
    NodeSpec,
    BuildFreeformInput,
    InsertNodeInput,
)
from ppt_com.tables import (
    MergeTableCellsInput,
    SetTableBordersInput,
    SetTableDataInput,
)
from ppt_com.advanced_ops import SetDefaultShapeStyleInput, CropPictureInput, SetPictureFormatInput
from ppt_com.shapes import AddShapeInput, UpdateShapeInput
from ppt_com.animation import AddAnimationInput, RemoveAnimationInput, UpdateAnimationInput
from ppt_com.connectors import FormatConnectorInput
from ppt_com.layout import SetSlideBackgroundInput
from ppt_com.text import GetAllTextInput
from utils.validation import font_size_warning


# ============================================================================
# freeform.py — NodeSpec
# ============================================================================

class TestNodeSpec:
    """Tests for NodeSpec model validator."""

    def test_line_segment_valid(self):
        """Line segment with x1/y1 is accepted."""
        node = NodeSpec(segment_type="line", x1=100, y1=200)
        assert node.segment_type == "line"

    def test_line_forces_editing_type_to_auto(self):
        """Line segment forces editing_type to 'auto' regardless of input."""
        node = NodeSpec(segment_type="line", editing_type="corner", x1=10, y1=20)
        assert node.editing_type == "auto"

    def test_curve_auto_valid(self):
        """Curve with editing_type='auto' and x1/y1 is accepted."""
        node = NodeSpec(segment_type="curve", editing_type="auto", x1=50, y1=60)
        assert node.segment_type == "curve"
        assert node.editing_type == "auto"

    def test_curve_corner_valid(self):
        """Curve with editing_type='corner' and all 6 coordinates is accepted."""
        node = NodeSpec(
            segment_type="curve", editing_type="corner",
            x1=10, y1=20, x2=30, y2=40, x3=50, y3=60,
        )
        assert node.editing_type == "corner"

    def test_curve_corner_missing_x2(self):
        """Curve corner missing x2 raises ValidationError."""
        with pytest.raises(ValidationError, match="Missing.*x2"):
            NodeSpec(
                segment_type="curve", editing_type="corner",
                x1=10, y1=20, y2=40, x3=50, y3=60,
            )

    def test_curve_corner_missing_y2(self):
        """Curve corner missing y2 raises ValidationError."""
        with pytest.raises(ValidationError, match="Missing.*y2"):
            NodeSpec(
                segment_type="curve", editing_type="corner",
                x1=10, y1=20, x2=30, x3=50, y3=60,
            )

    def test_curve_corner_missing_x3(self):
        """Curve corner missing x3 raises ValidationError."""
        with pytest.raises(ValidationError, match="Missing.*x3"):
            NodeSpec(
                segment_type="curve", editing_type="corner",
                x1=10, y1=20, x2=30, y2=40, y3=60,
            )

    def test_curve_corner_missing_y3(self):
        """Curve corner missing y3 raises ValidationError."""
        with pytest.raises(ValidationError, match="Missing.*y3"):
            NodeSpec(
                segment_type="curve", editing_type="corner",
                x1=10, y1=20, x2=30, y2=40, x3=50,
            )

    def test_curve_corner_missing_all_extra(self):
        """Curve corner with no extra coordinates lists all missing fields."""
        with pytest.raises(ValidationError, match="Missing"):
            NodeSpec(
                segment_type="curve", editing_type="corner",
                x1=10, y1=20,
            )

    def test_curve_smooth_rejected(self):
        """Curve with editing_type='smooth' is rejected for new freeforms."""
        with pytest.raises(ValidationError, match="must be 'auto' or 'corner'"):
            NodeSpec(segment_type="curve", editing_type="smooth", x1=10, y1=20)

    def test_curve_symmetric_rejected(self):
        """Curve with editing_type='symmetric' is rejected for new freeforms."""
        with pytest.raises(ValidationError, match="must be 'auto' or 'corner'"):
            NodeSpec(segment_type="curve", editing_type="symmetric", x1=10, y1=20)

    def test_invalid_segment_type(self):
        """Invalid segment_type raises ValidationError."""
        with pytest.raises(ValidationError, match="must be 'line' or 'curve'"):
            NodeSpec(segment_type="arc", x1=10, y1=20)

    def test_segment_type_case_insensitive(self):
        """segment_type is case-insensitive."""
        node = NodeSpec(segment_type="LINE", x1=100, y1=200)
        assert node.segment_type == "LINE"

    def test_editing_type_case_insensitive(self):
        """editing_type is case-insensitive (lowered internally)."""
        node = NodeSpec(segment_type="curve", editing_type="AUTO", x1=50, y1=60)
        assert node.editing_type == "AUTO"

    def test_whitespace_stripped(self):
        """Leading/trailing whitespace is stripped from string fields."""
        node = NodeSpec(segment_type="  line  ", x1=10, y1=20)
        assert node.segment_type == "line"

    def test_line_auto_default_editing_type(self):
        """Line segment defaults editing_type to 'auto'."""
        node = NodeSpec(segment_type="line", x1=10, y1=20)
        assert node.editing_type == "auto"

    def test_curve_auto_ignores_extra_coords(self):
        """Curve auto does not reject extra x2/y2/x3/y3 (they are just unused)."""
        node = NodeSpec(
            segment_type="curve", editing_type="auto",
            x1=10, y1=20, x2=30, y2=40, x3=50, y3=60,
        )
        assert node.x2 == 30


# ============================================================================
# freeform.py — BuildFreeformInput
# ============================================================================

class TestBuildFreeformInput:
    """Tests for BuildFreeformInput model validator."""

    def _make_node(self, **overrides):
        """Helper to create a minimal valid node dict."""
        base = {"segment_type": "line", "x1": 100, "y1": 200}
        base.update(overrides)
        return base

    def test_valid_minimal(self):
        """Minimal valid input with one line node is accepted."""
        inp = BuildFreeformInput(
            slide_index=1, start_x=0, start_y=0,
            nodes=[self._make_node()],
        )
        assert inp.start_editing_type == "corner"

    def test_start_editing_type_auto(self):
        """start_editing_type='auto' is accepted."""
        inp = BuildFreeformInput(
            slide_index=1, start_x=0, start_y=0,
            start_editing_type="auto",
            nodes=[self._make_node()],
        )
        assert inp.start_editing_type == "auto"

    def test_start_editing_type_corner(self):
        """start_editing_type='corner' is accepted."""
        inp = BuildFreeformInput(
            slide_index=1, start_x=0, start_y=0,
            start_editing_type="corner",
            nodes=[self._make_node()],
        )
        assert inp.start_editing_type == "corner"

    def test_start_editing_type_smooth_rejected(self):
        """start_editing_type='smooth' is rejected."""
        with pytest.raises(ValidationError, match="must be 'auto' or 'corner'"):
            BuildFreeformInput(
                slide_index=1, start_x=0, start_y=0,
                start_editing_type="smooth",
                nodes=[self._make_node()],
            )

    def test_start_editing_type_symmetric_rejected(self):
        """start_editing_type='symmetric' is rejected."""
        with pytest.raises(ValidationError, match="must be 'auto' or 'corner'"):
            BuildFreeformInput(
                slide_index=1, start_x=0, start_y=0,
                start_editing_type="symmetric",
                nodes=[self._make_node()],
            )

    def test_start_editing_type_invalid_value_rejected(self):
        """start_editing_type with arbitrary invalid value is rejected."""
        with pytest.raises(ValidationError, match="must be 'auto' or 'corner'"):
            BuildFreeformInput(
                slide_index=1, start_x=0, start_y=0,
                start_editing_type="bogus",
                nodes=[self._make_node()],
            )

    def test_start_editing_type_case_insensitive(self):
        """start_editing_type is case-insensitive."""
        inp = BuildFreeformInput(
            slide_index=1, start_x=0, start_y=0,
            start_editing_type="AUTO",
            nodes=[self._make_node()],
        )
        assert inp.start_editing_type == "auto"

    def test_empty_nodes_rejected(self):
        """Empty nodes list is rejected (min_length=1)."""
        with pytest.raises(ValidationError):
            BuildFreeformInput(
                slide_index=1, start_x=0, start_y=0,
                nodes=[],
            )

    def test_slide_index_zero_rejected(self):
        """slide_index=0 is rejected (ge=1)."""
        with pytest.raises(ValidationError):
            BuildFreeformInput(
                slide_index=0, start_x=0, start_y=0,
                nodes=[self._make_node()],
            )

    def test_multiple_nodes(self):
        """Multiple valid nodes are accepted."""
        inp = BuildFreeformInput(
            slide_index=1, start_x=0, start_y=0,
            nodes=[
                self._make_node(),
                {"segment_type": "curve", "editing_type": "auto", "x1": 50, "y1": 60},
            ],
        )
        assert len(inp.nodes) == 2

    def test_nested_node_validation_propagates(self):
        """Invalid node in nodes list raises ValidationError."""
        with pytest.raises(ValidationError, match="must be 'line' or 'curve'"):
            BuildFreeformInput(
                slide_index=1, start_x=0, start_y=0,
                nodes=[{"segment_type": "invalid", "x1": 10, "y1": 20}],
            )


# ============================================================================
# freeform.py — InsertNodeInput
# ============================================================================

class TestInsertNodeInput:
    """Tests for InsertNodeInput model validator."""

    def test_line_valid(self):
        """Line segment insert is accepted."""
        inp = InsertNodeInput(
            slide_index=1, shape_name="s1", after_index=1,
            segment_type="line", x1=100, y1=200,
        )
        assert inp.segment_type == "line"

    def test_line_forces_editing_type_to_auto(self):
        """Line insert forces editing_type to 'auto' regardless of input."""
        inp = InsertNodeInput(
            slide_index=1, shape_name="s1", after_index=1,
            segment_type="line", editing_type="corner", x1=10, y1=20,
        )
        assert inp.editing_type == "auto"

    def test_curve_auto_valid(self):
        """Curve auto insert is accepted."""
        inp = InsertNodeInput(
            slide_index=1, shape_name="s1", after_index=1,
            segment_type="curve", editing_type="auto", x1=50, y1=60,
        )
        assert inp.editing_type == "auto"

    def test_curve_corner_valid(self):
        """Curve corner with all coordinates is accepted."""
        inp = InsertNodeInput(
            slide_index=1, shape_name="s1", after_index=1,
            segment_type="curve", editing_type="corner",
            x1=10, y1=20, x2=30, y2=40, x3=50, y3=60,
        )
        assert inp.editing_type == "corner"

    def test_curve_corner_missing_coords(self):
        """Curve corner insert missing coordinates raises ValidationError."""
        with pytest.raises(ValidationError, match="Missing"):
            InsertNodeInput(
                slide_index=1, shape_name="s1", after_index=1,
                segment_type="curve", editing_type="corner",
                x1=10, y1=20,
            )

    def test_curve_smooth_rejected(self):
        """Curve with editing_type='smooth' is rejected for insert."""
        with pytest.raises(ValidationError, match="must be 'auto' or 'corner'"):
            InsertNodeInput(
                slide_index=1, shape_name="s1", after_index=1,
                segment_type="curve", editing_type="smooth", x1=10, y1=20,
            )

    def test_curve_symmetric_rejected(self):
        """Curve with editing_type='symmetric' is rejected for insert."""
        with pytest.raises(ValidationError, match="must be 'auto' or 'corner'"):
            InsertNodeInput(
                slide_index=1, shape_name="s1", after_index=1,
                segment_type="curve", editing_type="symmetric", x1=10, y1=20,
            )

    def test_invalid_segment_type(self):
        """Invalid segment_type raises ValidationError."""
        with pytest.raises(ValidationError, match="must be 'line' or 'curve'"):
            InsertNodeInput(
                slide_index=1, shape_name="s1", after_index=1,
                segment_type="spline", x1=10, y1=20,
            )

    def test_after_index_zero_rejected(self):
        """after_index=0 is rejected (ge=1)."""
        with pytest.raises(ValidationError):
            InsertNodeInput(
                slide_index=1, shape_name="s1", after_index=0,
                segment_type="line", x1=10, y1=20,
            )

    def test_slide_index_zero_rejected(self):
        """slide_index=0 is rejected (ge=1)."""
        with pytest.raises(ValidationError):
            InsertNodeInput(
                slide_index=0, shape_name="s1", after_index=1,
                segment_type="line", x1=10, y1=20,
            )


# ============================================================================
# tables.py — SetTableDataInput
# ============================================================================

class TestSetTableDataInput:
    """Tests for SetTableDataInput validators."""

    def test_valid_basic(self):
        """Basic 2D data array is accepted."""
        inp = SetTableDataInput(
            slide_index=1, shape_name_or_index="Table1",
            data=[["A", "B"], ["C", "D"]],
        )
        assert inp.data == [["A", "B"], ["C", "D"]]
        assert inp.start_row == 1
        assert inp.start_col == 1
        assert inp.bold_first_row is False

    def test_custom_start_position(self):
        """Custom start_row and start_col are accepted."""
        inp = SetTableDataInput(
            slide_index=1, shape_name_or_index="Table1",
            data=[["X"]], start_row=3, start_col=2,
        )
        assert inp.start_row == 3
        assert inp.start_col == 2

    def test_bold_first_row(self):
        """bold_first_row=True is accepted."""
        inp = SetTableDataInput(
            slide_index=1, shape_name_or_index="Table1",
            data=[["Header1", "Header2"], ["Data1", "Data2"]],
            bold_first_row=True,
        )
        assert inp.bold_first_row is True

    def test_empty_data_rejected(self):
        """Empty data list raises ValidationError."""
        with pytest.raises(ValidationError, match="data must contain at least one row"):
            SetTableDataInput(
                slide_index=1, shape_name_or_index="Table1",
                data=[],
            )

    def test_empty_first_row_rejected(self):
        """Data with empty first row raises ValidationError."""
        with pytest.raises(ValidationError, match="data rows must contain at least one value"):
            SetTableDataInput(
                slide_index=1, shape_name_or_index="Table1",
                data=[[]],
            )

    def test_zero_start_row_rejected(self):
        """start_row=0 is rejected (ge=1)."""
        with pytest.raises(ValidationError):
            SetTableDataInput(
                slide_index=1, shape_name_or_index="Table1",
                data=[["A"]], start_row=0,
            )

    def test_zero_start_col_rejected(self):
        """start_col=0 is rejected (ge=1)."""
        with pytest.raises(ValidationError):
            SetTableDataInput(
                slide_index=1, shape_name_or_index="Table1",
                data=[["A"]], start_col=0,
            )

    def test_shape_name_or_index_accepts_int(self):
        """shape_name_or_index accepts an integer."""
        inp = SetTableDataInput(
            slide_index=1, shape_name_or_index=5,
            data=[["A"]],
        )
        assert inp.shape_name_or_index == 5

    def test_single_row_data(self):
        """Single row data is accepted."""
        inp = SetTableDataInput(
            slide_index=1, shape_name_or_index="T",
            data=[["Q1", "$1.2M", "+12%"]],
        )
        assert len(inp.data) == 1
        assert len(inp.data[0]) == 3


# ============================================================================
# tables.py — MergeTableCellsInput
# ============================================================================

class TestMergeTableCellsInput:
    """Tests for MergeTableCellsInput range order validator."""

    def test_valid_single_cell(self):
        """start == end (single cell merge) is accepted."""
        inp = MergeTableCellsInput(
            slide_index=1, shape_name_or_index="Table1",
            start_row=2, start_col=3, end_row=2, end_col=3,
        )
        assert inp.start_row == 2

    def test_valid_range(self):
        """Normal range where end >= start is accepted."""
        inp = MergeTableCellsInput(
            slide_index=1, shape_name_or_index="Table1",
            start_row=1, start_col=1, end_row=3, end_col=4,
        )
        assert inp.end_row == 3

    def test_end_row_less_than_start_row(self):
        """end_row < start_row raises ValidationError."""
        with pytest.raises(ValidationError, match="end_row.*must be >= start_row"):
            MergeTableCellsInput(
                slide_index=1, shape_name_or_index="Table1",
                start_row=5, start_col=1, end_row=2, end_col=3,
            )

    def test_end_col_less_than_start_col(self):
        """end_col < start_col raises ValidationError."""
        with pytest.raises(ValidationError, match="end_col.*must be >= start_col"):
            MergeTableCellsInput(
                slide_index=1, shape_name_or_index="Table1",
                start_row=1, start_col=5, end_row=3, end_col=2,
            )

    def test_both_end_less_than_start(self):
        """Both end_row and end_col less than start raises ValidationError."""
        with pytest.raises(ValidationError):
            MergeTableCellsInput(
                slide_index=1, shape_name_or_index="Table1",
                start_row=5, start_col=5, end_row=1, end_col=1,
            )

    def test_shape_name_or_index_accepts_int(self):
        """shape_name_or_index accepts an integer index."""
        inp = MergeTableCellsInput(
            slide_index=1, shape_name_or_index=3,
            start_row=1, start_col=1, end_row=2, end_col=2,
        )
        assert inp.shape_name_or_index == 3

    def test_zero_row_rejected(self):
        """start_row=0 is rejected (ge=1)."""
        with pytest.raises(ValidationError):
            MergeTableCellsInput(
                slide_index=1, shape_name_or_index="T",
                start_row=0, start_col=1, end_row=1, end_col=1,
            )


# ============================================================================
# tables.py — SetTableBordersInput
# ============================================================================

class TestSetTableBordersInput:
    """Tests for SetTableBordersInput validators."""

    def _base(self, **overrides):
        """Helper to create a minimal valid input dict."""
        base = {
            "slide_index": 1,
            "shape_name_or_index": "Table1",
            "sides": ["top"],
            "visible": True,
        }
        base.update(overrides)
        return base

    def test_valid_minimal(self):
        """Minimal valid input is accepted."""
        inp = SetTableBordersInput(**self._base())
        assert inp.visible is True

    def test_valid_with_all_properties(self):
        """All optional properties provided at once is accepted."""
        inp = SetTableBordersInput(**self._base(
            color="#FF0000", weight=1.5, dash_style="solid",
        ))
        assert inp.color == "#FF0000"

    def test_no_property_rejected(self):
        """Providing no border property raises ValidationError."""
        with pytest.raises(ValidationError, match="At least one of"):
            SetTableBordersInput(
                slide_index=1, shape_name_or_index="Table1",
                sides=["top"],
            )

    def test_only_color_accepted(self):
        """Providing only color (no visible/weight/dash_style) is accepted."""
        inp = SetTableBordersInput(
            slide_index=1, shape_name_or_index="Table1",
            sides=["bottom"], color="#00FF00",
        )
        assert inp.color == "#00FF00"
        assert inp.visible is None

    def test_only_weight_accepted(self):
        """Providing only weight is accepted."""
        inp = SetTableBordersInput(
            slide_index=1, shape_name_or_index="Table1",
            sides=["left"], weight=2.0,
        )
        assert inp.weight == 2.0

    def test_only_dash_style_accepted(self):
        """Providing only dash_style is accepted."""
        inp = SetTableBordersInput(
            slide_index=1, shape_name_or_index="Table1",
            sides=["right"], dash_style="dash",
        )
        assert inp.dash_style == "dash"

    def test_end_row_less_than_start_row(self):
        """end_row < start_row raises ValidationError."""
        with pytest.raises(ValidationError, match="end_row.*must be >= start_row"):
            SetTableBordersInput(**self._base(start_row=5, end_row=2))

    def test_end_col_less_than_start_col(self):
        """end_col < start_col raises ValidationError."""
        with pytest.raises(ValidationError, match="end_col.*must be >= start_col"):
            SetTableBordersInput(**self._base(start_col=5, end_col=2))

    def test_end_row_none_skips_check(self):
        """end_row=None (default) skips the range-order check."""
        inp = SetTableBordersInput(**self._base(start_row=5))
        assert inp.end_row is None

    def test_end_col_none_skips_check(self):
        """end_col=None (default) skips the range-order check."""
        inp = SetTableBordersInput(**self._base(start_col=5))
        assert inp.end_col is None

    def test_end_row_equal_to_start_row(self):
        """end_row == start_row is accepted."""
        inp = SetTableBordersInput(**self._base(start_row=3, end_row=3))
        assert inp.end_row == 3

    def test_empty_sides_rejected(self):
        """Empty sides list is rejected (min_length=1)."""
        with pytest.raises(ValidationError):
            SetTableBordersInput(
                slide_index=1, shape_name_or_index="Table1",
                sides=[], visible=True,
            )

    def test_multiple_sides_accepted(self):
        """Multiple sides in list are accepted."""
        inp = SetTableBordersInput(**self._base(
            sides=["top", "bottom", "left", "right"],
        ))
        assert len(inp.sides) == 4

    def test_diagonal_sides_accepted(self):
        """Diagonal sides are accepted as valid values."""
        inp = SetTableBordersInput(**self._base(
            sides=["diagonal_down", "diagonal_up"],
        ))
        assert "diagonal_down" in inp.sides


# ============================================================================
# advanced_ops.py — SetDefaultShapeStyleInput
# ============================================================================

class TestSetDefaultShapeStyleInput:
    """Tests for the validate_mode cross-field validator."""

    def test_shape_based_both_fields_valid(self):
        """Shape-based mode with both fields is valid."""
        inp = SetDefaultShapeStyleInput(slide_index=1, shape_name_or_index="Rect 1")
        assert inp.slide_index == 1
        assert inp.shape_name_or_index == "Rect 1"

    def test_shape_based_int_index_valid(self):
        """shape_name_or_index accepts an integer."""
        inp = SetDefaultShapeStyleInput(slide_index=2, shape_name_or_index=3)
        assert inp.shape_name_or_index == 3

    def test_shape_based_only_slide_index_raises(self):
        """slide_index without shape_name_or_index raises ValidationError."""
        with pytest.raises(ValidationError):
            SetDefaultShapeStyleInput(slide_index=1)

    def test_shape_based_only_shape_name_raises(self):
        """shape_name_or_index without slide_index raises ValidationError."""
        with pytest.raises(ValidationError):
            SetDefaultShapeStyleInput(shape_name_or_index="Rect 1")

    def test_mixed_mode_raises(self):
        """Combining shape-based and property-based params raises ValidationError."""
        with pytest.raises(ValidationError):
            SetDefaultShapeStyleInput(
                slide_index=1, shape_name_or_index="Rect 1",
                fill_type="solid", fill_color="#FF0000",
            )

    def test_property_based_valid(self):
        """Property-based mode with valid fields is accepted."""
        inp = SetDefaultShapeStyleInput(
            fill_type="solid", fill_color="#FF0000",
            line_visible=False, font_bold=True,
        )
        assert inp.fill_type == "solid"
        assert inp.fill_color == "#FF0000"

    def test_fill_type_none_valid(self):
        """fill_type='none' without fill_color is valid."""
        inp = SetDefaultShapeStyleInput(fill_type="none")
        assert inp.fill_type == "none"

    def test_fill_type_solid_without_color_raises(self):
        """fill_type='solid' without fill_color raises ValidationError."""
        with pytest.raises(ValidationError):
            SetDefaultShapeStyleInput(fill_type="solid")

    def test_fill_type_invalid_raises(self):
        """Unknown fill_type value raises ValidationError."""
        with pytest.raises(ValidationError):
            SetDefaultShapeStyleInput(fill_type="gradient")

    def test_all_none_valid(self):
        """All-None input (no-op) is valid."""
        inp = SetDefaultShapeStyleInput()
        assert inp.fill_type is None
        assert inp.slide_index is None


# ============================================================================
# shapes.py — AddShapeInput corner_radius validation
# ============================================================================


class TestAddShapeCornerRadius:
    """Tests for AddShapeInput.corner_radius Pydantic field validation."""

    def test_corner_radius_valid_zero(self):
        """corner_radius=0.0 (square corners) is accepted."""
        inp = AddShapeInput(
            slide_index=1, shape_type="rounded_rectangle",
            left=0, top=0, width=100, height=50, corner_radius=0.0,
        )
        assert inp.corner_radius == 0.0

    def test_corner_radius_valid_one(self):
        """corner_radius=1.0 (maximum rounding) is accepted."""
        inp = AddShapeInput(
            slide_index=1, shape_type="rounded_rectangle",
            left=0, top=0, width=100, height=50, corner_radius=1.0,
        )
        assert inp.corner_radius == 1.0

    def test_corner_radius_valid_mid(self):
        """corner_radius=0.5 is accepted."""
        inp = AddShapeInput(
            slide_index=1, shape_type="rounded_rectangle",
            left=0, top=0, width=100, height=50, corner_radius=0.5,
        )
        assert inp.corner_radius == 0.5

    def test_corner_radius_none_by_default(self):
        """corner_radius defaults to None."""
        inp = AddShapeInput(
            slide_index=1, shape_type="rectangle",
            left=0, top=0, width=100, height=50,
        )
        assert inp.corner_radius is None

    def test_corner_radius_too_high(self):
        """corner_radius > 1.0 is rejected."""
        with pytest.raises(ValidationError):
            AddShapeInput(
                slide_index=1, shape_type="rounded_rectangle",
                left=0, top=0, width=100, height=50, corner_radius=1.1,
            )

    def test_corner_radius_negative(self):
        """corner_radius < 0.0 is rejected."""
        with pytest.raises(ValidationError):
            AddShapeInput(
                slide_index=1, shape_type="rounded_rectangle",
                left=0, top=0, width=100, height=50, corner_radius=-0.1,
            )

    def test_corner_radius_pt_valid(self):
        """corner_radius_pt with a positive value is accepted."""
        inp = AddShapeInput(
            slide_index=1, shape_type="rounded_rectangle",
            left=0, top=0, width=200, height=100, corner_radius_pt=10,
        )
        assert inp.corner_radius_pt == 10
        assert inp.corner_radius is None

    def test_corner_radius_pt_zero_rejected(self):
        """corner_radius_pt=0 is rejected (gt=0.0)."""
        with pytest.raises(ValidationError):
            AddShapeInput(
                slide_index=1, shape_type="rounded_rectangle",
                left=0, top=0, width=200, height=100, corner_radius_pt=0,
            )

    def test_corner_radius_pt_negative_rejected(self):
        """corner_radius_pt < 0 is rejected."""
        with pytest.raises(ValidationError):
            AddShapeInput(
                slide_index=1, shape_type="rounded_rectangle",
                left=0, top=0, width=200, height=100, corner_radius_pt=-5,
            )

    def test_both_corner_radius_and_pt_rejected(self):
        """Setting both corner_radius and corner_radius_pt raises."""
        with pytest.raises(ValidationError, match="mutually exclusive"):
            AddShapeInput(
                slide_index=1, shape_type="rounded_rectangle",
                left=0, top=0, width=200, height=100,
                corner_radius=0.5, corner_radius_pt=10,
            )

    def test_neither_corner_radius_valid(self):
        """Neither corner_radius nor corner_radius_pt is fine (both default None)."""
        inp = AddShapeInput(
            slide_index=1, shape_type="rounded_rectangle",
            left=0, top=0, width=200, height=100,
        )
        assert inp.corner_radius is None
        assert inp.corner_radius_pt is None


# ============================================================================
# layout.py — SetSlideBackgroundInput
# ============================================================================

class TestSetSlideBackgroundInput:
    """Tests for SetSlideBackgroundInput validate_slide_target validator."""

    def test_slide_index_only_valid(self):
        """slide_index alone is accepted."""
        inp = SetSlideBackgroundInput(
            slide_index=1, fill_type="solid", color="#FF0000",
        )
        assert inp.slide_index == 1
        assert inp.slide_indices is None

    def test_slide_indices_only_valid(self):
        """slide_indices alone is accepted."""
        inp = SetSlideBackgroundInput(
            slide_indices=[1, 2, 3], fill_type="solid", color="#FF0000",
        )
        assert inp.slide_indices == [1, 2, 3]
        assert inp.slide_index is None

    def test_both_slide_index_and_slide_indices_valid(self):
        """Both slide_index and slide_indices provided is accepted (slide_indices wins)."""
        inp = SetSlideBackgroundInput(
            slide_index=1, slide_indices=[2, 3],
            fill_type="solid", color="#FF0000",
        )
        assert inp.slide_index == 1
        assert inp.slide_indices == [2, 3]

    def test_neither_provided_raises(self):
        """Neither slide_index nor slide_indices raises ValidationError."""
        with pytest.raises(ValidationError, match="Either slide_index or slide_indices"):
            SetSlideBackgroundInput(fill_type="solid", color="#FF0000")

    def test_empty_slide_indices_raises(self):
        """Empty slide_indices list raises ValidationError."""
        with pytest.raises(ValidationError, match="slide_indices must not be empty"):
            SetSlideBackgroundInput(
                slide_indices=[], fill_type="solid", color="#FF0000",
            )

    def test_empty_slide_indices_with_slide_index_raises(self):
        """Empty slide_indices=[] with valid slide_index still raises."""
        with pytest.raises(ValidationError, match="slide_indices must not be empty"):
            SetSlideBackgroundInput(
                slide_index=1, slide_indices=[], fill_type="solid", color="#FF0000",
            )

    def test_slide_indices_with_zero_raises(self):
        """slide_indices containing 0 raises ValidationError."""
        with pytest.raises(ValidationError, match="must be >= 1"):
            SetSlideBackgroundInput(
                slide_indices=[1, 0, 3], fill_type="solid", color="#FF0000",
            )

    def test_slide_indices_with_negative_raises(self):
        """slide_indices containing negative value raises ValidationError."""
        with pytest.raises(ValidationError, match="must be >= 1"):
            SetSlideBackgroundInput(
                slide_indices=[-1, 2], fill_type="solid", color="#FF0000",
            )

    def test_slide_index_zero_raises(self):
        """slide_index=0 is rejected by ge=1."""
        with pytest.raises(ValidationError):
            SetSlideBackgroundInput(
                slide_index=0, fill_type="solid", color="#FF0000",
            )

    def test_single_slide_indices_valid(self):
        """Single-element slide_indices list is accepted."""
        inp = SetSlideBackgroundInput(
            slide_indices=[5], fill_type="none",
        )
        assert inp.slide_indices == [5]


# ============================================================================
# text.py — GetAllTextInput
# ============================================================================
class TestGetAllTextInput:
    """Tests for GetAllTextInput model validation."""

    def test_no_params_valid(self):
        """No parameters → extract all slides."""
        inp = GetAllTextInput()
        assert inp.slide_indices is None

    def test_slide_indices_valid(self):
        """Valid slide_indices list is accepted."""
        inp = GetAllTextInput(slide_indices=[1, 3, 5])
        assert inp.slide_indices == [1, 3, 5]

    def test_single_slide_index_valid(self):
        """Single-element list is accepted."""
        inp = GetAllTextInput(slide_indices=[1])
        assert inp.slide_indices == [1]

    def test_empty_slide_indices_raises(self):
        """Empty slide_indices list is rejected."""
        with pytest.raises(ValidationError, match="slide_indices must not be empty"):
            GetAllTextInput(slide_indices=[])

    def test_zero_slide_index_raises(self):
        """slide_indices with 0 is rejected."""
        with pytest.raises(ValidationError, match="must be >= 1"):
            GetAllTextInput(slide_indices=[0])

    def test_negative_slide_index_raises(self):
        """slide_indices with negative value is rejected."""
        with pytest.raises(ValidationError, match="must be >= 1"):
            GetAllTextInput(slide_indices=[-1])

    def test_mixed_valid_invalid_raises(self):
        """Mixed valid and invalid indices → rejected."""
        with pytest.raises(ValidationError, match="must be >= 1"):
            GetAllTextInput(slide_indices=[1, 0, 3])

    def test_output_path_default_none(self):
        """output_path defaults to None."""
        inp = GetAllTextInput()
        assert inp.output_path is None

    def test_output_path_valid_string(self):
        """String output_path is accepted."""
        inp = GetAllTextInput(output_path="slides.md")
        assert inp.output_path == "slides.md"

    def test_output_path_with_slide_indices(self):
        """output_path and slide_indices can be used together."""
        inp = GetAllTextInput(slide_indices=[1, 2], output_path="/tmp/out.md")
        assert inp.slide_indices == [1, 2]
        assert inp.output_path == "/tmp/out.md"


# ============================================================================
# utils/validation.py — font_size_warning
# ============================================================================
class TestFontSizeWarning:
    """Tests for font_size_warning helper."""

    def test_none_returns_none(self):
        assert font_size_warning(None) is None

    def test_large_size_returns_none(self):
        assert font_size_warning(20) is None

    def test_boundary_16_returns_none(self):
        assert font_size_warning(16) is None

    def test_just_below_16_returns_warning(self):
        result = font_size_warning(15.9)
        assert result is not None
        assert "15.9pt" in result
        assert "below the recommended minimum" in result

    def test_small_size_returns_warning(self):
        result = font_size_warning(8)
        assert result is not None
        assert "8pt" in result


# ============================================================================
# advanced_ops.py — CropPictureInput
# ============================================================================
class TestCropPictureInput:
    """Tests for CropPictureInput validators."""

    def test_crop_fit_alone_valid(self):
        m = CropPictureInput(slide_index=1, shape_name_or_index="pic", crop_fit="square")
        assert m.crop_fit == "square"
        assert m.crop_anchor == 0.5  # default

    def test_crop_fit_with_shape_and_anchor(self):
        m = CropPictureInput(
            slide_index=1, shape_name_or_index="pic",
            crop_fit="1:1", crop_shape="oval", crop_anchor=0.3,
        )
        assert m.crop_fit == "1:1"
        assert m.crop_anchor == 0.3

    def test_crop_fit_with_crop_left_raises(self):
        with pytest.raises(ValidationError, match="crop_fit cannot be combined"):
            CropPictureInput(
                slide_index=1, shape_name_or_index="pic",
                crop_fit="square", crop_left=10,
            )

    def test_crop_fit_with_crop_bottom_raises(self):
        with pytest.raises(ValidationError, match="crop_fit cannot be combined"):
            CropPictureInput(
                slide_index=1, shape_name_or_index="pic",
                crop_fit="square", crop_bottom=5,
            )

    def test_manual_crop_without_fit_valid(self):
        m = CropPictureInput(
            slide_index=1, shape_name_or_index="pic",
            crop_left=50, crop_right=50, crop_shape="oval",
        )
        assert m.crop_left == 50
        assert m.crop_fit is None

    def test_crop_anchor_bounds(self):
        with pytest.raises(ValidationError):
            CropPictureInput(
                slide_index=1, shape_name_or_index="pic",
                crop_fit="square", crop_anchor=1.5,
            )
        with pytest.raises(ValidationError):
            CropPictureInput(
                slide_index=1, shape_name_or_index="pic",
                crop_fit="square", crop_anchor=-0.1,
            )

    def test_corner_radius_pt_valid(self):
        m = CropPictureInput(
            slide_index=1, shape_name_or_index="pic",
            crop_shape="rounded_rectangle", corner_radius_pt=10,
        )
        assert m.corner_radius_pt == 10


# ============================================================================
# shapes.py — UpdateShapeInput
# ============================================================================
class TestUpdateShapeInput:
    """Tests for UpdateShapeInput with adjustments field."""

    def test_adjustments_dict_valid(self):
        m = UpdateShapeInput(
            slide_index=1, shape_name="Triangle 1",
            adjustments={1: 0.25},
        )
        assert m.adjustments == {1: 0.25}

    def test_adjustments_multiple_keys(self):
        m = UpdateShapeInput(
            slide_index=1, shape_name="Arrow 1",
            adjustments={1: 0.4, 2: 0.6},
        )
        assert m.adjustments[1] == 0.4
        assert m.adjustments[2] == 0.6

    def test_adjustments_none_by_default(self):
        m = UpdateShapeInput(slide_index=1, shape_name="Box")
        assert m.adjustments is None

    def test_adjustments_with_other_fields(self):
        m = UpdateShapeInput(
            slide_index=1, shape_name="Star 1",
            rotation=45, adjustments={1: 0.3},
        )
        assert m.rotation == 45
        assert m.adjustments == {1: 0.3}

    def test_adjustments_key_zero_rejected(self):
        """Adjustment index 0 is rejected (must be >= 1)."""
        with pytest.raises(ValidationError, match="must be >= 1"):
            UpdateShapeInput(
                slide_index=1, shape_name="Rect",
                adjustments={0: 0.5},
            )

    def test_adjustments_key_negative_rejected(self):
        """Negative adjustment index is rejected."""
        with pytest.raises(ValidationError, match="must be >= 1"):
            UpdateShapeInput(
                slide_index=1, shape_name="Rect",
                adjustments={-1: 0.5},
            )

    def test_adjustments_empty_dict_valid(self):
        """Empty adjustments dict is a valid no-op."""
        m = UpdateShapeInput(
            slide_index=1, shape_name="Box",
            adjustments={},
        )
        assert m.adjustments == {}


# ============================================================================
# connectors.py — FormatConnectorInput arrowhead size fields
# ============================================================================
class TestFormatConnectorInput:
    """Tests for FormatConnectorInput arrowhead size parameters."""

    def test_all_defaults_valid(self):
        """Minimal input with only required fields is accepted."""
        m = FormatConnectorInput(slide_index=1, shape_name_or_index="Line 1")
        assert m.begin_arrow_length is None
        assert m.end_arrow_width is None

    def test_arrow_length_values(self):
        """All valid begin_arrow_length values are accepted."""
        for val in ("short", "medium", "long"):
            m = FormatConnectorInput(
                slide_index=1, shape_name_or_index="c1",
                begin_arrow_length=val,
            )
            assert m.begin_arrow_length == val

    def test_arrow_width_values(self):
        """All valid begin_arrow_width values are accepted."""
        for val in ("narrow", "medium", "wide"):
            m = FormatConnectorInput(
                slide_index=1, shape_name_or_index="c1",
                begin_arrow_width=val,
            )
            assert m.begin_arrow_width == val

    def test_end_arrow_length_and_width(self):
        """end_arrow_length and end_arrow_width are accepted together."""
        m = FormatConnectorInput(
            slide_index=1, shape_name_or_index="c1",
            end_arrow_length="long", end_arrow_width="wide",
        )
        assert m.end_arrow_length == "long"
        assert m.end_arrow_width == "wide"

    def test_all_arrow_params_together(self):
        """All arrowhead params can be set in a single call."""
        m = FormatConnectorInput(
            slide_index=1, shape_name_or_index="c1",
            begin_arrow="triangle", begin_arrow_length="short", begin_arrow_width="narrow",
            end_arrow="stealth", end_arrow_length="long", end_arrow_width="wide",
            color="#FF0000", weight=2.0,
        )
        assert m.begin_arrow == "triangle"
        assert m.end_arrow_length == "long"
        assert m.weight == 2.0

    def test_shape_name_or_index_accepts_int(self):
        """shape_name_or_index accepts an integer index."""
        m = FormatConnectorInput(slide_index=1, shape_name_or_index=3)
        assert m.shape_name_or_index == 3

    def test_reconnect_begin_shape(self):
        """begin_shape and begin_site are accepted."""
        m = FormatConnectorInput(
            slide_index=1, shape_name_or_index="Connector 1",
            begin_shape="Rectangle 2", begin_site=3,
        )
        assert m.begin_shape == "Rectangle 2"
        assert m.begin_site == 3

    def test_reconnect_end_shape(self):
        """end_shape and end_site are accepted."""
        m = FormatConnectorInput(
            slide_index=1, shape_name_or_index="Connector 1",
            end_shape="Oval 1", end_site=2,
        )
        assert m.end_shape == "Oval 1"
        assert m.end_site == 2

    def test_reconnect_both_ends(self):
        """Both begin and end reconnection params together."""
        m = FormatConnectorInput(
            slide_index=1, shape_name_or_index="Connector 1",
            begin_shape="Rect 1", begin_site=1,
            end_shape="Rect 2", end_site=4,
        )
        assert m.begin_shape == "Rect 1"
        assert m.end_shape == "Rect 2"

    def test_reconnect_defaults_none(self):
        """Reconnection params default to None."""
        m = FormatConnectorInput(slide_index=1, shape_name_or_index="c1")
        assert m.begin_shape is None
        assert m.begin_site is None
        assert m.end_shape is None
        assert m.end_site is None

    def test_reconnect_site_must_be_positive(self):
        """begin_site and end_site must be >= 1."""
        with pytest.raises(ValidationError):
            FormatConnectorInput(
                slide_index=1, shape_name_or_index="c1",
                begin_shape="Rect 1", begin_site=0,
            )

    def test_reconnect_with_formatting(self):
        """Reconnection and formatting params can be combined."""
        m = FormatConnectorInput(
            slide_index=1, shape_name_or_index="Connector 1",
            begin_shape="Rect 3", begin_site=2,
            color="#0000FF", weight=3.0, end_arrow="triangle",
        )
        assert m.begin_shape == "Rect 3"
        assert m.color == "#0000FF"
        assert m.end_arrow == "triangle"

    def test_reconnect_end_site_must_be_positive(self):
        """end_site must be >= 1."""
        with pytest.raises(ValidationError):
            FormatConnectorInput(
                slide_index=1, shape_name_or_index="c1",
                end_shape="Rect 2", end_site=0,
            )

    def test_begin_site_without_begin_shape_rejected(self):
        """begin_site without begin_shape raises ValidationError."""
        with pytest.raises(ValidationError):
            FormatConnectorInput(
                slide_index=1, shape_name_or_index="c1",
                begin_site=2,
            )

    def test_end_site_without_end_shape_rejected(self):
        """end_site without end_shape raises ValidationError."""
        with pytest.raises(ValidationError):
            FormatConnectorInput(
                slide_index=1, shape_name_or_index="c1",
                end_site=3,
            )


# ============================================================================
# advanced_ops.py — SetPictureFormatInput
# ============================================================================
class TestSetPictureFormatInput:
    """Tests for SetPictureFormatInput validators."""

    def test_brightness_valid_zero(self):
        """brightness=0.0 is accepted."""
        inp = SetPictureFormatInput(
            slide_index=1, shape_name_or_index="pic", brightness=0.0,
        )
        assert inp.brightness == 0.0

    def test_brightness_valid_half(self):
        """brightness=0.5 is accepted."""
        inp = SetPictureFormatInput(
            slide_index=1, shape_name_or_index="pic", brightness=0.5,
        )
        assert inp.brightness == 0.5

    def test_brightness_valid_one(self):
        """brightness=1.0 is accepted."""
        inp = SetPictureFormatInput(
            slide_index=1, shape_name_or_index="pic", brightness=1.0,
        )
        assert inp.brightness == 1.0

    def test_brightness_too_low(self):
        """brightness=-0.1 is rejected."""
        with pytest.raises(ValidationError):
            SetPictureFormatInput(
                slide_index=1, shape_name_or_index="pic", brightness=-0.1,
            )

    def test_brightness_too_high(self):
        """brightness=1.1 is rejected."""
        with pytest.raises(ValidationError):
            SetPictureFormatInput(
                slide_index=1, shape_name_or_index="pic", brightness=1.1,
            )

    def test_contrast_valid_zero(self):
        """contrast=0.0 is accepted."""
        inp = SetPictureFormatInput(
            slide_index=1, shape_name_or_index="pic", contrast=0.0,
        )
        assert inp.contrast == 0.0

    def test_contrast_valid_one(self):
        """contrast=1.0 is accepted."""
        inp = SetPictureFormatInput(
            slide_index=1, shape_name_or_index="pic", contrast=1.0,
        )
        assert inp.contrast == 1.0

    def test_contrast_too_low(self):
        """contrast=-0.1 is rejected."""
        with pytest.raises(ValidationError):
            SetPictureFormatInput(
                slide_index=1, shape_name_or_index="pic", contrast=-0.1,
            )

    def test_contrast_too_high(self):
        """contrast=1.1 is rejected."""
        with pytest.raises(ValidationError):
            SetPictureFormatInput(
                slide_index=1, shape_name_or_index="pic", contrast=1.1,
            )

    def test_color_type_automatic(self):
        """color_type='automatic' is accepted."""
        inp = SetPictureFormatInput(
            slide_index=1, shape_name_or_index="pic", color_type="automatic",
        )
        assert inp.color_type == "automatic"

    def test_color_type_grayscale(self):
        """color_type='grayscale' is accepted."""
        inp = SetPictureFormatInput(
            slide_index=1, shape_name_or_index="pic", color_type="grayscale",
        )
        assert inp.color_type == "grayscale"

    def test_color_type_black_and_white(self):
        """color_type='black_and_white' is accepted."""
        inp = SetPictureFormatInput(
            slide_index=1, shape_name_or_index="pic", color_type="black_and_white",
        )
        assert inp.color_type == "black_and_white"

    def test_color_type_watermark(self):
        """color_type='watermark' is accepted."""
        inp = SetPictureFormatInput(
            slide_index=1, shape_name_or_index="pic", color_type="watermark",
        )
        assert inp.color_type == "watermark"

    def test_color_type_invalid(self):
        """Invalid color_type is rejected."""
        with pytest.raises(ValidationError, match="Unknown color_type"):
            SetPictureFormatInput(
                slide_index=1, shape_name_or_index="pic", color_type="sepia",
            )

    def test_transparent_color_valid(self):
        """Valid hex color for transparent_color is accepted."""
        inp = SetPictureFormatInput(
            slide_index=1, shape_name_or_index="pic",
            transparent_color="#FF0000",
        )
        assert inp.transparent_color == "#FF0000"

    def test_transparent_color_invalid_no_hash(self):
        """Hex color without '#' prefix is rejected."""
        with pytest.raises(ValidationError):
            SetPictureFormatInput(
                slide_index=1, shape_name_or_index="pic",
                transparent_color="FF0000",
            )

    def test_transparent_color_invalid_hex_chars(self):
        """Hex color with invalid characters is rejected."""
        with pytest.raises(ValidationError):
            SetPictureFormatInput(
                slide_index=1, shape_name_or_index="pic",
                transparent_color="#GG0000",
            )

    def test_transparent_color_invalid_format(self):
        """Non-hex string is rejected."""
        with pytest.raises(ValidationError):
            SetPictureFormatInput(
                slide_index=1, shape_name_or_index="pic",
                transparent_color="notahex",
            )

    def test_transparent_background_true(self):
        """transparent_background=True is accepted."""
        inp = SetPictureFormatInput(
            slide_index=1, shape_name_or_index="pic",
            transparent_background=True,
        )
        assert inp.transparent_background is True

    def test_transparent_background_false(self):
        """transparent_background=False is accepted."""
        inp = SetPictureFormatInput(
            slide_index=1, shape_name_or_index="pic",
            transparent_background=False,
        )
        assert inp.transparent_background is False

    def test_no_params_rejected(self):
        """No adjustment parameters raises ValidationError."""
        with pytest.raises(ValidationError, match="At least one adjustment"):
            SetPictureFormatInput(
                slide_index=1, shape_name_or_index="pic",
            )

    def test_multiple_params_valid(self):
        """Multiple parameters together are accepted."""
        inp = SetPictureFormatInput(
            slide_index=1, shape_name_or_index="pic",
            brightness=0.7, contrast=0.5, color_type="grayscale",
        )
        assert inp.brightness == 0.7
        assert inp.contrast == 0.5
        assert inp.color_type == "grayscale"

    def test_shape_name_or_index_accepts_int(self):
        """shape_name_or_index accepts an integer index."""
        inp = SetPictureFormatInput(
            slide_index=1, shape_name_or_index=3, brightness=0.5,
        )
        assert inp.shape_name_or_index == 3


# ============================================================================
# animation.py — UpdateAnimationInput
# ============================================================================
class TestUpdateAnimationInput:
    """Tests for UpdateAnimationInput model validator."""

    def test_valid_all_params(self):
        """All optional params provided is accepted."""
        inp = UpdateAnimationInput(
            slide_index=1, animation_index=2,
            effect="fade", trigger="with_previous",
            duration=1.5, delay=0.5, move_to=3,
        )
        assert inp.effect == "fade"
        assert inp.trigger == "with_previous"
        assert inp.duration == 1.5
        assert inp.delay == 0.5
        assert inp.move_to == 3

    def test_valid_effect_only(self):
        """Just effect change is accepted."""
        inp = UpdateAnimationInput(
            slide_index=1, animation_index=1, effect="zoom",
        )
        assert inp.effect == "zoom"
        assert inp.trigger is None

    def test_valid_move_to_only(self):
        """Just reorder is accepted."""
        inp = UpdateAnimationInput(
            slide_index=1, animation_index=1, move_to=5,
        )
        assert inp.move_to == 5

    def test_valid_duration_only(self):
        """Just duration change is accepted."""
        inp = UpdateAnimationInput(
            slide_index=1, animation_index=1, duration=2.0,
        )
        assert inp.duration == 2.0

    def test_no_params_raises(self):
        """No optional params raises ValidationError."""
        with pytest.raises(ValidationError, match="At least one optional parameter"):
            UpdateAnimationInput(slide_index=1, animation_index=1)

    def test_invalid_trigger(self):
        """Unknown trigger raises ValidationError."""
        with pytest.raises(ValidationError, match="Unknown trigger"):
            UpdateAnimationInput(
                slide_index=1, animation_index=1, trigger="invalid_trigger",
            )

    def test_invalid_effect_string(self):
        """Unknown effect name raises ValidationError."""
        with pytest.raises(ValidationError, match="Unknown effect"):
            UpdateAnimationInput(
                slide_index=1, animation_index=1, effect="nonexistent_effect",
            )

    def test_valid_effect_integer(self):
        """Effect as integer is accepted."""
        inp = UpdateAnimationInput(
            slide_index=1, animation_index=1, effect=42,
        )
        assert inp.effect == 42

    def test_move_to_zero_raises(self):
        """move_to=0 raises ValidationError."""
        with pytest.raises(ValidationError):
            UpdateAnimationInput(
                slide_index=1, animation_index=1, move_to=0,
            )

    def test_exit_only(self):
        """exit=True alone is a valid update."""
        inp = UpdateAnimationInput(
            slide_index=1, animation_index=1, exit=True,
        )
        assert inp.exit is True

    def test_direction_only(self):
        """direction alone is a valid update."""
        inp = UpdateAnimationInput(
            slide_index=1, animation_index=1, direction="left",
        )
        assert inp.direction == "left"

    def test_repeat_count_only(self):
        """repeat_count alone is a valid update."""
        inp = UpdateAnimationInput(
            slide_index=1, animation_index=1, repeat_count=5,
        )
        assert inp.repeat_count == 5


# ============================================================================
# animation.py — AddAnimationInput
# ============================================================================
class TestAddAnimationInput:
    """Tests for AddAnimationInput model."""

    def test_exit_default_false(self):
        """exit defaults to False."""
        inp = AddAnimationInput(
            slide_index=1, shape_name_or_index="Shape 1",
        )
        assert inp.exit is False

    def test_exit_true(self):
        """exit=True is accepted."""
        inp = AddAnimationInput(
            slide_index=1, shape_name_or_index="Shape 1",
            effect="fade", exit=True,
        )
        assert inp.exit is True

    def test_exit_with_emphasis_raises(self):
        """exit=True with emphasis effect is rejected."""
        with pytest.raises(ValidationError):
            AddAnimationInput(
                slide_index=1, shape_name_or_index="Shape 1",
                effect="teeter", exit=True,
            )

    def test_exit_with_motion_path_raises(self):
        """exit=True with motion path effect is rejected."""
        with pytest.raises(ValidationError):
            AddAnimationInput(
                slide_index=1, shape_name_or_index="Shape 1",
                effect="path_circle", exit=True,
            )

    def test_emphasis_effect_accepted(self):
        """Emphasis effect name is accepted."""
        inp = AddAnimationInput(
            slide_index=1, shape_name_or_index=1,
            effect="teeter",
        )
        assert inp.effect == "teeter"

    def test_motion_path_effect_accepted(self):
        """Motion path effect name is accepted."""
        inp = AddAnimationInput(
            slide_index=1, shape_name_or_index="Shape 1",
            effect="path_circle",
        )
        assert inp.effect == "path_circle"

    def test_direction_string(self):
        """Direction as friendly name string is accepted."""
        inp = AddAnimationInput(
            slide_index=1, shape_name_or_index="Shape 1",
            effect="fly", direction="left",
        )
        assert inp.direction == "left"

    def test_direction_invalid_string(self):
        """Invalid direction string is rejected."""
        with pytest.raises(ValidationError):
            AddAnimationInput(
                slide_index=1, shape_name_or_index="Shape 1",
                effect="fly", direction="diagonal",
            )

    def test_direction_integer(self):
        """Direction as MsoAnimDirection integer is accepted."""
        inp = AddAnimationInput(
            slide_index=1, shape_name_or_index="Shape 1",
            effect="fly", direction=4,
        )
        assert inp.direction == 4

    def test_repeat_count_valid(self):
        """repeat_count with valid value is accepted."""
        inp = AddAnimationInput(
            slide_index=1, shape_name_or_index="Shape 1",
            repeat_count=3,
        )
        assert inp.repeat_count == 3

    def test_repeat_count_negative_raises(self):
        """repeat_count with negative value is rejected."""
        with pytest.raises(ValidationError):
            AddAnimationInput(
                slide_index=1, shape_name_or_index="Shape 1",
                repeat_count=-1,
            )

    def test_auto_reverse_true(self):
        """auto_reverse=True is accepted."""
        inp = AddAnimationInput(
            slide_index=1, shape_name_or_index="Shape 1",
            auto_reverse=True,
        )
        assert inp.auto_reverse is True

    def test_smooth_start_end(self):
        """smooth_start and smooth_end are accepted together."""
        inp = AddAnimationInput(
            slide_index=1, shape_name_or_index="Shape 1",
            smooth_start=True, smooth_end=True,
        )
        assert inp.smooth_start is True
        assert inp.smooth_end is True

    def test_trigger_shape_with_on_shape_click(self):
        """trigger='on_shape_click' with trigger_shape is accepted."""
        inp = AddAnimationInput(
            slide_index=1, shape_name_or_index="Shape 1",
            trigger="on_shape_click", trigger_shape="Button 1",
        )
        assert inp.trigger == "on_shape_click"
        assert inp.trigger_shape == "Button 1"

    def test_trigger_shape_without_on_shape_click_raises(self):
        """trigger_shape without trigger='on_shape_click' is rejected."""
        with pytest.raises(ValidationError):
            AddAnimationInput(
                slide_index=1, shape_name_or_index="Shape 1",
                trigger="on_click", trigger_shape="Button 1",
            )

    def test_on_shape_click_without_trigger_shape_raises(self):
        """trigger='on_shape_click' without trigger_shape is rejected."""
        with pytest.raises(ValidationError):
            AddAnimationInput(
                slide_index=1, shape_name_or_index="Shape 1",
                trigger="on_shape_click",
            )

    def test_after_effect_valid(self):
        """after_effect='hide' is accepted."""
        inp = AddAnimationInput(
            slide_index=1, shape_name_or_index="Shape 1",
            after_effect="hide",
        )
        assert inp.after_effect == "hide"

    def test_after_effect_none(self):
        """after_effect='none' is accepted."""
        inp = AddAnimationInput(
            slide_index=1, shape_name_or_index="Shape 1",
            after_effect="none",
        )
        assert inp.after_effect == "none"

    def test_after_effect_hide_on_next_click(self):
        """after_effect='hide_on_next_click' is accepted."""
        inp = AddAnimationInput(
            slide_index=1, shape_name_or_index="Shape 1",
            after_effect="hide_on_next_click",
        )
        assert inp.after_effect == "hide_on_next_click"

    def test_after_effect_dim_with_color(self):
        """after_effect='dim' with dim_color is accepted."""
        inp = AddAnimationInput(
            slide_index=1, shape_name_or_index="Shape 1",
            after_effect="dim", dim_color="#808080",
        )
        assert inp.after_effect == "dim"
        assert inp.dim_color == "#808080"

    def test_after_effect_dim_without_color(self):
        """after_effect='dim' without dim_color is accepted (dim_color is optional)."""
        inp = AddAnimationInput(
            slide_index=1, shape_name_or_index="Shape 1",
            after_effect="dim",
        )
        assert inp.after_effect == "dim"
        assert inp.dim_color is None

    def test_after_effect_invalid(self):
        """Invalid after_effect raises ValidationError."""
        with pytest.raises(ValidationError, match="Unknown after_effect"):
            AddAnimationInput(
                slide_index=1, shape_name_or_index="Shape 1",
                after_effect="invalid",
            )

    def test_dim_color_without_dim(self):
        """dim_color without after_effect='dim' raises ValidationError."""
        with pytest.raises(ValidationError, match="dim_color can only be used"):
            AddAnimationInput(
                slide_index=1, shape_name_or_index="Shape 1",
                after_effect="hide", dim_color="#808080",
            )

    def test_dim_color_invalid_format(self):
        """dim_color with invalid format raises ValidationError."""
        with pytest.raises(ValidationError):
            AddAnimationInput(
                slide_index=1, shape_name_or_index="Shape 1",
                after_effect="dim", dim_color="red",
            )

    def test_build_level_valid(self):
        """build_level='first_level' is accepted."""
        inp = AddAnimationInput(
            slide_index=1, shape_name_or_index="Shape 1",
            build_level="first_level",
        )
        assert inp.build_level == "first_level"

    def test_build_level_invalid(self):
        """Invalid build_level raises ValidationError."""
        with pytest.raises(ValidationError, match="Unknown build_level"):
            AddAnimationInput(
                slide_index=1, shape_name_or_index="Shape 1",
                build_level="invalid",
            )

    def test_text_unit_effect_valid(self):
        """text_unit_effect='by_word' is accepted."""
        inp = AddAnimationInput(
            slide_index=1, shape_name_or_index="Shape 1",
            text_unit_effect="by_word",
        )
        assert inp.text_unit_effect == "by_word"

    def test_text_unit_effect_invalid(self):
        """Invalid text_unit_effect raises ValidationError."""
        with pytest.raises(ValidationError, match="Unknown text_unit_effect"):
            AddAnimationInput(
                slide_index=1, shape_name_or_index="Shape 1",
                text_unit_effect="invalid",
            )

    def test_animate_in_reverse(self):
        """animate_in_reverse=True is accepted."""
        inp = AddAnimationInput(
            slide_index=1, shape_name_or_index="Shape 1",
            animate_in_reverse=True,
        )
        assert inp.animate_in_reverse is True

    def test_animate_background(self):
        """animate_background=False is accepted."""
        inp = AddAnimationInput(
            slide_index=1, shape_name_or_index="Shape 1",
            animate_background=False,
        )
        assert inp.animate_background is False


# ============================================================================
# animation.py — UpdateAnimationInput (after_effect)
# ============================================================================
class TestUpdateAnimationInputAfterEffect:
    """Tests for UpdateAnimationInput after_effect fields."""

    def test_after_effect_valid(self):
        """after_effect='hide' is accepted."""
        inp = UpdateAnimationInput(
            slide_index=1, animation_index=1,
            after_effect="hide",
        )
        assert inp.after_effect == "hide"

    def test_after_effect_none(self):
        """after_effect='none' is accepted."""
        inp = UpdateAnimationInput(
            slide_index=1, animation_index=1,
            after_effect="none",
        )
        assert inp.after_effect == "none"

    def test_after_effect_hide_on_next_click(self):
        """after_effect='hide_on_next_click' is accepted."""
        inp = UpdateAnimationInput(
            slide_index=1, animation_index=1,
            after_effect="hide_on_next_click",
        )
        assert inp.after_effect == "hide_on_next_click"

    def test_after_effect_dim_with_color(self):
        """after_effect='dim' with dim_color is accepted."""
        inp = UpdateAnimationInput(
            slide_index=1, animation_index=1,
            after_effect="dim", dim_color="#808080",
        )
        assert inp.after_effect == "dim"
        assert inp.dim_color == "#808080"

    def test_after_effect_dim_without_color(self):
        """after_effect='dim' without dim_color is accepted."""
        inp = UpdateAnimationInput(
            slide_index=1, animation_index=1,
            after_effect="dim",
        )
        assert inp.after_effect == "dim"
        assert inp.dim_color is None

    def test_after_effect_invalid(self):
        """Invalid after_effect raises ValidationError."""
        with pytest.raises(ValidationError, match="Unknown after_effect"):
            UpdateAnimationInput(
                slide_index=1, animation_index=1,
                after_effect="invalid",
            )

    def test_dim_color_without_dim(self):
        """dim_color without after_effect='dim' raises ValidationError."""
        with pytest.raises(ValidationError, match="dim_color can only be used"):
            UpdateAnimationInput(
                slide_index=1, animation_index=1,
                after_effect="hide", dim_color="#808080",
            )

    def test_dim_color_invalid_format(self):
        """dim_color with invalid format raises ValidationError."""
        with pytest.raises(ValidationError):
            UpdateAnimationInput(
                slide_index=1, animation_index=1,
                after_effect="dim", dim_color="red",
            )

    def test_build_level_valid(self):
        """build_level='first_level' is accepted as sole update param."""
        inp = UpdateAnimationInput(
            slide_index=1, animation_index=1,
            build_level="first_level",
        )
        assert inp.build_level == "first_level"

    def test_build_level_invalid(self):
        """Invalid build_level raises ValidationError."""
        with pytest.raises(ValidationError, match="Unknown build_level"):
            UpdateAnimationInput(
                slide_index=1, animation_index=1,
                build_level="invalid",
            )

    def test_text_unit_effect_valid(self):
        """text_unit_effect='by_character' is accepted as sole update param."""
        inp = UpdateAnimationInput(
            slide_index=1, animation_index=1,
            text_unit_effect="by_character",
        )
        assert inp.text_unit_effect == "by_character"

    def test_text_unit_effect_invalid(self):
        """Invalid text_unit_effect raises ValidationError."""
        with pytest.raises(ValidationError, match="Unknown text_unit_effect"):
            UpdateAnimationInput(
                slide_index=1, animation_index=1,
                text_unit_effect="invalid",
            )

    def test_animate_in_reverse_valid(self):
        """animate_in_reverse=True is accepted as sole update param."""
        inp = UpdateAnimationInput(
            slide_index=1, animation_index=1,
            animate_in_reverse=True,
        )
        assert inp.animate_in_reverse is True

    def test_animate_background_valid(self):
        """animate_background=True is accepted as sole update param."""
        inp = UpdateAnimationInput(
            slide_index=1, animation_index=1,
            animate_background=True,
        )
        assert inp.animate_background is True


# ============================================================================
# animation.py — RemoveAnimationInput (sequence_index)
# ============================================================================
class TestRemoveAnimationInputSequenceIndex:
    """Tests for RemoveAnimationInput sequence_index field."""

    def test_sequence_index_valid(self):
        """sequence_index=1 is accepted."""
        inp = RemoveAnimationInput(
            slide_index=1, animation_index=1, sequence_index=1,
        )
        assert inp.sequence_index == 1

    def test_sequence_index_zero(self):
        """sequence_index=0 raises ValidationError (ge=1)."""
        with pytest.raises(ValidationError):
            RemoveAnimationInput(
                slide_index=1, animation_index=1, sequence_index=0,
            )

    def test_sequence_index_default_none(self):
        """Omitting sequence_index defaults to None (main sequence)."""
        inp = RemoveAnimationInput(slide_index=1, animation_index=1)
        assert inp.sequence_index is None


# ============================================================================
# animation.py — UpdateAnimationInput (sequence_index)
# ============================================================================
class TestUpdateAnimationInputSequenceIndex:
    """Tests for UpdateAnimationInput sequence_index field."""

    def test_sequence_index_valid(self):
        """sequence_index=1 with a change param is accepted."""
        inp = UpdateAnimationInput(
            slide_index=1, animation_index=1,
            sequence_index=1, duration=1.0,
        )
        assert inp.sequence_index == 1
        assert inp.duration == 1.0

    def test_sequence_index_alone_raises(self):
        """sequence_index alone (no change param) raises ValidationError."""
        with pytest.raises(ValidationError, match="At least one optional parameter"):
            UpdateAnimationInput(
                slide_index=1, animation_index=1, sequence_index=1,
            )


# ============================================================================
# OneDrive URL resolver
# ============================================================================
from utils.onedrive import resolve_local_path


class TestOneDriveResolver:
    """Tests for OneDrive URL to local path resolution."""

    def test_local_path_passthrough(self):
        """A regular local path is returned unchanged."""
        path = r"C:\Users\test\Documents\presentation.pptx"
        assert resolve_local_path(path) == path

    def test_local_path_unc_passthrough(self):
        """A UNC path is returned unchanged."""
        path = r"\\server\share\file.pptx"
        assert resolve_local_path(path) == path

    def test_none_on_unknown_url(self):
        """An unknown URL returns None when registry is empty and env vars unset."""
        from utils import onedrive
        with patch.object(onedrive.winreg, "OpenKey", side_effect=FileNotFoundError):
            with patch.dict("os.environ", {}, clear=True):
                result = resolve_local_path("https://unknown.example.com/file.pptx")
                assert result is None

    def test_url_decoding_spaces(self):
        """URL-encoded spaces (%20) are decoded correctly."""
        with patch("utils.onedrive._resolve_via_registry") as mock_reg:
            mock_reg.return_value = None
            with patch.dict(
                "os.environ",
                {"OneDriveConsumer": r"C:\Users\test\OneDrive"},
                clear=True,
            ):
                with patch("os.path.isdir", return_value=True):
                    result = resolve_local_path(
                        "https://d.docs.live.net/ABC123/My%20Documents/file.pptx"
                    )
                    assert result is not None
                    assert "My Documents" in result
                    assert "%20" not in result

    def test_url_decoding_japanese(self):
        """URL-encoded Japanese characters are decoded correctly."""
        with patch("utils.onedrive._resolve_via_registry") as mock_reg:
            mock_reg.return_value = None
            with patch.dict(
                "os.environ",
                {"OneDriveConsumer": r"C:\Users\test\OneDrive"},
                clear=True,
            ):
                with patch("os.path.isdir", return_value=True):
                    # Japanese "ドキュメント" URL-encoded
                    encoded = "%E3%83%89%E3%82%AD%E3%83%A5%E3%83%A1%E3%83%B3%E3%83%88"
                    result = resolve_local_path(
                        f"https://d.docs.live.net/ABC123/{encoded}/file.pptx"
                    )
                    assert result is not None
                    assert "ドキュメント" in result

    def test_registry_resolution(self):
        """Registry-based resolution replaces URL prefix with mount point."""
        from utils import onedrive

        mock_providers_key = MagicMock()
        mock_subkey = MagicMock()

        with patch.object(
            onedrive.winreg, "OpenKey",
            side_effect=[mock_providers_key, mock_subkey],
        ), patch.object(
            onedrive.winreg, "EnumKey",
            side_effect=["Personal", OSError],
        ), patch.object(
            onedrive.winreg, "QueryValueEx",
            side_effect=[
                ("https://d.docs.live.net/ABC123", None),  # UrlNamespace
                (r"C:\Users\test\OneDrive", None),  # MountPoint
            ],
        ), patch.object(
            onedrive.winreg, "CloseKey",
        ):
            result = onedrive._resolve_via_registry(
                "https://d.docs.live.net/ABC123/Documents/test.pptx"
            )
            assert result == r"C:\Users\test\OneDrive\Documents\test.pptx"

    def test_registry_resolution_without_cid(self):
        """Registry UrlNamespace without CID still resolves correctly."""
        from utils import onedrive

        mock_providers_key = MagicMock()
        mock_subkey = MagicMock()

        # UrlNamespace lacks the CID — just "https://d.docs.live.net"
        with patch.object(
            onedrive.winreg, "OpenKey",
            side_effect=[mock_providers_key, mock_subkey],
        ), patch.object(
            onedrive.winreg, "EnumKey",
            side_effect=["Personal", OSError],
        ), patch.object(
            onedrive.winreg, "QueryValueEx",
            side_effect=[
                ("https://d.docs.live.net", None),  # UrlNamespace (no CID)
                (r"C:\Users\test\OneDrive", None),  # MountPoint
            ],
        ), patch.object(
            onedrive.winreg, "CloseKey",
        ):
            result = onedrive._resolve_via_registry(
                "https://d.docs.live.net/45333604723378ea/Udemy%E8%AC%9B%E5%BA%A7/file.pptx"
            )
            # CID should be stripped, Japanese should be decoded
            assert result == r"C:\Users\test\OneDrive\Udemy講座\file.pptx"

    def test_empty_string_passthrough(self):
        """Empty string is returned as-is (not a URL)."""
        assert resolve_local_path("") == ""

    def test_exception_returns_none(self):
        """If an unexpected exception occurs, None is returned."""
        with patch("utils.onedrive._resolve_via_registry", side_effect=Exception("boom")):
            with patch("utils.onedrive._resolve_via_env", side_effect=Exception("boom")):
                # The outer try/except in resolve_local_path catches this
                result = resolve_local_path("https://d.docs.live.net/ABC/file.pptx")
                assert result is None


# ============================================================================
# export.py — ExportImagesInput (file_name)
# ============================================================================
from ppt_com.export import ExportImagesInput


class TestExportImagesInputFileName:
    """Tests for ExportImagesInput file_name field."""

    def test_file_name_default_none(self):
        """file_name defaults to None."""
        inp = ExportImagesInput(output_dir="C:/tmp")
        assert inp.file_name is None

    def test_file_name_with_slide_index(self):
        """file_name is accepted alongside slide_index."""
        inp = ExportImagesInput(
            output_dir="C:/tmp", slide_index=1, file_name="cover.png",
        )
        assert inp.file_name == "cover.png"
        assert inp.slide_index == 1

    def test_file_name_without_extension(self):
        """file_name without extension is accepted (extension added at runtime)."""
        inp = ExportImagesInput(
            output_dir="C:/tmp", slide_index=1, file_name="cover",
        )
        assert inp.file_name == "cover"

    def test_file_name_without_slide_index_raises(self):
        """file_name without slide_index raises ValidationError."""
        with pytest.raises(ValidationError, match="file_name requires slide_index"):
            ExportImagesInput(
                output_dir="C:/tmp", file_name="cover.png",
            )


# ============================================================================
# text.py — FormatTextInput / FormatTextRangeInput (highlight_color)
# ============================================================================
from ppt_com.text import FormatTextInput, FormatTextRangeInput


class TestFormatTextHighlightColor:
    """Tests for highlight_color field on FormatTextInput."""

    def test_highlight_color_default_none(self):
        """highlight_color defaults to None."""
        inp = FormatTextInput(slide_index=1, shape_name_or_index=1, bold=True)
        assert inp.highlight_color is None

    def test_highlight_color_valid(self):
        """highlight_color accepts hex string."""
        inp = FormatTextInput(
            slide_index=1, shape_name_or_index=1,
            highlight_color="#FFFF00",
        )
        assert inp.highlight_color == "#FFFF00"

    def test_highlight_color_clear(self):
        """highlight_color accepts 'clear' to remove highlight."""
        inp = FormatTextInput(
            slide_index=1, shape_name_or_index=1,
            highlight_color="clear",
        )
        assert inp.highlight_color == "clear"

    def test_highlight_color_invalid_rejected(self):
        """Invalid highlight_color string raises ValidationError."""
        with pytest.raises(ValidationError, match="highlight_color must be"):
            FormatTextInput(
                slide_index=1, shape_name_or_index=1,
                highlight_color="not-a-color",
            )

    def test_highlight_color_missing_hash_rejected(self):
        """Hex without '#' prefix is rejected."""
        with pytest.raises(ValidationError, match="highlight_color must be"):
            FormatTextInput(
                slide_index=1, shape_name_or_index=1,
                highlight_color="FFFF00",
            )


class TestFormatTextRangeHighlightColor:
    """Tests for highlight_color field on FormatTextRangeInput."""

    def test_highlight_color_default_none(self):
        """highlight_color defaults to None."""
        inp = FormatTextRangeInput(
            slide_index=1, shape_name_or_index=1,
            start=1, length=5, bold=True,
        )
        assert inp.highlight_color is None

    def test_highlight_color_valid(self):
        """highlight_color accepts hex string."""
        inp = FormatTextRangeInput(
            slide_index=1, shape_name_or_index=1,
            start=1, length=5, highlight_color="#00FF00",
        )
        assert inp.highlight_color == "#00FF00"

    def test_highlight_color_clear(self):
        """highlight_color accepts 'clear' for range."""
        inp = FormatTextRangeInput(
            slide_index=1, shape_name_or_index=1,
            start=1, length=5, highlight_color="clear",
        )
        assert inp.highlight_color == "clear"

    def test_highlight_color_invalid_rejected(self):
        """Invalid highlight_color string raises ValidationError."""
        with pytest.raises(ValidationError, match="highlight_color must be"):
            FormatTextRangeInput(
                slide_index=1, shape_name_or_index=1,
                start=1, length=5, highlight_color="bad",
            )


# ============================================================================
# smartart.py — English alias mapping and _resolve_layout
# ============================================================================
from ppt_com.smartart import (
    SMARTART_ENGLISH_ALIASES,
    _ENGLISH_NAME_TO_ID,
    _resolve_layout,
)


class TestSmartArtEnglishAliases:
    """Tests for the SMARTART_ENGLISH_ALIASES mapping and reverse mapping."""

    def test_alias_dict_not_empty(self):
        """The alias mapping should have a substantial number of entries."""
        assert len(SMARTART_ENGLISH_ALIASES) >= 80

    def test_reverse_mapping_same_size(self):
        """The reverse mapping should have the same number of entries (no duplicate English names)."""
        assert len(_ENGLISH_NAME_TO_ID) == len(SMARTART_ENGLISH_ALIASES)

    def test_reverse_mapping_lowercase_keys(self):
        """All keys in the reverse mapping should be lowercase."""
        for key in _ENGLISH_NAME_TO_ID:
            assert key == key.lower(), f"Key '{key}' is not lowercase"

    def test_known_aliases(self):
        """Spot-check some well-known layout aliases."""
        assert SMARTART_ENGLISH_ALIASES["urn:microsoft.com/office/officeart/2005/8/layout/process1"] == "Basic Process"
        assert SMARTART_ENGLISH_ALIASES["urn:microsoft.com/office/officeart/2005/8/layout/orgChart1"] == "Organization Chart"
        assert SMARTART_ENGLISH_ALIASES["urn:microsoft.com/office/officeart/2005/8/layout/venn1"] == "Basic Venn"
        assert SMARTART_ENGLISH_ALIASES["urn:microsoft.com/office/officeart/2005/8/layout/matrix1"] == "Basic Matrix"
        assert SMARTART_ENGLISH_ALIASES["urn:microsoft.com/office/officeart/2005/8/layout/pyramid1"] == "Basic Pyramid"
        assert SMARTART_ENGLISH_ALIASES["urn:microsoft.com/office/officeart/2005/8/layout/cycle1"] == "Basic Cycle"
        assert SMARTART_ENGLISH_ALIASES["urn:microsoft.com/office/officeart/2005/8/layout/hierarchy1"] == "Hierarchy"

    def test_reverse_lookup(self):
        """Reverse mapping should return the correct Id for English names."""
        assert _ENGLISH_NAME_TO_ID["basic process"] == "urn:microsoft.com/office/officeart/2005/8/layout/process1"
        assert _ENGLISH_NAME_TO_ID["organization chart"] == "urn:microsoft.com/office/officeart/2005/8/layout/orgChart1"

    def test_alias_lookup_known(self):
        """SMARTART_ENGLISH_ALIASES.get returns the English name for a known Id."""
        assert SMARTART_ENGLISH_ALIASES.get("urn:microsoft.com/office/officeart/2005/8/layout/process1") == "Basic Process"

    def test_alias_lookup_unknown(self):
        """SMARTART_ENGLISH_ALIASES.get returns None for an unknown Id."""
        assert SMARTART_ENGLISH_ALIASES.get("urn:unknown/layout") is None

    def test_all_ids_are_urns(self):
        """All layout Ids should start with 'urn:microsoft.com/'."""
        for layout_id in SMARTART_ENGLISH_ALIASES:
            assert layout_id.startswith("urn:microsoft.com/"), f"Invalid URN: {layout_id}"

    def test_all_english_names_non_empty(self):
        """All English names should be non-empty strings."""
        for layout_id, name in SMARTART_ENGLISH_ALIASES.items():
            assert isinstance(name, str) and len(name) > 0, f"Empty name for {layout_id}"


class TestResolveLayout:
    """Tests for _resolve_layout with mock COM objects."""

    def _make_mock_app(self, layouts):
        """Create a mock app with SmartArtLayouts collection.

        layouts: list of (Name, Id) tuples
        """
        app = MagicMock()
        app.SmartArtLayouts.Count = len(layouts)

        def layout_getter(idx):
            mock_layout = MagicMock()
            mock_layout.Name = layouts[idx - 1][0]
            mock_layout.Id = layouts[idx - 1][1]
            return mock_layout

        app.SmartArtLayouts.side_effect = layout_getter
        return app

    def test_locale_name_match(self):
        """Should match by locale-specific name first."""
        app = self._make_mock_app([
            ("基本プロセス", "urn:microsoft.com/office/officeart/2005/8/layout/process1"),
        ])
        result = _resolve_layout(app, "基本プロセス")
        assert result.Name == "基本プロセス"

    def test_locale_name_partial_match(self):
        """Should match partial locale name (contains)."""
        app = self._make_mock_app([
            ("基本プロセス", "urn:microsoft.com/office/officeart/2005/8/layout/process1"),
        ])
        result = _resolve_layout(app, "プロセス")
        assert result.Name == "基本プロセス"

    def test_english_alias_exact_match(self):
        """Should match English alias when locale name doesn't match."""
        app = self._make_mock_app([
            ("基本プロセス", "urn:microsoft.com/office/officeart/2005/8/layout/process1"),
        ])
        result = _resolve_layout(app, "Basic Process")
        assert result.Name == "基本プロセス"

    def test_english_alias_partial_match(self):
        """Should match partial English alias (contains)."""
        app = self._make_mock_app([
            ("組織図", "urn:microsoft.com/office/officeart/2005/8/layout/orgChart1"),
            ("基本プロセス", "urn:microsoft.com/office/officeart/2005/8/layout/process1"),
        ])
        result = _resolve_layout(app, "Organization")
        assert result.Id == "urn:microsoft.com/office/officeart/2005/8/layout/orgChart1"

    def test_english_alias_case_insensitive(self):
        """English alias matching should be case-insensitive."""
        app = self._make_mock_app([
            ("基本プロセス", "urn:microsoft.com/office/officeart/2005/8/layout/process1"),
        ])
        result = _resolve_layout(app, "basic process")
        assert result.Name == "基本プロセス"

    def test_not_found_raises_error(self):
        """Should raise ValueError when no layout matches."""
        app = self._make_mock_app([
            ("基本プロセス", "urn:microsoft.com/office/officeart/2005/8/layout/process1"),
        ])
        with pytest.raises(ValueError, match="not found"):
            _resolve_layout(app, "NonexistentLayout")

    def test_locale_name_takes_priority(self):
        """Locale name match should take priority over English alias."""
        app = self._make_mock_app([
            ("Basic Process", "urn:microsoft.com/office/officeart/2005/8/layout/process1"),
        ])
        result = _resolve_layout(app, "Basic Process")
        assert result.Name == "Basic Process"

    def test_english_alias_with_unmapped_layout_id(self):
        """Layout with unmapped Id should not match English alias search."""
        app = self._make_mock_app([
            ("カスタムレイアウト", "urn:custom/unknown-layout"),
        ])
        with pytest.raises(ValueError, match="not found"):
            _resolve_layout(app, "Basic Process")
