"""Tests for Pydantic model validators in ppt_com modules.

Covers all model_validator decorated methods in:
- freeform.py: NodeSpec, BuildFreeformInput, InsertNodeInput
- tables.py: MergeTableCellsInput, SetTableBordersInput
- advanced_ops.py: SetDefaultShapeStyleInput, CropPictureInput
- shapes.py: AddShapeInput
- layout.py: SetSlideBackgroundInput
- text.py: GetAllTextInput

These are pure Python tests — no COM or PowerPoint required.
"""

import sys

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
)
from ppt_com.advanced_ops import SetDefaultShapeStyleInput, CropPictureInput
from ppt_com.shapes import AddShapeInput, UpdateShapeInput
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
