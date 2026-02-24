"""Table operations for PowerPoint COM automation.

Handles creating tables, setting cell text/formatting, merging cells,
adding/deleting rows/columns, and applying table styles.
"""

import json
import logging
from typing import List, Optional, Union

from pydantic import BaseModel, Field, ConfigDict, model_validator

from utils.com_wrapper import ppt
from utils.color import hex_to_int, int_to_hex
from utils.navigation import goto_slide
from ppt_com.constants import (
    msoTrue, msoFalse,
    ppAlignLeft, ppAlignCenter, ppAlignRight, ppAlignJustify,
    msoAnchorTop, msoAnchorMiddle, msoAnchorBottom,
    ppBorderTop, ppBorderLeft, ppBorderBottom, ppBorderRight,
    ppBorderDiagonalDown, ppBorderDiagonalUp,
    msoLineSolid, msoLineRoundDot, msoLineDot, msoLineDash,
    msoLineDashDot, msoLineDashDotDot, msoLineLongDash, msoLineLongDashDot,
)

logger = logging.getLogger(__name__)

ALIGNMENT_MAP: dict[str, int] = {
    "left": ppAlignLeft,
    "center": ppAlignCenter,
    "right": ppAlignRight,
    "justify": ppAlignJustify,
}

VERTICAL_ALIGNMENT_MAP: dict[str, int] = {
    "top": msoAnchorTop,
    "middle": msoAnchorMiddle,
    "bottom": msoAnchorBottom,
}

BORDER_SIDE_MAP: dict[str, int] = {
    "top": ppBorderTop,
    "left": ppBorderLeft,
    "bottom": ppBorderBottom,
    "right": ppBorderRight,
    "diagonal_down": ppBorderDiagonalDown,
    "diagonal_up": ppBorderDiagonalUp,
}

DASH_STYLE_MAP: dict[str, int] = {
    "solid": msoLineSolid,
    "round_dot": msoLineRoundDot,
    "dot": msoLineDot,
    "dash": msoLineDash,
    "dash_dot": msoLineDashDot,
    "dash_dot_dot": msoLineDashDotDot,
    "long_dash": msoLineLongDash,
    "long_dash_dot": msoLineLongDashDot,
}

VERTICAL_ANCHOR_NAMES: dict[int, str] = {
    **{v: k for k, v in VERTICAL_ALIGNMENT_MAP.items()},
    2: "top_baseline",    # msoAnchorTopBaseline — valid but not settable via this tool
    5: "bottom_baseline", # msoAnchorBottomBaseLine — valid but not settable via this tool
}
ALIGNMENT_NAMES: dict[int, str] = {v: k for k, v in ALIGNMENT_MAP.items()}


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class AddTableInput(BaseModel):
    """Input for adding a table to a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    rows: int = Field(..., ge=1, description="Number of rows")
    cols: int = Field(..., ge=1, description="Number of columns")
    left: float = Field(default=50.0, description="Left position in points")
    top: float = Field(default=100.0, description="Top position in points")
    width: float = Field(default=600.0, description="Width in points")
    height: float = Field(default=300.0, description="Height in points")
    row_heights: Optional[List[float]] = Field(default=None, description="Row heights in points. If shorter than the row count, remaining rows keep their default height.")
    col_widths: Optional[List[float]] = Field(default=None, description="Column widths in points. If shorter than the column count, remaining columns keep their default width.")


class GetTableDataInput(BaseModel):
    """Input for getting table data."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Table shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    include_format: bool = Field(default=False, description="If True, also return a 'format' key containing a 2D array of cell formatting details (font, fill, alignment). 'data' always remains List[List[str]].")


class SetTableCellInput(BaseModel):
    """Input for setting text and formatting of a table cell."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Table shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    row: int = Field(..., ge=1, description="1-based row number")
    col: int = Field(..., ge=1, description="1-based column number")
    text: Optional[str] = Field(default=None, description="Cell text")
    font_name: Optional[str] = Field(default=None, description="Latin font name (e.g. 'Arial'). Also sets the East Asian font unless font_name_fareast is provided.")
    font_name_fareast: Optional[str] = Field(default=None, description="East Asian (CJK) font name (e.g. 'BIZ UDPゴシック'). Overrides the Far East font independently of font_name.")
    font_size: Optional[float] = Field(default=None, description="Font size in points")
    bold: Optional[bool] = Field(default=None, description="Bold on/off")
    italic: Optional[bool] = Field(default=None, description="Italic on/off")
    color: Optional[str] = Field(default=None, description="Font color as '#RRGGBB'")
    fill_color: Optional[str] = Field(default=None, description="Cell background color as '#RRGGBB'")
    alignment: Optional[str] = Field(
        default=None, description="Text alignment: 'left', 'center', 'right', or 'justify'"
    )
    vertical_alignment: Optional[str] = Field(
        default=None, description="Vertical text alignment: 'top', 'middle', or 'bottom'"
    )


class MergeTableCellsInput(BaseModel):
    """Input for merging table cells."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Table shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    start_row: int = Field(..., ge=1, description="Top-left cell row (1-based)")
    start_col: int = Field(..., ge=1, description="Top-left cell column (1-based)")
    end_row: int = Field(..., ge=1, description="Bottom-right cell row (1-based)")
    end_col: int = Field(..., ge=1, description="Bottom-right cell column (1-based)")

    @model_validator(mode="after")
    def _check_range_order(self) -> "MergeTableCellsInput":
        if self.end_row < self.start_row:
            raise ValueError(f"end_row ({self.end_row}) must be >= start_row ({self.start_row})")
        if self.end_col < self.start_col:
            raise ValueError(f"end_col ({self.end_col}) must be >= start_col ({self.start_col})")
        return self


class TableRowInput(BaseModel):
    """Input for adding or deleting a table row."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Table shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    position: Optional[int] = Field(
        default=None, ge=1,
        description="1-based row position. For add: insert before this row (omit to append). For delete: row to remove.",
    )
    height: Optional[float] = Field(default=None, description="Row height in points (only used when adding a row)")


class TableColumnInput(BaseModel):
    """Input for adding or deleting a table column."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Table shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    position: Optional[int] = Field(
        default=None, ge=1,
        description="1-based column position. For add: insert before this column (omit to append). For delete: column to remove.",
    )
    width: Optional[float] = Field(default=None, description="Column width in points (only used when adding a column)")


class SetTableStyleInput(BaseModel):
    """Input for applying a table style."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Table shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    style_id: Optional[str] = Field(
        default=None,
        description="Table style GUID (e.g. '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}')",
    )
    first_row: Optional[bool] = Field(default=None, description="Enable header row special formatting")
    last_row: Optional[bool] = Field(default=None, description="Enable total row special formatting")
    first_col: Optional[bool] = Field(default=None, description="Enable first column special formatting")
    last_col: Optional[bool] = Field(default=None, description="Enable last column special formatting")
    banding_rows: Optional[bool] = Field(default=None, description="Enable alternating row bands")
    banding_cols: Optional[bool] = Field(default=None, description="Enable alternating column bands")


class SetTableLayoutInput(BaseModel):
    """Input for setting row heights and/or column widths of an existing table."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Table shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    row_heights: Optional[List[float]] = Field(
        default=None,
        description="Heights for rows in points, indexed from row 1. If the list is shorter than the row count, remaining rows are unchanged."
    )
    col_widths: Optional[List[float]] = Field(
        default=None,
        description="Widths for columns in points, indexed from col 1. If the list is shorter than the column count, remaining columns are unchanged."
    )


class SplitTableCellsInput(BaseModel):
    """Input for splitting (unmerging) a merged table cell."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Table shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    row: int = Field(..., ge=1, description="1-based row of the cell to split")
    col: int = Field(..., ge=1, description="1-based column of the cell to split")
    num_rows: int = Field(default=1, ge=1, description="Number of rows in the resulting split (default 1 = simple unmerge)")
    num_cols: int = Field(default=1, ge=1, description="Number of columns in the resulting split (default 1 = simple unmerge)")


class SetTableBordersInput(BaseModel):
    """Input for setting borders on a range of table cells."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Table shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    start_row: int = Field(default=1, ge=1, description="First row of the target range (1-based, default=1)")
    start_col: int = Field(default=1, ge=1, description="First column of the target range (1-based, default=1)")
    end_row: Optional[int] = Field(default=None, ge=1, description="Last row of the target range (1-based, default=last row)")
    end_col: Optional[int] = Field(default=None, ge=1, description="Last column of the target range (1-based, default=last column)")
    sides: List[str] = Field(
        ..., min_length=1,
        description="Which borders to set: any of 'top', 'bottom', 'left', 'right', 'diagonal_down', 'diagonal_up'"
    )
    visible: Optional[bool] = Field(default=None, description="Show or hide the border")
    color: Optional[str] = Field(default=None, description="Border color as '#RRGGBB'")
    weight: Optional[float] = Field(default=None, description="Border line weight in points (e.g. 1.5)")
    dash_style: Optional[str] = Field(
        default=None,
        description="Border line style: 'solid', 'round_dot', 'dot', 'dash', 'dash_dot', 'dash_dot_dot', 'long_dash', 'long_dash_dot'"
    )

    @model_validator(mode="after")
    def _check_range_order(self) -> "SetTableBordersInput":
        if self.end_row is not None and self.end_row < self.start_row:
            raise ValueError(f"end_row ({self.end_row}) must be >= start_row ({self.start_row})")
        if self.end_col is not None and self.end_col < self.start_col:
            raise ValueError(f"end_col ({self.end_col}) must be >= start_col ({self.start_col})")
        return self

    @model_validator(mode="after")
    def _require_at_least_one_property(self) -> "SetTableBordersInput":
        if self.visible is None and self.color is None and self.weight is None and self.dash_style is None:
            raise ValueError(
                "At least one of visible, color, weight, or dash_style must be provided."
            )
        return self


# ---------------------------------------------------------------------------
# Helper: find a table shape
# ---------------------------------------------------------------------------
def _get_table_shape(slide, name_or_index: Union[str, int]):
    """Find a shape on a slide and verify it is a table.

    Args:
        slide: Slide COM object
        name_or_index: Shape name (str) or 1-based index (int)

    Returns:
        Shape COM object that contains a table

    Raises:
        ValueError: If shape not found or is not a table
    """
    if isinstance(name_or_index, int):
        if name_or_index < 1 or name_or_index > slide.Shapes.Count:
            raise ValueError(
                f"Shape index {name_or_index} out of range "
                f"(1-{slide.Shapes.Count})"
            )
        shape = slide.Shapes(name_or_index)
    else:
        shape = None
        for i in range(1, slide.Shapes.Count + 1):
            if slide.Shapes(i).Name == name_or_index:
                shape = slide.Shapes(i)
                break
        if shape is None:
            raise ValueError(f"Shape '{name_or_index}' not found on slide")

    if not shape.HasTable:
        raise ValueError(f"Shape '{shape.Name}' is not a table")
    return shape


# ---------------------------------------------------------------------------
# COM implementation functions (run on COM thread via ppt.execute)
# ---------------------------------------------------------------------------
def _add_table_impl(slide_index, rows, cols, left, top, width, height, row_heights, col_widths):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = slide.Shapes.AddTable(
        NumRows=rows, NumColumns=cols,
        Left=left, Top=top, Width=width, Height=height,
    )
    table = shape.Table
    if row_heights:
        for i, h in enumerate(row_heights, 1):
            if i <= table.Rows.Count:
                table.Rows(i).Height = h
    if col_widths:
        for i, w in enumerate(col_widths, 1):
            if i <= table.Columns.Count:
                table.Columns(i).Width = w
    return {
        "success": True,
        "shape_name": shape.Name,
        "shape_index": shape.ZOrderPosition,
        "rows": table.Rows.Count,
        "columns": table.Columns.Count,
    }


def _get_cell_format(cell) -> dict:
    """Extract formatting details from a table cell."""
    tf = cell.Shape.TextFrame
    tr = tf.TextRange
    font = tr.Font
    result = {}
    try:
        result["fill_color"] = int_to_hex(cell.Shape.Fill.ForeColor.RGB)
    except Exception:
        result["fill_color"] = None
    try:
        result["font_name"] = font.Name
    except Exception:
        result["font_name"] = None
    try:
        result["font_name_fareast"] = font.NameFarEast
    except Exception:
        result["font_name_fareast"] = None
    try:
        result["font_size"] = font.Size
    except Exception:
        result["font_size"] = None
    try:
        result["bold"] = bool(font.Bold == msoTrue)
    except Exception:
        result["bold"] = None
    try:
        result["italic"] = bool(font.Italic == msoTrue)
    except Exception:
        result["italic"] = None
    try:
        result["color"] = int_to_hex(font.Color.RGB)
    except Exception:
        result["color"] = None
    try:
        result["alignment"] = ALIGNMENT_NAMES.get(tr.ParagraphFormat.Alignment)
    except Exception:
        result["alignment"] = None
    try:
        result["vertical_alignment"] = VERTICAL_ANCHOR_NAMES.get(tf.VerticalAnchor)
    except Exception:
        result["vertical_alignment"] = None
    return result


def _get_table_data_impl(slide_index, shape_name_or_index, include_format):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.Table

    rows_count = table.Rows.Count
    cols_count = table.Columns.Count
    data = []
    fmt = [] if include_format else None
    for r in range(1, rows_count + 1):
        row_data = []
        row_fmt = [] if include_format else None
        for c in range(1, cols_count + 1):
            cell = table.Cell(r, c)
            row_data.append(cell.Shape.TextFrame.TextRange.Text)
            if include_format:
                row_fmt.append(_get_cell_format(cell))
        data.append(row_data)
        if include_format:
            fmt.append(row_fmt)

    result = {
        "success": True,
        "shape_name": shape.Name,
        "rows": rows_count,
        "columns": cols_count,
        "data": data,
    }
    if include_format:
        result["format"] = fmt
    return result


def _set_table_cell_impl(
    slide_index, shape_name_or_index,
    row, col, text,
    font_name, font_name_fareast, font_size, bold, italic, color,
    fill_color, alignment, vertical_alignment,
):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.Table
    cell = table.Cell(row, col)

    # Text access: Cell.Shape.TextFrame.TextRange
    if text is not None:
        cell.Shape.TextFrame.TextRange.Text = text.replace("\n", "\r")

    tr = cell.Shape.TextFrame.TextRange
    font = tr.Font

    if font_name is not None:
        font.Name = font_name
        font.NameFarEast = font_name  # default: match Latin unless overridden
    if font_name_fareast is not None:
        font.NameFarEast = font_name_fareast
    if font_size is not None:
        font.Size = font_size
    if bold is not None:
        font.Bold = msoTrue if bold else msoFalse
    if italic is not None:
        font.Italic = msoTrue if italic else msoFalse
    if color is not None:
        font.Color.RGB = hex_to_int(color)

    if alignment is not None:
        align_key = alignment.strip().lower()
        if align_key not in ALIGNMENT_MAP:
            raise ValueError(
                f"Unknown alignment '{alignment}'. Use: {', '.join(ALIGNMENT_MAP.keys())}"
            )
        tr.ParagraphFormat.Alignment = ALIGNMENT_MAP[align_key]

    # Cell fill
    if fill_color is not None:
        cell.Shape.Fill.Visible = msoTrue
        cell.Shape.Fill.ForeColor.RGB = hex_to_int(fill_color)

    if vertical_alignment is not None:
        va_key = vertical_alignment.strip().lower()
        if va_key not in VERTICAL_ALIGNMENT_MAP:
            raise ValueError(
                f"Unknown vertical_alignment '{vertical_alignment}'. Use: {', '.join(VERTICAL_ALIGNMENT_MAP.keys())}"
            )
        cell.Shape.TextFrame.VerticalAnchor = VERTICAL_ALIGNMENT_MAP[va_key]

    return {
        "success": True,
        "row": row,
        "col": col,
        "text": cell.Shape.TextFrame.TextRange.Text,
    }


def _merge_table_cells_impl(slide_index, shape_name_or_index, start_row, start_col, end_row, end_col):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.Table

    cell_from = table.Cell(start_row, start_col)
    cell_to = table.Cell(end_row, end_col)
    cell_from.Merge(cell_to)

    return {
        "success": True,
        "merged": f"Cell({start_row},{start_col}) to Cell({end_row},{end_col})",
    }


def _add_table_row_impl(slide_index, shape_name_or_index, position, height):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.Table

    if position is not None:
        new_row = table.Rows.Add(BeforeRow=position)
    else:
        new_row = table.Rows.Add()

    if height is not None:
        new_row.Height = height

    return {
        "success": True,
        "new_row_count": table.Rows.Count,
        "new_column_count": table.Columns.Count,
    }


def _delete_table_row_impl(slide_index, shape_name_or_index, position):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.Table

    if position is None:
        raise ValueError("position is required for deleting a row")
    table.Rows(position).Delete()

    return {
        "success": True,
        "new_row_count": table.Rows.Count,
        "new_column_count": table.Columns.Count,
    }


def _add_table_column_impl(slide_index, shape_name_or_index, position, width):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.Table

    if position is not None:
        new_col = table.Columns.Add(BeforeColumn=position)
    else:
        new_col = table.Columns.Add()

    if width is not None:
        new_col.Width = width

    return {
        "success": True,
        "new_row_count": table.Rows.Count,
        "new_column_count": table.Columns.Count,
    }


def _delete_table_column_impl(slide_index, shape_name_or_index, position):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.Table

    if position is None:
        raise ValueError("position is required for deleting a column")
    table.Columns(position).Delete()

    return {
        "success": True,
        "new_row_count": table.Rows.Count,
        "new_column_count": table.Columns.Count,
    }


def _set_table_style_impl(
    slide_index, shape_name_or_index, style_id,
    first_row, last_row, first_col, last_col,
    banding_rows, banding_cols,
):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.Table

    if style_id is not None:
        table.ApplyStyle(style_id, False)

    if first_row is not None:
        table.FirstRow = msoTrue if first_row else msoFalse
    if last_row is not None:
        table.LastRow = msoTrue if last_row else msoFalse
    if first_col is not None:
        table.FirstCol = msoTrue if first_col else msoFalse
    if last_col is not None:
        table.LastCol = msoTrue if last_col else msoFalse
    if banding_rows is not None:
        table.HorizBanding = msoTrue if banding_rows else msoFalse
    if banding_cols is not None:
        table.VertBanding = msoTrue if banding_cols else msoFalse

    return {
        "success": True,
        "style_applied": style_id is not None,
    }


def _set_table_layout_impl(slide_index, shape_name_or_index, row_heights, col_widths):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.Table

    if row_heights is not None:
        for i, h in enumerate(row_heights, 1):
            if i <= table.Rows.Count:
                table.Rows(i).Height = h

    if col_widths is not None:
        for i, w in enumerate(col_widths, 1):
            if i <= table.Columns.Count:
                table.Columns(i).Width = w

    return {
        "success": True,
        "row_heights": [table.Rows(i).Height for i in range(1, table.Rows.Count + 1)],
        "col_widths": [table.Columns(i).Width for i in range(1, table.Columns.Count + 1)],
    }


def _split_table_cells_impl(slide_index, shape_name_or_index, row, col, num_rows, num_cols):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.Table
    cell = table.Cell(row, col)
    cell.Split(num_rows, num_cols)
    return {
        "success": True,
        "row": row,
        "col": col,
        "num_rows": num_rows,
        "num_cols": num_cols,
    }


def _set_table_borders_impl(
    slide_index, shape_name_or_index,
    start_row, start_col, end_row, end_col,
    sides, visible, color, weight, dash_style,
):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.Table

    actual_end_row = end_row if end_row is not None else table.Rows.Count
    actual_end_col = end_col if end_col is not None else table.Columns.Count

    side_constants = []
    for side_name in sides:
        key = side_name.strip().lower()
        if key not in BORDER_SIDE_MAP:
            raise ValueError(
                f"Unknown border side '{side_name}'. Use: {', '.join(BORDER_SIDE_MAP.keys())}"
            )
        side_constants.append(BORDER_SIDE_MAP[key])

    color_int = hex_to_int(color) if color is not None else None

    dash_style_int = None
    if dash_style is not None:
        key = dash_style.strip().lower()
        if key not in DASH_STYLE_MAP:
            raise ValueError(
                f"Unknown dash_style '{dash_style}'. Use: {', '.join(DASH_STYLE_MAP.keys())}"
            )
        dash_style_int = DASH_STYLE_MAP[key]

    cells_updated = 0
    for r in range(start_row, actual_end_row + 1):
        for c in range(start_col, actual_end_col + 1):
            cell = table.Cell(r, c)
            for border_type in side_constants:
                border = cell.Borders.Item(border_type)
                if visible is not None:
                    border.Visible = msoTrue if visible else msoFalse
                if color_int is not None:
                    border.ForeColor.RGB = color_int
                if weight is not None:
                    border.Weight = weight
                if dash_style_int is not None:
                    border.DashStyle = dash_style_int
            if side_constants:
                cells_updated += 1

    return {
        "success": True,
        "cells_updated": cells_updated,
        "rows": f"{start_row}-{actual_end_row}",
        "cols": f"{start_col}-{actual_end_col}",
    }


# ---------------------------------------------------------------------------
# MCP tool functions (async wrappers that delegate to COM thread)
# ---------------------------------------------------------------------------
def add_table(params: AddTableInput) -> str:
    """Add a table to a slide.

    Args:
        params: Table parameters including rows, cols, position, and size.

    Returns:
        JSON with shape name, index, and table dimensions.
    """
    try:
        result = ppt.execute(
            _add_table_impl,
            params.slide_index, params.rows, params.cols,
            params.left, params.top, params.width, params.height,
            params.row_heights, params.col_widths,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to add table: {str(e)}"})


def get_table_data(params: GetTableDataInput) -> str:
    """Get all cell text values from a table as a 2D array.

    Args:
        params: Slide index and table shape identifier.

    Returns:
        JSON with row/column counts and 2D data array.
    """
    try:
        result = ppt.execute(
            _get_table_data_impl,
            params.slide_index, params.shape_name_or_index,
            params.include_format,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to get table data: {str(e)}"})


def set_table_cell(params: SetTableCellInput) -> str:
    """Set text and/or formatting for a table cell.

    Args:
        params: Cell location and optional text/formatting properties.

    Returns:
        JSON confirming the cell update.
    """
    try:
        result = ppt.execute(
            _set_table_cell_impl,
            params.slide_index, params.shape_name_or_index,
            params.row, params.col, params.text,
            params.font_name, params.font_name_fareast, params.font_size,
            params.bold, params.italic,
            params.color, params.fill_color, params.alignment,
            params.vertical_alignment,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to set table cell: {str(e)}"})


def merge_table_cells(params: MergeTableCellsInput) -> str:
    """Merge a range of table cells.

    Args:
        params: Cell range from (start_row, start_col) to (end_row, end_col).

    Returns:
        JSON confirming the merge.
    """
    try:
        result = ppt.execute(
            _merge_table_cells_impl,
            params.slide_index, params.shape_name_or_index,
            params.start_row, params.start_col,
            params.end_row, params.end_col,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to merge cells: {str(e)}"})


def add_table_row(params: TableRowInput) -> str:
    """Add a row to a table.

    Args:
        params: Table identifier and optional position to insert before.

    Returns:
        JSON with updated row/column counts.
    """
    try:
        result = ppt.execute(
            _add_table_row_impl,
            params.slide_index, params.shape_name_or_index,
            params.position, params.height,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to add table row: {str(e)}"})


def delete_table_row(params: TableRowInput) -> str:
    """Delete a row from a table.

    Args:
        params: Table identifier and row position to delete.

    Returns:
        JSON with updated row/column counts.
    """
    try:
        result = ppt.execute(
            _delete_table_row_impl,
            params.slide_index, params.shape_name_or_index,
            params.position,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to delete table row: {str(e)}"})


def add_table_column(params: TableColumnInput) -> str:
    """Add a column to a table.

    Args:
        params: Table identifier and optional position to insert before.

    Returns:
        JSON with updated row/column counts.
    """
    try:
        result = ppt.execute(
            _add_table_column_impl,
            params.slide_index, params.shape_name_or_index,
            params.position, params.width,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to add table column: {str(e)}"})


def delete_table_column(params: TableColumnInput) -> str:
    """Delete a column from a table.

    Args:
        params: Table identifier and column position to delete.

    Returns:
        JSON with updated row/column counts.
    """
    try:
        result = ppt.execute(
            _delete_table_column_impl,
            params.slide_index, params.shape_name_or_index,
            params.position,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to delete table column: {str(e)}"})


def set_table_style(params: SetTableStyleInput) -> str:
    """Apply a table style and configure banding options.

    Args:
        params: Style GUID and optional banding/header flags.

    Returns:
        JSON confirming the style application.
    """
    try:
        result = ppt.execute(
            _set_table_style_impl,
            params.slide_index, params.shape_name_or_index,
            params.style_id,
            params.first_row, params.last_row,
            params.first_col, params.last_col,
            params.banding_rows, params.banding_cols,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to set table style: {str(e)}"})


def set_table_layout(params: SetTableLayoutInput) -> str:
    """Set row heights and/or column widths for an existing table.

    Args:
        params: Table identifier and optional row_heights / col_widths arrays.

    Returns:
        JSON with the actual row heights and column widths after setting.
    """
    try:
        result = ppt.execute(
            _set_table_layout_impl,
            params.slide_index, params.shape_name_or_index,
            params.row_heights, params.col_widths,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to set table layout: {str(e)}"})


def split_table_cells(params: SplitTableCellsInput) -> str:
    """Split (unmerge) a merged table cell.

    Args:
        params: Cell location and target split dimensions.

    Returns:
        JSON confirming the split.
    """
    try:
        result = ppt.execute(
            _split_table_cells_impl,
            params.slide_index, params.shape_name_or_index,
            params.row, params.col,
            params.num_rows, params.num_cols,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to split table cells: {str(e)}"})


def set_table_borders(params: SetTableBordersInput) -> str:
    """Set borders on a range of table cells.

    Args:
        params: Cell range, sides to modify, and border properties.

    Returns:
        JSON with count of cells updated.
    """
    try:
        result = ppt.execute(
            _set_table_borders_impl,
            params.slide_index, params.shape_name_or_index,
            params.start_row, params.start_col,
            params.end_row, params.end_col,
            params.sides, params.visible,
            params.color, params.weight, params.dash_style,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to set table borders: {str(e)}"})


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register all table tools with the MCP server."""

    @mcp.tool(
        name="ppt_add_table",
        annotations={
            "title": "Add Table",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_add_table(params: AddTableInput) -> str:
        """Add a table to a slide.

        Creates a table with the specified number of rows and columns.
        All positions and sizes are in points (72 points = 1 inch).
        """
        return add_table(params)

    @mcp.tool(
        name="ppt_get_table_data",
        annotations={
            "title": "Get Table Data",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_get_table_data(params: GetTableDataInput) -> str:
        """Get all cell text values from a table.

        Returns a 2D array of cell text, plus row and column counts.
        Identify the table by shape name or 1-based shape index.
        """
        return get_table_data(params)

    @mcp.tool(
        name="ppt_set_table_cell",
        annotations={
            "title": "Set Table Cell",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_set_table_cell(params: SetTableCellInput) -> str:
        """Set text and/or formatting for a table cell.

        Access cell by 1-based row and column. Optionally set font properties,
        text alignment, and cell background color.
        """
        return set_table_cell(params)

    @mcp.tool(
        name="ppt_merge_table_cells",
        annotations={
            "title": "Merge Table Cells",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_merge_table_cells(params: MergeTableCellsInput) -> str:
        """Merge a range of table cells.

        Merges from (start_row, start_col) to (end_row, end_col).
        Uses Cell.Merge() which merges the entire rectangular range.
        """
        return merge_table_cells(params)

    @mcp.tool(
        name="ppt_add_table_row",
        annotations={
            "title": "Add Table Row",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_add_table_row(params: TableRowInput) -> str:
        """Add a row to a table.

        If position is provided, inserts before that row (1-based).
        If omitted, appends at the end.
        """
        return add_table_row(params)

    @mcp.tool(
        name="ppt_delete_table_row",
        annotations={
            "title": "Delete Table Row",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_delete_table_row(params: TableRowInput) -> str:
        """Delete a row from a table.

        Removes the row at the specified 1-based position.
        Remaining rows re-index automatically.
        """
        return delete_table_row(params)

    @mcp.tool(
        name="ppt_add_table_column",
        annotations={
            "title": "Add Table Column",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_add_table_column(params: TableColumnInput) -> str:
        """Add a column to a table.

        If position is provided, inserts before that column (1-based).
        If omitted, appends at the end.
        """
        return add_table_column(params)

    @mcp.tool(
        name="ppt_delete_table_column",
        annotations={
            "title": "Delete Table Column",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_delete_table_column(params: TableColumnInput) -> str:
        """Delete a column from a table.

        Removes the column at the specified 1-based position.
        Remaining columns re-index automatically.
        """
        return delete_table_column(params)

    @mcp.tool(
        name="ppt_set_table_style",
        annotations={
            "title": "Set Table Style",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_set_table_style(params: SetTableStyleInput) -> str:
        """Apply a table style and configure banding options.

        Use a style GUID to apply a built-in style (e.g.
        '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}' for Medium Style 2 - Accent 1).
        Optionally toggle header row, total row, banding rows/columns.
        """
        return set_table_style(params)

    @mcp.tool(
        name="ppt_set_table_layout",
        annotations={
            "title": "Set Table Layout",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_set_table_layout(params: SetTableLayoutInput) -> str:
        """Set row heights and/or column widths for an existing table.

        Provide row_heights (list of points per row) and/or col_widths
        (list of points per column), both 1-based and indexed from the first row/column.
        If either list is shorter than the table dimension, remaining rows/columns
        are left unchanged. Returns actual values after setting (PowerPoint may clamp
        to a minimum).
        """
        return set_table_layout(params)

    @mcp.tool(
        name="ppt_split_table_cells",
        annotations={
            "title": "Split Table Cells",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_split_table_cells(params: SplitTableCellsInput) -> str:
        """Split (unmerge) a merged table cell.

        Use num_rows=1, num_cols=1 (default) for a simple unmerge.
        Uses Cell.Split(NumRows, NumColumns) — the inverse of ppt_merge_table_cells.
        """
        return split_table_cells(params)

    @mcp.tool(
        name="ppt_set_table_borders",
        annotations={
            "title": "Set Table Borders",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_set_table_borders(params: SetTableBordersInput) -> str:
        """Set border style for a range of table cells.

        By default applies to the entire table (start_row=1, start_col=1,
        end_row/col=last row/col). Specify sides as a list of strings:
        'top', 'bottom', 'left', 'right', 'diagonal_down', 'diagonal_up'.
        Optionally set visible, color ('#RRGGBB'), weight (points), and dash_style.
        """
        return set_table_borders(params)
