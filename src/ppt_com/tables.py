"""Table operations for PowerPoint COM automation.

Handles creating tables, setting cell text/formatting, merging cells,
adding/deleting rows/columns, and applying table styles.
"""

import json
import logging
from typing import Optional, Union

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from utils.color import hex_to_int
from utils.navigation import goto_slide
from ppt_com.constants import (
    msoTrue, msoFalse,
    ppAlignLeft, ppAlignCenter, ppAlignRight,
)

logger = logging.getLogger(__name__)

ALIGNMENT_MAP: dict[str, int] = {
    "left": ppAlignLeft,
    "center": ppAlignCenter,
    "right": ppAlignRight,
}


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


class GetTableDataInput(BaseModel):
    """Input for getting table data."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Table shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )


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
    font_name: Optional[str] = Field(default=None, description="Font name (e.g. 'Arial')")
    font_size: Optional[float] = Field(default=None, description="Font size in points")
    bold: Optional[bool] = Field(default=None, description="Bold on/off")
    italic: Optional[bool] = Field(default=None, description="Italic on/off")
    color: Optional[str] = Field(default=None, description="Font color as '#RRGGBB'")
    fill_color: Optional[str] = Field(default=None, description="Cell background color as '#RRGGBB'")
    alignment: Optional[str] = Field(
        default=None, description="Text alignment: 'left', 'center', or 'right'"
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
def _add_table_impl(slide_index, rows, cols, left, top, width, height):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = slide.Shapes.AddTable(
        NumRows=rows, NumColumns=cols,
        Left=left, Top=top, Width=width, Height=height,
    )
    table = shape.Table
    return {
        "success": True,
        "shape_name": shape.Name,
        "shape_index": shape.ZOrderPosition,
        "rows": table.Rows.Count,
        "columns": table.Columns.Count,
    }


def _get_table_data_impl(slide_index, shape_name_or_index):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.Table

    rows_count = table.Rows.Count
    cols_count = table.Columns.Count
    data = []
    for r in range(1, rows_count + 1):
        row_data = []
        for c in range(1, cols_count + 1):
            cell = table.Cell(r, c)
            text = cell.Shape.TextFrame.TextRange.Text
            row_data.append(text)
        data.append(row_data)

    return {
        "success": True,
        "shape_name": shape.Name,
        "rows": rows_count,
        "columns": cols_count,
        "data": data,
    }


def _set_table_cell_impl(
    slide_index, shape_name_or_index,
    row, col, text,
    font_name, font_size, bold, italic, color,
    fill_color, alignment,
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


def _add_table_row_impl(slide_index, shape_name_or_index, position):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.Table

    if position is not None:
        table.Rows.Add(BeforeRow=position)
    else:
        table.Rows.Add()

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


def _add_table_column_impl(slide_index, shape_name_or_index, position):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.Table

    if position is not None:
        table.Columns.Add(BeforeColumn=position)
    else:
        table.Columns.Add()

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
            params.font_name, params.font_size, params.bold, params.italic,
            params.color, params.fill_color, params.alignment,
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
            params.position,
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
            params.position,
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
