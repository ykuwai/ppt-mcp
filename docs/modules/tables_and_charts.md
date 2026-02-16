# Module: Tables & Charts

## Overview

This module handles all operations related to PowerPoint tables and charts: creating tables, setting cell text and formatting, merging/splitting cells, table borders and backgrounds, row/column operations, table styles; creating charts, setting chart data via embedded Excel workbook, chart titles, legends, axes, series formatting, data labels, and chart type changes.

## Dependencies

- **Internal**: `utils.com_wrapper` (PowerPointCOMWrapper, safe_com_call), `utils.units` (inches_to_points, cm_to_points, points_to_inches, points_to_cm), `utils.color` (rgb_to_int, int_to_rgb, int_to_hex, hex_to_int), `ppt_com.constants` (all table/chart/border constants)
- **External**: `pywin32` (`win32com.client`, `pywintypes`)
- **Standard library**: `logging`

### Importing from Core Module

```python
from utils.com_wrapper import PowerPointCOMWrapper, safe_com_call
from utils.units import inches_to_points, points_to_inches, cm_to_points, points_to_cm
from utils.color import rgb_to_int, int_to_rgb, int_to_hex, hex_to_int
from ppt_com.constants import (
    # MsoTriState
    msoTrue, msoFalse,
    # Border types
    ppBorderTop, ppBorderLeft, ppBorderBottom, ppBorderRight,
    ppBorderDiagonalDown, ppBorderDiagonalUp,
    # Line dash styles
    msoLineSolid, msoLineDash, msoLineDot, msoLineDashDot,
    # Paragraph alignment
    ppAlignLeft, ppAlignCenter, ppAlignRight,
    # Chart types
    xlColumnClustered, xlColumnStacked, xlColumnStacked100,
    xlBarClustered, xlBarStacked,
    xlLine, xlLineMarkers, xlLineStacked,
    xlPie, xlPieExploded, xlDoughnut,
    xlArea, xlAreaStacked,
    xlXYScatter, xlXYScatterLines,
    xlRadar, xlBubble,
    xl3DColumnClustered, xl3DPie, xl3DLine,
    # Axis types
    xlCategory, xlValue,
    # Legend positions
    xlLegendPositionBottom, xlLegendPositionTop,
    xlLegendPositionLeft, xlLegendPositionRight,
)
import logging

logger = logging.getLogger(__name__)
```

## File Structure

```
ppt_com_mcp/
  ppt_com/
    tables.py    # Table operations
    charts.py    # Chart operations
```

---

## Constants Needed

These constants MUST be defined in `ppt_com/constants.py` (from the core module).

### Border Position Constants (ppBorderType)

| Name | Value | Description |
|------|-------|-------------|
| `ppBorderTop` | 1 | Top border |
| `ppBorderLeft` | 2 | Left border |
| `ppBorderBottom` | 3 | Bottom border |
| `ppBorderRight` | 4 | Right border |
| `ppBorderDiagonalDown` | 5 | Diagonal (top-left to bottom-right) |
| `ppBorderDiagonalUp` | 6 | Diagonal (bottom-left to top-right) |

### XlChartType (commonly used)

| Name | Value | Description |
|------|-------|-------------|
| `xlArea` | 1 | Area |
| `xlLine` | 4 | Line |
| `xlPie` | 5 | Pie |
| `xlBubble` | 15 | Bubble |
| `xlColumnClustered` | 51 | Clustered column |
| `xlColumnStacked` | 52 | Stacked column |
| `xlColumnStacked100` | 53 | 100% stacked column |
| `xl3DColumnClustered` | 54 | 3D clustered column |
| `xlBarClustered` | 57 | Clustered bar |
| `xlBarStacked` | 58 | Stacked bar |
| `xlLineStacked` | 63 | Stacked line |
| `xlLineMarkers` | 65 | Line with markers |
| `xlPieExploded` | 69 | Exploded pie |
| `xlXYScatterLines` | 74 | Scatter with lines |
| `xlAreaStacked` | 76 | Stacked area |
| `xlStockHLC` | 88 | Stock (High-Low-Close) |
| `xlXYScatter` | -4169 | Scatter |
| `xlDoughnut` | -4120 | Doughnut |
| `xl3DPie` | -4102 | 3D pie |
| `xl3DLine` | -4101 | 3D line |
| `xlRadar` | -4151 | Radar |

### XlAxisType

| Name | Value | Description |
|------|-------|-------------|
| `xlCategory` | 1 | Category axis (X axis) |
| `xlValue` | 2 | Value axis (Y axis) |
| `xlSeriesAxis` | 3 | Series axis (3D charts only) |

### XlLegendPosition

| Name | Value | Description |
|------|-------|-------------|
| `xlLegendPositionBottom` | -4107 | Bottom |
| `xlLegendPositionLeft` | -4131 | Left |
| `xlLegendPositionRight` | -4152 | Right |
| `xlLegendPositionTop` | -4160 | Top |
| `xlLegendPositionCorner` | 2 | Corner |

### Common Table Style GUIDs

| Style Name | GUID |
|-----------|------|
| No Style, No Grid | `{2D5ABB26-0587-4C30-8999-92F81FD0307C}` |
| Themed Style 1 - Accent 1 | `{3C2FFA5D-87B4-456A-9821-1D502468CF0F}` |
| Medium Style 2 - Accent 1 | `{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}` |
| Light Style 1 | `{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}` |
| Light Style 1 - Accent 1 | `{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}` |
| Light Style 2 | `{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}` |
| Dark Style 1 | `{E8034E78-7F5D-4C2E-B375-FC64B27BC917}` |
| Dark Style 2 | `{125E5076-3810-47DD-B79F-674D7AD40C01}` |

---

## File: `ppt_com/tables.py` - Table Operations

### Purpose

Provide MCP tools for creating and manipulating tables: add tables, set cell text, format cells (borders, fills, fonts), merge/split cells, add/delete rows/columns, resize rows/columns, and apply table styles.

---

### Tool: `add_table`

- **Description**: Add a table to a slide
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `rows` | int | Yes | Number of rows |
  | `columns` | int | Yes | Number of columns |
  | `left` | float | No | Left position in points. Default: 50 |
  | `top` | float | No | Top position in points. Default: 100 |
  | `width` | float | No | Width in points. Default: 600 |
  | `height` | float | No | Height in points. Default: 300 |
- **Returns**:
  ```json
  {
    "status": "success",
    "slide_index": 1,
    "shape_name": "Table 1",
    "shape_index": 3,
    "rows": 4,
    "columns": 3
  }
  ```
- **COM Implementation**:
  ```python
  def add_table(app, slide_index, rows, columns,
                left=50, top=100, width=600, height=300):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)

      shape = slide.Shapes.AddTable(
          NumRows=rows,
          NumColumns=columns,
          Left=left,
          Top=top,
          Width=width,
          Height=height,
      )
      # AddTable returns a Shape object; Table is accessed via shape.Table
      table = shape.Table

      return {
          "status": "success",
          "slide_index": slide_index,
          "shape_name": shape.Name,
          "shape_index": shape.ZOrderPosition,
          "rows": table.Rows.Count,
          "columns": table.Columns.Count,
      }
  ```
- **Error Cases**:
  - Invalid slide_index
  - rows or columns < 1

---

### Tool: `set_table_cell`

- **Description**: Set text and/or formatting for one or more table cells
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Table shape name or 1-based index |
  | `row` | int | Yes | 1-based row number |
  | `column` | int | Yes | 1-based column number |
  | `text` | string | No | Cell text |
  | `font_name` | string | No | Font name |
  | `font_size` | float | No | Font size in points |
  | `bold` | bool | No | Bold |
  | `italic` | bool | No | Italic |
  | `font_color` | string | No | Font color "#RRGGBB" |
  | `alignment` | int | No | 1=left, 2=center, 3=right |
  | `vertical_anchor` | int | No | 1=top, 3=middle, 4=bottom |
  | `fill_color` | string | No | Cell background color "#RRGGBB" |
  | `margin_left` | float | No | Cell left margin in points |
  | `margin_right` | float | No | Cell right margin in points |
  | `margin_top` | float | No | Cell top margin in points |
  | `margin_bottom` | float | No | Cell bottom margin in points |
- **Returns**:
  ```json
  {
    "status": "success",
    "row": 1,
    "column": 1,
    "text": "Header 1"
  }
  ```
- **COM Implementation**:
  ```python
  def set_table_cell(app, slide_index, shape_name_or_index,
                     row, column, text=None,
                     font_name=None, font_size=None,
                     bold=None, italic=None, font_color=None,
                     alignment=None, vertical_anchor=None,
                     fill_color=None,
                     margin_left=None, margin_right=None,
                     margin_top=None, margin_bottom=None):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = _get_table_shape(slide, shape_name_or_index)
      table = shape.Table
      cell = table.Cell(row, column)

      # Text access: Cell.Shape.TextFrame.TextRange
      if text is not None:
          cell.Shape.TextFrame.TextRange.Text = text.replace('\n', '\r')

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
      if font_color is not None:
          font.Color.RGB = hex_to_int(font_color)

      if alignment is not None:
          tr.ParagraphFormat.Alignment = alignment

      if vertical_anchor is not None:
          cell.Shape.TextFrame.VerticalAnchor = vertical_anchor

      # Cell fill
      if fill_color is not None:
          cell.Shape.Fill.Visible = msoTrue
          cell.Shape.Fill.ForeColor.RGB = hex_to_int(fill_color)

      # Cell margins
      tf = cell.Shape.TextFrame
      if margin_left is not None:
          tf.MarginLeft = margin_left
      if margin_right is not None:
          tf.MarginRight = margin_right
      if margin_top is not None:
          tf.MarginTop = margin_top
      if margin_bottom is not None:
          tf.MarginBottom = margin_bottom

      return {
          "status": "success",
          "row": row,
          "column": column,
          "text": text,
      }
  ```
- **Error Cases**:
  - Shape is not a table
  - Row or column out of range

### IMPORTANT: Table Cell Text Access Pattern

Table cells do NOT have a TextFrame directly. The access path is:

```
Table.Cell(row, col).Shape.TextFrame.TextRange
```

This is because each cell contains a Shape object that holds the text.

---

### Tool: `set_table_data`

- **Description**: Set data for multiple cells at once using a 2D array
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Table shape name or index |
  | `data` | list[list[string]] | Yes | 2D array of cell values. Outer list = rows, inner = columns |
  | `start_row` | int | No | Starting row (1-based). Default: 1 |
  | `start_column` | int | No | Starting column (1-based). Default: 1 |
- **Returns**:
  ```json
  {
    "status": "success",
    "cells_set": 12,
    "rows_affected": 4,
    "columns_affected": 3
  }
  ```
- **COM Implementation**:
  ```python
  def set_table_data(app, slide_index, shape_name_or_index,
                     data, start_row=1, start_column=1):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = _get_table_shape(slide, shape_name_or_index)
      table = shape.Table

      cells_set = 0
      for r_idx, row_data in enumerate(data):
          for c_idx, value in enumerate(row_data):
              row = start_row + r_idx
              col = start_column + c_idx
              if row <= table.Rows.Count and col <= table.Columns.Count:
                  cell = table.Cell(row, col)
                  cell.Shape.TextFrame.TextRange.Text = str(value)
                  cells_set += 1

      return {
          "status": "success",
          "cells_set": cells_set,
          "rows_affected": len(data),
          "columns_affected": max(len(r) for r in data) if data else 0,
      }
  ```
- **Error Cases**:
  - Shape is not a table
  - Data exceeds table dimensions (cells beyond bounds are silently skipped)

---

### Tool: `merge_cells`

- **Description**: Merge a range of table cells
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Table shape name or index |
  | `from_row` | int | Yes | Top-left cell row (1-based) |
  | `from_column` | int | Yes | Top-left cell column (1-based) |
  | `to_row` | int | Yes | Bottom-right cell row (1-based) |
  | `to_column` | int | Yes | Bottom-right cell column (1-based) |
- **Returns**:
  ```json
  {
    "status": "success",
    "merged": "Cell(1,1) to Cell(1,3)"
  }
  ```
- **COM Implementation**:
  ```python
  def merge_cells(app, slide_index, shape_name_or_index,
                  from_row, from_column, to_row, to_column):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = _get_table_shape(slide, shape_name_or_index)
      table = shape.Table

      # Merge from top-left to bottom-right
      cell_from = table.Cell(from_row, from_column)
      cell_to = table.Cell(to_row, to_column)
      cell_from.Merge(cell_to)

      return {
          "status": "success",
          "merged": f"Cell({from_row},{from_column}) to Cell({to_row},{to_column})",
      }
  ```
- **Error Cases**:
  - Shape is not a table
  - Cell coordinates out of range

---

### Tool: `split_cell`

- **Description**: Split a table cell into multiple rows/columns
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Table shape name or index |
  | `row` | int | Yes | Cell row (1-based) |
  | `column` | int | Yes | Cell column (1-based) |
  | `num_rows` | int | Yes | Number of rows to split into |
  | `num_columns` | int | Yes | Number of columns to split into |
- **Returns**:
  ```json
  {
    "status": "success",
    "split": "Cell(1,1) into 2x3"
  }
  ```
- **COM Implementation**:
  ```python
  def split_cell(app, slide_index, shape_name_or_index,
                 row, column, num_rows, num_columns):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = _get_table_shape(slide, shape_name_or_index)
      table = shape.Table

      cell = table.Cell(row, column)
      cell.Split(NumRows=num_rows, NumColumns=num_columns)

      return {
          "status": "success",
          "split": f"Cell({row},{column}) into {num_rows}x{num_columns}",
      }
  ```
- **Error Cases**:
  - Shape is not a table
  - Cell coordinates out of range

---

### Tool: `set_cell_borders`

- **Description**: Set border formatting for a table cell
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Table shape name or index |
  | `row` | int | Yes | Cell row (1-based) |
  | `column` | int | Yes | Cell column (1-based) |
  | `border_type` | int or string | No | Border: 1=top, 2=left, 3=bottom, 4=right, "all"=all four. Default: "all" |
  | `visible` | bool | No | Show/hide border |
  | `weight` | float | No | Border weight in points |
  | `color` | string | No | Border color "#RRGGBB" |
  | `dash_style` | int | No | 1=solid, 2=round dot, 3=dot, 4=dash, 5=dash-dot, 6=dash-dot-dot |
- **Returns**:
  ```json
  {
    "status": "success",
    "row": 1,
    "column": 1,
    "border_type": "all"
  }
  ```
- **COM Implementation**:
  ```python
  def set_cell_borders(app, slide_index, shape_name_or_index,
                       row, column, border_type="all",
                       visible=None, weight=None, color=None,
                       dash_style=None):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = _get_table_shape(slide, shape_name_or_index)
      table = shape.Table
      cell = table.Cell(row, column)

      # Determine which borders to set
      if border_type == "all":
          border_indices = [1, 2, 3, 4]  # top, left, bottom, right
      elif isinstance(border_type, int):
          border_indices = [border_type]
      else:
          raise ValueError(f"Invalid border_type: {border_type}")

      for bi in border_indices:
          border = cell.Borders(bi)
          if visible is not None:
              border.Visible = msoTrue if visible else msoFalse
          if weight is not None:
              border.Weight = weight
          if color is not None:
              border.ForeColor.RGB = hex_to_int(color)
          if dash_style is not None:
              border.DashStyle = dash_style

      return {
          "status": "success",
          "row": row,
          "column": column,
          "border_type": border_type,
      }
  ```
- **Error Cases**:
  - Shape is not a table
  - Cell coordinates out of range
  - Invalid border_type

---

### Tool: `modify_table_rows_columns`

- **Description**: Add/delete rows or columns, or resize them
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Table shape name or index |
  | `action` | string | Yes | "add_row", "add_column", "delete_row", "delete_column", "set_row_height", "set_column_width" |
  | `position` | int | No | For add: insert before this position. For delete/resize: target position. Omit for add = append at end |
  | `value` | float | No | For set_row_height/set_column_width: height/width in points |
- **Returns**:
  ```json
  {
    "status": "success",
    "action": "add_row",
    "position": 3,
    "new_row_count": 5,
    "new_column_count": 4
  }
  ```
- **COM Implementation**:
  ```python
  def modify_table_rows_columns(app, slide_index, shape_name_or_index,
                                action, position=None, value=None):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = _get_table_shape(slide, shape_name_or_index)
      table = shape.Table

      if action == "add_row":
          if position is not None:
              table.Rows.Add(BeforeRow=position)
          else:
              table.Rows.Add()  # Append at end
      elif action == "add_column":
          if position is not None:
              table.Columns.Add(BeforeColumn=position)
          else:
              table.Columns.Add()  # Append at end
      elif action == "delete_row":
          if position is None:
              raise ValueError("position is required for delete_row")
          table.Rows(position).Delete()
      elif action == "delete_column":
          if position is None:
              raise ValueError("position is required for delete_column")
          table.Columns(position).Delete()
      elif action == "set_row_height":
          if position is None or value is None:
              raise ValueError("position and value are required for set_row_height")
          table.Rows(position).Height = value
      elif action == "set_column_width":
          if position is None or value is None:
              raise ValueError("position and value are required for set_column_width")
          table.Columns(position).Width = value
      else:
          raise ValueError(f"Unknown action: {action}")

      return {
          "status": "success",
          "action": action,
          "position": position,
          "new_row_count": table.Rows.Count,
          "new_column_count": table.Columns.Count,
      }
  ```
- **Error Cases**:
  - Shape is not a table
  - Invalid action
  - Position out of range
  - Missing required parameters

---

### Tool: `apply_table_style`

- **Description**: Apply a table style and configure banding options
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Table shape name or index |
  | `style_id` | string | No | Table style GUID |
  | `save_formatting` | bool | No | Preserve existing cell formatting. Default: false |
  | `first_row` | bool | No | Enable header row special formatting |
  | `last_row` | bool | No | Enable total row special formatting |
  | `first_col` | bool | No | Enable first column special formatting |
  | `last_col` | bool | No | Enable last column special formatting |
  | `horiz_banding` | bool | No | Enable alternating row bands |
  | `vert_banding` | bool | No | Enable alternating column bands |
- **Returns**:
  ```json
  {
    "status": "success",
    "style_applied": true
  }
  ```
- **COM Implementation**:
  ```python
  def apply_table_style(app, slide_index, shape_name_or_index,
                        style_id=None, save_formatting=False,
                        first_row=None, last_row=None,
                        first_col=None, last_col=None,
                        horiz_banding=None, vert_banding=None):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = _get_table_shape(slide, shape_name_or_index)
      table = shape.Table

      if style_id is not None:
          table.ApplyStyle(style_id, save_formatting)

      if first_row is not None:
          table.FirstRow = msoTrue if first_row else msoFalse
      if last_row is not None:
          table.LastRow = msoTrue if last_row else msoFalse
      if first_col is not None:
          table.FirstCol = msoTrue if first_col else msoFalse
      if last_col is not None:
          table.LastCol = msoTrue if last_col else msoFalse
      if horiz_banding is not None:
          table.HorizBanding = msoTrue if horiz_banding else msoFalse
      if vert_banding is not None:
          table.VertBanding = msoTrue if vert_banding else msoFalse

      return {
          "status": "success",
          "style_applied": style_id is not None,
      }
  ```
- **Error Cases**:
  - Shape is not a table
  - Invalid style_id GUID

---

### Helper: `_get_table_shape`

```python
def _get_table_shape(slide, name_or_index):
    """Get a table shape and verify it is a table."""
    if isinstance(name_or_index, int):
        shape = slide.Shapes(name_or_index)
    else:
        shape = None
        for i in range(1, slide.Shapes.Count + 1):
            if slide.Shapes(i).Name == name_or_index:
                shape = slide.Shapes(i)
                break
        if shape is None:
            raise ValueError(f"Shape '{name_or_index}' not found")

    if not shape.HasTable:
        raise ValueError(f"Shape '{shape.Name}' is not a table")
    return shape
```

---

## File: `ppt_com/charts.py` - Chart Operations

### Purpose

Provide MCP tools for creating charts and manipulating chart data, titles, axes, legends, series formatting, and data labels.

---

### Tool: `add_chart`

- **Description**: Add a chart to a slide
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `chart_type` | int | Yes | XlChartType value (e.g., 51 for clustered column) |
  | `left` | float | No | Left position in points. Default: 50 |
  | `top` | float | No | Top position in points. Default: 50 |
  | `width` | float | No | Width in points. Default: 500 |
  | `height` | float | No | Height in points. Default: 350 |
  | `style` | int | No | Chart style number. Default: -1 (default) |
- **Returns**:
  ```json
  {
    "status": "success",
    "slide_index": 1,
    "shape_name": "Chart 1",
    "shape_index": 2,
    "chart_type": 51
  }
  ```
- **COM Implementation**:
  ```python
  def add_chart(app, slide_index, chart_type,
                left=50, top=50, width=500, height=350,
                style=-1):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)

      # AddChart2 is the modern method (Office 2013+)
      chart_shape = slide.Shapes.AddChart2(
          Style=style,
          Type=chart_type,
          Left=left,
          Top=top,
          Width=width,
          Height=height,
          NewLayout=True,
      )

      return {
          "status": "success",
          "slide_index": slide_index,
          "shape_name": chart_shape.Name,
          "shape_index": chart_shape.ZOrderPosition,
          "chart_type": chart_type,
      }
  ```
- **Error Cases**:
  - Invalid slide_index
  - Invalid chart_type

**IMPORTANT**: AddChart2 internally launches Excel to create the chart data workbook. This can take a moment and leaves an Excel process running briefly.

---

### Tool: `set_chart_data`

- **Description**: Set chart data by writing to the embedded Excel workbook
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Chart shape name or index |
  | `categories` | list[string] | Yes | Category labels (X axis) |
  | `series` | list[object] | Yes | Array of series objects: [{"name": "Sales", "values": [100, 200, 300]}, ...] |
- **Returns**:
  ```json
  {
    "status": "success",
    "categories_count": 4,
    "series_count": 2
  }
  ```
- **COM Implementation**:
  ```python
  def set_chart_data(app, slide_index, shape_name_or_index,
                     categories, series):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = _get_chart_shape(slide, shape_name_or_index)
      chart = shape.Chart

      # Activate chart data to access the embedded Excel workbook
      chart.ChartData.Activate()
      wb = chart.ChartData.Workbook
      ws = wb.Worksheets(1)

      try:
          # Clear existing data
          ws.Cells.Clear()

          # Write category labels (column A, starting from A2)
          for i, cat in enumerate(categories):
              ws.Range(f"A{i + 2}").Value = cat

          # Write series data
          for s_idx, s in enumerate(series):
              col_letter = chr(ord('B') + s_idx)  # B, C, D, ...

              # Series name in row 1
              ws.Range(f"{col_letter}1").Value = s["name"]

              # Series values starting from row 2
              for v_idx, val in enumerate(s["values"]):
                  ws.Range(f"{col_letter}{v_idx + 2}").Value = val

          # Set the data source range
          last_col = chr(ord('B') + len(series) - 1)
          last_row = len(categories) + 1
          data_range = ws.Range(f"A1:{last_col}{last_row}")
          chart.SetSourceData(data_range)

      finally:
          # CRITICAL: Always close the workbook to prevent Excel process leak
          wb.Close(SaveChanges=False)

      return {
          "status": "success",
          "categories_count": len(categories),
          "series_count": len(series),
      }
  ```
- **Error Cases**:
  - Shape is not a chart
  - Excel workbook activation fails
  - Data arrays have inconsistent lengths

### CRITICAL: Always Close the Workbook

After calling `chart.ChartData.Activate()`, you MUST call `wb.Close(SaveChanges=False)` in a finally block. Failing to do so will leave an orphaned Excel process that consumes memory and may lock files.

---

### Tool: `set_chart_title`

- **Description**: Set or remove the chart title
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Chart shape name or index |
  | `title_text` | string | No | Title text. Set to null/empty to remove title |
  | `font_size` | float | No | Title font size |
  | `bold` | bool | No | Title bold |
  | `color` | string | No | Title color "#RRGGBB" |
- **Returns**:
  ```json
  {
    "status": "success",
    "has_title": true,
    "title": "Monthly Sales Report"
  }
  ```
- **COM Implementation**:
  ```python
  def set_chart_title(app, slide_index, shape_name_or_index,
                      title_text=None, font_size=None,
                      bold=None, color=None):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = _get_chart_shape(slide, shape_name_or_index)
      chart = shape.Chart

      if title_text is not None and title_text != "":
          chart.HasTitle = True
          chart.ChartTitle.Text = title_text
          if font_size is not None:
              chart.ChartTitle.Font.Size = font_size
          if bold is not None:
              chart.ChartTitle.Font.Bold = bold
          if color is not None:
              chart.ChartTitle.Font.Color.RGB = hex_to_int(color)
      elif title_text == "":
          chart.HasTitle = False

      return {
          "status": "success",
          "has_title": chart.HasTitle,
          "title": chart.ChartTitle.Text if chart.HasTitle else None,
      }
  ```
- **Error Cases**:
  - Shape is not a chart

---

### Tool: `set_chart_legend`

- **Description**: Configure chart legend
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Chart shape name or index |
  | `visible` | bool | No | Show/hide legend |
  | `position` | int | No | Legend position: -4107=bottom, -4160=top, -4131=left, -4152=right, 2=corner |
  | `font_size` | float | No | Legend font size |
- **Returns**:
  ```json
  {
    "status": "success",
    "has_legend": true,
    "position": -4107
  }
  ```
- **COM Implementation**:
  ```python
  def set_chart_legend(app, slide_index, shape_name_or_index,
                       visible=None, position=None, font_size=None):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = _get_chart_shape(slide, shape_name_or_index)
      chart = shape.Chart

      if visible is not None:
          chart.HasLegend = visible

      if chart.HasLegend:
          if position is not None:
              chart.Legend.Position = position
          if font_size is not None:
              chart.Legend.Font.Size = font_size

      return {
          "status": "success",
          "has_legend": chart.HasLegend,
          "position": chart.Legend.Position if chart.HasLegend else None,
      }
  ```
- **Error Cases**:
  - Shape is not a chart

---

### Tool: `set_chart_axis`

- **Description**: Configure chart axis properties
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Chart shape name or index |
  | `axis_type` | int | Yes | 1=category (X), 2=value (Y) |
  | `title_text` | string | No | Axis title text |
  | `min_scale` | float | No | Minimum scale value (value axis only) |
  | `max_scale` | float | No | Maximum scale value (value axis only) |
  | `major_unit` | float | No | Major unit interval (value axis only) |
  | `number_format` | string | No | Number format for tick labels (e.g., "#,##0") |
  | `font_size` | float | No | Tick label font size |
  | `major_gridlines` | bool | No | Show/hide major gridlines |
  | `minor_gridlines` | bool | No | Show/hide minor gridlines |
- **Returns**:
  ```json
  {
    "status": "success",
    "axis_type": 2,
    "has_title": true,
    "title": "Revenue ($)"
  }
  ```
- **COM Implementation**:
  ```python
  def set_chart_axis(app, slide_index, shape_name_or_index,
                     axis_type, title_text=None,
                     min_scale=None, max_scale=None, major_unit=None,
                     number_format=None, font_size=None,
                     major_gridlines=None, minor_gridlines=None):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = _get_chart_shape(slide, shape_name_or_index)
      chart = shape.Chart

      axis = chart.Axes(axis_type)

      if title_text is not None:
          axis.HasTitle = True
          axis.AxisTitle.Text = title_text

      if min_scale is not None:
          axis.MinimumScale = min_scale
      if max_scale is not None:
          axis.MaximumScale = max_scale
      if major_unit is not None:
          axis.MajorUnit = major_unit

      if number_format is not None:
          axis.TickLabels.NumberFormat = number_format
      if font_size is not None:
          axis.TickLabels.Font.Size = font_size

      if major_gridlines is not None:
          axis.HasMajorGridlines = major_gridlines
      if minor_gridlines is not None:
          axis.HasMinorGridlines = minor_gridlines

      return {
          "status": "success",
          "axis_type": axis_type,
          "has_title": axis.HasTitle,
          "title": axis.AxisTitle.Text if axis.HasTitle else None,
      }
  ```
- **Error Cases**:
  - Shape is not a chart
  - Invalid axis_type
  - Setting min/max on category axis (only works for value axis)

---

### Tool: `format_chart_series`

- **Description**: Format a specific data series (color, line weight, markers)
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Chart shape name or index |
  | `series_index` | int | Yes | 1-based series index |
  | `fill_color` | string | No | Series fill color "#RRGGBB" |
  | `line_color` | string | No | Series line/border color "#RRGGBB" |
  | `line_weight` | float | No | Line weight in points |
  | `show_data_labels` | bool | No | Show/hide data labels |
  | `data_label_show_value` | bool | No | Show values in data labels |
  | `data_label_show_category` | bool | No | Show category names in data labels |
  | `data_label_show_percentage` | bool | No | Show percentages (pie charts) |
  | `data_label_font_size` | float | No | Data label font size |
  | `data_label_number_format` | string | No | Data label number format |
- **Returns**:
  ```json
  {
    "status": "success",
    "series_index": 1,
    "series_name": "Sales"
  }
  ```
- **COM Implementation**:
  ```python
  def format_chart_series(app, slide_index, shape_name_or_index,
                          series_index, fill_color=None, line_color=None,
                          line_weight=None, show_data_labels=None,
                          data_label_show_value=None,
                          data_label_show_category=None,
                          data_label_show_percentage=None,
                          data_label_font_size=None,
                          data_label_number_format=None):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = _get_chart_shape(slide, shape_name_or_index)
      chart = shape.Chart
      series = chart.SeriesCollection(series_index)

      if fill_color is not None:
          series.Format.Fill.ForeColor.RGB = hex_to_int(fill_color)
      if line_color is not None:
          series.Format.Line.ForeColor.RGB = hex_to_int(line_color)
      if line_weight is not None:
          series.Format.Line.Weight = line_weight

      if show_data_labels is not None:
          series.HasDataLabels = show_data_labels

      if series.HasDataLabels:
          labels = series.DataLabels()
          if data_label_show_value is not None:
              labels.ShowValue = data_label_show_value
          if data_label_show_category is not None:
              labels.ShowCategoryName = data_label_show_category
          if data_label_show_percentage is not None:
              labels.ShowPercentage = data_label_show_percentage
          if data_label_font_size is not None:
              labels.Font.Size = data_label_font_size
          if data_label_number_format is not None:
              labels.NumberFormat = data_label_number_format

      return {
          "status": "success",
          "series_index": series_index,
          "series_name": series.Name,
      }
  ```
- **Error Cases**:
  - Shape is not a chart
  - series_index out of range
  - Setting data label properties when show_data_labels=False

---

### Tool: `change_chart_type`

- **Description**: Change the chart type of an existing chart
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Chart shape name or index |
  | `chart_type` | int | Yes | New XlChartType value |
- **Returns**:
  ```json
  {
    "status": "success",
    "new_chart_type": 4
  }
  ```
- **COM Implementation**:
  ```python
  def change_chart_type(app, slide_index, shape_name_or_index, chart_type):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = _get_chart_shape(slide, shape_name_or_index)
      chart = shape.Chart

      chart.ChartType = chart_type

      return {
          "status": "success",
          "new_chart_type": chart_type,
      }
  ```
- **Error Cases**:
  - Shape is not a chart
  - Incompatible chart type conversion (e.g., pie chart from multi-series data)

---

### Helper: `_get_chart_shape`

```python
def _get_chart_shape(slide, name_or_index):
    """Get a chart shape and verify it contains a chart."""
    if isinstance(name_or_index, int):
        shape = slide.Shapes(name_or_index)
    else:
        shape = None
        for i in range(1, slide.Shapes.Count + 1):
            if slide.Shapes(i).Name == name_or_index:
                shape = slide.Shapes(i)
                break
        if shape is None:
            raise ValueError(f"Shape '{name_or_index}' not found")

    if not shape.HasChart:
        raise ValueError(f"Shape '{shape.Name}' is not a chart")
    return shape
```

---

## Implementation Notes

### 1. Table Cell Access Path

Table cells use a unique access path for text: `Table.Cell(row, col).Shape.TextFrame.TextRange`. The Cell object itself does NOT have a TextFrame -- you must go through `Cell.Shape`:

```python
# CORRECT:
cell = table.Cell(1, 1)
cell.Shape.TextFrame.TextRange.Text = "Hello"
cell.Shape.TextFrame.TextRange.Font.Bold = msoTrue

# WRONG (will fail):
# cell.TextFrame.TextRange.Text = "Hello"
```

### 2. Table Cell Fill Access Path

Similarly, cell fill is accessed via `Cell.Shape.Fill`, NOT `Cell.Fill`:

```python
cell.Shape.Fill.Visible = msoTrue
cell.Shape.Fill.ForeColor.RGB = hex_to_int("#FF0000")
```

### 3. Table Cell Borders Access Path

Borders, however, are accessed directly from the Cell object:

```python
border = cell.Borders(3)  # ppBorderBottom
border.Visible = msoTrue
border.Weight = 2.0
border.ForeColor.RGB = hex_to_int("#000000")
```

### 4. Chart Data Requires Excel

Chart data is stored in an embedded Excel workbook. To modify data:
1. Call `chart.ChartData.Activate()` -- this starts an Excel process
2. Access `chart.ChartData.Workbook` to get the Workbook object
3. Modify cells via `ws.Range("A1").Value = ...`
4. Call `chart.SetSourceData(range)` to update the chart
5. **ALWAYS** call `wb.Close(SaveChanges=False)` in a finally block

### 5. Column Letter Calculation

For more than 26 series/columns, you need a proper column letter function:

```python
def col_letter(n):
    """Convert 1-based column number to Excel column letter(s)."""
    result = ""
    while n > 0:
        n -= 1
        result = chr(ord('A') + (n % 26)) + result
        n //= 26
    return result
```

### 6. Table Style GUIDs

Table style GUIDs are not well-documented. The GUIDs listed in the Constants section are the most commonly used ones. They come from the Open XML specification and are consistent across PowerPoint versions.

### 7. Merge/Split Complexity

Complex merge/split operations can make the table structure unstable. Best practices:
- Merge from top-left to bottom-right direction
- Avoid deeply nested merge/split sequences
- After extensive merging, verify the table structure

### 8. Chart Type Compatibility

Not all chart types are interchangeable. For example:
- Pie charts typically work with only 1 series
- Converting from a multi-series chart to pie may lose data
- 3D chart types may not support all formatting options

### 9. All Table/Chart Indices Are 1-Based

Consistent with all PowerPoint COM collections:
- `Table.Cell(1, 1)` = first cell (top-left)
- `Table.Rows(1)` = first row
- `chart.SeriesCollection(1)` = first series
- `chart.Axes(1)` = category axis

### 10. RGB Colors in Charts

Chart colors use the same BGR encoding as everywhere else in PowerPoint COM. Always use the `hex_to_int` helper from `utils/color.py`.
