"""Table tools, on Apple Events.

Mirrors ``ppt_com/tables.py``. Same function names, same signatures, same
returned shapes; what differs is the walk through PowerPoint's object model.

Three things about tables on this side are worth knowing before reading on.

**Cells are reached with a command, not by index.** ``get cell from`` takes the
``table object`` of a ``shape table`` plus a row and a column. Element indexing
under ``shape table`` did not line up in earlier testing (MACOS_PORT section 9,
item 2), so every cell in this module goes through the command, which did.

**Row and column counts come from the shape, not from counting elements.**
``number of rows`` and ``number of columns`` are declared on ``shape table``
and are one cheap read each, which keeps the flaky part of the object model out
of the answer.

**``line format`` has no ``visible``.** Borders are hidden by making them fully
transparent instead, and the result says so rather than pretending the field
was honoured as asked.
"""

import logging

from appscript import k
from appscript.reference import CommandError

from backend.mac_ae import count, elements, is_missing, ppt, raw, shapes_of
from backend.mac_enums import (
    MsoLineDashStyle,
    MsoVerticalAnchor,
    PpBorderType,
    PpParagraphAlignment,
    to_keyword,
)
from utils.color import hex_to_rgb_list, rgb_list_to_hex
from utils.navigation import goto_slide
from ppt_com.constants import msoLineDot

logger = logging.getLogger(__name__)

# `dash style` is both a property of `line format` and a family of enumerator
# names, and appscript resolves the collision toward the enumerators, so the
# property is only reachable by its raw four character code.
_DASH_STYLE_CODE = b'LFds'

# scripts/gen_mac_enums.py pairs by name, and Windows calls this one `msoLineDot`
# while macOS calls it `line dash style square dot`, so the pairing was missed
# and the generated table has no entry for 3. macOS does have the style, so this
# is a gap in the generator rather than a gap in PowerPoint, and it should
# disappear the next time the table is regenerated.
_DASH_STYLES = dict(MsoLineDashStyle)
_DASH_STYLES.setdefault(msoLineDot, k.line_dash_style_square_dot)


def _windows_constant(table, keyword, default=None):
    """Turn a macOS enumerator back into the Windows constant it stands for."""
    for value, word in table.items():
        if word == keyword:
            return value
    return default


def _get_table_shape(slide, name_or_index):
    """Find a shape on a slide and verify it is a table."""
    shapes = shapes_of(slide)
    if isinstance(name_or_index, int):
        if name_or_index < 1 or name_or_index > len(shapes):
            raise ValueError(
                f"Shape index {name_or_index} out of range "
                f"(1-{len(shapes)})"
            )
        shape = shapes[name_or_index - 1]
    else:
        shape = None
        for candidate in shapes:
            if candidate.name() == name_or_index:
                shape = candidate
                break
        if shape is None:
            raise ValueError(f"Shape '{name_or_index}' not found on slide")

    if not shape.has_table():
        raise ValueError(f"Shape '{shape.name()}' is not a table")
    return shape


def _dimensions(shape):
    """Return (rows, columns) for a table shape."""
    return shape.number_of_rows(), shape.number_of_columns()


def _cell(table, row, col):
    """Reach one cell. See the module docstring for why this is a command."""
    return table.get_cell_from(row=row, column=col)


def _cell_text(cell):
    """Read a cell's text, treating `missing value` as an empty cell."""
    content = cell.shape.text_frame.text_range.content()
    return "" if is_missing(content) else content


def _refusal(tool_name: str, reason: str, alternatives=None) -> dict:
    """The body a tool returns when macOS genuinely cannot do it.

    ``backend.unsupported.unsupported`` builds the same payload but returns it
    already encoded, and these functions hand a dict back to a caller that
    encodes it. Same keys, same reading, one less round of JSON.
    """
    payload = {
        "error": f"{tool_name} is not available on macOS",
        "reason": reason,
        "platform": "macOS",
    }
    if alternatives:
        payload["alternatives"] = alternatives
    return payload


def _apply_cell_font(cell, font_name, font_name_fareast, font_size, bold, italic, color):
    """Apply font properties to a cell, through `font of text range`.

    Reaching the font any other way answers -10006.
    """
    font = cell.shape.text_frame.text_range.font
    if font_name is not None:
        font.font_name.set(font_name)
        font.east_asian_name.set(font_name)  # default: match Latin unless overridden
    if font_name_fareast is not None:
        font.east_asian_name.set(font_name_fareast)
    if font_size is not None:
        font.font_size.set(font_size)
    if bold is not None:
        font.bold.set(bool(bold))
    if italic is not None:
        font.italic.set(bool(italic))
    if color is not None:
        font.font_color.set(hex_to_rgb_list(color))


# ---------------------------------------------------------------------------
# Apple Event implementation functions
# ---------------------------------------------------------------------------
def _add_table_impl(slide_index, rows, cols, left, top, width, height, row_heights, col_widths):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]

    before = count(slide.shapes)
    # The insertion location is the slide itself, never its shapes. `at` given
    # as `slide.shapes.end` raises -1708, and the error does not say why.
    # `number of rows` and `number of columns` are marked read only in the
    # dictionary yet `make` accepts them, which is the only route to a table of
    # the requested size.
    app.make(
        new=k.shape_table,
        at=slide.end,
        with_properties={
            k.number_of_rows: rows,
            k.number_of_columns: cols,
            k.left_position: left,
            k.top: top,
            k.width: width,
            k.height: height,
        },
    )

    # Nothing is trusted because it did not raise.
    shapes_now = shapes_of(slide)
    if len(shapes_now) != before + 1:
        return _refusal(
            "ppt_add_table",
            "PowerPoint reported success but the slide gained no shape, which "
            "is the silent no-op recorded in MACOS_PORT section 5.",
        )

    # `make`'s own return value is not used. Some of PowerPoint's Standard
    # Suite verbs hand back something appscript cannot unpack, and a new shape
    # lands at the end of the z order anyway, so the shape is fetched again by
    # a route that is known to work.
    shape = shapes_now[-1]
    try:
        is_table = shape.has_table()
    except CommandError:
        is_table = False
    if not is_table:
        return _refusal(
            "ppt_add_table",
            "PowerPoint created a shape but not a table. A `make` that falls "
            "through leaves an empty autoshape behind, so nothing was written.",
        )

    table = shape.table_object
    if row_heights:
        table_rows = elements(table.rows)
        for i, h in enumerate(row_heights, 1):
            if i <= len(table_rows):
                table_rows[i - 1].height.set(h)
    if col_widths:
        table_cols = elements(table.columns)
        for i, w in enumerate(col_widths, 1):
            if i <= len(table_cols):
                table_cols[i - 1].width.set(w)

    actual_rows, actual_cols = _dimensions(shape)
    return {
        "success": True,
        "shape_name": shape.name(),
        "shape_index": shape.z_order_position(),
        "rows": actual_rows,
        "columns": actual_cols,
    }


def _get_cell_format(cell) -> dict:
    """Extract formatting details from a table cell."""
    tf = cell.shape.text_frame
    font = tf.text_range.font
    result = {}
    try:
        result["fill_color_hex"] = rgb_list_to_hex(cell.shape.fill_format.fore_color())
    except Exception:
        result["fill_color_hex"] = None
    try:
        result["font_name"] = font.font_name()
    except Exception:
        result["font_name"] = None
    try:
        result["font_name_fareast"] = font.east_asian_name()
    except Exception:
        result["font_name_fareast"] = None
    try:
        size = font.font_size()
        result["font_size"] = None if is_missing(size) else size
    except Exception:
        result["font_size"] = None
    try:
        result["bold"] = bool(font.bold())
    except Exception:
        result["bold"] = None
    try:
        result["italic"] = bool(font.italic())
    except Exception:
        result["italic"] = None
    try:
        result["font_color_hex"] = rgb_list_to_hex(font.font_color())
    except Exception:
        result["font_color_hex"] = None
    try:
        from ppt_com.tables import ALIGNMENT_NAMES

        result["alignment"] = ALIGNMENT_NAMES.get(
            _windows_constant(PpParagraphAlignment, tf.text_range.paragraph_format.alignment())
        )
    except Exception:
        result["alignment"] = None
    try:
        from ppt_com.tables import VERTICAL_ANCHOR_NAMES

        result["vertical_alignment"] = VERTICAL_ANCHOR_NAMES.get(
            _windows_constant(MsoVerticalAnchor, tf.vertical_anchor())
        )
    except Exception:
        result["vertical_alignment"] = None
    return result


def _get_table_data_impl(slide_index, shape_name_or_index, include_format):
    ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.table_object

    rows_count, cols_count = _dimensions(shape)
    data = []
    fmt = [] if include_format else None
    for r in range(1, rows_count + 1):
        row_data = []
        row_fmt = [] if include_format else None
        for c in range(1, cols_count + 1):
            cell = _cell(table, r, c)
            row_data.append(_cell_text(cell))
            if include_format:
                row_fmt.append(_get_cell_format(cell))
        data.append(row_data)
        if include_format:
            fmt.append(row_fmt)

    result = {
        "success": True,
        "shape_name": shape.name(),
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
    # Lazy import: ppt_com/tables.py imports this module at the bottom of its
    # own file, so importing it back at module scope would let an
    # "import ppt_mac.tables first" ordering run that swap block against a
    # module that has defined nothing yet, and the swap would silently not
    # happen. By call time both modules are fully loaded.
    from ppt_com.tables import ALIGNMENT_MAP, VERTICAL_ALIGNMENT_MAP

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.table_object
    cell = _cell(table, row, col)

    # Text access: cell's shape's text frame's text range
    if text is not None:
        cell.shape.text_frame.text_range.content.set(text.replace("\n", "\r"))

    tr = cell.shape.text_frame.text_range
    _apply_cell_font(cell, font_name, font_name_fareast, font_size, bold, italic, color)

    if alignment is not None:
        align_key = alignment.strip().lower()
        if align_key not in ALIGNMENT_MAP:
            raise ValueError(
                f"Unknown alignment '{alignment}'. Use: {', '.join(ALIGNMENT_MAP.keys())}"
            )
        tr.paragraph_format.alignment.set(
            to_keyword(PpParagraphAlignment, ALIGNMENT_MAP[align_key], "paragraph alignment")
        )

    # Cell fill
    if fill_color is not None:
        cell.shape.fill_format.visible.set(True)
        cell.shape.fill_format.fore_color.set(hex_to_rgb_list(fill_color))

    if vertical_alignment is not None:
        va_key = vertical_alignment.strip().lower()
        if va_key not in VERTICAL_ALIGNMENT_MAP:
            raise ValueError(
                f"Unknown vertical_alignment '{vertical_alignment}'. Use: {', '.join(VERTICAL_ALIGNMENT_MAP.keys())}"
            )
        cell.shape.text_frame.vertical_anchor.set(
            to_keyword(MsoVerticalAnchor, VERTICAL_ALIGNMENT_MAP[va_key], "vertical anchor")
        )

    return {
        "success": True,
        "row": row,
        "col": col,
        # Read back rather than echoed, so a write that did not land shows up.
        "text": _cell_text(cell),
    }


def _set_table_data_impl(
    slide_index, shape_name_or_index, data,
    start_row, start_col, bold_first_row,
):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.table_object

    rows_count, cols_count = _dimensions(shape)

    cells_set = 0
    rows_written = 0
    for r_idx, row_data in enumerate(data):
        target_row = start_row + r_idx
        if target_row > rows_count:
            break
        row_had_writes = False
        for c_idx, cell_text in enumerate(row_data):
            target_col = start_col + c_idx
            if target_col > cols_count:
                break
            cell = _cell(table, target_row, target_col)
            cell.shape.text_frame.text_range.content.set(
                str(cell_text).replace("\n", "\r")
            )
            if bold_first_row and r_idx == 0:
                cell.shape.text_frame.text_range.font.bold.set(True)
            cells_set += 1
            row_had_writes = True
        if row_had_writes:
            rows_written += 1

    return {
        "success": True,
        "shape_name": shape.name(),
        "cells_set": cells_set,
        "rows_written": rows_written,
        "table_rows": rows_count,
        "table_columns": cols_count,
    }


def _merge_table_cells_impl(slide_index, shape_name_or_index, start_row, start_col, end_row, end_col):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.table_object

    cell_from = _cell(table, start_row, start_col)
    cell_to = _cell(table, end_row, end_col)
    cell_from.merge(merge_with=cell_to)

    return {
        "success": True,
        "merged": f"Cell({start_row},{start_col}) to Cell({end_row},{end_col})",
    }


def _add_table_row_impl(slide_index, shape_name_or_index, position, height):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.table_object

    before_rows, _ = _dimensions(shape)

    # PowerPoint's dictionary declares no command for adding a row, so this is
    # the Standard Suite `make`, which is also how shapes and slides get
    # created here even though that suite is absent too. It is attempted rather
    # than refused on principle, and the row count decides whether it worked.
    if position is not None:
        location = elements(table.rows)[position - 1].before
    else:
        location = table.end
    new_row = app.make(new=k.row, at=location)

    after_rows, after_cols = _dimensions(shape)
    if after_rows != before_rows + 1:
        return _refusal(
            "ppt_add_table_row",
            "PowerPoint reported no error but the table still has "
            f"{after_rows} rows. Adding a row is not scriptable here; rebuild "
            "the table at the size you need instead.",
            ["ppt_add_table"],
        )

    if height is not None:
        try:
            new_row.height.set(height)
        except CommandError:
            # The row exists either way, so the height is the only part lost.
            logger.warning("Could not set the height of the new row", exc_info=True)

    return {
        "success": True,
        "new_row_count": after_rows,
        "new_column_count": after_cols,
    }


def _delete_table_row_impl(slide_index, shape_name_or_index, position):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.table_object

    if position is None:
        raise ValueError("position is required for deleting a row")

    before_rows, _ = _dimensions(shape)
    rows = elements(table.rows)
    if position < 1 or position > len(rows):
        raise ValueError(f"Row {position} out of range (1-{len(rows)})")
    rows[position - 1].delete()

    after_rows, after_cols = _dimensions(shape)
    if after_rows != before_rows - 1:
        return _refusal(
            "ppt_delete_table_row",
            "PowerPoint reported no error but the table still has "
            f"{after_rows} rows. Deleting a row is not scriptable here.",
        )

    return {
        "success": True,
        "new_row_count": after_rows,
        "new_column_count": after_cols,
    }


def _add_table_column_impl(slide_index, shape_name_or_index, position, width):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.table_object

    _, before_cols = _dimensions(shape)

    if position is not None:
        location = elements(table.columns)[position - 1].before
    else:
        location = table.end
    new_col = app.make(new=k.column, at=location)

    after_rows, after_cols = _dimensions(shape)
    if after_cols != before_cols + 1:
        return _refusal(
            "ppt_add_table_column",
            "PowerPoint reported no error but the table still has "
            f"{after_cols} columns. Adding a column is not scriptable here; "
            "rebuild the table at the size you need instead.",
            ["ppt_add_table"],
        )

    if width is not None:
        try:
            new_col.width.set(width)
        except CommandError:
            logger.warning("Could not set the width of the new column", exc_info=True)

    return {
        "success": True,
        "new_row_count": after_rows,
        "new_column_count": after_cols,
    }


def _delete_table_column_impl(slide_index, shape_name_or_index, position):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.table_object

    if position is None:
        raise ValueError("position is required for deleting a column")

    _, before_cols = _dimensions(shape)
    cols = elements(table.columns)
    if position < 1 or position > len(cols):
        raise ValueError(f"Column {position} out of range (1-{len(cols)})")
    cols[position - 1].delete()

    after_rows, after_cols = _dimensions(shape)
    if after_cols != before_cols - 1:
        return _refusal(
            "ppt_delete_table_column",
            "PowerPoint reported no error but the table still has "
            f"{after_cols} columns. Deleting a column is not scriptable here.",
        )

    return {
        "success": True,
        "new_row_count": after_rows,
        "new_column_count": after_cols,
    }


def _set_table_style_impl(
    slide_index, shape_name_or_index, style_id,
    first_row, last_row, first_col, last_col,
    banding_rows, banding_cols,
):
    """Refuse, because macOS exposes no table style at all.

    Windows drives ``Table.ApplyStyle`` plus ``FirstRow``, ``LastRow``,
    ``FirstCol``, ``LastCol``, ``HorizBanding`` and ``VertBanding``. The whole
    ``table`` class on macOS carries one property, ``table direction``, so
    there is nothing here to apply a style GUID to and nothing to band.
    """
    return _refusal(
        "ppt_set_table_style",
        "PowerPoint for Mac's `table` class exposes only its text direction to "
        "Apple Events. There is no style, no header or total row flag and no "
        "banding, so a style GUID has nothing to apply to. Format the cells "
        "directly instead.",
        ["ppt_set_table_cell", "ppt_set_table_borders"],
    )


def _set_table_layout_impl(slide_index, shape_name_or_index, row_heights, col_widths):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.table_object

    table_rows = elements(table.rows)
    table_cols = elements(table.columns)

    if row_heights is not None:
        for i, h in enumerate(row_heights, 1):
            if i <= len(table_rows):
                table_rows[i - 1].height.set(h)

    if col_widths is not None:
        for i, w in enumerate(col_widths, 1):
            if i <= len(table_cols):
                table_cols[i - 1].width.set(w)

    # Read back rather than echoed. PowerPoint clamps to a minimum, and this is
    # also the check that the writes above landed at all.
    return {
        "success": True,
        "row_heights": [row.height() for row in elements(table.rows)],
        "col_widths": [col.width() for col in elements(table.columns)],
    }


def _split_table_cells_impl(slide_index, shape_name_or_index, row, col, num_rows, num_cols):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.table_object
    cell = _cell(table, row, col)
    cell.split(number_of_rows=num_rows, number_of_columns=num_cols)
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
    from ppt_com.tables import BORDER_SIDE_MAP, DASH_STYLE_MAP

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]
    shape = _get_table_shape(slide, shape_name_or_index)
    table = shape.table_object

    rows_count, cols_count = _dimensions(shape)
    actual_end_row = end_row if end_row is not None else rows_count
    actual_end_col = end_col if end_col is not None else cols_count

    side_keywords = []
    for side_name in sides:
        key = side_name.strip().lower()
        if key not in BORDER_SIDE_MAP:
            raise ValueError(
                f"Unknown border side '{side_name}'. Use: {', '.join(BORDER_SIDE_MAP.keys())}"
            )
        side_keywords.append(
            to_keyword(PpBorderType, BORDER_SIDE_MAP[key], "border side")
        )

    rgb = hex_to_rgb_list(color) if color is not None else None

    dash_keyword = None
    if dash_style is not None:
        key = dash_style.strip().lower()
        if key not in DASH_STYLE_MAP:
            raise ValueError(
                f"Unknown dash_style '{dash_style}'. Use: {', '.join(DASH_STYLE_MAP.keys())}"
            )
        dash_keyword = to_keyword(
            _DASH_STYLES, DASH_STYLE_MAP[key], "line dash style"
        )

    cells_updated = 0
    for r in range(start_row, actual_end_row + 1):
        for c in range(start_col, actual_end_col + 1):
            cell = _cell(table, r, c)
            for edge in side_keywords:
                border = cell.get_border(edge=edge)
                if visible is not None:
                    # `line format` has no `visible` property on macOS. Full
                    # transparency is the stand-in that was measured to apply
                    # cleanly, and it works on diagonals too, which the COM
                    # path cannot manage.
                    border.transparency.set(0.0 if visible else 1.0)
                if rgb is not None:
                    border.fore_color.set(rgb)
                if weight is not None:
                    border.line_weight.set(weight)
                if dash_keyword is not None:
                    raw(border, _DASH_STYLE_CODE).set(dash_keyword)
            if side_keywords:
                cells_updated += 1

    result = {
        "success": True,
        "cells_updated": cells_updated,
        "rows": f"{start_row}-{actual_end_row}",
        "cols": f"{start_col}-{actual_end_col}",
    }
    if visible is not None:
        result["warnings"] = [
            "PowerPoint for Mac's line format has no visible property, so "
            f"visible={visible} was applied as transparency "
            f"{0.0 if visible else 1.0}. The border still exists; it is drawn "
            "fully transparent."
        ]
    return result
