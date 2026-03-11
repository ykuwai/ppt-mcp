"""Text content, formatting, and manipulation tools for PowerPoint COM automation."""

import json
import logging
import os
import re
import time
import unicodedata
from typing import List, Optional, Union

from pydantic import BaseModel, Field, ConfigDict, field_validator, model_validator

from utils.com_wrapper import ppt
from utils.navigation import goto_slide
from utils.color import hex_to_int, int_to_hex, int_to_rgb, get_theme_color_index
from utils.validation import font_size_warning
from ppt_com.constants import (
    msoTrue, msoFalse, msoTriStateMixed,
    msoPlaceholder, msoGroup,
    ppAlignLeft, ppAlignCenter, ppAlignRight, ppAlignJustify, ppAlignDistribute,
    ppAutoSizeNone, ppAutoSizeShapeToFitText, ppAutoSizeTextToFitShape,
    ppBulletNone, ppBulletUnnumbered, ppBulletNumbered,
    msoTextOrientationHorizontal, msoTextOrientationVertical,
    msoTextOrientationUpward, msoTextOrientationDownward,
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------
def _get_shape(slide, name_or_index):
    """Find shape by name (str) or index (int).

    Args:
        slide: Slide COM object
        name_or_index: Shape name (str) or 1-based index (int)

    Returns:
        Shape COM object

    Raises:
        ValueError: If shape not found
    """
    if isinstance(name_or_index, int):
        if name_or_index < 1 or name_or_index > slide.Shapes.Count:
            raise ValueError(
                f"Shape index {name_or_index} out of range "
                f"(1-{slide.Shapes.Count})"
            )
        return slide.Shapes(name_or_index)
    else:
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            if shape.Name == name_or_index:
                return shape
        raise ValueError(f"Shape '{name_or_index}' not found on slide")


# ---------------------------------------------------------------------------
# Alignment / orientation / auto-size maps
# ---------------------------------------------------------------------------
ALIGNMENT_MAP = {
    "left": ppAlignLeft,
    "center": ppAlignCenter,
    "right": ppAlignRight,
    "justify": ppAlignJustify,
    "distribute": ppAlignDistribute,
}

ORIENTATION_MAP = {
    "horizontal": msoTextOrientationHorizontal,
    "vertical": msoTextOrientationVertical,
    "upward": msoTextOrientationUpward,
    "downward": msoTextOrientationDownward,
}

AUTO_SIZE_MAP = {
    "none": ppAutoSizeNone,
    "shape_to_fit": ppAutoSizeShapeToFitText,
    "shrink_to_fit": ppAutoSizeTextToFitShape,
}

BULLET_TYPE_MAP = {
    "none": ppBulletNone,
    "unnumbered": ppBulletUnnumbered,
    "numbered": ppBulletNumbered,
}


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class SetTextInput(BaseModel):
    """Input for setting text content of a shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    text: str = Field(
        ...,
        description=(
            "Text content. Use \\n for paragraph breaks (Enter) "
            "and \\v for line breaks within the same paragraph (Shift+Enter). "
            "Example: 'First paragraph\\nSecond paragraph' or "
            "'Line one\\vLine two' (same paragraph, no bullet/indent change)."
        ),
    )


class GetTextInput(BaseModel):
    """Input for getting text from a shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )


class FormatTextInput(BaseModel):
    """Input for formatting all text in a shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    font_name: Optional[str] = Field(default=None, description="Latin font name (e.g. 'Arial'). Also sets the East Asian font unless font_name_fareast is provided.")
    font_name_fareast: Optional[str] = Field(default=None, description="East Asian (CJK) font name (e.g. 'BIZ UDPゴシック'). Overrides the Far East font independently of font_name.")
    font_size: Optional[float] = Field(default=None, description="Font size in points")
    bold: Optional[bool] = Field(default=None, description="Bold on/off")
    italic: Optional[bool] = Field(default=None, description="Italic on/off")
    underline: Optional[bool] = Field(default=None, description="Underline on/off")
    color: Optional[str] = Field(default=None, description="Color as '#RRGGBB' hex string")
    font_color_theme: Optional[str] = Field(
        default=None,
        description="Theme color name (e.g. 'accent1', 'dark1')"
    )
    highlight_color: Optional[str] = Field(
        default=None,
        description="Text highlight (marker) color as '#RRGGBB' hex string, or 'clear' to remove highlight. Requires Office 2019+.",
    )

    @field_validator("highlight_color")
    @classmethod
    def validate_highlight_color(cls, v):
        if v is None:
            return v
        if v.lower() == "clear":
            return v
        if not re.fullmatch(r"#[0-9A-Fa-f]{6}", v):
            raise ValueError("highlight_color must be '#RRGGBB' hex string or 'clear'")
        return v


class FormatTextRangeInput(BaseModel):
    """Input for formatting a specific character range within a shape.

    Range can be specified either by start/length or by search_text.
    When search_text is provided, the matching text position is used
    automatically (start and length must not be set).
    """
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    start: Optional[int] = Field(default=None, description="1-based character start position (mutually exclusive with search_text)")
    length: Optional[int] = Field(default=None, description="Number of characters to format (mutually exclusive with search_text)")
    search_text: Optional[str] = Field(default=None, description="Text to search for in the shape. The matching range is formatted automatically. Mutually exclusive with start/length.")
    occurrence: int = Field(default=1, description="Which occurrence of search_text to target (1 = first). Only used with search_text.", ge=1)
    font_name: Optional[str] = Field(default=None, description="Latin font name. Also sets the East Asian font unless font_name_fareast is provided.")
    font_name_fareast: Optional[str] = Field(default=None, description="East Asian (CJK) font name (e.g. 'BIZ UDPゴシック'). Overrides the Far East font independently of font_name.")
    font_size: Optional[float] = Field(default=None, description="Font size in points")
    bold: Optional[bool] = Field(default=None, description="Bold on/off")
    italic: Optional[bool] = Field(default=None, description="Italic on/off")
    underline: Optional[bool] = Field(default=None, description="Underline on/off")
    color: Optional[str] = Field(default=None, description="Color as '#RRGGBB' hex string")
    font_color_theme: Optional[str] = Field(
        default=None,
        description="Theme color name (e.g. 'accent1', 'dark1')"
    )
    highlight_color: Optional[str] = Field(
        default=None,
        description="Text highlight (marker) color as '#RRGGBB' hex string, or 'clear' to remove highlight. Requires Office 2019+.",
    )

    @field_validator("highlight_color")
    @classmethod
    def validate_highlight_color(cls, v):
        if v is None:
            return v
        if v.lower() == "clear":
            return v
        if not re.fullmatch(r"#[0-9A-Fa-f]{6}", v):
            raise ValueError("highlight_color must be '#RRGGBB' hex string or 'clear'")
        return v

    @field_validator("search_text")
    @classmethod
    def validate_search_text_not_empty(cls, v):
        if v is not None and v == "":
            raise ValueError("search_text must not be empty")
        return v

    @model_validator(mode="after")
    def validate_range_specification(self):
        """Ensure either start/length or search_text is provided, not both."""
        has_start = self.start is not None
        has_length = self.length is not None
        has_search = self.search_text is not None

        if has_search:
            if has_start or has_length:
                raise ValueError(
                    "search_text is mutually exclusive with start/length. "
                    "Use either search_text or start+length, not both."
                )
        else:
            if not has_start or not has_length:
                raise ValueError(
                    "Either search_text or both start and length must be provided."
                )
            if self.occurrence != 1:
                raise ValueError(
                    "occurrence is only valid with search_text, not with start/length."
                )
        return self


class SetParagraphFormatInput(BaseModel):
    """Input for setting paragraph formatting."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    paragraph_index: Optional[int] = Field(
        default=None, description="1-based paragraph index. Omit to format all paragraphs."
    )
    alignment: Optional[str] = Field(
        default=None, description="'left', 'center', 'right', 'justify', or 'distribute'"
    )
    line_spacing: Optional[float] = Field(default=None, description="Line spacing multiplier")
    space_before: Optional[float] = Field(default=None, description="Space before paragraph in points")
    space_after: Optional[float] = Field(default=None, description="Space after paragraph in points")
    indent_level: Optional[int] = Field(default=None, ge=1, le=9, description="Indent level (1-9)")
    first_line_indent: Optional[float] = Field(default=None, description="First line indent in points")


class SetBulletInput(BaseModel):
    """Input for setting bullet/numbering on paragraphs."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    paragraph_index: Optional[int] = Field(
        default=None, description="1-based paragraph index. Omit to set for all paragraphs."
    )
    bullet_type: str = Field(
        ..., description="'none', 'unnumbered', or 'numbered'"
    )
    bullet_char: Optional[str] = Field(
        default=None, description="Bullet character (e.g. '●', '■')"
    )
    bullet_start_value: Optional[int] = Field(
        default=None, description="Starting number for numbered bullets"
    )
    indent_level: Optional[int] = Field(
        default=None, ge=1, le=9,
        description="Indent level 1-9. Sets the nesting depth of the bullet. "
        "Level 1 = top-level bullet, level 2 = first sub-bullet, etc.",
    )


class FindReplaceTextInput(BaseModel):
    """Input for find and replace text."""
    model_config = ConfigDict(str_strip_whitespace=True)

    find_text: str = Field(..., description="Text to find")
    replace_text: str = Field(..., description="Replacement text")
    slide_index: Optional[int] = Field(
        default=None, description="1-based slide index. Omit to search all slides."
    )


class SetTextframeInput(BaseModel):
    """Input for configuring text frame properties (auto-fit, margins, etc.)."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    auto_size: Optional[str] = Field(
        default=None,
        description=(
            "Text auto-fit mode: "
            "'none' (no auto-fit), "
            "'shrink_to_fit' (shrink text to fit the shape), "
            "'shape_to_fit' (resize shape to fit text)"
        ),
    )
    word_wrap: Optional[bool] = Field(
        default=None, description="Enable/disable word wrapping"
    )
    margin_left: Optional[float] = Field(default=None, description="Left margin in points")
    margin_right: Optional[float] = Field(default=None, description="Right margin in points")
    margin_top: Optional[float] = Field(default=None, description="Top margin in points")
    margin_bottom: Optional[float] = Field(default=None, description="Bottom margin in points")
    orientation: Optional[str] = Field(
        default=None, description="'horizontal', 'vertical', 'upward', or 'downward'"
    )
    vertical_anchor: Optional[str] = Field(
        default=None,
        description="Vertical text anchor: 'top', 'middle', or 'bottom'.",
    )


class GetAllTextInput(BaseModel):
    """Input for extracting all text from the presentation as pseudo-Markdown."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_indices: Optional[List[int]] = Field(
        default=None,
        description=(
            "1-based slide indices to extract. "
            "Omit to extract all slides."
        ),
    )
    output_path: Optional[str] = Field(
        default=None,
        description=(
            "File path to write the markdown text to (UTF-8). "
            "When provided, the result is written to the file and a JSON "
            "confirmation is returned instead of the text itself. "
            "Relative paths are resolved from the MCP server's working "
            "directory. Parent directories must already exist."
        ),
    )

    @model_validator(mode="after")
    def check_slide_indices(self):
        """Validate slide_indices values."""
        if self.slide_indices is not None:
            if len(self.slide_indices) == 0:
                raise ValueError("slide_indices must not be empty")
            for idx in self.slide_indices:
                if idx < 1:
                    raise ValueError(
                        f"slide_indices values must be >= 1, got {idx}"
                    )
        return self


class CheckTypographyInput(BaseModel):
    """Input for checking typographic issues (widow lines) on slides."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: Optional[int] = Field(
        default=None, ge=1,
        description="1-based slide index to check. Omit to check all slides.",
    )
    max_chars: int = Field(
        default=3, ge=1, le=10,
        description="Max characters for a line to be flagged as a widow (default 3).",
    )
    max_words: int = Field(
        default=2, ge=1, le=5,
        description="Max words for an English line to be flagged as a widow (default 2).",
    )
    fix: bool = Field(
        default=False,
        description="If true, attempt to fix widows by widening shapes to the right. "
        "Left edge stays fixed. Expansion is limited by neighboring shapes.",
    )
    max_expand_pt: float = Field(
        default=20.0, ge=1, le=50,
        description="Max points to expand shape width when fix=true (default 20).",
    )


# ---------------------------------------------------------------------------
# Placeholder types to skip (non-content)
# ---------------------------------------------------------------------------
_SKIP_PLACEHOLDER_TYPES = {13, 14, 15, 16}  # SlideNumber, Header, Footer, Date
_TITLE_PLACEHOLDER_TYPES = {1, 3, 5}  # Title, CenterTitle, VerticalTitle
_SUBTITLE_PLACEHOLDER_TYPES = {4}  # Subtitle

# Max slides per COM batch — keep under the 30-second COM call timeout
_GET_ALL_TEXT_BATCH_SIZE = 15


# ---------------------------------------------------------------------------
# Helpers for ppt_get_all_text
# ---------------------------------------------------------------------------
def _is_all_bold(shape) -> bool:
    """Check if ALL text runs in a shape are bold."""
    try:
        tr = shape.TextFrame.TextRange
        text = tr.Text.strip()
        if not text:
            return False
        run_count = tr.Runs().Count
        if run_count == 0:
            return False
        for i in range(1, run_count + 1):
            run = tr.Runs(i)
            # Skip whitespace-only runs
            if not run.Text.strip():
                continue
            if run.Font.Bold != msoTrue:
                return False
        return True
    except Exception:
        return False


def _runs_to_markdown(paragraph) -> str:
    """Convert a paragraph's runs to Markdown with bold/italic markers.

    Merges consecutive runs with identical formatting to avoid
    fragmented markers like **word1****word2**.
    """
    raw = []
    try:
        run_count = paragraph.Runs().Count
    except Exception:
        return paragraph.Text.replace("\r", "").replace("\v", "\n")

    if run_count == 0:
        return paragraph.Text.replace("\r", "").replace("\v", "\n")

    for i in range(1, run_count + 1):
        try:
            run = paragraph.Runs(i)
            text = run.Text.replace("\r", "").replace("\v", "\n")
            if not text:
                continue
            font = run.Font
            is_bold = font.Bold == msoTrue
            is_italic = font.Italic == msoTrue
            raw.append({"text": text, "bold": is_bold, "italic": is_italic})
        except Exception:
            continue

    # Merge consecutive runs with identical formatting
    merged = []
    for r in raw:
        if merged and merged[-1]["bold"] == r["bold"] and merged[-1]["italic"] == r["italic"]:
            merged[-1]["text"] += r["text"]
        else:
            merged.append(dict(r))

    # Format
    parts = []
    for m in merged:
        t = m["text"]
        if m["bold"] and m["italic"]:
            parts.append(f"***{t}***")
        elif m["bold"]:
            parts.append(f"**{t}**")
        elif m["italic"]:
            parts.append(f"*{t}*")
        else:
            parts.append(t)

    return "".join(parts)


def _plain_text(text_range) -> str:
    """Extract plain text from a TextRange, stripping formatting markers."""
    try:
        return text_range.Text.replace("\r", " ").replace("\v", " ").strip()
    except Exception:
        return ""


def _shape_paragraphs_to_markdown(shape, as_heading: str = "") -> str:
    """Convert a shape's paragraphs to Markdown text.

    Args:
        shape: COM shape with text frame
        as_heading: If set (e.g. "#" or "##"), render as heading
    """
    try:
        tr = shape.TextFrame.TextRange
    except Exception:
        return ""

    if as_heading:
        # For headings, use plain text (bold markers are redundant for # and ##)
        text = _plain_text(tr)
        if not text:
            return ""
        return f"{as_heading} {text}"

    lines = []
    para_count = tr.Paragraphs().Count
    for i in range(1, para_count + 1):
        para = tr.Paragraphs(i)
        text = _runs_to_markdown(para).strip()
        if not text:
            # Preserve empty paragraph as blank line
            lines.append("")
            continue

        # Detect bullet
        indent_level = para.IndentLevel
        bullet_prefix = ""
        try:
            bullet = para.ParagraphFormat.Bullet
            if bullet.Visible == msoTrue:
                indent = "  " * max(0, indent_level - 1)
                if bullet.Type == ppBulletNumbered:
                    bullet_prefix = f"{indent}1. "
                else:
                    bullet_prefix = f"{indent}- "
        except Exception:
            pass

        lines.append(f"{bullet_prefix}{text}")

    return "\n".join(lines)


def _table_to_markdown(shape) -> str:
    """Convert a table shape to a Markdown table.

    Note: Cell text is extracted as plain text; inline bold/italic
    formatting within table cells is not preserved.
    """
    try:
        table = shape.Table
        rows = table.Rows.Count
        cols = table.Columns.Count

        md_rows = []
        for r in range(1, rows + 1):
            cells = []
            for c in range(1, cols + 1):
                try:
                    text = table.Cell(r, c).Shape.TextFrame.TextRange.Text
                    text = text.replace("\r", " ").replace("\v", " ").replace("|", "\\|").strip()
                except Exception:
                    text = ""
                cells.append(text)
            md_rows.append("| " + " | ".join(cells) + " |")

            # Add header separator after first row
            if r == 1:
                md_rows.append("| " + " | ".join(["---"] * cols) + " |")

        return "\n".join(md_rows)
    except Exception:
        return ""


def _collect_text_shapes(slide) -> list:
    """Collect all text-bearing shapes from a slide with position info.

    Returns a list of dicts with keys:
        shape, top, left, width, height, is_title, is_subtitle,
        has_table, is_group
    Skips SlideNumber, Header, Footer, Date placeholders.
    """
    shapes = []

    def _process_shape(shape, offset_top=0.0, offset_left=0.0):
        """Process a single shape (may be called recursively for groups).

        Args:
            shape: COM shape object
            offset_top: Accumulated Y offset from parent groups
            offset_left: Accumulated X offset from parent groups
        """
        # Check placeholder skip / classify in a single COM read
        is_title = False
        is_subtitle = False
        if shape.Type == msoPlaceholder:
            try:
                ph_type = shape.PlaceholderFormat.Type
                if ph_type in _SKIP_PLACEHOLDER_TYPES:
                    return
                is_title = ph_type in _TITLE_PLACEHOLDER_TYPES
                is_subtitle = ph_type in _SUBTITLE_PLACEHOLDER_TYPES
            except Exception:
                pass

        # Recurse into groups early (no need to build info dict)
        # Pass group's position as offset since child coordinates are
        # relative to the group, not the slide.
        if shape.Type == msoGroup:
            try:
                g_top = shape.Top
                g_left = shape.Left
                for gi in range(1, shape.GroupItems.Count + 1):
                    _process_shape(
                        shape.GroupItems(gi),
                        offset_top + g_top,
                        offset_left + g_left,
                    )
            except Exception:
                pass
            return

        info = {
            "shape": shape,
            "top": shape.Top + offset_top,
            "left": shape.Left + offset_left,
            "width": shape.Width,
            "height": shape.Height,
            "is_title": is_title,
            "is_subtitle": is_subtitle,
            "has_table": False,
        }

        # Check for table
        try:
            if shape.HasTable:
                info["has_table"] = True
                shapes.append(info)
                return
        except Exception:
            pass

        # Check for text
        try:
            if shape.HasTextFrame:
                if shape.TextFrame.HasText:
                    shapes.append(info)
        except Exception:
            pass

    for i in range(1, slide.Shapes.Count + 1):
        _process_shape(slide.Shapes(i))

    return shapes


def _group_into_rows(shape_infos: list, threshold: float = 0.4) -> list:
    """Group shapes into rows based on vertical overlap.

    Two shapes are in the same row if their vertical overlap exceeds
    `threshold` of the shorter shape's height.

    Returns a list of rows, each row is a list of shape_infos sorted
    left-to-right. Rows are sorted top-to-bottom.
    """
    if not shape_infos:
        return []

    # Sort by top position
    sorted_shapes = sorted(shape_infos, key=lambda s: (s["top"], s["left"]))

    rows = []
    used = set()

    for i, s in enumerate(sorted_shapes):
        if i in used:
            continue
        row = [s]
        used.add(i)

        s_top = s["top"]
        s_bottom = s_top + s["height"]

        for j in range(i + 1, len(sorted_shapes)):
            if j in used:
                continue
            other = sorted_shapes[j]
            o_top = other["top"]
            o_bottom = o_top + other["height"]

            # Calculate vertical overlap
            overlap_top = max(s_top, o_top)
            overlap_bottom = min(s_bottom, o_bottom)
            overlap = max(0, overlap_bottom - overlap_top)

            shorter_height = min(s["height"], other["height"])
            if shorter_height > 0 and overlap / shorter_height >= threshold:
                row.append(other)
                used.add(j)

        # Sort row by left position
        row.sort(key=lambda x: x["left"])
        rows.append(row)

    # Sort rows by average top position
    rows.sort(key=lambda row: sum(s["top"] for s in row) / len(row))
    return rows


def _shape_info_to_markdown(info: dict, subheading_level: str = "##") -> str:
    """Convert a single shape_info dict to Markdown text.

    Args:
        info: Shape info dict from _collect_text_shapes.
        subheading_level: Heading prefix for all-bold non-title shapes.
            Use "##" for full-width (default), "###" for column shapes.
    """
    shape = info["shape"]

    # Table
    if info["has_table"]:
        return _table_to_markdown(shape)

    # Title
    if info["is_title"]:
        return _shape_paragraphs_to_markdown(shape, as_heading="#")

    # Subtitle — plain text (no heading marker)
    if info["is_subtitle"]:
        return _shape_paragraphs_to_markdown(shape)

    # All-bold → subheading (level depends on context)
    if _is_all_bold(shape):
        return _shape_paragraphs_to_markdown(shape, as_heading=subheading_level)

    # Regular text
    return _shape_paragraphs_to_markdown(shape)


def _group_into_columns(shapes: list, threshold: float = 50.0) -> list:
    """Group shapes by X-position proximity into columns.

    Args:
        shapes: Flat list of shape_info dicts (column shapes only).
        threshold: Maximum difference in Left values (points) for
            shapes to be considered the same column.

    Returns:
        List of columns (each a list of shape_infos sorted top-to-bottom),
        columns sorted left-to-right.
    """
    if not shapes:
        return []

    sorted_shapes = sorted(shapes, key=lambda s: s["left"])

    columns = []
    current_col = [sorted_shapes[0]]
    col_avg_left = sorted_shapes[0]["left"]

    for s in sorted_shapes[1:]:
        if abs(s["left"] - col_avg_left) <= threshold:
            current_col.append(s)
            # Update rolling average so drifting X values stay grouped
            col_avg_left = sum(x["left"] for x in current_col) / len(current_col)
        else:
            columns.append(current_col)
            current_col = [s]
            col_avg_left = s["left"]
    columns.append(current_col)

    # Sort each column by Y position (top-to-bottom)
    for col in columns:
        col.sort(key=lambda s: s["top"])

    # Sort columns left-to-right
    columns.sort(key=lambda col: sum(s["left"] for s in col) / len(col))
    return columns


def _slide_to_markdown(slide, slide_index: int) -> str:
    """Convert a single slide to pseudo-Markdown.

    Layout algorithm:
    - Rows are processed in Y-order to preserve vertical position.
    - Single-shape rows: rendered inline at their natural position.
    - Consecutive multi-shape rows: collected together, then grouped
      by X-position into columns so heading + body from the same
      column appear together.  All-bold shapes in columns use ###.
    """
    # Check if slide is hidden
    hidden = ""
    try:
        if slide.SlideShowTransition.Hidden:
            hidden = " (hidden)"
    except Exception:
        pass
    parts = [f"== Slide {slide_index}{hidden} =="]

    shape_infos = _collect_text_shapes(slide)
    if not shape_infos:
        parts.append("(no text)")
        return "\n".join(parts)

    rows = _group_into_rows(shape_infos)
    has_multi_shape_rows = any(len(row) > 1 for row in rows)

    if not has_multi_shape_rows:
        # Simple layout: all single-shape rows
        for row in rows:
            md = _shape_info_to_markdown(row[0])
            if md.strip():
                parts.append(md)
    else:
        # Mixed layout: interleave full-width and column groups
        # in original Y-order.  Consecutive multi-shape rows are
        # collected and flushed as a column group together.
        pending_column_shapes = []

        def _flush_columns():
            """Group pending column shapes by X and append to parts."""
            if not pending_column_shapes:
                return
            columns = _group_into_columns(pending_column_shapes)
            for col_idx, col in enumerate(columns):
                if col_idx > 0:
                    parts.append("")  # blank line between columns
                for info in col:
                    md = _shape_info_to_markdown(info, subheading_level="###")
                    if md.strip():
                        parts.append(md)
            pending_column_shapes.clear()

        for row in rows:
            if len(row) == 1:
                # Flush any pending column shapes before this full-width row
                _flush_columns()
                md = _shape_info_to_markdown(row[0])
                if md.strip():
                    parts.append(md)
            else:
                # Collect column shapes from consecutive multi-shape rows
                pending_column_shapes.extend(row)

        # Flush remaining column shapes at the end
        _flush_columns()

    return "\n".join(parts)


def _get_all_text_impl(slide_indices) -> str:
    """Extract all text from the presentation as pseudo-Markdown.

    Runs on the COM thread.
    """
    pres = ppt._get_pres_impl()
    total_slides = pres.Slides.Count

    if slide_indices is None:
        indices = list(range(1, total_slides + 1))
    else:
        indices = slide_indices

    slide_parts = []
    for idx in indices:
        if idx < 1 or idx > total_slides:
            slide_parts.append(
                f"== Slide {idx} ==\n(invalid slide index, "
                f"presentation has {total_slides} slides)"
            )
            continue
        slide = pres.Slides(idx)
        slide_parts.append(_slide_to_markdown(slide, idx))

    return "\n\n".join(slide_parts)


# ---------------------------------------------------------------------------
# COM implementation functions (run on COM thread)
# ---------------------------------------------------------------------------
def _set_text_impl(slide_index: int, shape_name_or_index, text: str) -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    if not shape.HasTextFrame:
        raise ValueError(f"Shape '{shape.Name}' does not have a text frame")

    tf = shape.TextFrame
    tr = tf.TextRange
    text = text.replace('\n', '\r')  # \n -> paragraph break (Enter)
    # \v (vertical tab) -> line break (Shift+Enter) — passed through as-is
    tr.Text = text

    return {
        "status": "success",
        "slide_index": slide_index,
        "shape_name": shape.Name,
        "text_length": tr.Length,
        "paragraph_count": tr.Paragraphs().Count,
    }


def _get_text_impl(slide_index: int, shape_name_or_index) -> dict:
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    if not shape.HasTextFrame:
        raise ValueError(f"Shape '{shape.Name}' does not have a text frame")

    tf = shape.TextFrame
    tr = tf.TextRange

    result = {
        "status": "success",
        "shape_name": shape.Name,
        "text": tr.Text,
        "text_length": tr.Length,
        "paragraph_count": tr.Paragraphs().Count,
    }

    paragraphs = []
    for i in range(1, tr.Paragraphs().Count + 1):
        para = tr.Paragraphs(i)
        para_info = {
            "index": i,
            "text": para.Text,
            "indent_level": para.IndentLevel,
            "alignment": para.ParagraphFormat.Alignment,
        }
        paragraphs.append(para_info)
    result["paragraphs"] = paragraphs

    runs = []
    run_count = tr.Runs().Count
    for i in range(1, run_count + 1):
        run = tr.Runs(i)
        font = run.Font
        color_int = font.Color.RGB
        run_info = {
            "index": i,
            "text": run.Text,
            "start": run.Start,
            "length": run.Length,
            "font_name": font.Name,
            "font_name_fareast": font.NameFarEast,
            "font_size": font.Size,
            "bold": font.Bold == msoTrue,
            "italic": font.Italic == msoTrue,
            "underline": font.Underline == msoTrue,
            "color_rgb": int_to_hex(color_int),
        }
        runs.append(run_info)
    result["runs"] = runs

    return result


def _apply_highlight(shape, highlight_color, start=None, length=None):
    """Apply or clear text highlight color via TextFrame2.

    Args:
        shape: PowerPoint Shape COM object.
        highlight_color: Hex color string (e.g. '#FFFF00'), or 'clear' to remove.
        start: 1-based start position (None for full range).
        length: Number of characters (None for full range).
    """
    if highlight_color.lower() == "clear":
        _clear_highlight(shape, start, length)
        return

    # Imported locally — only available on Windows where COM is used
    import win32com.client
    tr2 = shape.TextFrame2.TextRange
    if start is not None and length is not None:
        target = _get_textrange2_characters(tr2, start, length)
    else:
        target = tr2
    target.Font.Highlight.RGB = hex_to_int(highlight_color)


def _get_textrange2_characters(tr2, start, length):
    """Get a TextRange2.Characters sub-range via InvokeTypes.

    pywin32 late-binding dispatches Characters(start, length) incorrectly,
    so we call the COM method directly via InvokeTypes.
    """
    import win32com.client
    oleobj = tr2._oleobj_
    dispid = oleobj.GetIDsOfNames('Characters')
    # InvokeTypes args:
    #   2 = DISPATCH_PROPERTYGET
    #   (9, 0) = return type VT_DISPATCH, no flags
    #   (12, 17) = VT_VARIANT, PARAMFLAG_FIN|PARAMFLAG_FHASDEFAULT (optional input)
    result = oleobj.InvokeTypes(
        dispid, 0, 2,
        (9, 0),
        ((12, 17), (12, 17)),
        start, length,
    )
    return win32com.client.Dispatch(result)


def _clear_highlight(shape, start=None, length=None):
    """Clear text highlight by using ClearFormatting and restoring font properties.

    COM does not expose a direct method to remove highlights. This workaround:
    1. Saves per-run font formatting of the target range.
    2. Selects the text and executes ClearFormatting (clears highlight + all formatting).
    3. Restores the saved formatting so only the highlight is removed.

    Note: Requires the PowerPoint window to be visible and active, because
    Select() + ExecuteMso("ClearFormatting") operates through the UI layer.
    """
    app = shape.Application
    tr = shape.TextFrame.TextRange

    if start is not None and length is not None:
        target = tr.Characters(start, length)
    else:
        target = tr

    # Step 1: Save per-run formatting
    runs = _save_run_formatting(target)

    # Step 2: Select text and clear formatting (removes highlight + all formatting).
    # Sleep durations give the COM/UI layer time to process the selection and
    # ribbon command. These are empirically chosen minimums that work reliably
    # on typical hardware; very slow machines may need longer.
    target.Select()
    time.sleep(0.15)
    app.CommandBars.ExecuteMso("ClearFormatting")
    time.sleep(0.05)

    # Step 3: Restore saved formatting
    # Re-fetch the range after ClearFormatting (COM object may be stale)
    if start is not None and length is not None:
        target = shape.TextFrame.TextRange.Characters(start, length)
    else:
        target = shape.TextFrame.TextRange
    _restore_run_formatting(target, runs)


def _save_run_formatting(text_range):
    """Save per-run font formatting for later restoration.

    Groups consecutive characters with identical formatting into runs
    to minimize the number of COM calls during restore.
    """
    total = text_range.Length
    if total == 0:
        return []

    runs = []
    i = 1
    while i <= total:
        ch = text_range.Characters(i, 1)
        font = ch.Font
        fmt = {
            'start': i,
            'bold': font.Bold,
            'italic': font.Italic,
            'underline': font.Underline,
            'size': font.Size,
            'color_rgb': font.Color.RGB,
            'name': font.Name,
        }
        try:
            fmt['name_far_east'] = font.NameFarEast
        except Exception as e:
            logger.debug("NameFarEast unavailable at char %d: %s", i, e)

        # Extend run while formatting matches
        j = i + 1
        while j <= total:
            ch2 = text_range.Characters(j, 1)
            f2 = ch2.Font
            if not (f2.Bold == fmt['bold'] and
                    f2.Italic == fmt['italic'] and
                    f2.Underline == fmt['underline'] and
                    f2.Size == fmt['size'] and
                    f2.Color.RGB == fmt['color_rgb'] and
                    f2.Name == fmt['name']):
                break
            # Compare NameFarEast if available
            if 'name_far_east' in fmt:
                try:
                    if f2.NameFarEast != fmt['name_far_east']:
                        break
                except Exception as e:
                    logger.debug("NameFarEast comparison failed at char %d: %s", j, e)
                    break
            j += 1

        fmt['length'] = j - i
        runs.append(fmt)
        i = j

    return runs


def _restore_run_formatting(text_range, runs):
    """Restore per-run font formatting saved by _save_run_formatting."""
    for r in runs:
        chars = text_range.Characters(r['start'], r['length'])
        font = chars.Font
        font.Bold = r['bold']
        font.Italic = r['italic']
        font.Underline = r['underline']
        font.Size = r['size']
        font.Color.RGB = r['color_rgb']
        font.Name = r['name']
        if 'name_far_east' in r:
            try:
                font.NameFarEast = r['name_far_east']
            except Exception as e:
                logger.debug("Failed to restore NameFarEast: %s", e)


def _apply_font_props(font, font_name, font_name_fareast, font_size, bold, italic, underline, color, font_color_theme):
    """Apply font properties to a Font COM object."""
    if font_name is not None:
        font.Name = font_name
        font.NameFarEast = font_name  # default: match Latin unless overridden
    if font_name_fareast is not None:
        font.NameFarEast = font_name_fareast  # override East Asian font independently
    if font_size is not None:
        font.Size = font_size
    if bold is not None:
        font.Bold = msoTrue if bold else msoFalse
    if italic is not None:
        font.Italic = msoTrue if italic else msoFalse
    if underline is not None:
        font.Underline = msoTrue if underline else msoFalse
    if color is not None:
        font.Color.RGB = hex_to_int(color)
    if font_color_theme is not None:
        font.Color.ObjectThemeColor = get_theme_color_index(font_color_theme)


def _format_text_impl(slide_index, shape_name_or_index,
                       font_name, font_name_fareast, font_size, bold, italic, underline,
                       color, font_color_theme, highlight_color) -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    if not shape.HasTextFrame:
        raise ValueError(f"Shape '{shape.Name}' does not have a text frame")

    tr = shape.TextFrame.TextRange
    _apply_font_props(tr.Font, font_name, font_name_fareast, font_size, bold, italic, underline, color, font_color_theme)

    if highlight_color is not None:
        _apply_highlight(shape, highlight_color)

    return {
        "status": "success",
        "shape_name": shape.Name,
        "formatted_text": tr.Text,
        "start": tr.Start,
        "length": tr.Length,
    }


def _format_text_range_impl(slide_index, shape_name_or_index, start, length,
                              search_text, occurrence,
                              font_name, font_name_fareast, font_size, bold, italic, underline,
                              color, font_color_theme, highlight_color) -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    if not shape.HasTextFrame:
        raise ValueError(f"Shape '{shape.Name}' does not have a text frame")

    tr = shape.TextFrame.TextRange

    # Resolve search_text to start/length if provided
    if search_text is not None:
        full_text = tr.Text
        pos = -1
        search_from = 0
        for i in range(occurrence):
            pos = full_text.find(search_text, search_from)
            if pos == -1:
                if i == 0:
                    raise ValueError(
                        f"search_text '{search_text}' not found in shape '{shape.Name}'"
                    )
                else:
                    raise ValueError(
                        f"search_text '{search_text}' has only {i} occurrence(s) "
                        f"in shape '{shape.Name}', but occurrence={occurrence} was requested"
                    )
            search_from = pos + len(search_text)
        # COM Characters() is 1-based
        start = pos + 1
        length = len(search_text)

    target = tr.Characters(Start=start, Length=length)
    _apply_font_props(target.Font, font_name, font_name_fareast, font_size, bold, italic, underline, color, font_color_theme)

    if highlight_color is not None:
        _apply_highlight(shape, highlight_color, start, length)

    return {
        "status": "success",
        "shape_name": shape.Name,
        "formatted_text": target.Text,
        "start": target.Start,
        "length": target.Length,
    }


def _set_paragraph_format_impl(slide_index, shape_name_or_index, paragraph_index,
                                 alignment, line_spacing, space_before, space_after,
                                 indent_level, first_line_indent) -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    if not shape.HasTextFrame:
        raise ValueError(f"Shape '{shape.Name}' does not have a text frame")

    tr = shape.TextFrame.TextRange

    if paragraph_index is not None:
        target = tr.Paragraphs(paragraph_index)
    else:
        target = tr

    pf = target.ParagraphFormat

    if alignment is not None:
        align_val = ALIGNMENT_MAP.get(alignment)
        if align_val is None:
            raise ValueError(
                f"Invalid alignment '{alignment}'. "
                f"Valid values: {list(ALIGNMENT_MAP.keys())}"
            )
        pf.Alignment = align_val

    if line_spacing is not None:
        pf.LineRuleWithin = msoTrue
        pf.SpaceWithin = line_spacing

    if space_before is not None:
        pf.LineRuleBefore = msoFalse
        pf.SpaceBefore = space_before

    if space_after is not None:
        pf.LineRuleAfter = msoFalse
        pf.SpaceAfter = space_after

    if indent_level is not None:
        target.IndentLevel = indent_level

    if first_line_indent is not None:
        pf.FirstLineIndent = first_line_indent

    return {
        "status": "success",
        "shape_name": shape.Name,
        "paragraph_index": paragraph_index or "all",
    }


def _set_bullet_impl(slide_index, shape_name_or_index, paragraph_index,
                       bullet_type, bullet_char, bullet_start_value,
                       indent_level) -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    if not shape.HasTextFrame:
        raise ValueError(f"Shape '{shape.Name}' does not have a text frame")

    tr = shape.TextFrame.TextRange

    if paragraph_index is not None:
        target = tr.Paragraphs(paragraph_index)
    else:
        target = tr

    bullet_type_val = BULLET_TYPE_MAP.get(bullet_type)
    if bullet_type_val is None:
        raise ValueError(
            f"Invalid bullet_type '{bullet_type}'. "
            f"Valid values: {list(BULLET_TYPE_MAP.keys())}"
        )

    pf = target.ParagraphFormat
    bullet = pf.Bullet

    if bullet_type_val == ppBulletNone:
        bullet.Visible = msoFalse
    else:
        bullet.Visible = msoTrue
        bullet.Type = bullet_type_val

    if bullet_char is not None:
        bullet.Character = ord(bullet_char[0])

    if bullet_start_value is not None:
        bullet.StartValue = bullet_start_value

    if indent_level is not None:
        target.IndentLevel = indent_level

    return {
        "status": "success",
        "shape_name": shape.Name,
        "paragraph_index": paragraph_index or "all",
        "bullet_type": bullet_type,
        "indent_level": indent_level,
    }


def _find_replace_text_impl(find_text, replace_text, slide_index) -> dict:
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()

    replacements = []

    if slide_index is not None:
        slides_to_search = [pres.Slides(slide_index)]
    else:
        slides_to_search = [pres.Slides(i) for i in range(1, pres.Slides.Count + 1)]

    for slide in slides_to_search:
        for si in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(si)
            if not shape.HasTextFrame:
                continue
            tr = shape.TextFrame.TextRange
            while True:
                result = tr.Replace(
                    FindWhat=find_text,
                    ReplaceWhat=replace_text,
                    After=0,
                    MatchCase=msoFalse,
                    WholeWords=msoFalse,
                )
                if result is None:
                    break
                replacements.append({
                    "slide_index": slide.SlideIndex,
                    "shape_name": shape.Name,
                    "start": result.Start,
                    "length": result.Length,
                })

    return {
        "status": "success",
        "find_text": find_text,
        "replace_text": replace_text,
        "replacement_count": len(replacements),
        "replacements": replacements,
    }


def _set_textframe_impl(slide_index, shape_name_or_index,
                        auto_size, word_wrap,
                        margin_left, margin_right, margin_top, margin_bottom,
                        orientation, vertical_anchor) -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    if not shape.HasTextFrame:
        raise ValueError(f"Shape '{shape.Name}' does not have a text frame")

    tf = shape.TextFrame

    if margin_left is not None:
        tf.MarginLeft = margin_left
    if margin_right is not None:
        tf.MarginRight = margin_right
    if margin_top is not None:
        tf.MarginTop = margin_top
    if margin_bottom is not None:
        tf.MarginBottom = margin_bottom
    if word_wrap is not None:
        tf.WordWrap = msoTrue if word_wrap else msoFalse
    if orientation is not None:
        orient_val = ORIENTATION_MAP.get(orientation)
        if orient_val is None:
            raise ValueError(
                f"Invalid orientation '{orientation}'. "
                f"Valid values: {list(ORIENTATION_MAP.keys())}"
            )
        tf.Orientation = orient_val

    # Use TextFrame2 for auto_size (supports shrink_to_fit)
    if auto_size is not None:
        auto_size_val = AUTO_SIZE_MAP.get(auto_size)
        if auto_size_val is None:
            raise ValueError(
                f"Invalid auto_size '{auto_size}'. "
                f"Valid values: {list(AUTO_SIZE_MAP.keys())}"
            )
        shape.TextFrame2.AutoSize = auto_size_val

    if vertical_anchor is not None:
        VERTICAL_ANCHOR_MAP = {
            "top": 1,       # msoAnchorTop
            "middle": 3,    # msoAnchorMiddle
            "bottom": 4,    # msoAnchorBottom
        }
        anchor_val = VERTICAL_ANCHOR_MAP.get(vertical_anchor.lower())
        if anchor_val is None:
            raise ValueError(
                f"Invalid vertical_anchor '{vertical_anchor}'. "
                f"Must be one of: {sorted(VERTICAL_ANCHOR_MAP)}"
            )
        tf.VerticalAnchor = anchor_val

    return {
        "status": "success",
        "shape_name": shape.Name,
    }


# ---------------------------------------------------------------------------
# MCP tool functions
# ---------------------------------------------------------------------------
def set_text(params: SetTextInput) -> str:
    """Set the entire text content of a shape."""
    try:
        result = ppt.execute(
            _set_text_impl, params.slide_index, params.shape_name_or_index, params.text
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def get_text(params: GetTextInput) -> str:
    """Get text content and formatting info from a shape."""
    try:
        result = ppt.execute(
            _get_text_impl, params.slide_index, params.shape_name_or_index
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def format_text(params: FormatTextInput) -> str:
    """Format all text in a shape."""
    try:
        result = ppt.execute(
            _format_text_impl,
            params.slide_index, params.shape_name_or_index,
            params.font_name, params.font_name_fareast,
            params.font_size, params.bold, params.italic,
            params.underline, params.color, params.font_color_theme,
            params.highlight_color,
        )
        warn = font_size_warning(params.font_size)
        if warn:
            result["warning"] = warn
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def format_text_range(params: FormatTextRangeInput) -> str:
    """Format a specific character range within a shape's text."""
    try:
        result = ppt.execute(
            _format_text_range_impl,
            params.slide_index, params.shape_name_or_index,
            params.start, params.length,
            params.search_text, params.occurrence,
            params.font_name, params.font_name_fareast,
            params.font_size, params.bold, params.italic,
            params.underline, params.color, params.font_color_theme,
            params.highlight_color,
        )
        warn = font_size_warning(params.font_size)
        if warn:
            result["warning"] = warn
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def set_paragraph_format(params: SetParagraphFormatInput) -> str:
    """Set paragraph formatting for a shape."""
    try:
        result = ppt.execute(
            _set_paragraph_format_impl,
            params.slide_index, params.shape_name_or_index, params.paragraph_index,
            params.alignment, params.line_spacing, params.space_before,
            params.space_after, params.indent_level, params.first_line_indent,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def set_bullet(params: SetBulletInput) -> str:
    """Set bullet/numbering for paragraphs in a shape."""
    try:
        result = ppt.execute(
            _set_bullet_impl,
            params.slide_index, params.shape_name_or_index, params.paragraph_index,
            params.bullet_type, params.bullet_char, params.bullet_start_value,
            params.indent_level,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def find_replace_text(params: FindReplaceTextInput) -> str:
    """Find and replace text across all slides or a specific slide."""
    try:
        result = ppt.execute(
            _find_replace_text_impl,
            params.find_text, params.replace_text, params.slide_index,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def set_textframe(params: SetTextframeInput) -> str:
    """Configure text frame properties (auto-fit, word wrap, margins, orientation)."""
    try:
        result = ppt.execute(
            _set_textframe_impl,
            params.slide_index, params.shape_name_or_index,
            params.auto_size, params.word_wrap,
            params.margin_left, params.margin_right, params.margin_top, params.margin_bottom,
            params.orientation, params.vertical_anchor,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def get_all_text(params: GetAllTextInput) -> str:
    """Extract all text from the presentation as pseudo-Markdown.

    Batches COM calls to avoid the 30-second timeout on large presentations.
    Optionally writes the result to a file if output_path is specified.
    """
    try:
        if params.slide_indices is not None:
            indices = params.slide_indices
        else:
            # Get total slide count first
            total = ppt.execute(lambda: ppt._get_pres_impl().Slides.Count)
            indices = list(range(1, total + 1))

        # Process in batches to stay under the 30s COM timeout
        all_parts = []
        for i in range(0, len(indices), _GET_ALL_TEXT_BATCH_SIZE):
            batch = indices[i:i + _GET_ALL_TEXT_BATCH_SIZE]
            part = ppt.execute(_get_all_text_impl, batch)
            all_parts.append(part)

        text = "\n\n".join(all_parts)

        if params.output_path:
            abs_path = os.path.abspath(params.output_path)
            with open(abs_path, "w", encoding="utf-8") as f:
                f.write(text)
            return json.dumps({
                "status": "success",
                "output_path": abs_path,
                "slide_count": len(indices),
            })

        return text
    except Exception as e:
        return json.dumps({"error": str(e)})


# ---------------------------------------------------------------------------
# Typography check (widow line detection)
# ---------------------------------------------------------------------------

def _is_latin(text: str) -> bool:
    """Return True if text is primarily Latin characters."""
    latin = sum(1 for c in text if c.isascii() and c.isalpha())
    return latin > len(text) * 0.5


def _char_type(c):
    """Classify a character for word-boundary detection."""
    if c in "\u3001\u3002\uff0c\uff0e\uff01\uff1f\uff09\u300d\u300f\u3011\u3009\u300b\u30fb":
        return "punct_close"
    if c in "\uff08\u300c\u300e\u3010\u3008\u300a":
        return "punct_open"
    cp = ord(c)
    if cp < 0x80:
        return "latin" if c.isalpha() else ("digit" if c.isdigit() else "ascii")
    if 0x30A0 <= cp <= 0x30FF or 0x31F0 <= cp <= 0x31FF:
        return "katakana"
    if 0x3040 <= cp <= 0x309F:
        return "hiragana"
    if unicodedata.category(c).startswith("Lo"):
        return "kanji"
    return "other"


def _find_best_vbreak(prev_line_text, widow_text):
    """Find the best position in prev_line_text to insert \\v.

    Returns (break_pos, before, after) or None.
    The break is chosen at a character-type transition in the second half
    of prev_line_text, preferring punctuation > type changes.
    """
    candidates = []
    for i in range(1, len(prev_line_text)):
        pt = _char_type(prev_line_text[i - 1])
        ct = _char_type(prev_line_text[i])
        score = 0
        if pt == "punct_close":
            score = 5
        elif ct == "punct_open":
            score = 5
        elif pt != ct:
            if pt == "hiragana" and ct == "katakana":
                score = 4
            elif pt == "hiragana" and ct == "kanji":
                score = 3
            elif pt == "katakana" and ct in ("hiragana", "kanji"):
                score = 3
            elif pt == "kanji" and ct == "hiragana":
                score = 2
            elif pt == "latin" and ct in ("hiragana", "katakana", "kanji"):
                score = 3
            elif pt in ("hiragana", "katakana", "kanji") and ct == "latin":
                score = 3
            else:
                score = 1
        if score > 0:
            candidates.append((i, score))

    if not candidates:
        return None

    min_pos = len(prev_line_text) // 2
    valid = [(p, s) for p, s in candidates if p >= min_pos]
    if not valid:
        valid = [(p, s) for p, s in candidates if p >= len(prev_line_text) // 3]
    if not valid:
        return None

    valid.sort(key=lambda x: (x[1], x[0]), reverse=True)
    pos = valid[0][0]
    return pos, prev_line_text[:pos], prev_line_text[pos:] + widow_text


def _get_widows(shape, max_chars, max_words):
    """Return list of widow issues for a single shape."""
    tr = shape.TextFrame.TextRange
    lines_count = tr.Lines().Count
    if lines_count < 2:
        return []

    widows = []
    for li in range(2, lines_count + 1):
        prev_line = tr.Lines(li - 1)
        prev_text = prev_line.Text
        # Explicit break (\r = paragraph, \n = soft return) — not a widow.
        if prev_text.endswith("\r") or prev_text.endswith("\n"):
            continue

        cur_line = tr.Lines(li)
        cur_text = cur_line.Text.rstrip("\r\n")
        if not cur_text:
            continue

        is_widow = False
        if _is_latin(cur_text):
            if len(cur_text.split()) <= max_words:
                is_widow = True
        else:
            if len(cur_text) <= max_chars:
                is_widow = True

        if is_widow:
            widows.append({
                "line_index": li,
                "line_text": cur_text,
                "char_count": len(cur_text),
                "prev_line_text": prev_text.rstrip("\r\n"),
            })
    return widows


def _get_short_vbreaks(shape, max_chars, max_words):
    """Return list of lines after an explicit \\v that are too short.

    These are the opposite of widows — explicit breaks that leave the
    following line unnecessarily sparse. Removing the \\v would let text
    flow naturally.
    """
    tr = shape.TextFrame.TextRange
    lines_count = tr.Lines().Count
    if lines_count < 2:
        return []

    short_breaks = []
    for li in range(2, lines_count + 1):
        prev_line = tr.Lines(li - 1)
        prev_text = prev_line.Text
        # Only flag lines after an explicit \v (shows as \n in COM).
        # Skip \r (paragraph break) — those are intentional structural breaks.
        if not prev_text.endswith("\n"):
            continue

        cur_line = tr.Lines(li)
        cur_text = cur_line.Text.rstrip("\r\n")
        if not cur_text:
            continue

        is_short = False
        if _is_latin(cur_text):
            if len(cur_text.split()) <= max_words:
                is_short = True
        else:
            if len(cur_text) <= max_chars:
                is_short = True

        if is_short:
            short_breaks.append({
                "line_index": li,
                "line_text": cur_text,
                "char_count": len(cur_text),
                "prev_line_text": prev_text.rstrip("\r\n"),
                "issue_type": "short_after_vbreak",
            })
    return short_breaks


def _right_neighbor_gap(shape, slide):
    """Find the gap (pt) to the nearest shape on the right that vertically overlaps."""
    s_right = shape.Left + shape.Width
    s_top = shape.Top
    s_bottom = shape.Top + shape.Height
    min_gap = float("inf")

    for j in range(1, slide.Shapes.Count + 1):
        other = slide.Shapes(j)
        if other.Name == shape.Name:
            continue
        # Must vertically overlap
        if other.Top + other.Height <= s_top or other.Top >= s_bottom:
            continue
        # Must be to the right
        if other.Left > s_right - 1:
            gap = other.Left - s_right
            if gap < min_gap:
                min_gap = gap

    return min_gap


def _check_typography_impl(slide_indices, max_chars, max_words,
                           fix, max_expand_pt):
    """Scan shapes for widow lines; optionally fix by widening."""
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    issues = []
    fixed = []

    for si in slide_indices:
        if si < 1 or si > pres.Slides.Count:
            continue
        goto_slide(app, si)
        slide = pres.Slides(si)

        for j in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(j)
            if not shape.HasTextFrame:
                continue
            tr = shape.TextFrame.TextRange
            if not tr.Text.strip():
                continue

            # Detect auto-shrink (text compressed to fit shape)
            try:
                auto_size = shape.TextFrame2.AutoSize
                if auto_size == ppAutoSizeTextToFitShape:
                    issues.append({
                        "slide_index": si,
                        "shape_name": shape.Name,
                        "shape_width": round(shape.Width, 2),
                        "type": "auto_shrink",
                        "fixable": False,
                    })
            except Exception:
                logger.debug("Cannot read AutoSize for shape '%s'",
                             shape.Name, exc_info=True)

            widows = _get_widows(shape, max_chars, max_words)

            # Detect short lines after explicit \v breaks
            # (when fix=True, we re-check after fix and report then)
            if not fix:
                vbreak_shorts = _get_short_vbreaks(shape, max_chars, max_words)
                for vb in vbreak_shorts:
                    issues.append({
                        "slide_index": si,
                        "shape_name": shape.Name,
                        "shape_width": round(shape.Width, 2),
                        **vb,
                    })

            if not widows:
                continue

            if fix:
                # Calculate safe expansion room
                gap = _right_neighbor_gap(shape, slide)
                room = min(gap - 2, max_expand_pt)  # 2pt margin
                if room < 1:
                    room = 0  # skip widen step, go straight to \v

                original_width = shape.Width
                resolved = False
                for step in range(1, int(room) + 1):
                    shape.Width = original_width + step
                    remaining = _get_widows(shape, max_chars, max_words)
                    if not remaining:
                        fixed.append({
                            "slide_index": si,
                            "shape_name": shape.Name,
                            "old_width": round(original_width, 2),
                            "new_width": round(shape.Width, 2),
                            "expanded_by": step,
                        })
                        resolved = True
                        break

                if not resolved:
                    # Revert width — try soft-return insertion instead
                    shape.Width = original_width
                    remaining = widows
                    # Strategy 2: insert \v at word boundary
                    # Process widows in reverse order (later positions first)
                    # so that earlier character positions remain valid.
                    remaining.sort(
                        key=lambda w: w["line_index"], reverse=True,
                    )
                    vbreak_applied = False
                    for w in remaining:
                        brk = _find_best_vbreak(
                            w["prev_line_text"], w["line_text"],
                        )
                        if brk is None:
                            issues.append({
                                "slide_index": si,
                                "shape_name": shape.Name,
                                "shape_width": round(original_width, 2),
                                "fix_status": "no_break_point",
                                **w,
                            })
                            continue
                        # Find the fragment in the full text and get
                        # the COM character position for insertion.
                        # NOTE: find() returns the first occurrence of old_frag.
                        # If identical text appears multiple times in the shape,
                        # the wrong position may be used. This is rare in practice.
                        brk_pos, before, after = brk
                        old_frag = w["prev_line_text"] + w["line_text"]
                        full_text = tr.Text
                        idx = full_text.find(old_frag)
                        if idx == -1:
                            issues.append({
                                "slide_index": si,
                                "shape_name": shape.Name,
                                "shape_width": round(original_width, 2),
                                "fix_status": "text_not_found",
                                **w,
                            })
                            continue
                        # COM position (1-based) for the break point
                        com_pos = idx + brk_pos + 1
                        tr.Characters(com_pos, 0).InsertBefore("\v")
                        vbreak_applied = True
                        fixed.append({
                            "slide_index": si,
                            "shape_name": shape.Name,
                            "fix_method": "soft_return",
                            "before": before,
                            "after": after,
                        })
                    # After \v insertions, re-check for remaining widows
                    # and new short_after_vbreak issues
                    if vbreak_applied:
                        still_remaining = _get_widows(
                            shape, max_chars, max_words,
                        )
                        for w in still_remaining:
                            issues.append({
                                "slide_index": si,
                                "shape_name": shape.Name,
                                "shape_width": round(shape.Width, 2),
                                "fix_status": "remaining",
                                **w,
                            })

                # Always report short_after_vbreak in fix mode
                # (covers both width-expanded and \v-inserted shapes)
                post_fix_vbreaks = _get_short_vbreaks(
                    shape, max_chars, max_words,
                )
                for vb in post_fix_vbreaks:
                    issues.append({
                        "slide_index": si,
                        "shape_name": shape.Name,
                        "shape_width": round(shape.Width, 2),
                        **vb,
                    })
            else:
                for w in widows:
                    issues.append({
                        "slide_index": si,
                        "shape_name": shape.Name,
                        "shape_width": round(shape.Width, 2),
                        **w,
                    })

    result = {"issues": issues, "total": len(issues)}
    if fix:
        result["fixed"] = fixed
        result["fixed_count"] = len(fixed)
        result["remaining"] = len(issues)
    return result


def check_typography(params: CheckTypographyInput) -> str:
    """Check slides for typographic widow lines."""
    try:
        if params.slide_index is not None:
            indices = [params.slide_index]
        else:
            total = ppt.execute(lambda: ppt._get_pres_impl().Slides.Count)
            indices = list(range(1, total + 1))

        result = ppt.execute(
            _check_typography_impl, indices,
            params.max_chars, params.max_words,
            params.fix, params.max_expand_pt,
        )
        return json.dumps(result, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register all text tools with the MCP server."""

    @mcp.tool(
        name="ppt_set_text",
        annotations={
            "title": "Set Shape Text",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_ppt_set_text(params: SetTextInput) -> str:
        """Set the entire text content of a shape.

        Replaces all existing text.
        \\n = paragraph break (Enter) — starts a new paragraph with its own
        bullet/numbering and indent level.
        \\v = line break (Shift+Enter) — soft return within the same paragraph,
        preserving bullet/indent. Use \\v for wrapping at natural word
        boundaries within one paragraph.
        """
        return set_text(params)

    @mcp.tool(
        name="ppt_get_text",
        annotations={
            "title": "Get Shape Text",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_ppt_get_text(params: GetTextInput) -> str:
        """Get text content from a shape, including paragraph and run details.

        Returns the full text, paragraph info (alignment, indent level),
        and per-run formatting (font, size, bold, italic, color).
        """
        return get_text(params)

    @mcp.tool(
        name="ppt_format_text",
        annotations={
            "title": "Format All Text in Shape",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_ppt_format_text(params: FormatTextInput) -> str:
        """Apply formatting to ALL text in a shape.

        Sets font properties (name, size, bold, italic, underline, color)
        for the entire text content of the shape.
        """
        return format_text(params)

    @mcp.tool(
        name="ppt_format_text_range",
        annotations={
            "title": "Format Partial Text (Characters)",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_ppt_format_text_range(params: FormatTextRangeInput) -> str:
        """Format a specific character range within a shape's text.

        Target range can be specified in two ways (mutually exclusive):
        1. **start + length**: Characters(start, length) — start is 1-based.
           Example: to bold characters 3 through 7, use start=3, length=5.
        2. **search_text**: Search for the text and format the matching range.
           Use occurrence to target the Nth match (default: 1st).
        """
        return format_text_range(params)

    @mcp.tool(
        name="ppt_set_paragraph_format",
        annotations={
            "title": "Set Paragraph Format",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_ppt_set_paragraph_format(params: SetParagraphFormatInput) -> str:
        """Set paragraph-level formatting for a shape.

        Applies alignment, line spacing, space before/after, indent level,
        and first-line indent. Omit paragraph_index to format all paragraphs.
        """
        return set_paragraph_format(params)

    @mcp.tool(
        name="ppt_set_bullet",
        annotations={
            "title": "Set Bullet/Numbering",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_ppt_set_bullet(params: SetBulletInput) -> str:
        """Set bullet or numbering for paragraphs in a shape.

        bullet_type can be 'none', 'unnumbered', or 'numbered'.
        Use bullet_char for custom bullet characters (e.g. '●').
        Use bullet_start_value to set the starting number.
        Use indent_level (1-9) to set nesting depth in one call.

        Nested bullet example — first use ppt_set_text to create paragraphs
        separated by \\n, then call ppt_set_bullet per paragraph
        (slide_index and shape_name_or_index required on each call):
          ppt_set_bullet(..., paragraph_index=1, bullet_type='unnumbered', indent_level=1)
          ppt_set_bullet(..., paragraph_index=2, bullet_type='unnumbered', indent_level=2)
          ppt_set_bullet(..., paragraph_index=3, bullet_type='unnumbered', indent_level=3)
        """
        return set_bullet(params)

    @mcp.tool(
        name="ppt_find_replace_text",
        annotations={
            "title": "Find and Replace Text",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_ppt_find_replace_text(params: FindReplaceTextInput) -> str:
        """Find and replace text across all slides or a specific slide.

        Searches all text-containing shapes. If slide_index is omitted,
        searches the entire presentation.
        """
        return find_replace_text(params)

    @mcp.tool(
        name="ppt_set_textframe",
        annotations={
            "title": "Set TextFrame Properties",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_ppt_set_textframe(params: SetTextframeInput) -> str:
        """Configure text frame auto-fit, word wrap, margins, orientation, and vertical anchor.

        Controls how text fits within a shape:
        - auto_size='shrink_to_fit': shrink text font to fit the shape
        - auto_size='shape_to_fit': resize the shape to fit all text
        - auto_size='none': no auto-fitting (text may overflow)
        - word_wrap: enable/disable text wrapping at shape boundary
        - vertical_anchor: 'top', 'middle', or 'bottom' — controls vertical text alignment
        Also sets inner margins (points) and text orientation.
        """
        return set_textframe(params)

    @mcp.tool(
        name="ppt_get_all_text",
        annotations={
            "title": "Get All Text as Markdown",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_ppt_get_all_text(params: GetAllTextInput) -> str:
        """Extract all text from the presentation as pseudo-Markdown.

        Returns a structured overview of every slide's content:
        - `# Heading` for slide titles
        - `## Subheading` for all-bold full-width shapes
        - `### Subheading` for all-bold shapes in multi-column layouts
        - `**bold**` and `*italic*` inline formatting
        - `- bullet` items with indentation
        - Markdown tables for table shapes

        Column shapes are grouped by X-position so heading + body from
        the same column appear together.

        Set output_path to write the result to a UTF-8 file instead of
        returning the text directly.

        Omit slide_indices to get all slides.
        """
        return get_all_text(params)

    @mcp.tool(
        name="ppt_check_typography",
        annotations={
            "title": "Check Typography (Widow Lines)",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_ppt_check_typography(params: CheckTypographyInput) -> str:
        """Detect typography issues: widow lines and auto-shrunk text.

        Scans shapes for: (1) widow lines where text wrapping pushed only
        a few characters (≤ max_chars, default 3) or words (≤ max_words,
        default 2 for English text) to the next visual line,
        (2) auto-shrunk text where shrink_to_fit compresses text to fit
        (reported with fixable=false — no auto-fix available).

        With fix=false (default), detection is read-only — no changes
        are made. Set fix=true to auto-fix widows: first tries widening
        shapes (left edge fixed, stops at neighbors with 2pt margin),
        then inserts soft returns (\\v) at word boundaries. Unfixable
        shapes are reported with fix_status='no_break_point' or
        'text_not_found'.
        """
        return check_typography(params)
