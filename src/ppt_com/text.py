"""Text content, formatting, and manipulation tools for PowerPoint COM automation."""

import json
import logging
from typing import Optional, Union

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from utils.navigation import goto_slide
from utils.color import hex_to_int, int_to_hex, int_to_rgb, get_theme_color_index
from ppt_com.constants import (
    msoTrue, msoFalse, msoTriStateMixed,
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
        ..., description="Shape name (string) or 1-based index (int)"
    )
    text: str = Field(..., description="Text content. Use \\n for paragraph breaks.")


class GetTextInput(BaseModel):
    """Input for getting text from a shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (string) or 1-based index (int)"
    )


class FormatTextInput(BaseModel):
    """Input for formatting all text in a shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (string) or 1-based index (int)"
    )
    font_name: Optional[str] = Field(default=None, description="Font name (e.g. 'Arial')")
    font_size: Optional[float] = Field(default=None, description="Font size in points")
    bold: Optional[bool] = Field(default=None, description="Bold on/off")
    italic: Optional[bool] = Field(default=None, description="Italic on/off")
    underline: Optional[bool] = Field(default=None, description="Underline on/off")
    color: Optional[str] = Field(default=None, description="Color as '#RRGGBB' hex string")
    font_color_theme: Optional[str] = Field(
        default=None,
        description="Theme color name (e.g. 'accent1', 'dark1')"
    )


class FormatTextRangeInput(BaseModel):
    """Input for formatting a specific character range within a shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (string) or 1-based index (int)"
    )
    start: int = Field(..., description="1-based character start position")
    length: int = Field(..., description="Number of characters to format")
    font_name: Optional[str] = Field(default=None, description="Font name")
    font_size: Optional[float] = Field(default=None, description="Font size in points")
    bold: Optional[bool] = Field(default=None, description="Bold on/off")
    italic: Optional[bool] = Field(default=None, description="Italic on/off")
    underline: Optional[bool] = Field(default=None, description="Underline on/off")
    color: Optional[str] = Field(default=None, description="Color as '#RRGGBB' hex string")
    font_color_theme: Optional[str] = Field(
        default=None,
        description="Theme color name (e.g. 'accent1', 'dark1')"
    )


class SetParagraphFormatInput(BaseModel):
    """Input for setting paragraph formatting."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (string) or 1-based index (int)"
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
    indent_level: Optional[int] = Field(default=None, description="Indent level (1-9)")
    first_line_indent: Optional[float] = Field(default=None, description="First line indent in points")


class SetBulletInput(BaseModel):
    """Input for setting bullet/numbering on paragraphs."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (string) or 1-based index (int)"
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
        ..., description="Shape name (string) or 1-based index (int)"
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


# ---------------------------------------------------------------------------
# COM implementation functions (run on COM thread)
# ---------------------------------------------------------------------------
def _set_text_impl(slide_index: int, shape_name_or_index, text: str) -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    if not shape.HasTextFrame:
        raise ValueError(f"Shape '{shape.Name}' does not have a text frame")

    tf = shape.TextFrame
    tr = tf.TextRange
    text = text.replace('\n', '\r')
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
    pres = app.ActivePresentation
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
            "font_size": font.Size,
            "bold": font.Bold == msoTrue,
            "italic": font.Italic == msoTrue,
            "underline": font.Underline == msoTrue,
            "color_rgb": int_to_hex(color_int),
        }
        runs.append(run_info)
    result["runs"] = runs

    return result


def _apply_font_props(font, font_name, font_size, bold, italic, underline, color, font_color_theme):
    """Apply font properties to a Font COM object."""
    if font_name is not None:
        font.Name = font_name
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
                       font_name, font_size, bold, italic, underline,
                       color, font_color_theme) -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    if not shape.HasTextFrame:
        raise ValueError(f"Shape '{shape.Name}' does not have a text frame")

    tr = shape.TextFrame.TextRange
    _apply_font_props(tr.Font, font_name, font_size, bold, italic, underline, color, font_color_theme)

    return {
        "status": "success",
        "shape_name": shape.Name,
        "formatted_text": tr.Text,
        "start": tr.Start,
        "length": tr.Length,
    }


def _format_text_range_impl(slide_index, shape_name_or_index, start, length,
                              font_name, font_size, bold, italic, underline,
                              color, font_color_theme) -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    if not shape.HasTextFrame:
        raise ValueError(f"Shape '{shape.Name}' does not have a text frame")

    tr = shape.TextFrame.TextRange
    target = tr.Characters(Start=start, Length=length)
    _apply_font_props(target.Font, font_name, font_size, bold, italic, underline, color, font_color_theme)

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
    pres = app.ActivePresentation
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
                       bullet_type, bullet_char, bullet_start_value) -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
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

    bullet = target.ParagraphFormat.Bullet

    if bullet_type_val == ppBulletNone:
        bullet.Visible = msoFalse
    else:
        bullet.Visible = msoTrue
        bullet.Type = bullet_type_val

    if bullet_char is not None:
        bullet.Character = ord(bullet_char[0])

    if bullet_start_value is not None:
        bullet.StartValue = bullet_start_value

    return {
        "status": "success",
        "shape_name": shape.Name,
        "paragraph_index": paragraph_index or "all",
        "bullet_type": bullet_type,
    }


def _find_replace_text_impl(find_text, replace_text, slide_index) -> dict:
    app = ppt._get_app_impl()
    pres = app.ActivePresentation

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
                        orientation) -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
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
            params.font_name, params.font_size, params.bold, params.italic,
            params.underline, params.color, params.font_color_theme,
        )
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
            params.font_name, params.font_size, params.bold, params.italic,
            params.underline, params.color, params.font_color_theme,
        )
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
            params.orientation,
        )
        return json.dumps(result)
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

        Replaces all existing text. Use \\n for paragraph breaks
        (they are converted to \\r internally for PowerPoint).
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

        This is the KEY feature for partial text formatting. Uses
        Characters(start, length) to target specific characters.
        Start is 1-based. For example, to bold characters 3 through 7,
        use start=3, length=5.
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
        """Configure text frame auto-fit, word wrap, margins, and orientation.

        Controls how text fits within a shape:
        - auto_size='shrink_to_fit': shrink text font to fit the shape
        - auto_size='shape_to_fit': resize the shape to fit all text
        - auto_size='none': no auto-fitting (text may overflow)
        - word_wrap: enable/disable text wrapping at shape boundary
        Also sets inner margins (points) and text orientation.
        """
        return set_textframe(params)
