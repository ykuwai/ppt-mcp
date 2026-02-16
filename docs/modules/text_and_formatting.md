# Module: Text & Formatting

## Overview

This module handles all text operations and formatting within PowerPoint shapes: setting and getting text, partial text formatting via Characters/Words/Paragraphs, font properties (name, size, bold, italic, underline, color, subscript, superscript), paragraph formatting (alignment, line spacing, indentation), bullet/numbering configuration, text find/replace, and TextFrame properties (margins, auto-size, word wrap, orientation). This is the CRITICAL module for rich text manipulation -- the most commonly requested feature by users.

## Dependencies

- **Internal**: `utils.com_wrapper` (PowerPointCOMWrapper, safe_com_call), `utils.units` (inches_to_points, cm_to_points, points_to_inches, points_to_cm), `utils.color` (rgb_to_int, int_to_rgb, int_to_hex, hex_to_int), `ppt_com.constants` (all text/paragraph/bullet constants)
- **External**: `pywin32` (`win32com.client`, `pywintypes`)
- **Standard library**: `logging`

### Importing from Core Module

```python
from utils.com_wrapper import PowerPointCOMWrapper, safe_com_call
from utils.units import inches_to_points, points_to_inches, cm_to_points, points_to_cm
from utils.color import rgb_to_int, int_to_rgb, int_to_hex, hex_to_int
from ppt_com.constants import (
    # MsoTriState
    msoTrue, msoFalse, msoTriStateMixed,
    # Text orientation
    msoTextOrientationHorizontal, msoTextOrientationVertical,
    msoTextOrientationVerticalFarEast, msoTextOrientationUpward,
    msoTextOrientationDownward,
    # AutoSize
    ppAutoSizeNone, ppAutoSizeShapeToFitText,
    # Paragraph alignment
    ppAlignLeft, ppAlignCenter, ppAlignRight, ppAlignJustify, ppAlignDistribute,
    # Bullet type
    ppBulletNone, ppBulletUnnumbered, ppBulletNumbered, ppBulletPicture,
    # Numbered bullet style
    ppBulletArabicParenRight, ppBulletArabicPeriod, ppBulletArabicParenBoth,
    ppBulletRomanUCPeriod, ppBulletRomanLCPeriod,
    ppBulletAlphaUCPeriod, ppBulletAlphaLCPeriod,
    # Theme colors
    msoThemeColorDark1, msoThemeColorLight1, msoThemeColorAccent1,
    msoThemeColorAccent2, msoThemeColorAccent3, msoThemeColorAccent4,
    msoThemeColorAccent5, msoThemeColorAccent6,
)
import logging

logger = logging.getLogger(__name__)
```

## File Structure

```
ppt_com_mcp/
  ppt_com/
    text.py   # Text operations and formatting
```

---

## Constants Needed

These constants MUST be defined in `ppt_com/constants.py` (from the core module). Listed here for reference:

### MsoTriState

| Name | Value | Description |
|------|-------|-------------|
| `msoTrue` | -1 | True |
| `msoFalse` | 0 | False |
| `msoTriStateMixed` | -2 | Mixed (multiple values) |

### PpParagraphAlignment

| Name | Value | Description |
|------|-------|-------------|
| `ppAlignLeft` | 1 | Left aligned |
| `ppAlignCenter` | 2 | Center aligned |
| `ppAlignRight` | 3 | Right aligned |
| `ppAlignJustify` | 4 | Justified |
| `ppAlignDistribute` | 5 | Distributed |
| `ppAlignmentMixed` | -2 | Mixed |

### PpAutoSize

| Name | Value | Description |
|------|-------|-------------|
| `ppAutoSizeNone` | 0 | No auto-sizing |
| `ppAutoSizeMixed` | -2 | Mixed |
| `ppAutoSizeShapeToFitText` | 1 | Resize shape to fit text |

### MsoTextOrientation

| Name | Value | Description |
|------|-------|-------------|
| `msoTextOrientationHorizontal` | 1 | Horizontal (default) |
| `msoTextOrientationUpward` | 2 | Bottom to top |
| `msoTextOrientationDownward` | 3 | Top to bottom |
| `msoTextOrientationVertical` | 5 | Vertical (top-down, right-to-left) |
| `msoTextOrientationVerticalFarEast` | 6 | Vertical (East Asian) |

### PpBulletType

| Name | Value | Description |
|------|-------|-------------|
| `ppBulletNone` | 0 | No bullet |
| `ppBulletUnnumbered` | 1 | Symbol bullet |
| `ppBulletNumbered` | 2 | Numbered bullet |
| `ppBulletPicture` | 3 | Picture bullet |
| `ppBulletMixed` | -2 | Mixed |

### PpNumberedBulletStyle (commonly used)

| Name | Value | Description |
|------|-------|-------------|
| `ppBulletArabicParenRight` | 2 | 1) 2) 3) |
| `ppBulletArabicPeriod` | 3 | 1. 2. 3. |
| `ppBulletArabicParenBoth` | 12 | (1) (2) (3) |
| `ppBulletRomanUCPeriod` | 4 | I. II. III. |
| `ppBulletRomanLCPeriod` | 5 | i. ii. iii. |
| `ppBulletAlphaUCPeriod` | 6 | A. B. C. |
| `ppBulletAlphaLCPeriod` | 7 | a. b. c. |

### MsoThemeColorIndex (for font colors)

| Name | Value | Description |
|------|-------|-------------|
| `msoThemeColorDark1` | 1 | Dark 1 (usually black) |
| `msoThemeColorLight1` | 2 | Light 1 (usually white) |
| `msoThemeColorDark2` | 3 | Dark 2 |
| `msoThemeColorLight2` | 4 | Light 2 |
| `msoThemeColorAccent1` | 5 | Accent 1 |
| `msoThemeColorAccent2` | 6 | Accent 2 |
| `msoThemeColorAccent3` | 7 | Accent 3 |
| `msoThemeColorAccent4` | 8 | Accent 4 |
| `msoThemeColorAccent5` | 9 | Accent 5 |
| `msoThemeColorAccent6` | 10 | Accent 6 |
| `msoThemeColorHyperlink` | 11 | Hyperlink |
| `msoThemeColorFollowedHyperlink` | 12 | Followed hyperlink |

---

## File: `ppt_com/text.py` - Text Operations

### Purpose

Provide MCP tools for all text-related operations: set/get text, partial formatting via Characters(), font settings, paragraph formatting, bullet configuration, find/replace, and TextFrame properties.

---

### Tool: `set_text`

- **Description**: Set the entire text content of a shape
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Shape name (string) or 1-based index (int) |
  | `text` | string | Yes | Text content. Use `\r` for paragraph breaks (NOT `\n`) |
- **Returns**:
  ```json
  {
    "status": "success",
    "slide_index": 1,
    "shape_name": "TextBox 1",
    "text_length": 25,
    "paragraph_count": 3
  }
  ```
- **COM Implementation**:
  ```python
  def set_text(app, slide_index, shape_name_or_index, text):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = _get_shape(slide, shape_name_or_index)

      if not shape.HasTextFrame:
          raise ValueError(f"Shape '{shape.Name}' does not have a text frame")

      tf = shape.TextFrame
      tr = tf.TextRange

      # IMPORTANT: PowerPoint uses \r (CR) for paragraph breaks, NOT \n (LF)
      # Convert any \n to \r for consistency
      text = text.replace('\n', '\r')

      tr.Text = text

      return {
          "status": "success",
          "slide_index": slide_index,
          "shape_name": shape.Name,
          "text_length": tr.Length,
          "paragraph_count": tr.Paragraphs().Count,
      }
  ```
- **Error Cases**:
  - Shape not found on slide
  - Shape does not have a text frame (e.g., picture, line)

---

### Tool: `get_text`

- **Description**: Get the text content of a shape, optionally with formatting information
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Shape name or 1-based index |
  | `include_formatting` | bool | No | If true, return per-run formatting info. Default: false |
- **Returns**:
  ```json
  {
    "status": "success",
    "shape_name": "TextBox 1",
    "text": "Hello World",
    "text_length": 11,
    "paragraph_count": 1,
    "paragraphs": [
      {
        "index": 1,
        "text": "Hello World",
        "indent_level": 1,
        "alignment": 1
      }
    ],
    "runs": [
      {
        "index": 1,
        "text": "Hello ",
        "start": 1,
        "length": 6,
        "font_name": "Arial",
        "font_size": 18.0,
        "bold": false,
        "italic": false,
        "underline": false,
        "color_rgb": "#FF0000"
      }
    ]
  }
  ```
- **COM Implementation**:
  ```python
  def get_text(app, slide_index, shape_name_or_index, include_formatting=False):
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

      # Always include paragraph info
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

      if include_formatting:
          runs = []
          run_count = tr.Runs().Count
          for i in range(1, run_count + 1):
              run = tr.Runs(i)
              font = run.Font
              color_int = font.Color.RGB
              r, g, b = int_to_rgb(color_int)
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
  ```
- **Error Cases**:
  - Shape not found on slide
  - Shape does not have a text frame

---

### Tool: `format_text`

- **Description**: Apply formatting to all text in a shape, or to a specific character range using `Characters(start, length)`. This is the MOST IMPORTANT tool for partial text formatting.
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Shape name or 1-based index |
  | `start` | int | No | 1-based character start position. Omit to format ALL text |
  | `length` | int | No | Number of characters. Omit with start to format 1 char at start pos |
  | `font_name` | string | No | Font name (e.g., "Arial", "Yu Gothic") |
  | `font_size` | float | No | Font size in points |
  | `bold` | bool | No | Bold on/off |
  | `italic` | bool | No | Italic on/off |
  | `underline` | bool | No | Underline on/off |
  | `color` | string | No | Color as "#RRGGBB" hex string |
  | `theme_color` | int | No | Theme color index (1-12, see MsoThemeColorIndex) |
  | `shadow` | bool | No | Text shadow on/off |
  | `emboss` | bool | No | Emboss on/off |
  | `subscript` | bool | No | Subscript on/off |
  | `superscript` | bool | No | Superscript on/off |
  | `baseline_offset` | float | No | Baseline offset (-1.0 to 1.0, negative=subscript, positive=superscript) |
  | `font_name_ascii` | string | No | Font for ASCII characters |
  | `font_name_far_east` | string | No | Font for East Asian characters |
- **Returns**:
  ```json
  {
    "status": "success",
    "shape_name": "TextBox 1",
    "formatted_text": "Hello",
    "start": 1,
    "length": 5
  }
  ```
- **COM Implementation**:
  ```python
  def format_text(app, slide_index, shape_name_or_index,
                  start=None, length=None,
                  font_name=None, font_size=None,
                  bold=None, italic=None, underline=None,
                  color=None, theme_color=None,
                  shadow=None, emboss=None,
                  subscript=None, superscript=None,
                  baseline_offset=None,
                  font_name_ascii=None, font_name_far_east=None):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = _get_shape(slide, shape_name_or_index)

      if not shape.HasTextFrame:
          raise ValueError(f"Shape '{shape.Name}' does not have a text frame")

      tr = shape.TextFrame.TextRange

      # Determine which TextRange to format
      # KEY PATTERN: Characters(Start, Length) is the core of partial formatting
      if start is not None and length is not None:
          # Format specific character range
          target = tr.Characters(start, length)
      elif start is not None:
          # Format single character at position
          target = tr.Characters(start)
      else:
          # Format ALL text
          target = tr

      font = target.Font

      # Apply font properties
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
          # color is "#RRGGBB" format, convert to PowerPoint BGR integer
          font.Color.RGB = hex_to_int(color)
      if theme_color is not None:
          font.Color.ObjectThemeColor = theme_color
      if shadow is not None:
          font.Shadow = msoTrue if shadow else msoFalse
      if emboss is not None:
          font.Emboss = msoTrue if emboss else msoFalse
      if subscript is not None:
          font.Subscript = msoTrue if subscript else msoFalse
      if superscript is not None:
          font.Superscript = msoTrue if superscript else msoFalse
      if baseline_offset is not None:
          font.BaselineOffset = baseline_offset
      if font_name_ascii is not None:
          font.NameAscii = font_name_ascii
      if font_name_far_east is not None:
          font.NameFarEast = font_name_far_east

      return {
          "status": "success",
          "shape_name": shape.Name,
          "formatted_text": target.Text,
          "start": target.Start,
          "length": target.Length,
      }
  ```
- **Error Cases**:
  - Shape not found
  - Shape has no text frame
  - Start position exceeds text length (COM auto-adjusts to last char, may produce unexpected results)
  - Both `color` and `theme_color` specified (theme_color takes precedence as it is set last)

### CRITICAL PATTERN: Partial Text Formatting with Characters()

The `Characters(Start, Length)` method is the fundamental mechanism for partial text formatting in PowerPoint COM. Unlike python-pptx which requires pre-splitting runs, COM allows formatting ANY arbitrary character range on the fly. PowerPoint internally splits runs as needed.

```python
# Example: Rainbow text
tr = shape.TextFrame.TextRange
tr.Text = "RAINBOW"
tr.Font.Size = 48

colors = [
    rgb_to_int(255, 0, 0),     # R - Red
    rgb_to_int(255, 165, 0),   # A - Orange
    rgb_to_int(255, 255, 0),   # I - Yellow
    rgb_to_int(0, 128, 0),     # N - Green
    rgb_to_int(0, 0, 255),     # B - Blue
    rgb_to_int(75, 0, 130),    # O - Indigo
    rgb_to_int(238, 130, 238), # W - Violet
]
for i, c in enumerate(colors):
    tr.Characters(i + 1, 1).Font.Color.RGB = c

# Example: Bold+Red a keyword
tr.Text = "Error: File not found"
found = tr.Find("Error")
if found is not None:
    found.Font.Color.RGB = rgb_to_int(255, 0, 0)
    found.Font.Bold = msoTrue  # -1

# Example: Build formatted text incrementally using InsertAfter
tr.Text = ""
part1 = tr.InsertAfter("Sales Report: ")
part1.Font.Size = 20
part1.Font.Bold = msoTrue
part2 = tr.InsertAfter("+15%")
part2.Font.Size = 20
part2.Font.Bold = msoTrue
part2.Font.Color.RGB = rgb_to_int(0, 128, 0)  # Green
```

**Key Rules for Characters()**:
1. Start is **1-based** (not 0-based like Python)
2. If Start > text length, it wraps to the last character
3. If Length > remaining text, it covers until end of text
4. Both Start and Length omitted = all characters
5. Only Start specified = single character at that position

---

### Tool: `format_paragraph`

- **Description**: Apply paragraph-level formatting to all paragraphs or a specific paragraph range
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Shape name or 1-based index |
  | `paragraph_index` | int | No | 1-based paragraph index. Omit to format ALL paragraphs |
  | `paragraph_count` | int | No | Number of paragraphs from paragraph_index. Default: 1 |
  | `alignment` | int | No | PpParagraphAlignment: 1=left, 2=center, 3=right, 4=justify, 5=distribute |
  | `space_before` | float | No | Space before paragraph (points or line multiple) |
  | `space_after` | float | No | Space after paragraph (points or line multiple) |
  | `space_within` | float | No | Line spacing within paragraph (points or line multiple) |
  | `line_rule_before` | bool | No | If true, space_before is line multiple; if false, it is points |
  | `line_rule_after` | bool | No | If true, space_after is line multiple; if false, it is points |
  | `line_rule_within` | bool | No | If true, space_within is line multiple; if false, it is points |
  | `indent_level` | int | No | Indent level (1-9) |
  | `text_direction` | int | No | Text direction (1=LTR, 2=RTL) |
- **Returns**:
  ```json
  {
    "status": "success",
    "shape_name": "TextBox 1",
    "paragraph_index": 1,
    "paragraph_count": 1
  }
  ```
- **COM Implementation**:
  ```python
  def format_paragraph(app, slide_index, shape_name_or_index,
                       paragraph_index=None, paragraph_count=1,
                       alignment=None, space_before=None, space_after=None,
                       space_within=None, line_rule_before=None,
                       line_rule_after=None, line_rule_within=None,
                       indent_level=None, text_direction=None):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = _get_shape(slide, shape_name_or_index)

      if not shape.HasTextFrame:
          raise ValueError(f"Shape '{shape.Name}' does not have a text frame")

      tr = shape.TextFrame.TextRange

      # Get the target paragraph range
      if paragraph_index is not None:
          target = tr.Paragraphs(paragraph_index, paragraph_count)
      else:
          target = tr  # All paragraphs
          paragraph_index = 1
          paragraph_count = tr.Paragraphs().Count

      pf = target.ParagraphFormat

      if alignment is not None:
          pf.Alignment = alignment
      if indent_level is not None:
          target.IndentLevel = indent_level

      # Line spacing rules: LineRule* must be set BEFORE the corresponding Space* value
      if line_rule_before is not None:
          pf.LineRuleBefore = msoTrue if line_rule_before else msoFalse
      if space_before is not None:
          pf.SpaceBefore = space_before

      if line_rule_after is not None:
          pf.LineRuleAfter = msoTrue if line_rule_after else msoFalse
      if space_after is not None:
          pf.SpaceAfter = space_after

      if line_rule_within is not None:
          pf.LineRuleWithin = msoTrue if line_rule_within else msoFalse
      if space_within is not None:
          pf.SpaceWithin = space_within

      if text_direction is not None:
          pf.TextDirection = text_direction

      return {
          "status": "success",
          "shape_name": shape.Name,
          "paragraph_index": paragraph_index,
          "paragraph_count": paragraph_count,
      }
  ```
- **Error Cases**:
  - Shape not found
  - Shape has no text frame
  - paragraph_index exceeds number of paragraphs

### Line Spacing Rules

The `LineRuleWithin`, `LineRuleBefore`, and `LineRuleAfter` properties control how the corresponding `SpaceWithin`, `SpaceBefore`, and `SpaceAfter` values are interpreted:

| LineRule value | Space value interpretation |
|-------|-------------|
| `msoTrue` (-1) | Line multiple (e.g., 1.5 = 1.5x line height) |
| `msoFalse` (0) | Points (e.g., 12 = 12 points) |

IMPORTANT: Set the LineRule* property BEFORE setting the corresponding Space* value. If they are set in the wrong order, the value may be misinterpreted.

---

### Tool: `set_bullet`

- **Description**: Configure bullet/numbering for paragraphs
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Shape name or 1-based index |
  | `paragraph_index` | int | No | 1-based paragraph index. Omit for all paragraphs |
  | `paragraph_count` | int | No | Number of paragraphs. Default: 1 |
  | `bullet_type` | int | No | 0=none, 1=unnumbered(symbol), 2=numbered |
  | `bullet_character` | int | No | Unicode code point for bullet symbol (e.g., 8226 for bullet dot) |
  | `bullet_font_name` | string | No | Font for bullet character |
  | `bullet_color` | string | No | Bullet color as "#RRGGBB" |
  | `bullet_relative_size` | float | No | Bullet size relative to text (0.25-4.0) |
  | `numbered_style` | int | No | PpNumberedBulletStyle (e.g., 3 for "1. 2. 3.") |
  | `start_value` | int | No | Starting number for numbered bullets (1-32767) |
  | `use_text_color` | bool | No | Use text color for bullet |
  | `use_text_font` | bool | No | Use text font for bullet |
- **Returns**:
  ```json
  {
    "status": "success",
    "shape_name": "Content Placeholder 2",
    "paragraph_index": 1,
    "bullet_type": 1
  }
  ```
- **COM Implementation**:
  ```python
  def set_bullet(app, slide_index, shape_name_or_index,
                 paragraph_index=None, paragraph_count=1,
                 bullet_type=None, bullet_character=None,
                 bullet_font_name=None, bullet_color=None,
                 bullet_relative_size=None, numbered_style=None,
                 start_value=None, use_text_color=None,
                 use_text_font=None):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = _get_shape(slide, shape_name_or_index)

      if not shape.HasTextFrame:
          raise ValueError(f"Shape '{shape.Name}' does not have a text frame")

      tr = shape.TextFrame.TextRange

      if paragraph_index is not None:
          target = tr.Paragraphs(paragraph_index, paragraph_count)
      else:
          target = tr

      bullet = target.ParagraphFormat.Bullet

      if bullet_type is not None:
          if bullet_type == 0:
              bullet.Visible = msoFalse  # No bullet
          else:
              bullet.Visible = msoTrue
              bullet.Type = bullet_type

      if bullet_character is not None:
          bullet.Character = bullet_character

      if bullet_font_name is not None:
          bullet.Font.Name = bullet_font_name

      if bullet_color is not None:
          bullet.Font.Color.RGB = hex_to_int(bullet_color)

      if bullet_relative_size is not None:
          bullet.RelativeSize = bullet_relative_size

      if numbered_style is not None:
          bullet.Style = numbered_style

      if start_value is not None:
          bullet.StartValue = start_value

      if use_text_color is not None:
          bullet.UseTextColor = msoTrue if use_text_color else msoFalse

      if use_text_font is not None:
          bullet.UseTextFont = msoTrue if use_text_font else msoFalse

      return {
          "status": "success",
          "shape_name": shape.Name,
          "paragraph_index": paragraph_index or "all",
          "bullet_type": bullet_type,
      }
  ```
- **Error Cases**:
  - Shape not found
  - Shape has no text frame
  - Invalid bullet_type value

---

### Tool: `find_and_replace`

- **Description**: Find text and optionally replace it within a shape. Can also apply formatting to found text.
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Shape name or 1-based index |
  | `find_what` | string | Yes | Text to search for |
  | `replace_with` | string | No | Replacement text. Omit to just find (no replace) |
  | `match_case` | bool | No | Case-sensitive search. Default: false |
  | `whole_words` | bool | No | Match whole words only. Default: false |
  | `color` | string | No | Apply this color to found/replaced text ("#RRGGBB") |
  | `bold` | bool | No | Apply bold to found/replaced text |
  | `font_size` | float | No | Apply font size to found/replaced text |
- **Returns**:
  ```json
  {
    "status": "success",
    "found": true,
    "occurrences": [
      {
        "start": 1,
        "length": 5,
        "text": "Hello"
      }
    ],
    "replaced": true,
    "replacement_text": "World"
  }
  ```
- **COM Implementation**:
  ```python
  def find_and_replace(app, slide_index, shape_name_or_index,
                       find_what, replace_with=None,
                       match_case=False, whole_words=False,
                       color=None, bold=None, font_size=None):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = _get_shape(slide, shape_name_or_index)

      if not shape.HasTextFrame:
          raise ValueError(f"Shape '{shape.Name}' does not have a text frame")

      tr = shape.TextFrame.TextRange
      occurrences = []

      if replace_with is not None:
          # Replace mode: use Replace method
          # Replace replaces only the FIRST occurrence
          # To replace all, we loop
          replaced_count = 0
          while True:
              result = tr.Replace(
                  FindWhat=find_what,
                  ReplaceWhat=replace_with,
                  After=0,
                  MatchCase=msoTrue if match_case else msoFalse,
                  WholeWords=msoTrue if whole_words else msoFalse,
              )
              if result is None:
                  break
              replaced_count += 1
              # Apply formatting to the replaced text if requested
              if color is not None:
                  result.Font.Color.RGB = hex_to_int(color)
              if bold is not None:
                  result.Font.Bold = msoTrue if bold else msoFalse
              if font_size is not None:
                  result.Font.Size = font_size
              occurrences.append({
                  "start": result.Start,
                  "length": result.Length,
                  "text": result.Text,
              })

          return {
              "status": "success",
              "found": replaced_count > 0,
              "occurrences": occurrences,
              "replaced": True,
              "replacement_text": replace_with,
          }
      else:
          # Find mode: use Find method to locate all occurrences
          after = 0
          while True:
              found = tr.Find(
                  FindWhat=find_what,
                  After=after,
                  MatchCase=msoTrue if match_case else msoFalse,
                  WholeWords=msoTrue if whole_words else msoFalse,
              )
              if found is None:
                  break
              # Apply formatting to found text if requested
              if color is not None:
                  found.Font.Color.RGB = hex_to_int(color)
              if bold is not None:
                  found.Font.Bold = msoTrue if bold else msoFalse
              if font_size is not None:
                  found.Font.Size = font_size
              occurrences.append({
                  "start": found.Start,
                  "length": found.Length,
                  "text": found.Text,
              })
              after = found.Start + found.Length - 1
              if after >= tr.Length:
                  break

          return {
              "status": "success",
              "found": len(occurrences) > 0,
              "occurrences": occurrences,
              "replaced": False,
          }
  ```
- **Error Cases**:
  - Shape not found
  - Shape has no text frame
  - find_what is empty string

---

### Tool: `set_textframe_properties`

- **Description**: Configure TextFrame properties: margins, word wrap, auto-size, orientation, vertical anchor
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Shape name or 1-based index |
  | `margin_left` | float | No | Left margin in points |
  | `margin_right` | float | No | Right margin in points |
  | `margin_top` | float | No | Top margin in points |
  | `margin_bottom` | float | No | Bottom margin in points |
  | `word_wrap` | bool | No | Enable/disable word wrap |
  | `auto_size` | int | No | 0=none, 1=shape-to-fit-text |
  | `orientation` | int | No | MsoTextOrientation value (1=horizontal, 5=vertical, etc.) |
  | `vertical_anchor` | int | No | MsoVerticalAnchor value (1=top, 3=middle, 4=bottom) |
- **Returns**:
  ```json
  {
    "status": "success",
    "shape_name": "TextBox 1"
  }
  ```
- **COM Implementation**:
  ```python
  def set_textframe_properties(app, slide_index, shape_name_or_index,
                               margin_left=None, margin_right=None,
                               margin_top=None, margin_bottom=None,
                               word_wrap=None, auto_size=None,
                               orientation=None, vertical_anchor=None):
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
      if auto_size is not None:
          tf.AutoSize = auto_size  # ppAutoSizeNone=0, ppAutoSizeShapeToFitText=1
      if orientation is not None:
          tf.Orientation = orientation
      if vertical_anchor is not None:
          tf.VerticalAnchor = vertical_anchor

      return {
          "status": "success",
          "shape_name": shape.Name,
      }
  ```
- **Error Cases**:
  - Shape not found
  - Shape has no text frame

---

### Tool: `insert_text`

- **Description**: Insert text before or after existing text, or at a specific position. The inserted text range is returned so it can be separately formatted.
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Shape name or 1-based index |
  | `text` | string | Yes | Text to insert (use \r for paragraph breaks) |
  | `position` | string | No | "before" or "after" (default: "after") |
  | `font_name` | string | No | Font for inserted text |
  | `font_size` | float | No | Font size for inserted text |
  | `bold` | bool | No | Bold for inserted text |
  | `color` | string | No | Color for inserted text as "#RRGGBB" |
- **Returns**:
  ```json
  {
    "status": "success",
    "shape_name": "TextBox 1",
    "inserted_start": 12,
    "inserted_length": 6,
    "full_text": "Hello World Added!"
  }
  ```
- **COM Implementation**:
  ```python
  def insert_text(app, slide_index, shape_name_or_index,
                  text, position="after",
                  font_name=None, font_size=None,
                  bold=None, color=None):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = _get_shape(slide, shape_name_or_index)

      if not shape.HasTextFrame:
          raise ValueError(f"Shape '{shape.Name}' does not have a text frame")

      tr = shape.TextFrame.TextRange
      text = text.replace('\n', '\r')

      if position == "before":
          inserted = tr.InsertBefore(text)
      else:
          inserted = tr.InsertAfter(text)

      # InsertBefore/InsertAfter return the TextRange of the inserted text
      # Apply formatting to the inserted portion
      if font_name is not None:
          inserted.Font.Name = font_name
      if font_size is not None:
          inserted.Font.Size = font_size
      if bold is not None:
          inserted.Font.Bold = msoTrue if bold else msoFalse
      if color is not None:
          inserted.Font.Color.RGB = hex_to_int(color)

      return {
          "status": "success",
          "shape_name": shape.Name,
          "inserted_start": inserted.Start,
          "inserted_length": inserted.Length,
          "full_text": shape.TextFrame.TextRange.Text,
      }
  ```
- **Error Cases**:
  - Shape not found
  - Shape has no text frame
  - Invalid position value (not "before" or "after")

---

### Tool: `set_indent_and_tabs`

- **Description**: Configure ruler-level indentation and tab stops for the text frame
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Shape name or 1-based index |
  | `level` | int | No | Indent level to configure (1-5). If omitted, sets levels via indent_settings |
  | `first_margin` | float | No | First-line indent in points for the specified level |
  | `left_margin` | float | No | Left margin in points for the specified level |
  | `tab_stops` | list | No | List of tab stops: [{"type": 1, "position": 72.0}, ...] type: 1=left, 2=center, 3=right, 4=decimal |
- **Returns**:
  ```json
  {
    "status": "success",
    "shape_name": "TextBox 1",
    "level": 1
  }
  ```
- **COM Implementation**:
  ```python
  def set_indent_and_tabs(app, slide_index, shape_name_or_index,
                          level=None, first_margin=None, left_margin=None,
                          tab_stops=None):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = _get_shape(slide, shape_name_or_index)

      if not shape.HasTextFrame:
          raise ValueError(f"Shape '{shape.Name}' does not have a text frame")

      ruler = shape.TextFrame.Ruler

      if level is not None:
          ruler_level = ruler.Levels(level)
          if first_margin is not None:
              ruler_level.FirstMargin = first_margin
          if left_margin is not None:
              ruler_level.LeftMargin = left_margin

      if tab_stops is not None:
          tab_stops_obj = ruler.TabStops
          for tab in tab_stops:
              tab_stops_obj.Add(tab["type"], tab["position"])

      return {
          "status": "success",
          "shape_name": shape.Name,
          "level": level,
      }
  ```
- **Error Cases**:
  - Shape not found
  - Shape has no text frame
  - level not in range 1-5
  - tab_stops items missing required keys

---

### Helper Function: `_get_shape`

```python
def _get_shape(slide, name_or_index):
    """
    Get a shape by name (string) or 1-based index (int).

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
        # Search by name
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            if shape.Name == name_or_index:
                return shape
        raise ValueError(f"Shape '{name_or_index}' not found on slide")
```

---

## Implementation Notes

### 1. Paragraph Breaks: Use \r, NOT \n

PowerPoint COM uses carriage return (`\r`, CR) for paragraph breaks. Line feed (`\n`, LF) does NOT create paragraph breaks. Always convert `\n` to `\r` before setting text:

```python
text = text.replace('\n', '\r')
```

### 2. Characters() Index is 1-Based

COM's `Characters(Start, Length)` uses 1-based indexing. When converting from Python 0-based indices, add 1:

```python
# Python: position 0, length 5
# COM: Characters(1, 5)
python_start = 0
com_start = python_start + 1
```

### 3. RGB Color Values are BGR-Encoded

PowerPoint COM stores RGB as `R + G*256 + B*65536` (BGR byte order). This means `0x0000FF` is RED (not blue). Always use the `rgb_to_int` and `hex_to_int` helper functions from `utils/color.py`:

```python
# "#FF0000" (red) becomes: 255 + 0*256 + 0*65536 = 255 (0x0000FF in hex)
color_int = hex_to_int("#FF0000")  # Returns 255
```

### 4. MsoTriState Values

COM uses MsoTriState for boolean-like properties:
- `msoTrue` = -1 (True)
- `msoFalse` = 0 (False)
- `msoTriStateMixed` = -2 (Mixed/ambiguous)

When reading: compare with constants, do NOT use Python truthy/falsy:
```python
# WRONG: if font.Bold:  (msoTrue=-1 is truthy, but msoFalse=0 is falsy - happens to work but not reliable)
# RIGHT:
is_bold = font.Bold == msoTrue
```

### 5. Font.Color.RGB Sets Color Type Automatically

Setting `Font.Color.RGB` automatically changes `Font.Color.Type` to `msoColorTypeRGB` (1). If you want to use theme colors instead, set `Font.Color.ObjectThemeColor` after any RGB setting.

### 6. InsertBefore/InsertAfter Return Value

Both `InsertBefore(text)` and `InsertAfter(text)` return a `TextRange` object representing the newly inserted text. This is the recommended way to apply formatting to inserted text:

```python
inserted = tr.InsertAfter("NEW TEXT")
inserted.Font.Bold = msoTrue  # Bold only the inserted text
```

After insertion, the positions of existing Characters() references change. Re-fetch any character ranges after inserting.

### 7. Find Method Returns None When Not Found

`TextRange.Find(FindWhat)` returns `None` (not an error) when the text is not found. Always check:

```python
found = tr.Find("keyword")
if found is not None:
    found.Font.Bold = msoTrue
```

### 8. Replace Method Replaces Only First Occurrence

`TextRange.Replace(FindWhat, ReplaceWhat)` replaces only the FIRST occurrence and returns the TextRange of the replacement. Loop to replace all:

```python
while True:
    result = tr.Replace(FindWhat="old", ReplaceWhat="new")
    if result is None:
        break
```

### 9. TextFrame.HasText Check

Before reading text, check if the TextFrame has text:
```python
if shape.HasTextFrame:
    tf = shape.TextFrame
    if tf.HasText:  # msoTrue = -1
        text = tf.TextRange.Text
```

### 10. Performance with Large Text

COM calls have overhead per call. For operations on many characters (e.g., coloring each character individually), batch operations where possible. For example, format contiguous ranges rather than individual characters:

```python
# Slow: one call per character
for i in range(1, 1000):
    tr.Characters(i, 1).Font.Color.RGB = some_color

# Fast: one call for the entire range
tr.Characters(1, 999).Font.Color.RGB = some_color
```

### 11. Mixed Font Properties

When reading font properties from a TextRange that spans multiple runs with different formatting, the property returns `msoTriStateMixed` (-2) for boolean properties or a special "mixed" value. Handle this in the return data:

```python
bold_val = tr.Font.Bold
if bold_val == msoTriStateMixed:
    bold = "mixed"
elif bold_val == msoTrue:
    bold = True
else:
    bold = False
```
