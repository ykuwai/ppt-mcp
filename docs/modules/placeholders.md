# Module: Placeholder Operations

## Overview

This module handles all operations related to PowerPoint placeholders: listing placeholders on slides and layouts, getting placeholder info, setting placeholder text and formatting, working with SlideMaster/CustomLayout placeholders, managing the inheritance chain (Master > Layout > Slide), and controlling HeadersFooters (footer text, slide numbers, date/time). Placeholders are a KEY FEATURE because they are the primary mechanism for structured content in PowerPoint templates.

## Dependencies

- **Internal**: `utils.com_wrapper` (PowerPointCOMWrapper, safe_com_call), `utils.units` (inches_to_points, cm_to_points, points_to_inches, points_to_cm), `utils.color` (rgb_to_int, int_to_rgb, int_to_hex, hex_to_int), `ppt_com.constants` (all placeholder/layout/master constants)
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
    # Placeholder types
    ppPlaceholderTitle, ppPlaceholderBody, ppPlaceholderCenterTitle,
    ppPlaceholderSubtitle, ppPlaceholderVerticalTitle,
    ppPlaceholderVerticalBody, ppPlaceholderObject,
    ppPlaceholderChart, ppPlaceholderBitmap, ppPlaceholderMediaClip,
    ppPlaceholderOrgChart, ppPlaceholderTable,
    ppPlaceholderSlideNumber, ppPlaceholderHeader,
    ppPlaceholderFooter, ppPlaceholderDate,
    ppPlaceholderVerticalObject, ppPlaceholderPicture,
    # Shape types
    msoPlaceholder, msoAutoShape, msoTable, msoChart,
    msoPicture, msoSmartArt, msoMedia,
    # Date formats
    ppDateTimeMdyy, ppDateTimeddddMMMMddyyyy,
    ppDateTimedMMMMyyyy, ppDateTimeMMMMdyyyy,
    ppDateTimeHmm, ppDateTimeHmmss,
)
import logging

logger = logging.getLogger(__name__)
```

## File Structure

```
ppt_com_mcp/
  ppt_com/
    placeholders.py   # Placeholder operations
```

---

## Constants Needed

These constants MUST be defined in `ppt_com/constants.py` (from the core module). Listed here for reference:

### PpPlaceholderType - Complete List

| Name | Value | Description |
|------|-------|-------------|
| `ppPlaceholderMixed` | -2 | Mixed (multiple selection) |
| `ppPlaceholderTitle` | 1 | Title |
| `ppPlaceholderBody` | 2 | Body / Content |
| `ppPlaceholderCenterTitle` | 3 | Center title (title slide) |
| `ppPlaceholderSubtitle` | 4 | Subtitle (title slide) |
| `ppPlaceholderVerticalTitle` | 5 | Vertical title |
| `ppPlaceholderVerticalBody` | 6 | Vertical body |
| `ppPlaceholderObject` | 7 | Object |
| `ppPlaceholderChart` | 8 | Chart |
| `ppPlaceholderBitmap` | 9 | Bitmap |
| `ppPlaceholderMediaClip` | 10 | Media clip |
| `ppPlaceholderOrgChart` | 11 | Organization chart |
| `ppPlaceholderTable` | 12 | Table |
| `ppPlaceholderSlideNumber` | 13 | Slide number |
| `ppPlaceholderHeader` | 14 | Header (notes/handout only) |
| `ppPlaceholderFooter` | 15 | Footer |
| `ppPlaceholderDate` | 16 | Date |
| `ppPlaceholderVerticalObject` | 17 | Vertical object |
| `ppPlaceholderPicture` | 18 | Picture |
| `ppPlaceholderCameo` | 19 | Cameo (PowerPoint 365 live camera) |

### Placeholder Type Name Map (for readable output)

```python
PLACEHOLDER_TYPE_NAMES = {
    -2: "Mixed",
    1: "Title",
    2: "Body",
    3: "CenterTitle",
    4: "Subtitle",
    5: "VerticalTitle",
    6: "VerticalBody",
    7: "Object",
    8: "Chart",
    9: "Bitmap",
    10: "MediaClip",
    11: "OrgChart",
    12: "Table",
    13: "SlideNumber",
    14: "Header",
    15: "Footer",
    16: "Date",
    17: "VerticalObject",
    18: "Picture",
    19: "Cameo",
}
```

### ContainedType Values (MsoShapeType for what is inside a placeholder)

| Name | Value | Description |
|------|-------|-------------|
| `msoAutoShape` | 1 | AutoShape (text content / empty) |
| `msoChart` | 3 | Chart |
| `msoLinkedPicture` | 11 | Linked picture |
| `msoPicture` | 13 | Picture |
| `msoPlaceholder` | 14 | Empty placeholder |
| `msoMedia` | 16 | Media |
| `msoTable` | 19 | Table |
| `msoSmartArt` | 24 | SmartArt |

### PpDateTimeFormat (commonly used)

| Name | Value | Description | Example |
|------|-------|-------------|---------|
| `ppDateTimeMdyy` | 1 | Month/Day/Year (short) | 2/16/26 |
| `ppDateTimeddddMMMMddyyyy` | 2 | Day Month Day Year | Monday, February 16, 2026 |
| `ppDateTimedMMMMyyyy` | 3 | Day Month Year | 16 February 2026 |
| `ppDateTimeMMMMdyyyy` | 4 | Month Day Year | February 16, 2026 |
| `ppDateTimedMMMyy` | 5 | Day Month Year (short) | 16-Feb-26 |
| `ppDateTimeMMMMyy` | 6 | Month Year | February 26 |
| `ppDateTimeMMyy` | 7 | Month/Year | 2/26 |
| `ppDateTimeHmm` | 10 | Hour:Minute (24h) | 14:30 |
| `ppDateTimeHmmss` | 11 | Hour:Minute:Second (24h) | 14:30:45 |
| `ppDateTimehmmAMPM` | 12 | Hour:Minute AM/PM | 2:30 PM |
| `ppDateTimeFigureOut` | 14 | Auto-detect | - |

---

## File: `ppt_com/placeholders.py` - Placeholder Operations

### Purpose

Provide MCP tools for listing, querying, and manipulating placeholders on slides, layouts, and masters. Also handles HeadersFooters management and the inheritance chain.

---

### Tool: `list_placeholders`

- **Description**: List all placeholders on a slide with their type, name, position, size, and content info
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
- **Returns**:
  ```json
  {
    "status": "success",
    "slide_index": 1,
    "placeholder_count": 3,
    "placeholders": [
      {
        "index": 1,
        "name": "Title 1",
        "type": 1,
        "type_name": "Title",
        "contained_type": 1,
        "contained_type_name": "AutoShape",
        "left": 457.2,
        "top": 274.6,
        "width": 8229.6,
        "height": 1143.0,
        "has_text_frame": true,
        "has_text": true,
        "text": "Presentation Title"
      },
      {
        "index": 2,
        "name": "Content Placeholder 2",
        "type": 2,
        "type_name": "Body",
        "contained_type": 1,
        "contained_type_name": "AutoShape",
        "left": 457.2,
        "top": 1600.2,
        "width": 8229.6,
        "height": 4525.9,
        "has_text_frame": true,
        "has_text": false,
        "text": null
      }
    ]
  }
  ```
- **COM Implementation**:
  ```python
  CONTAINED_TYPE_NAMES = {
      1: "AutoShape",
      3: "Chart",
      11: "LinkedPicture",
      13: "Picture",
      14: "Placeholder",
      16: "Media",
      19: "Table",
      24: "SmartArt",
  }

  def list_placeholders(app, slide_index):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      phs = slide.Shapes.Placeholders

      placeholders = []
      for i in range(1, phs.Count + 1):
          ph = phs(i)
          pf = ph.PlaceholderFormat

          info = {
              "index": i,
              "name": ph.Name,
              "type": pf.Type,
              "type_name": PLACEHOLDER_TYPE_NAMES.get(pf.Type, f"Unknown({pf.Type})"),
              "contained_type": pf.ContainedType,
              "contained_type_name": CONTAINED_TYPE_NAMES.get(
                  pf.ContainedType, f"Unknown({pf.ContainedType})"
              ),
              "left": round(ph.Left, 1),
              "top": round(ph.Top, 1),
              "width": round(ph.Width, 1),
              "height": round(ph.Height, 1),
              "has_text_frame": bool(ph.HasTextFrame),
              "has_text": False,
              "text": None,
          }

          if ph.HasTextFrame:
              tf = ph.TextFrame
              has_text = tf.HasText
              info["has_text"] = bool(has_text)
              if has_text:
                  info["text"] = tf.TextRange.Text

          placeholders.append(info)

      return {
          "status": "success",
          "slide_index": slide_index,
          "placeholder_count": phs.Count,
          "placeholders": placeholders,
      }
  ```
- **Error Cases**:
  - Invalid slide_index

---

### Tool: `get_placeholder`

- **Description**: Get detailed information about a specific placeholder, including formatting details
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `placeholder_index` | int | No | 1-based placeholder index (from list_placeholders) |
  | `placeholder_type` | int | No | PpPlaceholderType value. If set, finds first placeholder of this type |
- **Returns**:
  ```json
  {
    "status": "success",
    "placeholder": {
      "index": 1,
      "name": "Title 1",
      "type": 1,
      "type_name": "Title",
      "contained_type": 1,
      "left": 457.2,
      "top": 274.6,
      "width": 8229.6,
      "height": 1143.0,
      "rotation": 0.0,
      "has_text_frame": true,
      "text": "My Title",
      "paragraph_count": 1,
      "font_name": "Calibri",
      "font_size": 44.0,
      "alignment": 2
    }
  }
  ```
- **COM Implementation**:
  ```python
  def get_placeholder(app, slide_index, placeholder_index=None,
                      placeholder_type=None):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)

      ph = None
      if placeholder_index is not None:
          ph = slide.Shapes.Placeholders(placeholder_index)
      elif placeholder_type is not None:
          ph = _find_placeholder_by_type(slide, placeholder_type)
          if ph is None:
              raise ValueError(
                  f"No placeholder of type {placeholder_type} "
                  f"({PLACEHOLDER_TYPE_NAMES.get(placeholder_type, 'Unknown')}) "
                  f"found on slide {slide_index}"
              )
      else:
          raise ValueError("Must specify either placeholder_index or placeholder_type")

      pf = ph.PlaceholderFormat
      info = {
          "index": pf.Type,  # PlaceholderFormat index
          "name": ph.Name,
          "type": pf.Type,
          "type_name": PLACEHOLDER_TYPE_NAMES.get(pf.Type, f"Unknown({pf.Type})"),
          "contained_type": pf.ContainedType,
          "left": round(ph.Left, 1),
          "top": round(ph.Top, 1),
          "width": round(ph.Width, 1),
          "height": round(ph.Height, 1),
          "rotation": ph.Rotation,
          "has_text_frame": bool(ph.HasTextFrame),
      }

      if ph.HasTextFrame:
          tf = ph.TextFrame
          tr = tf.TextRange
          info["text"] = tr.Text if tf.HasText else None
          info["paragraph_count"] = tr.Paragraphs().Count
          if tf.HasText:
              info["font_name"] = tr.Font.Name
              try:
                  info["font_size"] = tr.Font.Size
              except Exception:
                  info["font_size"] = None  # Mixed sizes
              info["alignment"] = tr.ParagraphFormat.Alignment

      return {
          "status": "success",
          "placeholder": info,
      }
  ```
- **Error Cases**:
  - Invalid slide_index
  - Placeholder not found at given index
  - No placeholder of given type on slide
  - Neither placeholder_index nor placeholder_type provided

---

### Tool: `set_placeholder_text`

- **Description**: Set the text content of a placeholder. Supports finding by index or by type.
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `placeholder_index` | int | No | 1-based placeholder index |
  | `placeholder_type` | int | No | PpPlaceholderType value (1=Title, 2=Body, 3=CenterTitle, 4=Subtitle) |
  | `text` | string | Yes | Text content (use \r for paragraph breaks) |
- **Returns**:
  ```json
  {
    "status": "success",
    "slide_index": 1,
    "placeholder_name": "Title 1",
    "placeholder_type": 1,
    "text_length": 15
  }
  ```
- **COM Implementation**:
  ```python
  def set_placeholder_text(app, slide_index,
                           placeholder_index=None, placeholder_type=None,
                           text=""):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)

      ph = _resolve_placeholder(slide, placeholder_index, placeholder_type)

      if not ph.HasTextFrame:
          raise ValueError(
              f"Placeholder '{ph.Name}' does not have a text frame "
              f"(type={ph.PlaceholderFormat.Type})"
          )

      text = text.replace('\n', '\r')
      ph.TextFrame.TextRange.Text = text

      return {
          "status": "success",
          "slide_index": slide_index,
          "placeholder_name": ph.Name,
          "placeholder_type": ph.PlaceholderFormat.Type,
          "text_length": ph.TextFrame.TextRange.Length,
      }
  ```
- **Error Cases**:
  - Placeholder not found
  - Placeholder has no text frame (e.g., picture placeholder)
  - Neither placeholder_index nor placeholder_type provided

---

### Tool: `list_layouts`

- **Description**: List all custom layouts available in the presentation, including their placeholders
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `design_index` | int | No | 1-based design (master) index. Default: 1 (first/default design) |
- **Returns**:
  ```json
  {
    "status": "success",
    "design_name": "Office Theme",
    "layout_count": 11,
    "layouts": [
      {
        "index": 1,
        "name": "Title Slide",
        "matching_name": "title",
        "placeholder_count": 2,
        "placeholders": [
          {
            "type": 3,
            "type_name": "CenterTitle",
            "name": "Title 1"
          },
          {
            "type": 4,
            "type_name": "Subtitle",
            "name": "Subtitle 2"
          }
        ]
      },
      {
        "index": 2,
        "name": "Title and Content",
        "matching_name": "objTx",
        "placeholder_count": 2,
        "placeholders": [
          {
            "type": 1,
            "type_name": "Title",
            "name": "Title 1"
          },
          {
            "type": 2,
            "type_name": "Body",
            "name": "Content Placeholder 2"
          }
        ]
      }
    ]
  }
  ```
- **COM Implementation**:
  ```python
  def list_layouts(app, design_index=1):
      pres = app.ActivePresentation
      design = pres.Designs(design_index)
      master = design.SlideMaster
      layouts_col = master.CustomLayouts

      layouts = []
      for i in range(1, layouts_col.Count + 1):
          layout = layouts_col(i)
          phs = layout.Shapes.Placeholders

          ph_list = []
          for j in range(1, phs.Count + 1):
              ph = phs(j)
              pf = ph.PlaceholderFormat
              ph_list.append({
                  "type": pf.Type,
                  "type_name": PLACEHOLDER_TYPE_NAMES.get(pf.Type, f"Unknown({pf.Type})"),
                  "name": ph.Name,
              })

          layouts.append({
              "index": i,
              "name": layout.Name,
              "matching_name": layout.MatchingName,
              "placeholder_count": phs.Count,
              "placeholders": ph_list,
          })

      return {
          "status": "success",
          "design_name": design.Name,
          "layout_count": layouts_col.Count,
          "layouts": layouts,
      }
  ```
- **Error Cases**:
  - Invalid design_index

---

### Tool: `list_masters`

- **Description**: List all slide masters (designs) in the presentation
- **Parameters**: None
- **Returns**:
  ```json
  {
    "status": "success",
    "master_count": 1,
    "masters": [
      {
        "design_index": 1,
        "design_name": "Office Theme",
        "master_name": "Office Theme",
        "layout_count": 11,
        "has_title_master": false
      }
    ]
  }
  ```
- **COM Implementation**:
  ```python
  def list_masters(app):
      pres = app.ActivePresentation
      designs = pres.Designs

      masters = []
      for i in range(1, designs.Count + 1):
          design = designs(i)
          master = design.SlideMaster
          masters.append({
              "design_index": i,
              "design_name": design.Name,
              "master_name": master.Name,
              "layout_count": master.CustomLayouts.Count,
              "has_title_master": bool(design.HasTitleMaster),
          })

      return {
          "status": "success",
          "master_count": designs.Count,
          "masters": masters,
      }
  ```
- **Error Cases**:
  - No presentation open

---

### Tool: `get_slide_layout_info`

- **Description**: Get the layout and master information for a specific slide
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
- **Returns**:
  ```json
  {
    "status": "success",
    "slide_index": 1,
    "layout_name": "Title Slide",
    "design_name": "Office Theme",
    "master_name": "Office Theme",
    "follow_master_background": true,
    "display_master_shapes": true
  }
  ```
- **COM Implementation**:
  ```python
  def get_slide_layout_info(app, slide_index):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)

      layout = slide.CustomLayout
      design = slide.Design
      master = slide.Master  # Direct access to slide's master

      return {
          "status": "success",
          "slide_index": slide_index,
          "layout_name": layout.Name,
          "design_name": design.Name,
          "master_name": master.Name,
          "follow_master_background": slide.FollowMasterBackground == msoTrue,
          "display_master_shapes": slide.DisplayMasterShapes == msoTrue,
      }
  ```
- **Error Cases**:
  - Invalid slide_index

---

### Tool: `set_slide_layout`

- **Description**: Change the layout of a slide
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `layout_index` | int | No | 1-based layout index (from list_layouts) |
  | `layout_name` | string | No | Layout name to search for |
  | `design_index` | int | No | 1-based design index. Default: use slide's current design |
- **Returns**:
  ```json
  {
    "status": "success",
    "slide_index": 1,
    "new_layout_name": "Title and Content"
  }
  ```
- **COM Implementation**:
  ```python
  def set_slide_layout(app, slide_index,
                       layout_index=None, layout_name=None,
                       design_index=None):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)

      # Determine which master to use
      if design_index is not None:
          master = pres.Designs(design_index).SlideMaster
      else:
          master = slide.Master  # Current slide's master

      layouts = master.CustomLayouts

      target_layout = None
      if layout_index is not None:
          target_layout = layouts(layout_index)
      elif layout_name is not None:
          for i in range(1, layouts.Count + 1):
              if layouts(i).Name == layout_name:
                  target_layout = layouts(i)
                  break
          if target_layout is None:
              raise ValueError(f"Layout '{layout_name}' not found")
      else:
          raise ValueError("Must specify either layout_index or layout_name")

      # Setting CustomLayout changes the slide's layout
      # Existing content is preserved as much as possible
      slide.CustomLayout = target_layout

      return {
          "status": "success",
          "slide_index": slide_index,
          "new_layout_name": target_layout.Name,
      }
  ```
- **Error Cases**:
  - Invalid slide_index
  - Layout not found
  - Neither layout_index nor layout_name specified

---

### Tool: `set_headers_footers`

- **Description**: Configure headers, footers, slide numbers, and date/time display
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | No | 1-based slide index. Omit to set on master (affects all slides) |
  | `footer_visible` | bool | No | Show/hide footer |
  | `footer_text` | string | No | Footer text content |
  | `slide_number_visible` | bool | No | Show/hide slide number |
  | `date_visible` | bool | No | Show/hide date/time |
  | `date_use_format` | bool | No | true = auto-updating date, false = fixed text |
  | `date_format` | int | No | PpDateTimeFormat value (1-14). Only when date_use_format=true |
  | `date_text` | string | No | Fixed date text. Only when date_use_format=false |
  | `display_on_title_slide` | bool | No | Show on title slides (master-level only) |
- **Returns**:
  ```json
  {
    "status": "success",
    "target": "master",
    "footer_visible": true,
    "footer_text": "Confidential",
    "slide_number_visible": true,
    "date_visible": true
  }
  ```
- **COM Implementation**:
  ```python
  def set_headers_footers(app, slide_index=None,
                          footer_visible=None, footer_text=None,
                          slide_number_visible=None,
                          date_visible=None, date_use_format=None,
                          date_format=None, date_text=None,
                          display_on_title_slide=None):
      pres = app.ActivePresentation

      if slide_index is not None:
          # Slide-level settings
          target = pres.Slides(slide_index).HeadersFooters
          target_label = f"slide {slide_index}"
      else:
          # Master-level settings (affects all slides without individual overrides)
          target = pres.SlideMaster.HeadersFooters
          target_label = "master"

      if footer_visible is not None:
          target.Footer.Visible = msoTrue if footer_visible else msoFalse
      if footer_text is not None:
          target.Footer.Text = footer_text

      if slide_number_visible is not None:
          target.SlideNumber.Visible = msoTrue if slide_number_visible else msoFalse

      if date_visible is not None:
          target.DateAndTime.Visible = msoTrue if date_visible else msoFalse

      if date_use_format is not None:
          target.DateAndTime.UseFormat = msoTrue if date_use_format else msoFalse

      if date_format is not None:
          # Only applicable when UseFormat = msoTrue
          target.DateAndTime.Format = date_format

      if date_text is not None:
          # Only applicable when UseFormat = msoFalse (fixed text mode)
          target.DateAndTime.Text = date_text

      if display_on_title_slide is not None and slide_index is None:
          # DisplayOnTitleSlide is only available on master-level HeadersFooters
          target.DisplayOnTitleSlide = msoTrue if display_on_title_slide else msoFalse

      return {
          "status": "success",
          "target": target_label,
          "footer_visible": footer_visible,
          "footer_text": footer_text,
          "slide_number_visible": slide_number_visible,
          "date_visible": date_visible,
      }
  ```
- **Error Cases**:
  - Invalid slide_index
  - Setting display_on_title_slide on slide-level (only works on master)

---

### Tool: `get_headers_footers`

- **Description**: Get current headers/footers settings for a slide or master
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | No | 1-based slide index. Omit to get master settings |
- **Returns**:
  ```json
  {
    "status": "success",
    "target": "slide 1",
    "footer": {
      "visible": true,
      "text": "Confidential"
    },
    "slide_number": {
      "visible": true
    },
    "date_and_time": {
      "visible": true,
      "use_format": true,
      "format": 1,
      "text": ""
    },
    "display_on_title_slide": true
  }
  ```
- **COM Implementation**:
  ```python
  def get_headers_footers(app, slide_index=None):
      pres = app.ActivePresentation

      if slide_index is not None:
          hf = pres.Slides(slide_index).HeadersFooters
          target_label = f"slide {slide_index}"
      else:
          hf = pres.SlideMaster.HeadersFooters
          target_label = "master"

      result = {
          "status": "success",
          "target": target_label,
          "footer": {
              "visible": hf.Footer.Visible == msoTrue,
              "text": hf.Footer.Text,
          },
          "slide_number": {
              "visible": hf.SlideNumber.Visible == msoTrue,
          },
          "date_and_time": {
              "visible": hf.DateAndTime.Visible == msoTrue,
              "use_format": hf.DateAndTime.UseFormat == msoTrue,
              "format": hf.DateAndTime.Format,
              "text": hf.DateAndTime.Text,
          },
      }

      if slide_index is None:
          result["display_on_title_slide"] = hf.DisplayOnTitleSlide == msoTrue

      return result
  ```
- **Error Cases**:
  - Invalid slide_index

---

### Tool: `clear_slide_headers_footers`

- **Description**: Clear individual slide header/footer overrides, reverting to master settings
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
- **Returns**:
  ```json
  {
    "status": "success",
    "slide_index": 1,
    "message": "Headers/footers cleared, now following master settings"
  }
  ```
- **COM Implementation**:
  ```python
  def clear_slide_headers_footers(app, slide_index):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      slide.HeadersFooters.Clear()

      return {
          "status": "success",
          "slide_index": slide_index,
          "message": "Headers/footers cleared, now following master settings",
      }
  ```
- **Error Cases**:
  - Invalid slide_index

---

### Tool: `set_master_text_style`

- **Description**: Modify the master-level text styles (title, body, default) that are inherited by all slides
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `style_type` | int | Yes | 1=Title, 2=Body, 3=Default |
  | `level` | int | Yes | Indent level (1-5) |
  | `font_name` | string | No | Font name |
  | `font_size` | float | No | Font size in points |
  | `bold` | bool | No | Bold on/off |
  | `italic` | bool | No | Italic on/off |
  | `color` | string | No | Color as "#RRGGBB" |
  | `alignment` | int | No | PpParagraphAlignment value |
  | `design_index` | int | No | 1-based design index. Default: 1 |
- **Returns**:
  ```json
  {
    "status": "success",
    "style_type": 1,
    "level": 1,
    "font_name": "Yu Gothic UI Semibold",
    "font_size": 36.0
  }
  ```
- **COM Implementation**:
  ```python
  def set_master_text_style(app, style_type, level,
                            font_name=None, font_size=None,
                            bold=None, italic=None, color=None,
                            alignment=None, design_index=1):
      pres = app.ActivePresentation
      master = pres.Designs(design_index).SlideMaster

      # style_type: 1=ppTitleStyle, 2=ppBodyStyle, 3=ppDefaultStyle
      style = master.TextStyles(style_type)
      lvl = style.Levels(level)
      font = lvl.Font

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
          lvl.ParagraphFormat.Alignment = alignment

      return {
          "status": "success",
          "style_type": style_type,
          "level": level,
          "font_name": font_name,
          "font_size": font_size,
      }
  ```
- **Error Cases**:
  - Invalid style_type (must be 1, 2, or 3)
  - Invalid level (must be 1-5)
  - Invalid design_index

---

### Tool: `get_theme_colors`

- **Description**: Get the current theme color scheme
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `design_index` | int | No | 1-based design index. Default: 1 |
- **Returns**:
  ```json
  {
    "status": "success",
    "colors": [
      {"index": 1, "name": "Dark1", "rgb": "#000000"},
      {"index": 2, "name": "Light1", "rgb": "#FFFFFF"},
      {"index": 3, "name": "Dark2", "rgb": "#44546A"},
      {"index": 4, "name": "Light2", "rgb": "#E7E6E6"},
      {"index": 5, "name": "Accent1", "rgb": "#4472C4"},
      {"index": 6, "name": "Accent2", "rgb": "#ED7D31"},
      {"index": 7, "name": "Accent3", "rgb": "#A5A5A5"},
      {"index": 8, "name": "Accent4", "rgb": "#FFC000"},
      {"index": 9, "name": "Accent5", "rgb": "#5B9BD5"},
      {"index": 10, "name": "Accent6", "rgb": "#70AD47"},
      {"index": 11, "name": "Hyperlink", "rgb": "#0563C1"},
      {"index": 12, "name": "FollowedHyperlink", "rgb": "#954F72"}
    ]
  }
  ```
- **COM Implementation**:
  ```python
  THEME_COLOR_NAMES = [
      "Dark1", "Light1", "Dark2", "Light2",
      "Accent1", "Accent2", "Accent3", "Accent4",
      "Accent5", "Accent6", "Hyperlink", "FollowedHyperlink",
  ]

  def get_theme_colors(app, design_index=1):
      pres = app.ActivePresentation
      master = pres.Designs(design_index).SlideMaster
      tcs = master.Theme.ThemeColorScheme

      colors = []
      for i in range(1, min(tcs.Count + 1, 13)):
          color_val = tcs(i).RGB
          hex_color = int_to_hex(color_val)
          colors.append({
              "index": i,
              "name": THEME_COLOR_NAMES[i - 1] if i <= 12 else f"Color{i}",
              "rgb": hex_color,
          })

      return {
          "status": "success",
          "colors": colors,
      }
  ```
- **Error Cases**:
  - Invalid design_index

---

### Helper Functions

```python
def _find_placeholder_by_type(slide, placeholder_type):
    """
    Find the first placeholder of a given type on a slide.

    Args:
        slide: Slide COM object
        placeholder_type: PpPlaceholderType value

    Returns:
        Shape COM object, or None if not found

    Note: For title slides, ppPlaceholderTitle (1) may not exist --
          use ppPlaceholderCenterTitle (3) instead.
    """
    phs = slide.Shapes.Placeholders
    for i in range(1, phs.Count + 1):
        ph = phs(i)
        if ph.PlaceholderFormat.Type == placeholder_type:
            return ph
    return None


def _resolve_placeholder(slide, placeholder_index=None, placeholder_type=None):
    """
    Resolve a placeholder by index or type.

    Args:
        slide: Slide COM object
        placeholder_index: 1-based placeholder index
        placeholder_type: PpPlaceholderType value

    Returns:
        Shape COM object

    Raises:
        ValueError: If placeholder not found or neither parameter specified
    """
    if placeholder_index is not None:
        return slide.Shapes.Placeholders(placeholder_index)
    elif placeholder_type is not None:
        ph = _find_placeholder_by_type(slide, placeholder_type)
        if ph is None:
            # For title: try CenterTitle as fallback
            if placeholder_type == ppPlaceholderTitle:
                ph = _find_placeholder_by_type(slide, ppPlaceholderCenterTitle)
            # For subtitle: try Body as fallback
            elif placeholder_type == ppPlaceholderSubtitle:
                ph = _find_placeholder_by_type(slide, ppPlaceholderBody)

            if ph is None:
                raise ValueError(
                    f"No placeholder of type "
                    f"{PLACEHOLDER_TYPE_NAMES.get(placeholder_type, placeholder_type)} "
                    f"found on slide"
                )
        return ph
    else:
        raise ValueError("Must specify either placeholder_index or placeholder_type")
```

---

## Implementation Notes

### 1. Placeholders vs Shapes

Placeholders ARE shapes (Shape objects) but with an additional `PlaceholderFormat` property. A shape is a placeholder if its `Type` property equals `msoPlaceholder` (14):

```python
if shape.Type == 14:  # msoPlaceholder
    pf = shape.PlaceholderFormat
    ph_type = pf.Type  # PpPlaceholderType value
```

### 2. Placeholders Collection vs Shapes Collection

The `Slide.Shapes.Placeholders` collection contains ONLY placeholder shapes. The `Slide.Shapes` collection contains ALL shapes including placeholders. Placeholders have their own indexing that is independent of the Shapes collection index.

### 3. Inheritance Chain: Master > Layout > Slide

The formatting inheritance flows as:

```
SlideMaster (TextStyles, Background, HeadersFooters)
    |
    v
CustomLayout (inherits from master, can override)
    |
    v
Slide (inherits from layout, can override)
```

When a property is explicitly set at the slide level, it "overrides" the inherited value. To revert to inherited values:
- Background: `slide.FollowMasterBackground = msoTrue`
- HeadersFooters: `slide.HeadersFooters.Clear()`
- Master shapes visibility: `slide.DisplayMasterShapes = msoTrue`

Text formatting overrides on individual placeholders CANNOT be easily reset via COM -- the only way is to delete the placeholder and re-apply the layout.

### 4. Title Slides Use CenterTitle, Not Title

Standard title slides use `ppPlaceholderCenterTitle` (3) and `ppPlaceholderSubtitle` (4), NOT `ppPlaceholderTitle` (1). Content slides use `ppPlaceholderTitle` (1) and `ppPlaceholderBody` (2). The `_resolve_placeholder` helper handles this by falling back to CenterTitle when Title is not found.

### 5. Placeholder Indices Can Have Gaps

If a placeholder is deleted from a slide, the remaining placeholders keep their original indices. This means indices may not be contiguous (e.g., 1, 3, 5 instead of 1, 2, 3). Always iterate using a for-loop, never assume contiguous indices.

### 6. Image Placeholders (ppPlaceholderPicture = 18)

COM API does NOT have a direct method to insert an image INTO a picture placeholder (the way the UI does). The practical approach is:

1. Get the placeholder's position and size (Left, Top, Width, Height)
2. Add a picture using `Shapes.AddPicture` at that position
3. Optionally delete the original placeholder

```python
ph = slide.Shapes.Placeholders(2)
left, top, width, height = ph.Left, ph.Top, ph.Width, ph.Height
pic = slide.Shapes.AddPicture(
    FileName=r"C:\path\to\image.jpg",
    LinkToFile=0,       # msoFalse
    SaveWithDocument=-1, # msoTrue
    Left=left, Top=top, Width=width, Height=height
)
```

### 7. Shapes.Title Shortcut

`Slide.Shapes.Title` is a shortcut to access the title placeholder directly:

```python
title = slide.Shapes.Title
if title is not None:
    title.TextFrame.TextRange.Text = "My Title"
```

This is equivalent to `Placeholders(1)` on most layouts, but more readable.

### 8. HeadersFooters Scope

- **Master-level**: Settings apply to all slides that don't have individual overrides
- **Slide-level**: Override master settings for that specific slide
- `HeadersFooters.Clear()`: Remove slide-level overrides, revert to master
- **Header property**: Only available on NotesMaster and HandoutMaster, NOT on slides

### 9. Multiple Designs (Masters)

A presentation can have multiple designs, each with its own SlideMaster and set of CustomLayouts. Access via `Presentation.Designs` collection:

```python
for i in range(1, pres.Designs.Count + 1):
    design = pres.Designs(i)
    master = design.SlideMaster
```

A slide's design/master can be queried via `Slide.Design` and `Slide.Master`.

### 10. Layout Names Are Locale-Dependent

Standard layout names like "Title Slide", "Title and Content", etc., may be localized. In Japanese environments, they might be "タイトル スライド", "タイトルとコンテンツ", etc. When searching by name, consider using `MatchingName` property which is locale-independent, or search by placeholder types instead.
