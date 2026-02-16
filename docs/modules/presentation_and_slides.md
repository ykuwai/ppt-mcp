# Module: Presentation & Slide Operations

## Overview

This module handles all operations related to PowerPoint presentations (creating, opening, saving, closing, properties) and slides (adding, deleting, duplicating, moving, listing, notes, backgrounds, transitions, and copying between presentations). It builds on the core infrastructure from `utils/com_wrapper.py`, `utils/units.py`, `utils/color.py`, and `ppt_com/constants.py`.

## Dependencies

- **Internal**: `utils.com_wrapper` (PowerPointCOMWrapper, safe_com_call), `utils.units` (points/inches/cm conversions), `utils.color` (rgb_to_int, int_to_rgb, int_to_hex), `ppt_com.constants` (all enumeration constants)
- **External**: `pywin32` (`win32com.client`, `pywintypes`)
- **Standard library**: `os`, `logging`

### Importing from Core Module

```python
from utils.com_wrapper import PowerPointCOMWrapper, safe_com_call
from utils.units import inches_to_points, points_to_inches, cm_to_points, points_to_cm
from utils.color import rgb_to_int, int_to_rgb, int_to_hex
from ppt_com.constants import (
    ppLayoutTitle, ppLayoutText, ppLayoutBlank, ppLayoutTitleOnly,
    ppLayoutSectionHeader, ppLayoutComparison, ppLayoutCustom,
    ppSaveAsOpenXMLPresentation, ppSaveAsPDF, ppSaveAsPNG, ppSaveAsJPG,
    ppSaveAsDefault, ppFixedFormatTypePDF, ppFixedFormatTypeXPS,
    ppFixedFormatIntentScreen, ppFixedFormatIntentPrint,
    ppTransitionSpeedFast, ppTransitionSpeedMedium, ppTransitionSpeedSlow,
    ppEffectNone, ppEffectFade, ppEffectPush, ppEffectWipe,
    msoTrue, msoFalse,
)
import logging
import os

logger = logging.getLogger(__name__)
```

## File Structure

```
ppt_com_mcp/
  ppt_com/
    presentation.py   # Presentation-level operations
    slides.py         # Slide-level operations
```

---

## File: `ppt_com/presentation.py` - Presentation Operations

### Purpose

Provide MCP tools for creating, opening, saving, closing, and querying PowerPoint presentations. Also covers page setup (slide size), document properties, template/theme application, and section management.

---

### Tool: `create_presentation`

- **Description**: Create a new empty PowerPoint presentation
- **Parameters**: None
- **Returns**:
  ```json
  {
    "success": true,
    "name": "Presentation1",
    "slides_count": 0
  }
  ```
- **COM Implementation**:
  ```python
  def create_presentation(app):
      """Create a new empty presentation."""
      pres = app.Presentations.Add()
      return {
          "success": True,
          "name": pres.Name,
          "slides_count": pres.Slides.Count,
      }
  ```
- **Error Cases**: COM connection lost (reconnect)

---

### Tool: `open_presentation`

- **Description**: Open an existing presentation file
- **Parameters**:
  - `file_path` (str, required): Full path to the .pptx/.pptm/.ppt file
  - `read_only` (bool, optional, default=False): Open in read-only mode
  - `untitled` (bool, optional, default=False): Open as untitled copy
  - `with_window` (bool, optional, default=True): Whether to show a window
- **Returns**:
  ```json
  {
    "success": true,
    "name": "MyPresentation.pptx",
    "full_name": "C:\\path\\to\\MyPresentation.pptx",
    "slides_count": 15,
    "read_only": false
  }
  ```
- **COM Implementation**:
  ```python
  def open_presentation(app, file_path, read_only=False, untitled=False, with_window=True):
      """
      Open an existing presentation.

      Presentations.Open parameters:
        FileName (str): File path (required)
        ReadOnly (MsoTriState): Open read-only
        Untitled (MsoTriState): Open as untitled copy
        WithWindow (MsoTriState): Show window (default msoTrue)

      Supported formats: .pptx, .pptm, .ppt, .potx, .potm, .ppsx, .ppsm,
                         .pps, .ppam, .htm, .rtf, .odp, etc.
      """
      if not os.path.exists(file_path):
          raise FileNotFoundError(f"File not found: {file_path}")

      pres = app.Presentations.Open(
          FileName=file_path,
          ReadOnly=msoTrue if read_only else msoFalse,
          Untitled=msoTrue if untitled else msoFalse,
          WithWindow=msoTrue if with_window else msoFalse,
      )
      return {
          "success": True,
          "name": pres.Name,
          "full_name": pres.FullName,
          "slides_count": pres.Slides.Count,
          "read_only": bool(pres.ReadOnly),
      }
  ```
- **Error Cases**:
  - File not found: Raise FileNotFoundError before COM call
  - Password-protected file: No password parameter in `Presentations.Open`; a dialog may appear. Consider warning the user.
  - Invalid file format: COM error

---

### Tool: `save_presentation`

- **Description**: Save the active presentation (overwrite)
- **Parameters**: None (saves the active presentation)
- **Returns**: `{"success": true, "name": "MyPresentation.pptx", "saved": true}`
- **COM Implementation**:
  ```python
  def save_presentation(app):
      """Save the active presentation."""
      pres = app.ActivePresentation
      pres.Save()
      return {
          "success": True,
          "name": pres.Name,
          "saved": bool(pres.Saved),
      }
  ```
- **Error Cases**: No presentation open, file is read-only, file path invalid

---

### Tool: `save_presentation_as`

- **Description**: Save the active presentation with a new name and/or format
- **Parameters**:
  - `file_path` (str, required): Target file path
  - `file_format` (int, optional, default=24): PpSaveAsFileType constant. Common values:
    - `24` = ppSaveAsOpenXMLPresentation (.pptx)
    - `32` = ppSaveAsPDF (.pdf)
    - `18` = ppSaveAsPNG (each slide as .png)
    - `17` = ppSaveAsJPG (each slide as .jpg)
    - `1` = ppSaveAsPresentation (.ppt legacy)
    - `28` = ppSaveAsOpenXMLShow (.ppsx)
    - `39` = ppSaveAsMP4 (.mp4)
    - `37` = ppSaveAsWMV (.wmv)
    - `11` = ppSaveAsDefault
  - `embed_fonts` (bool, optional, default=False): Embed fonts in the file
- **Returns**: `{"success": true, "name": "NewName.pptx", "full_name": "C:\\path\\NewName.pptx"}`
- **COM Implementation**:
  ```python
  def save_presentation_as(app, file_path, file_format=24, embed_fonts=False):
      """
      Save the active presentation with a new name/format.

      IMPORTANT: SaveAs changes the presentation's FullName to the new path.
      To save a copy without changing the name, use save_copy_as instead.

      When file_format is an image type (PNG=18, JPG=17, BMP=19, GIF=16, TIF=21, EMF=23),
      a folder is created at the specified path and individual slide images are saved inside.
      """
      pres = app.ActivePresentation
      kwargs = {
          "FileName": file_path,
          "FileFormat": file_format,
      }
      if embed_fonts:
          kwargs["EmbedFonts"] = msoTrue
      pres.SaveAs(**kwargs)
      return {
          "success": True,
          "name": pres.Name,
          "full_name": pres.FullName,
      }
  ```
- **Error Cases**: No presentation open, invalid path, unsupported format

---

### Tool: `save_copy_as`

- **Description**: Save a copy of the active presentation without changing the current file name
- **Parameters**:
  - `file_path` (str, required): Target file path for the copy
- **Returns**: `{"success": true, "copy_path": "C:\\path\\backup.pptx"}`
- **COM Implementation**:
  ```python
  def save_copy_as(app, file_path):
      """
      Save a copy without changing the presentation's FullName.
      NOTE: SaveCopyAs does NOT support file format conversion.
      The copy will be in the same format as the original.
      """
      pres = app.ActivePresentation
      pres.SaveCopyAs(file_path)
      return {
          "success": True,
          "copy_path": file_path,
      }
  ```

---

### Tool: `export_as_pdf`

- **Description**: Export the active presentation as PDF with detailed control
- **Parameters**:
  - `file_path` (str, required): Output PDF file path
  - `intent` (str, optional, default="print"): "print" (high quality) or "screen" (smaller file)
  - `include_hidden_slides` (bool, optional, default=False): Include hidden slides
  - `frame_slides` (bool, optional, default=False): Draw frame around slides
- **Returns**: `{"success": true, "pdf_path": "C:\\output\\file.pdf"}`
- **COM Implementation**:
  ```python
  def export_as_pdf(app, file_path, intent="print", include_hidden_slides=False, frame_slides=False):
      """
      Export using ExportAsFixedFormat for fine-grained PDF control.

      ExportAsFixedFormat parameters:
        Path (str): Output path
        FixedFormatType: ppFixedFormatTypePDF=2 or ppFixedFormatTypeXPS=1
        Intent: ppFixedFormatIntentPrint=2 or ppFixedFormatIntentScreen=1
        FrameSlides (MsoTriState): Draw frame around slides
        PrintHiddenSlides (MsoTriState): Include hidden slides
        RangeType: ppPrintAll=1 (all slides)
      """
      pres = app.ActivePresentation
      intent_val = 2 if intent == "print" else 1  # ppFixedFormatIntentPrint / Screen
      pres.ExportAsFixedFormat(
          Path=file_path,
          FixedFormatType=2,  # ppFixedFormatTypePDF
          Intent=intent_val,
          FrameSlides=msoTrue if frame_slides else msoFalse,
          PrintHiddenSlides=msoTrue if include_hidden_slides else msoFalse,
          RangeType=1,  # ppPrintAll
          PrintRange=None,
          IncludeDocProperties=True,
          DocStructureTags=True,
          BitmapMissingFonts=True,
          UseISO19005_1=False,
      )
      return {"success": True, "pdf_path": file_path}
  ```

---

### Tool: `close_presentation`

- **Description**: Close a presentation
- **Parameters**:
  - `save_changes` (bool, optional, default=False): Whether to save before closing. If False, unsaved changes are discarded.
- **Returns**: `{"success": true, "closed": "MyPresentation.pptx"}`
- **COM Implementation**:
  ```python
  def close_presentation(app, save_changes=False):
      """
      Close the active presentation.

      To avoid save dialogs, set pres.Saved = True before Close()
      when save_changes is False.
      """
      pres = app.ActivePresentation
      name = pres.Name
      if save_changes:
          pres.Save()
      else:
          pres.Saved = True  # Mark as "no changes" to suppress dialog
      pres.Close()
      return {"success": True, "closed": name}
  ```

---

### Tool: `get_presentation_info`

- **Description**: Get detailed information about the active presentation
- **Parameters**: None
- **Returns**:
  ```json
  {
    "name": "MyPresentation.pptx",
    "full_name": "C:\\path\\to\\MyPresentation.pptx",
    "path": "C:\\path\\to",
    "slides_count": 15,
    "read_only": false,
    "saved": true,
    "slide_width": 960.0,
    "slide_height": 540.0,
    "slide_width_inches": 13.333,
    "slide_height_inches": 7.5,
    "first_slide_number": 1,
    "template_name": "Office Theme"
  }
  ```
- **COM Implementation**:
  ```python
  def get_presentation_info(app):
      pres = app.ActivePresentation
      page = pres.PageSetup
      return {
          "name": pres.Name,
          "full_name": pres.FullName,
          "path": pres.Path,
          "slides_count": pres.Slides.Count,
          "read_only": bool(pres.ReadOnly),
          "saved": bool(pres.Saved),
          "slide_width": page.SlideWidth,
          "slide_height": page.SlideHeight,
          "slide_width_inches": round(page.SlideWidth / 72.0, 3),
          "slide_height_inches": round(page.SlideHeight / 72.0, 3),
          "first_slide_number": page.FirstSlideNumber,
          "template_name": pres.TemplateName,
      }
  ```

---

### Tool: `set_slide_size`

- **Description**: Set the slide dimensions of the active presentation
- **Parameters**:
  - `width` (float, required): Width in points (72 points = 1 inch)
  - `height` (float, required): Height in points
  - `unit` (str, optional, default="points"): Unit system: "points", "inches", or "cm"
- **Returns**: `{"success": true, "width_points": 960.0, "height_points": 540.0}`
- **COM Implementation**:
  ```python
  def set_slide_size(app, width, height, unit="points"):
      """
      Set slide dimensions.

      Standard sizes:
        16:9 = 960x540 pt (13.333x7.5 inches)
        4:3 = 720x540 pt (10x7.5 inches)
        A4 landscape = 842x595 pt (~11.693x8.268 inches)

      PageSetup properties:
        SlideWidth (Single, R/W): Slide width in points
        SlideHeight (Single, R/W): Slide height in points
        FirstSlideNumber (Long, R/W): Starting slide number
      """
      pres = app.ActivePresentation
      if unit == "inches":
          width = width * 72.0
          height = height * 72.0
      elif unit == "cm":
          width = width * (72.0 / 2.54)
          height = height * (72.0 / 2.54)
      pres.PageSetup.SlideWidth = width
      pres.PageSetup.SlideHeight = height
      return {
          "success": True,
          "width_points": pres.PageSetup.SlideWidth,
          "height_points": pres.PageSetup.SlideHeight,
      }
  ```

---

### Tool: `get_document_properties`

- **Description**: Get built-in document properties of the active presentation
- **Parameters**: None
- **Returns**:
  ```json
  {
    "title": "My Title",
    "author": "John",
    "subject": "Topic",
    "keywords": "ppt, automation",
    "comments": "",
    "category": "Report",
    "company": "Acme",
    "manager": ""
  }
  ```
- **COM Implementation**:
  ```python
  def get_document_properties(app):
      """
      Read BuiltInDocumentProperties.

      Available property names: "Title", "Subject", "Author", "Keywords",
      "Comments", "Template", "Last Author", "Revision Number",
      "Application Name", "Last Print Date", "Creation Date",
      "Last Save Time", "Total Editing Time", "Number of Pages",
      "Number of Words", "Number of Characters", "Security",
      "Category", "Manager", "Company"
      """
      pres = app.ActivePresentation
      props = pres.BuiltInDocumentProperties
      result = {}
      for name in ["Title", "Author", "Subject", "Keywords", "Comments", "Category", "Company", "Manager"]:
          try:
              result[name.lower()] = str(props(name).Value)
          except Exception:
              result[name.lower()] = ""
      return result
  ```

---

### Tool: `set_document_properties`

- **Description**: Set built-in document properties of the active presentation
- **Parameters**:
  - `properties` (dict, required): Dictionary of property name to value. Valid keys: "title", "author", "subject", "keywords", "comments", "category", "company", "manager"
- **Returns**: `{"success": true, "updated": ["title", "author"]}`
- **COM Implementation**:
  ```python
  def set_document_properties(app, properties):
      """
      Set BuiltInDocumentProperties.
      Property names are case-insensitive in this function but must match
      the exact BuiltInDocumentProperties names when passed to COM.
      """
      PROP_MAP = {
          "title": "Title", "author": "Author", "subject": "Subject",
          "keywords": "Keywords", "comments": "Comments",
          "category": "Category", "company": "Company", "manager": "Manager",
      }
      pres = app.ActivePresentation
      props = pres.BuiltInDocumentProperties
      updated = []
      for key, value in properties.items():
          prop_name = PROP_MAP.get(key.lower())
          if prop_name:
              props(prop_name).Value = str(value)
              updated.append(key.lower())
      return {"success": True, "updated": updated}
  ```

---

### Tool: `apply_template`

- **Description**: Apply a template or theme file to the active presentation
- **Parameters**:
  - `file_path` (str, required): Path to template (.potx) or theme (.thmx) file
- **Returns**: `{"success": true}`
- **COM Implementation**:
  ```python
  def apply_template(app, file_path):
      """
      Apply a template or theme.
      .potx files: use pres.ApplyTemplate(path)
      .thmx files: use pres.ApplyTheme(path)
      """
      pres = app.ActivePresentation
      ext = os.path.splitext(file_path)[1].lower()
      if ext == ".thmx":
          pres.ApplyTheme(file_path)
      else:
          pres.ApplyTemplate(file_path)
      return {"success": True}
  ```

---

### Tool: `list_sections`

- **Description**: List all sections in the active presentation
- **Parameters**: None
- **Returns**:
  ```json
  {
    "sections": [
      {
        "index": 1,
        "name": "Introduction",
        "first_slide": 1,
        "slides_count": 3
      }
    ]
  }
  ```
- **COM Implementation**:
  ```python
  def list_sections(app):
      """
      List sections via SectionProperties.

      SectionProperties methods:
        Count: Number of sections
        Name(sectionIndex): Section name
        FirstSlide(sectionIndex): First slide number in section
        SlidesCount(sectionIndex): Number of slides in section
        SectionID(sectionIndex): Section unique ID
      """
      pres = app.ActivePresentation
      sections = pres.SectionProperties
      result = []
      for i in range(1, sections.Count + 1):
          result.append({
              "index": i,
              "name": sections.Name(i),
              "first_slide": sections.FirstSlide(i),
              "slides_count": sections.SlidesCount(i),
          })
      return {"sections": result}
  ```

---

### Tool: `add_section`

- **Description**: Add a section to the presentation
- **Parameters**:
  - `name` (str, required): Section name
  - `before_slide` (int, optional): Insert section before this slide index. If omitted, adds at the end.
- **Returns**: `{"success": true, "section_index": 2}`
- **COM Implementation**:
  ```python
  def add_section(app, name, before_slide=None):
      """
      Add a section.

      SectionProperties.AddBeforeSlide(slideIndex, sectionName): Add before specified slide
      SectionProperties.AddSection(sectionIndex, sectionName): Add at section position

      Max 512 sections per presentation.
      """
      pres = app.ActivePresentation
      sections = pres.SectionProperties
      if before_slide is not None:
          idx = sections.AddBeforeSlide(before_slide, name)
      else:
          idx = sections.AddSection(sections.Count + 1, name)
      return {"success": True, "section_index": idx}
  ```

---

### Tool: `rename_section`

- **Description**: Rename an existing section
- **Parameters**:
  - `section_index` (int, required): 1-based section index
  - `new_name` (str, required): New section name
- **Returns**: `{"success": true}`
- **COM Implementation**:
  ```python
  def rename_section(app, section_index, new_name):
      pres = app.ActivePresentation
      pres.SectionProperties.Rename(section_index, new_name)
      return {"success": True}
  ```

---

### Tool: `delete_section`

- **Description**: Delete a section
- **Parameters**:
  - `section_index` (int, required): 1-based section index
  - `delete_slides` (bool, optional, default=False): If True, also delete slides in the section
- **Returns**: `{"success": true}`
- **COM Implementation**:
  ```python
  def delete_section(app, section_index, delete_slides=False):
      pres = app.ActivePresentation
      pres.SectionProperties.Delete(section_index, delete_slides)
      return {"success": True}
  ```

---

## File: `ppt_com/slides.py` - Slide Operations

### Purpose

Provide MCP tools for managing individual slides: adding, deleting, duplicating, moving, listing, notes management, background settings, transitions, and copying between presentations.

---

### Tool: `add_slide`

- **Description**: Add a new slide to the active presentation
- **Parameters**:
  - `position` (int, optional, default=end): 1-based position for the new slide. If omitted, adds at the end.
  - `layout` (int, optional, default=12): PpSlideLayout constant. Common values:
    - `1` = ppLayoutTitle (title slide)
    - `2` = ppLayoutText (title + text)
    - `11` = ppLayoutTitleOnly (title only)
    - `12` = ppLayoutBlank (blank)
    - `33` = ppLayoutSectionHeader
    - `34` = ppLayoutComparison
    - `35` = ppLayoutContentWithCaption
    - `36` = ppLayoutPictureWithCaption
  - `layout_name` (str, optional): Custom layout name to use instead of layout integer. Searched from SlideMaster.CustomLayouts by Name property.
- **Returns**:
  ```json
  {
    "success": true,
    "slide_index": 3,
    "slide_id": 258,
    "layout": 12
  }
  ```
- **COM Implementation**:
  ```python
  def add_slide(app, position=None, layout=12, layout_name=None):
      """
      Add a slide using either a PpSlideLayout integer or a custom layout name.

      Two methods available:
        Slides.Add(Index, Layout): Legacy API, uses PpSlideLayout integer
        Slides.AddSlide(Index, pCustomLayout): New API, uses CustomLayout object

      If layout_name is provided, search for it in SlideMaster.CustomLayouts:
        for i in range(1, master.CustomLayouts.Count + 1):
            if master.CustomLayouts(i).Name == layout_name:
                return master.CustomLayouts(i)

      PpSlideLayout constants (from ppt_com.constants):
        ppLayoutTitle=1, ppLayoutText=2, ppLayoutTwoColumnText=3,
        ppLayoutTable=4, ppLayoutTitleOnly=11, ppLayoutBlank=12,
        ppLayoutCustom=32, ppLayoutSectionHeader=33, ppLayoutComparison=34,
        ppLayoutContentWithCaption=35, ppLayoutPictureWithCaption=36
      """
      pres = app.ActivePresentation
      if position is None:
          position = pres.Slides.Count + 1

      if layout_name:
          # Search for custom layout by name
          master = pres.SlideMaster
          custom_layout = None
          for i in range(1, master.CustomLayouts.Count + 1):
              if master.CustomLayouts(i).Name == layout_name:
                  custom_layout = master.CustomLayouts(i)
                  break
          if custom_layout is None:
              raise ValueError(
                  f"Layout '{layout_name}' not found. "
                  f"Use list_layouts to see available layouts."
              )
          slide = pres.Slides.AddSlide(Index=position, pCustomLayout=custom_layout)
      else:
          slide = pres.Slides.Add(Index=position, Layout=layout)

      return {
          "success": True,
          "slide_index": slide.SlideIndex,
          "slide_id": slide.SlideID,
          "layout": slide.Layout,
      }
  ```
- **Error Cases**: Position out of range, invalid layout constant, layout_name not found

---

### Tool: `delete_slide`

- **Description**: Delete a slide by index
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
- **Returns**: `{"success": true, "deleted_index": 3, "remaining_count": 14}`
- **COM Implementation**:
  ```python
  def delete_slide(app, slide_index):
      """
      Delete a slide. After deletion, remaining slides re-index automatically.
      When deleting multiple slides in a loop, delete from the end to avoid
      index shifting issues.
      """
      pres = app.ActivePresentation
      if slide_index < 1 or slide_index > pres.Slides.Count:
          raise ValueError(f"Slide index {slide_index} out of range (1-{pres.Slides.Count})")
      pres.Slides(slide_index).Delete()
      return {
          "success": True,
          "deleted_index": slide_index,
          "remaining_count": pres.Slides.Count,
      }
  ```

---

### Tool: `duplicate_slide`

- **Description**: Duplicate a slide (the copy is inserted immediately after the original)
- **Parameters**:
  - `slide_index` (int, required): 1-based index of the slide to duplicate
- **Returns**: `{"success": true, "new_slide_index": 4, "new_slide_id": 260}`
- **COM Implementation**:
  ```python
  def duplicate_slide(app, slide_index):
      """
      Duplicate a slide. The duplicate is inserted immediately after the original.
      Returns a SlideRange. To move the duplicate, use move_slide afterwards.

      slide.Duplicate() returns a SlideRange object.
      """
      pres = app.ActivePresentation
      dup_range = pres.Slides(slide_index).Duplicate()
      new_slide = dup_range(1)
      return {
          "success": True,
          "new_slide_index": new_slide.SlideIndex,
          "new_slide_id": new_slide.SlideID,
      }
  ```

---

### Tool: `move_slide`

- **Description**: Move a slide to a new position
- **Parameters**:
  - `slide_index` (int, required): Current 1-based index of the slide
  - `new_position` (int, required): Target 1-based position
- **Returns**: `{"success": true, "moved_from": 3, "moved_to": 1}`
- **COM Implementation**:
  ```python
  def move_slide(app, slide_index, new_position):
      """
      Move a slide using Slide.MoveTo(toPos).
      toPos is the target position (1-based).
      """
      pres = app.ActivePresentation
      pres.Slides(slide_index).MoveTo(toPos=new_position)
      return {
          "success": True,
          "moved_from": slide_index,
          "moved_to": new_position,
      }
  ```

---

### Tool: `list_slides`

- **Description**: List all slides with their key properties
- **Parameters**: None
- **Returns**:
  ```json
  {
    "slides_count": 3,
    "slides": [
      {
        "index": 1,
        "slide_id": 256,
        "name": "Slide1",
        "layout": 1,
        "layout_name": "Title Slide",
        "hidden": false,
        "has_notes": true
      }
    ]
  }
  ```
- **COM Implementation**:
  ```python
  def list_slides(app):
      """
      List all slides in the active presentation.

      Key Slide properties:
        SlideIndex: 1-based position in the Slides collection
        SlideID: Unique immutable ID (survives add/delete)
        SlideNumber: Display slide number (affected by FirstSlideNumber)
        Name: Slide name (default "Slide1", "Slide2", etc.)
        Layout: PpSlideLayout integer
        CustomLayout.Name: Readable layout name
        SlideShowTransition.Hidden: Whether slide is hidden
      """
      pres = app.ActivePresentation
      slides = []
      for i in range(1, pres.Slides.Count + 1):
          slide = pres.Slides(i)
          layout_name = ""
          try:
              layout_name = slide.CustomLayout.Name
          except Exception:
              pass

          has_notes = False
          try:
              notes_text = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
              has_notes = len(notes_text.strip()) > 0
          except Exception:
              pass

          slides.append({
              "index": slide.SlideIndex,
              "slide_id": slide.SlideID,
              "name": slide.Name,
              "layout": slide.Layout,
              "layout_name": layout_name,
              "hidden": bool(slide.SlideShowTransition.Hidden),
              "has_notes": has_notes,
          })
      return {"slides_count": pres.Slides.Count, "slides": slides}
  ```

---

### Tool: `get_slide_info`

- **Description**: Get detailed information about a specific slide
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
- **Returns**:
  ```json
  {
    "index": 1,
    "slide_id": 256,
    "slide_number": 1,
    "name": "Slide1",
    "layout": 1,
    "layout_name": "Title Slide",
    "hidden": false,
    "shapes_count": 3,
    "has_title": true,
    "title_text": "My Presentation",
    "notes_text": "Speaker notes here",
    "follow_master_background": true,
    "transition_effect": 0,
    "advance_on_click": true,
    "advance_on_time": false,
    "advance_time": 0,
    "design_name": "Office Theme"
  }
  ```
- **COM Implementation**:
  ```python
  def get_slide_info(app, slide_index):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      trans = slide.SlideShowTransition

      layout_name = ""
      try:
          layout_name = slide.CustomLayout.Name
      except Exception:
          pass

      title_text = ""
      has_title = bool(slide.Shapes.HasTitle)
      if has_title:
          try:
              title_text = slide.Shapes.Title.TextFrame.TextRange.Text
          except Exception:
              pass

      notes_text = ""
      try:
          notes_text = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
      except Exception:
          pass

      design_name = ""
      try:
          design_name = slide.Design.Name
      except Exception:
          pass

      return {
          "index": slide.SlideIndex,
          "slide_id": slide.SlideID,
          "slide_number": slide.SlideNumber,
          "name": slide.Name,
          "layout": slide.Layout,
          "layout_name": layout_name,
          "hidden": bool(trans.Hidden),
          "shapes_count": slide.Shapes.Count,
          "has_title": has_title,
          "title_text": title_text,
          "notes_text": notes_text,
          "follow_master_background": bool(slide.FollowMasterBackground),
          "transition_effect": trans.EntryEffect,
          "advance_on_click": bool(trans.AdvanceOnClick),
          "advance_on_time": bool(trans.AdvanceOnTime),
          "advance_time": trans.AdvanceTime,
          "design_name": design_name,
      }
  ```

---

### Tool: `set_slide_notes`

- **Description**: Set the speaker notes for a slide
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
  - `notes_text` (str, required): The notes text to set
- **Returns**: `{"success": true}`
- **COM Implementation**:
  ```python
  def set_slide_notes(app, slide_index, notes_text):
      """
      Set speaker notes.

      Notes object hierarchy:
        Slide.NotesPage (SlideRange)
          -> Shapes
              -> Placeholders(1): Slide thumbnail image
              -> Placeholders(2): Notes text placeholder
                  -> TextFrame.TextRange.Text
      """
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = notes_text
      return {"success": True}
  ```

---

### Tool: `get_slide_notes`

- **Description**: Get the speaker notes for a slide
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
- **Returns**: `{"slide_index": 1, "notes_text": "Speaker notes here"}`
- **COM Implementation**:
  ```python
  def get_slide_notes(app, slide_index):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      try:
          notes_text = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
      except Exception:
          notes_text = ""
      return {"slide_index": slide_index, "notes_text": notes_text}
  ```

---

### Tool: `set_slide_background`

- **Description**: Set the background of a specific slide
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
  - `color` (str, optional): Hex color string like "#FF0000" for solid fill
  - `rgb` (list[int], optional): [R, G, B] values (0-255 each) for solid fill
  - `picture_path` (str, optional): Path to image file for picture fill
  - `follow_master` (bool, optional, default=False): If True, revert to master background
- **Returns**: `{"success": true}`
- **COM Implementation**:
  ```python
  def set_slide_background(app, slide_index, color=None, rgb=None, picture_path=None, follow_master=False):
      """
      Set slide background.

      IMPORTANT: Must set slide.FollowMasterBackground = False before
      setting individual background. Otherwise the setting won't take effect.

      Background.Fill methods:
        .Solid() -> set ForeColor.RGB for solid color
        .UserPicture(PictureFile) -> set picture background
        .PresetGradient(Style, Variant, PresetGradientType) -> gradient
        .Patterned(Pattern) -> pattern fill
        .PresetTextured(PresetTexture) -> texture fill

      Color format: PowerPoint uses BGR integer = R + (G*256) + (B*65536)
      Use rgb_to_int(r, g, b) from utils.color.
      """
      from utils.color import rgb_to_int, hex_to_int

      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)

      if follow_master:
          slide.FollowMasterBackground = True
          return {"success": True}

      slide.FollowMasterBackground = False
      bg = slide.Background.Fill

      if picture_path:
          bg.UserPicture(PictureFile=picture_path)
      elif color:
          bg.Solid()
          bg.ForeColor.RGB = hex_to_int(color)
      elif rgb:
          bg.Solid()
          bg.ForeColor.RGB = rgb_to_int(rgb[0], rgb[1], rgb[2])

      return {"success": True}
  ```

---

### Tool: `set_slide_transition`

- **Description**: Set the transition effect for a slide
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
  - `effect` (int, optional, default=0): PpEntryEffect constant. Common values:
    - `0` = ppEffectNone
    - `3844` = ppEffectFade
    - `3845` = ppEffectPush
    - `3846` = ppEffectWipe
    - `3847` = ppEffectSplit
    - `257` = ppEffectCut
    - `513` = ppEffectRandom
  - `speed` (int, optional, default=2): PpTransitionSpeed (1=slow, 2=medium, 3=fast)
  - `duration` (float, optional): Transition duration in seconds
  - `advance_on_click` (bool, optional, default=True): Advance on mouse click
  - `advance_on_time` (bool, optional, default=False): Auto-advance after delay
  - `advance_time` (int, optional, default=0): Auto-advance delay in seconds
- **Returns**: `{"success": true}`
- **COM Implementation**:
  ```python
  def set_slide_transition(app, slide_index, effect=0, speed=2, duration=None,
                           advance_on_click=True, advance_on_time=False, advance_time=0):
      """
      Set slide transition.

      SlideShowTransition properties:
        EntryEffect (PpEntryEffect): Transition effect type
        Speed (PpTransitionSpeed): 1=slow, 2=medium, 3=fast
        Duration (Single): Duration in seconds
        AdvanceOnClick (Boolean): Advance on click
        AdvanceOnTime (Boolean): Auto-advance
        AdvanceTime (Long): Auto-advance delay in seconds
        Hidden (Boolean): Hide slide in slideshow
      """
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      trans = slide.SlideShowTransition
      trans.EntryEffect = effect
      trans.Speed = speed
      if duration is not None:
          trans.Duration = duration
      trans.AdvanceOnClick = advance_on_click
      trans.AdvanceOnTime = advance_on_time
      trans.AdvanceTime = advance_time
      return {"success": True}
  ```

---

### Tool: `set_slide_hidden`

- **Description**: Show or hide a slide in the slideshow
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
  - `hidden` (bool, required): True to hide, False to show
- **Returns**: `{"success": true}`
- **COM Implementation**:
  ```python
  def set_slide_hidden(app, slide_index, hidden):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      slide.SlideShowTransition.Hidden = hidden
      return {"success": True}
  ```

---

### Tool: `copy_slide_between_presentations`

- **Description**: Copy slides from one presentation to another
- **Parameters**:
  - `source_path` (str, required): Path to the source presentation file
  - `slide_start` (int, optional, default=1): First slide to copy (1-based)
  - `slide_end` (int, optional): Last slide to copy (defaults to slide_start)
  - `insert_position` (int, optional): Position in the active presentation to insert (defaults to end)
- **Returns**: `{"success": true, "slides_copied": 3, "insert_position": 5}`
- **COM Implementation**:
  ```python
  def copy_slide_between_presentations(app, source_path, slide_start=1, slide_end=None, insert_position=None):
      """
      Copy slides from a source file into the active presentation.

      Uses Slides.InsertFromFile which is more efficient than Copy+Paste
      because it doesn't use the clipboard.

      Slides.InsertFromFile parameters:
        FileName (str): Source file path
        Index (Long): Insert position (inserts AFTER this index)
        SlideStart (Long): First slide number to copy
        SlideEnd (Long): Last slide number to copy

      NOTE: InsertFromFile inserts after the specified Index.
      So Index=0 inserts at the beginning, Index=Slides.Count inserts at the end.
      """
      pres = app.ActivePresentation
      if slide_end is None:
          slide_end = slide_start
      if insert_position is None:
          insert_position = pres.Slides.Count

      pres.Slides.InsertFromFile(
          FileName=source_path,
          Index=insert_position,
          SlideStart=slide_start,
          SlideEnd=slide_end,
      )
      slides_copied = slide_end - slide_start + 1
      return {
          "success": True,
          "slides_copied": slides_copied,
          "insert_position": insert_position,
      }
  ```

---

### Tool: `list_layouts`

- **Description**: List all available slide layouts in the active presentation
- **Parameters**: None
- **Returns**:
  ```json
  {
    "layouts": [
      {
        "index": 1,
        "name": "Title Slide",
        "placeholders_count": 2
      }
    ]
  }
  ```
- **COM Implementation**:
  ```python
  def list_layouts(app):
      """
      List all CustomLayouts from the SlideMaster.
      Each layout has a Name, Index, and Shapes.Placeholders collection.
      """
      pres = app.ActivePresentation
      master = pres.SlideMaster
      layouts = []
      for i in range(1, master.CustomLayouts.Count + 1):
          layout = master.CustomLayouts(i)
          ph_count = 0
          try:
              ph_count = layout.Shapes.Placeholders.Count
          except Exception:
              pass
          layouts.append({
              "index": i,
              "name": layout.Name,
              "placeholders_count": ph_count,
          })
      return {"layouts": layouts}
  ```

---

### Tool: `export_slide_as_image`

- **Description**: Export a specific slide as an image file
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
  - `file_path` (str, required): Output image file path
  - `format` (str, optional, default="PNG"): Image format: "PNG", "JPG", "BMP", "TIF", "GIF", "EMF", "WMF"
  - `width` (int, optional, default=1920): Output image width in pixels
  - `height` (int, optional, default=1080): Output image height in pixels
- **Returns**: `{"success": true, "file_path": "C:\\output\\slide_1.png"}`
- **COM Implementation**:
  ```python
  def export_slide_as_image(app, slide_index, file_path, format="PNG", width=1920, height=1080):
      """
      Export a single slide as an image.

      Slide.Export(PathName, FilterName, ScaleWidth, ScaleHeight)
        PathName: Output file path
        FilterName: "PNG", "JPG", "BMP", "TIF", "GIF", "EMF", "WMF"
        ScaleWidth: Output width in pixels
        ScaleHeight: Output height in pixels

      If ScaleWidth/ScaleHeight are omitted, default resolution (96dpi) is used.
      """
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      slide.Export(file_path, format.upper(), width, height)
      return {"success": True, "file_path": file_path}
  ```

---

## Constants Needed

All constants are imported from `ppt_com.constants`. Key constants for this module:

### PpSlideLayout (for add_slide)
| Constant | Value | Description |
|----------|-------|-------------|
| ppLayoutTitle | 1 | Title slide |
| ppLayoutText | 2 | Title + text |
| ppLayoutTitleOnly | 11 | Title only |
| ppLayoutBlank | 12 | Blank |
| ppLayoutSectionHeader | 33 | Section header |
| ppLayoutComparison | 34 | Comparison |
| ppLayoutContentWithCaption | 35 | Content with caption |
| ppLayoutPictureWithCaption | 36 | Picture with caption |
| ppLayoutCustom | 32 | Custom layout |

### PpSaveAsFileType (for save_presentation_as)
| Constant | Value | Description |
|----------|-------|-------------|
| ppSaveAsPresentation | 1 | .ppt (legacy) |
| ppSaveAsOpenXMLPresentation | 24 | .pptx |
| ppSaveAsOpenXMLShow | 28 | .ppsx |
| ppSaveAsPDF | 32 | .pdf |
| ppSaveAsPNG | 18 | .png (folder of images) |
| ppSaveAsJPG | 17 | .jpg (folder of images) |
| ppSaveAsMP4 | 39 | .mp4 |
| ppSaveAsWMV | 37 | .wmv |
| ppSaveAsDefault | 11 | Default format |

### PpTransitionSpeed (for set_slide_transition)
| Constant | Value | Description |
|----------|-------|-------------|
| ppTransitionSpeedSlow | 1 | Slow |
| ppTransitionSpeedMedium | 2 | Medium |
| ppTransitionSpeedFast | 3 | Fast |

### PpEntryEffect (common, for set_slide_transition)
| Constant | Value | Description |
|----------|-------|-------------|
| ppEffectNone | 0 | No transition |
| ppEffectCut | 257 | Cut |
| ppEffectRandom | 513 | Random |
| ppEffectFade | 3844 | Fade |
| ppEffectPush | 3845 | Push |
| ppEffectWipe | 3846 | Wipe |
| ppEffectSplit | 3847 | Split |

### MsoTriState (for boolean COM parameters)
| Constant | Value |
|----------|-------|
| msoTrue | -1 |
| msoFalse | 0 |

---

## Implementation Notes

1. **1-Based Indexing**: All PowerPoint COM collections use 1-based indexing. `Slides(1)` is the first slide, `Shapes(1)` is the first shape.

2. **Slide Index vs Slide Number vs Slide ID**:
   - `SlideIndex`: Position in the Slides collection (1-based, changes when slides are reordered)
   - `SlideNumber`: Display number (affected by `PageSetup.FirstSlideNumber`)
   - `SlideID`: Unique immutable integer per slide (survives add/delete/reorder). Use `Slides.FindBySlideID(id)` to locate.

3. **SaveAs vs SaveCopyAs**: `SaveAs` changes the presentation's `FullName` to the new path. `SaveCopyAs` preserves the original name but does not support format conversion.

4. **Image Export via SaveAs**: When using `SaveAs` with an image format (PNG, JPG, etc.), PowerPoint creates a **folder** at the specified path and saves individual slide images inside it. The folder name is taken from the path.

5. **InsertFromFile vs Copy+Paste**: `InsertFromFile` is preferred for copying slides between presentations because it does not use the clipboard and is more reliable for automation.

6. **FollowMasterBackground**: When setting individual slide backgrounds, you MUST set `slide.FollowMasterBackground = False` first. Otherwise the custom background will not be visible.

7. **Notes Access**: Notes text is at `slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text`. Placeholders(1) is the slide thumbnail, Placeholders(2) is the notes text.

8. **SlideShowTransition.Hidden**: This property controls whether a slide is hidden in the slideshow. It is on the `SlideShowTransition` object, not directly on the `Slide` object.

9. **Color Format**: PowerPoint COM uses BGR integer format: `R + (G * 256) + (B * 65536)`. Always use `utils.color.rgb_to_int()` or `utils.color.hex_to_int()` for conversions. Never pass standard `0xRRGGBB` hex values directly.

10. **Points as Unit**: All position and size values are in points (72 points = 1 inch). Use `utils.units` for conversions to/from inches and centimeters.

11. **COM Error Recovery**: Wrap all COM calls with try/except for `pywintypes.com_error`. On disconnection errors, clear the app reference and reconnect through `PowerPointCOMWrapper.get_app()`.
