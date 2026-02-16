# Module: Shape Operations

## Overview

This module handles all operations related to shapes on PowerPoint slides: adding shapes (rectangles, ovals, textboxes, lines, connectors, pictures, freeforms, callouts, word art), listing shapes, modifying shape properties (position, size, rotation, visibility, z-order, name, flip), grouping/ungrouping, duplicating, deleting, and exporting shapes as images. It also handles shape-level fill, line, shadow, and effect formatting.

## Dependencies

- **Internal**: `utils.com_wrapper` (PowerPointCOMWrapper, safe_com_call), `utils.units` (inches_to_points, cm_to_points), `utils.color` (rgb_to_int, int_to_rgb, int_to_hex, hex_to_int), `ppt_com.constants` (all shape/fill/line/effect constants)
- **External**: `pywin32` (`win32com.client`, `pywintypes`)
- **Standard library**: `os`, `logging`

### Importing from Core Module

```python
from utils.com_wrapper import PowerPointCOMWrapper, safe_com_call
from utils.units import inches_to_points, points_to_inches, cm_to_points, points_to_cm
from utils.color import rgb_to_int, int_to_rgb, int_to_hex, hex_to_int
from ppt_com.constants import (
    # Shape types
    msoAutoShape, msoCallout, msoChart, msoFreeform, msoGroup,
    msoLine, msoPicture, msoPlaceholder, msoTextEffect, msoMedia,
    msoTextBox, msoTable, msoSmartArt, msoEmbeddedOLEObject,
    # AutoShape types
    msoShapeRectangle, msoShapeRoundedRectangle, msoShapeOval,
    msoShapeDiamond, msoShapeIsoscelesTriangle, msoShapeRightArrow,
    msoShapeLeftArrow, msoShapeUpArrow, msoShapeDownArrow,
    msoShape5pointStar, msoShapeHeart, msoShapeCross,
    msoShapeFlowchartProcess, msoShapeFlowchartDecision,
    msoShapeFlowchartTerminator, msoShapeCloud,
    # Text orientation
    msoTextOrientationHorizontal, msoTextOrientationVertical,
    # Z-Order
    msoBringToFront, msoSendToBack, msoBringForward, msoSendBackward,
    # Flip
    msoFlipHorizontal, msoFlipVertical,
    # Connector types
    msoConnectorStraight, msoConnectorElbow, msoConnectorCurve,
    # Fill types
    msoFillSolid, msoFillGradient, msoFillPatterned, msoFillPicture,
    # Gradient styles
    msoGradientHorizontal, msoGradientVertical, msoGradientDiagonalUp,
    msoGradientDiagonalDown, msoGradientFromCorner, msoGradientFromCenter,
    # Line styles
    msoLineSolid, msoLineDash, msoLineDot, msoLineDashDot,
    msoLineSingle, msoLineThinThin, msoLineThinThick,
    # Arrow styles
    msoArrowheadNone, msoArrowheadTriangle, msoArrowheadOpen,
    msoArrowheadStealth, msoArrowheadDiamond, msoArrowheadOval,
    msoArrowheadShort, msoArrowheadLengthMedium, msoArrowheadLong,
    msoArrowheadNarrow, msoArrowheadWidthMedium, msoArrowheadWide,
    # Callout
    msoCalloutOne, msoCalloutTwo, msoCalloutThree, msoCalloutFour,
    # Freeform
    msoSegmentLine, msoSegmentCurve, msoEditingAuto, msoEditingCorner,
    # Effects
    msoShadowStyleOuterShadow, msoShadowStyleInnerShadow,
    msoSoftEdgeTypeNone, msoReflectionTypeNone,
    msoBevelNone, msoBevelCircle, msoBevelRelaxedInset,
    msoMaterialMatte, msoMaterialPlastic, msoMaterialMetal,
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
    shapes.py      # All shape operations
```

---

## MCP Tools

### Tool: `list_shapes`

- **Description**: List all shapes on a specific slide
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
- **Returns**:
  ```json
  {
    "slide_index": 1,
    "shapes_count": 5,
    "shapes": [
      {
        "index": 1,
        "name": "Title 1",
        "type": 14,
        "type_name": "msoPlaceholder",
        "left": 50.0,
        "top": 30.0,
        "width": 620.0,
        "height": 80.0,
        "rotation": 0.0,
        "visible": true,
        "has_text_frame": true,
        "text_preview": "Slide Title",
        "z_order_position": 1
      }
    ]
  }
  ```
- **COM Implementation**:
  ```python
  # MsoShapeType name map for readable output
  SHAPE_TYPE_NAMES = {
      1: "msoAutoShape", 2: "msoCallout", 3: "msoChart", 4: "msoComment",
      5: "msoFreeform", 6: "msoGroup", 7: "msoEmbeddedOLEObject",
      8: "msoFormControl", 9: "msoLine", 10: "msoLinkedOLEObject",
      11: "msoLinkedPicture", 12: "msoOLEControlObject", 13: "msoPicture",
      14: "msoPlaceholder", 15: "msoTextEffect", 16: "msoMedia",
      17: "msoTextBox", 19: "msoTable", 24: "msoSmartArt",
      -2: "msoShapeTypeMixed",
  }

  def list_shapes(app, slide_index):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shapes = []
      for i in range(1, slide.Shapes.Count + 1):
          shape = slide.Shapes(i)
          text_preview = ""
          has_text = False
          try:
              if shape.HasTextFrame:
                  has_text = True
                  if shape.TextFrame.HasText:
                      full_text = shape.TextFrame.TextRange.Text
                      text_preview = full_text[:100] + ("..." if len(full_text) > 100 else "")
          except Exception:
              pass

          shapes.append({
              "index": i,
              "name": shape.Name,
              "type": shape.Type,
              "type_name": SHAPE_TYPE_NAMES.get(shape.Type, f"unknown({shape.Type})"),
              "left": round(shape.Left, 2),
              "top": round(shape.Top, 2),
              "width": round(shape.Width, 2),
              "height": round(shape.Height, 2),
              "rotation": round(shape.Rotation, 2),
              "visible": bool(shape.Visible),
              "has_text_frame": has_text,
              "text_preview": text_preview,
              "z_order_position": shape.ZOrderPosition,
          })
      return {"slide_index": slide_index, "shapes_count": slide.Shapes.Count, "shapes": shapes}
  ```

---

### Tool: `add_shape`

- **Description**: Add an auto shape to a slide
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
  - `shape_type` (int, required): MsoAutoShapeType constant. Common values:
    - `1` = msoShapeRectangle
    - `5` = msoShapeRoundedRectangle
    - `9` = msoShapeOval
    - `4` = msoShapeDiamond
    - `7` = msoShapeIsoscelesTriangle
    - `33` = msoShapeRightArrow
    - `92` = msoShape5pointStar
    - `61` = msoShapeFlowchartProcess
    - `63` = msoShapeFlowchartDecision
    - `179` = msoShapeCloud
  - `left` (float, required): Left position in points
  - `top` (float, required): Top position in points
  - `width` (float, required): Width in points
  - `height` (float, required): Height in points
  - `name` (str, optional): Custom name for the shape
- **Returns**:
  ```json
  {
    "success": true,
    "shape_index": 3,
    "shape_name": "Rectangle 1",
    "shape_type": 1
  }
  ```
- **COM Implementation**:
  ```python
  def add_shape(app, slide_index, shape_type, left, top, width, height, name=None):
      """
      Shapes.AddShape(Type, Left, Top, Width, Height)
        Type: MsoAutoShapeType integer
        Left, Top, Width, Height: in points (72 points = 1 inch)
      Returns: Shape object
      """
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = slide.Shapes.AddShape(
          Type=shape_type, Left=left, Top=top, Width=width, Height=height
      )
      if name:
          shape.Name = name
      return {
          "success": True,
          "shape_index": shape.ZOrderPosition,
          "shape_name": shape.Name,
          "shape_type": shape.AutoShapeType,
      }
  ```

---

### Tool: `add_textbox`

- **Description**: Add a text box to a slide
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
  - `left` (float, required): Left position in points
  - `top` (float, required): Top position in points
  - `width` (float, required): Width in points
  - `height` (float, required): Height in points
  - `text` (str, optional): Initial text content
  - `orientation` (int, optional, default=1): MsoTextOrientation (1=horizontal, 5=vertical, 2=upward, 3=downward, 6=verticalFarEast)
  - `name` (str, optional): Custom name for the textbox
- **Returns**: `{"success": true, "shape_index": 4, "shape_name": "TextBox 1"}`
- **COM Implementation**:
  ```python
  def add_textbox(app, slide_index, left, top, width, height, text=None, orientation=1, name=None):
      """
      Shapes.AddTextbox(Orientation, Left, Top, Width, Height)
        Orientation: MsoTextOrientation integer
          1 = msoTextOrientationHorizontal
          5 = msoTextOrientationVertical
          2 = msoTextOrientationUpward
          3 = msoTextOrientationDownward
          6 = msoTextOrientationVerticalFarEast
      """
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      textbox = slide.Shapes.AddTextbox(
          Orientation=orientation, Left=left, Top=top, Width=width, Height=height
      )
      if text:
          textbox.TextFrame.TextRange.Text = text
      if name:
          textbox.Name = name
      return {
          "success": True,
          "shape_index": textbox.ZOrderPosition,
          "shape_name": textbox.Name,
      }
  ```

---

### Tool: `add_picture`

- **Description**: Add an image to a slide
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
  - `file_path` (str, required): Path to image file
  - `left` (float, required): Left position in points
  - `top` (float, required): Top position in points
  - `width` (float, optional, default=-1): Width in points. -1 for original size.
  - `height` (float, optional, default=-1): Height in points. -1 for original size.
  - `link_to_file` (bool, optional, default=False): If True, link to file instead of embedding
  - `name` (str, optional): Custom name
- **Returns**: `{"success": true, "shape_name": "Picture 1", "width": 300.0, "height": 200.0}`
- **COM Implementation**:
  ```python
  def add_picture(app, slide_index, file_path, left, top, width=-1, height=-1,
                  link_to_file=False, name=None):
      """
      Shapes.AddPicture(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
        FileName: Image file path
        LinkToFile: MsoTriState - link to external file
        SaveWithDocument: MsoTriState - embed in document
        Width/Height: -1 for original size

      RULES:
        LinkToFile=False, SaveWithDocument=True -> embedded (recommended)
        LinkToFile=True, SaveWithDocument=True -> linked + embedded backup
        LinkToFile=True, SaveWithDocument=False -> linked only (fragile)
        LinkToFile=False, SaveWithDocument=False -> ERROR
      """
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      save_with = not link_to_file  # If linked, don't embed; otherwise embed
      pic = slide.Shapes.AddPicture(
          FileName=file_path,
          LinkToFile=msoTrue if link_to_file else msoFalse,
          SaveWithDocument=msoTrue if save_with or not link_to_file else msoFalse,
          Left=left, Top=top, Width=width, Height=height,
      )
      if name:
          pic.Name = name
      return {
          "success": True,
          "shape_name": pic.Name,
          "width": round(pic.Width, 2),
          "height": round(pic.Height, 2),
      }
  ```

---

### Tool: `add_line`

- **Description**: Add a line shape to a slide
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
  - `begin_x` (float, required): Start X position in points
  - `begin_y` (float, required): Start Y position in points
  - `end_x` (float, required): End X position in points
  - `end_y` (float, required): End Y position in points
  - `weight` (float, optional, default=1.0): Line weight in points
  - `color` (str, optional): Hex color like "#000000"
  - `dash_style` (int, optional, default=1): MsoLineDashStyle (1=solid, 4=dash, 3=dot, 5=dashDot)
  - `begin_arrow` (int, optional, default=1): MsoArrowheadStyle for start (1=none, 2=triangle, 3=open)
  - `end_arrow` (int, optional, default=1): MsoArrowheadStyle for end
  - `name` (str, optional): Custom name
- **Returns**: `{"success": true, "shape_name": "Line 1"}`
- **COM Implementation**:
  ```python
  def add_line(app, slide_index, begin_x, begin_y, end_x, end_y,
               weight=1.0, color=None, dash_style=1, begin_arrow=1, end_arrow=1, name=None):
      """
      Shapes.AddLine(BeginX, BeginY, EndX, EndY)
      Then set Line properties: Weight, ForeColor.RGB, DashStyle,
      BeginArrowheadStyle, EndArrowheadStyle
      """
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      line = slide.Shapes.AddLine(
          BeginX=begin_x, BeginY=begin_y, EndX=end_x, EndY=end_y
      )
      line.Line.Weight = weight
      line.Line.DashStyle = dash_style
      line.Line.BeginArrowheadStyle = begin_arrow
      line.Line.EndArrowheadStyle = end_arrow
      if color:
          line.Line.ForeColor.RGB = hex_to_int(color)
      if name:
          line.Name = name
      return {"success": True, "shape_name": line.Name}
  ```

---

### Tool: `add_connector`

- **Description**: Add a connector between two shapes
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
  - `connector_type` (int, optional, default=2): MsoConnectorType (1=straight, 2=elbow, 3=curve)
  - `begin_shape_name` (str, required): Name of the starting shape
  - `begin_site` (int, optional, default=1): Connection site on the begin shape (1-based)
  - `end_shape_name` (str, required): Name of the ending shape
  - `end_site` (int, optional, default=3): Connection site on the end shape (1-based)
- **Returns**: `{"success": true, "connector_name": "Connector 1"}`
- **COM Implementation**:
  ```python
  def add_connector(app, slide_index, begin_shape_name, end_shape_name,
                    connector_type=2, begin_site=1, end_site=3):
      """
      Shapes.AddConnector(Type, BeginX, BeginY, EndX, EndY)
      Then: connector.ConnectorFormat.BeginConnect(ConnectedShape, ConnectionSite)
            connector.ConnectorFormat.EndConnect(ConnectedShape, ConnectionSite)
            connector.RerouteConnections()

      ConnectionSite is 1-based. Check Shape.ConnectionSiteCount for valid range.
      RerouteConnections() optimizes the connector path.

      MsoConnectorType: 1=straight, 2=elbow, 3=curve
      """
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      begin_shape = slide.Shapes(begin_shape_name)
      end_shape = slide.Shapes(end_shape_name)

      connector = slide.Shapes.AddConnector(
          Type=connector_type, BeginX=0, BeginY=0, EndX=100, EndY=100
      )
      connector.ConnectorFormat.BeginConnect(ConnectedShape=begin_shape, ConnectionSite=begin_site)
      connector.ConnectorFormat.EndConnect(ConnectedShape=end_shape, ConnectionSite=end_site)
      connector.RerouteConnections()
      return {"success": True, "connector_name": connector.Name}
  ```

---

### Tool: `modify_shape`

- **Description**: Modify properties of an existing shape
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
  - `shape_name` (str, required): Name of the shape to modify (or shape_index)
  - `left` (float, optional): New left position in points
  - `top` (float, optional): New top position in points
  - `width` (float, optional): New width in points
  - `height` (float, optional): New height in points
  - `rotation` (float, optional): Rotation in degrees (clockwise, 0-360)
  - `name` (str, optional): New name for the shape
  - `visible` (bool, optional): Show/hide the shape
  - `lock_aspect_ratio` (bool, optional): Lock aspect ratio
- **Returns**: `{"success": true, "shape_name": "MyShape"}`
- **COM Implementation**:
  ```python
  def modify_shape(app, slide_index, shape_name, left=None, top=None, width=None, height=None,
                   rotation=None, name=None, visible=None, lock_aspect_ratio=None):
      """
      Shape properties (all in points, 72 pt = 1 inch):
        Left, Top: Position from slide top-left corner
        Width, Height: Size
        Rotation: Clockwise degrees (0-360). Negative values auto-convert (e.g., -45 -> 315)
        Name: Shape name (should be unique per slide)
        Visible: MsoTriState (-1=visible, 0=hidden)
        LockAspectRatio: MsoTriState - when locked, changing Width auto-adjusts Height

      Incremental movement: shape.IncrementLeft(delta), shape.IncrementTop(delta)
      Incremental rotation: shape.IncrementRotation(degrees)
      """
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = slide.Shapes(shape_name)
      if left is not None:
          shape.Left = left
      if top is not None:
          shape.Top = top
      if width is not None:
          shape.Width = width
      if height is not None:
          shape.Height = height
      if rotation is not None:
          shape.Rotation = rotation
      if visible is not None:
          shape.Visible = msoTrue if visible else msoFalse
      if lock_aspect_ratio is not None:
          shape.LockAspectRatio = msoTrue if lock_aspect_ratio else msoFalse
      if name is not None:
          shape.Name = name
      return {"success": True, "shape_name": shape.Name}
  ```

---

### Tool: `set_shape_z_order`

- **Description**: Change the z-order (stacking) of a shape
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
  - `shape_name` (str, required): Name of the shape
  - `z_order_cmd` (int, required): MsoZOrderCmd constant:
    - `0` = msoBringToFront
    - `1` = msoSendToBack
    - `2` = msoBringForward (one level up)
    - `3` = msoSendBackward (one level down)
- **Returns**: `{"success": true, "new_z_position": 5}`
- **COM Implementation**:
  ```python
  def set_shape_z_order(app, slide_index, shape_name, z_order_cmd):
      """
      shape.ZOrder(ZOrderCmd)
        msoBringToFront=0, msoSendToBack=1,
        msoBringForward=2, msoSendBackward=3

      shape.ZOrderPosition: Read-only, 1-based. Higher = more in front.
      NOTE: msoBringInFrontOfText(4) and msoSendBehindText(5) are Word-only.
      """
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = slide.Shapes(shape_name)
      shape.ZOrder(z_order_cmd)
      return {"success": True, "new_z_position": shape.ZOrderPosition}
  ```

---

### Tool: `flip_shape`

- **Description**: Flip a shape horizontally or vertically
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
  - `shape_name` (str, required): Name of the shape
  - `direction` (str, required): "horizontal" or "vertical"
- **Returns**: `{"success": true, "horizontal_flip": true, "vertical_flip": false}`
- **COM Implementation**:
  ```python
  def flip_shape(app, slide_index, shape_name, direction):
      """
      shape.Flip(FlipCmd)
        msoFlipHorizontal=0, msoFlipVertical=1

      Read-only properties:
        shape.HorizontalFlip (MsoTriState)
        shape.VerticalFlip (MsoTriState)
      """
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = slide.Shapes(shape_name)
      if direction == "horizontal":
          shape.Flip(0)  # msoFlipHorizontal
      else:
          shape.Flip(1)  # msoFlipVertical
      return {
          "success": True,
          "horizontal_flip": bool(shape.HorizontalFlip),
          "vertical_flip": bool(shape.VerticalFlip),
      }
  ```

---

### Tool: `delete_shape`

- **Description**: Delete a shape from a slide
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
  - `shape_name` (str, required): Name of the shape to delete
- **Returns**: `{"success": true, "deleted": "MyShape"}`
- **COM Implementation**:
  ```python
  def delete_shape(app, slide_index, shape_name):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = slide.Shapes(shape_name)
      deleted_name = shape.Name
      shape.Delete()
      return {"success": True, "deleted": deleted_name}
  ```

---

### Tool: `duplicate_shape`

- **Description**: Duplicate a shape on the same slide
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
  - `shape_name` (str, required): Name of the shape to duplicate
  - `offset_left` (float, optional, default=20): Horizontal offset for the duplicate
  - `offset_top` (float, optional, default=20): Vertical offset for the duplicate
- **Returns**: `{"success": true, "new_shape_name": "Rectangle 2"}`
- **COM Implementation**:
  ```python
  def duplicate_shape(app, slide_index, shape_name, offset_left=20, offset_top=20):
      """
      shape.Duplicate() returns a ShapeRange.
      The duplicate is placed near the original with a slight offset.
      """
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = slide.Shapes(shape_name)
      dup = shape.Duplicate()
      new_shape = dup(1)
      new_shape.Left = shape.Left + offset_left
      new_shape.Top = shape.Top + offset_top
      return {"success": True, "new_shape_name": new_shape.Name}
  ```

---

### Tool: `group_shapes`

- **Description**: Group multiple shapes together
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
  - `shape_names` (list[str], required): Names of the shapes to group (minimum 2)
- **Returns**: `{"success": true, "group_name": "Group 1", "items_count": 3}`
- **COM Implementation**:
  ```python
  def group_shapes(app, slide_index, shape_names):
      """
      slide.Shapes.Range(names_array).Group()
      Requires at least 2 shapes.
      Returns a Shape object (the group).

      Group access: group.GroupItems(i) to access individual shapes.
      group.GroupItems.Count for the number of grouped shapes.
      Properties of individual shapes can be changed through GroupItems.
      """
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      if len(shape_names) < 2:
          raise ValueError("At least 2 shapes are required for grouping")
      shape_range = slide.Shapes.Range(shape_names)
      group = shape_range.Group()
      return {
          "success": True,
          "group_name": group.Name,
          "items_count": group.GroupItems.Count,
      }
  ```

---

### Tool: `ungroup_shapes`

- **Description**: Ungroup a grouped shape
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
  - `group_name` (str, required): Name of the group shape
- **Returns**: `{"success": true, "ungrouped_shapes": ["Rect1", "Oval1"]}`
- **COM Implementation**:
  ```python
  def ungroup_shapes(app, slide_index, group_name):
      """
      group.Ungroup() returns a ShapeRange of the ungrouped shapes.
      Nested groups are recursively ungrouped.
      Use ShapeRange.Regroup() to re-group previously ungrouped shapes.
      """
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      group = slide.Shapes(group_name)
      ungrouped = group.Ungroup()
      names = []
      for i in range(1, ungrouped.Count + 1):
          names.append(ungrouped(i).Name)
      return {"success": True, "ungrouped_shapes": names}
  ```

---

### Tool: `set_shape_fill`

- **Description**: Set the fill (background) of a shape
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
  - `shape_name` (str, required): Name of the shape
  - `fill_type` (str, required): "solid", "gradient", "picture", "no_fill"
  - `color` (str, optional): Hex color for solid fill (e.g., "#FF0000")
  - `rgb` (list[int], optional): [R, G, B] for solid fill
  - `transparency` (float, optional, default=0): Transparency 0.0-1.0
  - `gradient_style` (int, optional, default=1): MsoGradientStyle for gradient
  - `gradient_variant` (int, optional, default=1): Gradient variant (1-4)
  - `gradient_color2` (str, optional): Second hex color for 2-color gradient
  - `picture_path` (str, optional): Image file path for picture fill
  - `theme_color` (str, optional): Theme color name (e.g., "accent1")
- **Returns**: `{"success": true}`
- **COM Implementation**:
  ```python
  def set_shape_fill(app, slide_index, shape_name, fill_type,
                     color=None, rgb=None, transparency=0,
                     gradient_style=1, gradient_variant=1, gradient_color2=None,
                     picture_path=None, theme_color=None):
      """
      FillFormat methods:
        .Solid() -> then set ForeColor.RGB
        .TwoColorGradient(Style, Variant) -> set ForeColor.RGB and BackColor.RGB
        .OneColorGradient(Style, Variant, Degree)
        .PresetGradient(Style, Variant, PresetGradientType)
        .UserPicture(PictureFile)
        .Visible = msoFalse for no fill

      FillFormat properties:
        .ForeColor.RGB: Primary color (BGR integer)
        .BackColor.RGB: Secondary color (for gradients/patterns)
        .Transparency: 0.0 (opaque) to 1.0 (transparent)
        .Visible: MsoTriState
        .ForeColor.ObjectThemeColor: Theme color index
        .ForeColor.Brightness: -1.0 to 1.0

      MsoGradientStyle: 1=horizontal, 2=vertical, 3=diagonalUp,
        4=diagonalDown, 5=fromCorner, 7=fromCenter
      """
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = slide.Shapes(shape_name)
      fill = shape.Fill

      if fill_type == "no_fill":
          fill.Visible = msoFalse
      elif fill_type == "solid":
          fill.Solid()
          if theme_color:
              from utils.color import get_theme_color_index
              fill.ForeColor.ObjectThemeColor = get_theme_color_index(theme_color)
          elif color:
              fill.ForeColor.RGB = hex_to_int(color)
          elif rgb:
              fill.ForeColor.RGB = rgb_to_int(rgb[0], rgb[1], rgb[2])
          fill.Transparency = transparency
      elif fill_type == "gradient":
          if gradient_color2:
              fill.TwoColorGradient(gradient_style, gradient_variant)
              if color:
                  fill.ForeColor.RGB = hex_to_int(color)
              fill.BackColor.RGB = hex_to_int(gradient_color2)
          else:
              fill.OneColorGradient(gradient_style, gradient_variant, 0.5)
              if color:
                  fill.ForeColor.RGB = hex_to_int(color)
          fill.Transparency = transparency
      elif fill_type == "picture":
          if picture_path:
              fill.UserPicture(picture_path)

      return {"success": True}
  ```

---

### Tool: `set_shape_line`

- **Description**: Set the line (border) formatting of a shape
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
  - `shape_name` (str, required): Name of the shape
  - `visible` (bool, optional, default=True): Show/hide the line
  - `color` (str, optional): Hex color like "#000000"
  - `weight` (float, optional): Line weight in points
  - `dash_style` (int, optional): MsoLineDashStyle (1=solid, 2=roundDot, 3=dot, 4=dash, 5=dashDot, 6=dashDotDot, 7=longDash, 8=longDashDot)
  - `style` (int, optional): MsoLineStyle (1=single, 2=thinThin, 3=thinThick, 4=thickThin, 5=thickBetweenThin)
  - `transparency` (float, optional): Transparency 0.0-1.0
- **Returns**: `{"success": true}`
- **COM Implementation**:
  ```python
  def set_shape_line(app, slide_index, shape_name, visible=True,
                     color=None, weight=None, dash_style=None,
                     style=None, transparency=None):
      """
      LineFormat properties:
        .Visible (MsoTriState)
        .ForeColor.RGB (BGR integer)
        .Weight (Single, points)
        .DashStyle (MsoLineDashStyle)
        .Style (MsoLineStyle)
        .Transparency (0.0-1.0)
        .BeginArrowheadStyle, .EndArrowheadStyle (MsoArrowheadStyle)
        .BeginArrowheadLength, .EndArrowheadLength (MsoArrowheadLength)
        .BeginArrowheadWidth, .EndArrowheadWidth (MsoArrowheadWidth)
      """
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = slide.Shapes(shape_name)
      line = shape.Line
      line.Visible = msoTrue if visible else msoFalse
      if color:
          line.ForeColor.RGB = hex_to_int(color)
      if weight is not None:
          line.Weight = weight
      if dash_style is not None:
          line.DashStyle = dash_style
      if style is not None:
          line.Style = style
      if transparency is not None:
          line.Transparency = transparency
      return {"success": True}
  ```

---

### Tool: `set_shape_shadow`

- **Description**: Set shadow effect on a shape
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
  - `shape_name` (str, required): Name of the shape
  - `visible` (bool, optional, default=True): Enable/disable shadow
  - `offset_x` (float, optional, default=5): Horizontal offset in points
  - `offset_y` (float, optional, default=5): Vertical offset in points
  - `blur` (float, optional, default=8): Blur radius in points
  - `color` (str, optional): Shadow color as hex
  - `transparency` (float, optional, default=0.5): Shadow transparency 0.0-1.0
  - `style` (int, optional, default=2): MsoShadowStyle (1=inner, 2=outer)
- **Returns**: `{"success": true}`
- **COM Implementation**:
  ```python
  def set_shape_shadow(app, slide_index, shape_name, visible=True,
                       offset_x=5, offset_y=5, blur=8, color=None,
                       transparency=0.5, style=2):
      """
      ShadowFormat properties:
        .Visible, .OffsetX, .OffsetY, .Blur, .Transparency, .Style
        .ForeColor.RGB, .Size, .RotateWithShape
      MsoShadowStyle: 1=inner, 2=outer
      """
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = slide.Shapes(shape_name)
      shadow = shape.Shadow
      shadow.Visible = msoTrue if visible else msoFalse
      shadow.OffsetX = offset_x
      shadow.OffsetY = offset_y
      shadow.Blur = blur
      shadow.Transparency = transparency
      shadow.Style = style
      if color:
          shadow.ForeColor.RGB = hex_to_int(color)
      return {"success": True}
  ```

---

### Tool: `set_shape_effects`

- **Description**: Set glow, reflection, soft edge, and 3D effects on a shape
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
  - `shape_name` (str, required): Name of the shape
  - `glow_radius` (float, optional): Glow radius in points (0 to remove)
  - `glow_color` (str, optional): Glow hex color
  - `glow_transparency` (float, optional): Glow transparency 0.0-1.0
  - `reflection_type` (int, optional): MsoReflectionType (0=none, 1-9=presets)
  - `soft_edge_type` (int, optional): MsoSoftEdgeType (0=none, 1-6=presets)
  - `bevel_top_type` (int, optional): MsoBevelType for top bevel (1=none, 3=circle, etc.)
  - `bevel_top_depth` (float, optional): Bevel depth in points
  - `three_d_rotation_x` (float, optional): X rotation (-90 to 90)
  - `three_d_rotation_y` (float, optional): Y rotation (-90 to 90)
  - `three_d_material` (int, optional): MsoPresetMaterial (1=matte, 2=plastic, 3=metal)
- **Returns**: `{"success": true}`
- **COM Implementation**:
  ```python
  def set_shape_effects(app, slide_index, shape_name,
                        glow_radius=None, glow_color=None, glow_transparency=None,
                        reflection_type=None, soft_edge_type=None,
                        bevel_top_type=None, bevel_top_depth=None,
                        three_d_rotation_x=None, three_d_rotation_y=None,
                        three_d_material=None):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = slide.Shapes(shape_name)

      # Glow: shape.Glow.Color.RGB, .Radius, .Transparency
      if glow_radius is not None:
          shape.Glow.Radius = glow_radius
      if glow_color:
          shape.Glow.Color.RGB = hex_to_int(glow_color)
      if glow_transparency is not None:
          shape.Glow.Transparency = glow_transparency

      # Reflection: shape.Reflection.Type (MsoReflectionType 0-9)
      if reflection_type is not None:
          shape.Reflection.Type = reflection_type

      # Soft Edge: shape.SoftEdge.Type (MsoSoftEdgeType 0-6)
      if soft_edge_type is not None:
          shape.SoftEdge.Type = soft_edge_type

      # 3D (ThreeDFormat): shape.ThreeD.*
      if bevel_top_type is not None:
          shape.ThreeD.BevelTopType = bevel_top_type
      if bevel_top_depth is not None:
          shape.ThreeD.BevelTopDepth = bevel_top_depth
      if three_d_rotation_x is not None:
          shape.ThreeD.RotationX = three_d_rotation_x
      if three_d_rotation_y is not None:
          shape.ThreeD.RotationY = three_d_rotation_y
      if three_d_material is not None:
          shape.ThreeD.PresetMaterial = three_d_material

      return {"success": True}
  ```

---

### Tool: `copy_shape_format`

- **Description**: Copy formatting from one shape and apply to another
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
  - `source_shape_name` (str, required): Name of the source shape
  - `target_shape_name` (str, required): Name of the target shape
- **Returns**: `{"success": true}`
- **COM Implementation**:
  ```python
  def copy_shape_format(app, slide_index, source_shape_name, target_shape_name):
      """
      shape.PickUp() copies formatting (fill, line, effects) but NOT text formatting.
      shape.Apply() applies the picked-up formatting to another shape.
      """
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      slide.Shapes(source_shape_name).PickUp()
      slide.Shapes(target_shape_name).Apply()
      return {"success": True}
  ```

---

### Tool: `export_shape_as_image`

- **Description**: Export a specific shape as an image
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
  - `shape_name` (str, required): Name of the shape
  - `file_path` (str, required): Output file path
  - `format` (int, optional, default=2): ppShapeFormat (1=JPG, 2=PNG, 3=BMP)
- **Returns**: `{"success": true, "file_path": "C:\\output\\shape.png"}`
- **COM Implementation**:
  ```python
  def export_shape_as_image(app, slide_index, shape_name, file_path, format=2):
      """
      shape.Export(PathName, Filter)
        Filter: ppShapeFormatJPG=1, ppShapeFormatPNG=2, ppShapeFormatBMP=3
      """
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = slide.Shapes(shape_name)
      shape.Export(PathName=file_path, Filter=format)
      return {"success": True, "file_path": file_path}
  ```

---

### Tool: `manage_shape_tags`

- **Description**: Add, get, or delete custom tags on a shape
- **Parameters**:
  - `slide_index` (int, required): 1-based slide index
  - `shape_name` (str, required): Name of the shape
  - `action` (str, required): "add", "get", "delete", or "list"
  - `tag_name` (str, optional): Tag name (required for add/get/delete)
  - `tag_value` (str, optional): Tag value (required for add)
- **Returns**: Varies by action
- **COM Implementation**:
  ```python
  def manage_shape_tags(app, slide_index, shape_name, action, tag_name=None, tag_value=None):
      """
      Tags are key-value string pairs stored as hidden metadata on shapes.
      Tags.Add(Name, Value): Add or update a tag
      Tags(Name): Get tag value by name
      Tags.Delete(Name): Delete a tag
      Tags.Name(i), Tags.Value(i): Access by 1-based index
      Tags.Count: Number of tags

      NOTE: Tag names are automatically converted to UPPERCASE internally.
      Tag values are strings only (convert numbers to strings before storing).
      """
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = slide.Shapes(shape_name)
      tags = shape.Tags

      if action == "add":
          tags.Add(tag_name, tag_value)
          return {"success": True, "tag": tag_name.upper(), "value": tag_value}
      elif action == "get":
          value = tags(tag_name)
          return {"tag": tag_name.upper(), "value": value}
      elif action == "delete":
          tags.Delete(tag_name)
          return {"success": True, "deleted": tag_name.upper()}
      elif action == "list":
          all_tags = {}
          for i in range(1, tags.Count + 1):
              all_tags[tags.Name(i)] = tags.Value(i)
          return {"tags": all_tags, "count": tags.Count}
  ```

---

## Constants Needed

All constants imported from `ppt_com.constants`. Key constants for this module:

### MsoAutoShapeType (commonly used)
| Constant | Value | Description |
|----------|-------|-------------|
| msoShapeRectangle | 1 | Rectangle |
| msoShapeRoundedRectangle | 5 | Rounded rectangle |
| msoShapeOval | 9 | Oval/circle |
| msoShapeDiamond | 4 | Diamond |
| msoShapeIsoscelesTriangle | 7 | Triangle |
| msoShapeRightArrow | 33 | Right arrow |
| msoShapeLeftArrow | 34 | Left arrow |
| msoShapeUpArrow | 35 | Up arrow |
| msoShapeDownArrow | 36 | Down arrow |
| msoShape5pointStar | 92 | 5-point star |
| msoShapeHeart | 21 | Heart |
| msoShapeFlowchartProcess | 61 | Flowchart process |
| msoShapeFlowchartDecision | 63 | Flowchart decision |
| msoShapeFlowchartTerminator | 69 | Flowchart terminator |
| msoShapeCloud | 179 | Cloud |

### MsoShapeType (for type identification)
| Constant | Value | Description |
|----------|-------|-------------|
| msoAutoShape | 1 | Auto shape |
| msoCallout | 2 | Callout |
| msoChart | 3 | Chart |
| msoFreeform | 5 | Freeform |
| msoGroup | 6 | Group |
| msoLine | 9 | Line |
| msoPicture | 13 | Picture |
| msoPlaceholder | 14 | Placeholder |
| msoTextBox | 17 | Text box |
| msoTable | 19 | Table |
| msoMedia | 16 | Media |
| msoSmartArt | 24 | SmartArt |

### MsoZOrderCmd
| Constant | Value | Description |
|----------|-------|-------------|
| msoBringToFront | 0 | Bring to front |
| msoSendToBack | 1 | Send to back |
| msoBringForward | 2 | Bring forward one level |
| msoSendBackward | 3 | Send backward one level |

### MsoFlipCmd
| Constant | Value |
|----------|-------|
| msoFlipHorizontal | 0 |
| msoFlipVertical | 1 |

---

## Implementation Notes

1. **Shape Access**: Shapes can be accessed by 1-based index `slide.Shapes(1)` or by name `slide.Shapes("MyShape")`. Name-based access is preferred for MCP tools.

2. **Default Names**: PowerPoint auto-generates names like "Rectangle 1", "TextBox 2". These names are not guaranteed unique if shapes are renamed.

3. **HasTextFrame / HasTable / HasChart**: Always check these boolean properties before accessing TextFrame, Table, or Chart to avoid COM errors.

4. **Type Identification**: Use `shape.Type` (MsoShapeType) to identify the kind of shape. For auto shapes, also check `shape.AutoShapeType` (MsoAutoShapeType).

5. **Points Unit**: All positions and sizes are in points (72 points = 1 inch). Standard 16:9 slide is 960x540 points.

6. **BGR Color Format**: PowerPoint COM uses `R + (G * 256) + (B * 65536)`. Always use `utils.color` functions.

7. **Connector Sites**: Each shape has numbered connection sites. `shape.ConnectionSiteCount` returns the number. Standard shapes typically have 4 sites (top, right, bottom, left = 1, 2, 3, 4).

8. **Group Operations**: GroupItems allow property modification but not adding/removing shapes from a group. To change group membership, ungroup, modify, and regroup.

9. **Tags**: Tags are invisible to users. Names auto-uppercase. Values are strings only. Very useful for MCP metadata tracking.

10. **PickUp/Apply**: Copies shape formatting (fill, line, effects) but NOT text formatting. For text formatting, use the text_and_formatting module.
