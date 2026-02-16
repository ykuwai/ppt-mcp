# Missing Operations Audit: PowerPoint MCP Server

> Comprehensive audit of commonly-needed PowerPoint operations not yet implemented
> in the ppt-com-mcp server (currently 95 tools). Covers COM API gaps, ExecuteMso
> operations, clipboard operations, and undo management.

---

## Table of Contents

1. [HIGH Priority: Shape Alignment and Distribution](#1-shape-alignment-and-distribution)
2. [HIGH Priority: Slide Size / Page Setup](#2-slide-size--page-setup)
3. [HIGH Priority: Copy Formatting (Format Painter)](#3-copy-formatting-format-painter)
4. [HIGH Priority: Clipboard Operations (Copy/Paste Shapes)](#4-clipboard-operations-copypaste-shapes)
5. [HIGH Priority: Undo / Redo](#5-undo--redo)
6. [HIGH Priority: Image Cropping](#6-image-cropping)
7. [HIGH Priority: Slide Background Formatting](#7-slide-background-formatting)
8. [MEDIUM Priority: Shape Effects (Glow, Reflection, Soft Edge, 3D)](#8-shape-effects-glow-reflection-soft-edge-3d)
9. [MEDIUM Priority: Merge Shapes](#9-merge-shapes)
10. [MEDIUM Priority: Shape Flip](#10-shape-flip)
11. [MEDIUM Priority: Comments](#11-comments)
12. [MEDIUM Priority: Tags (Custom Metadata)](#12-tags-custom-metadata)
13. [MEDIUM Priority: Replace Fonts Presentation-Wide](#13-replace-fonts-presentation-wide)
14. [MEDIUM Priority: Lock/Unlock Aspect Ratio](#14-lockunlock-aspect-ratio)
15. [MEDIUM Priority: Export Shape as Image](#15-export-shape-as-image)
16. [LOW Priority: Selection Operations](#16-selection-operations)
17. [LOW Priority: StartNewUndoEntry](#17-startnewundoentry)
18. [LOW Priority: View Control](#18-view-control)
19. [LOW Priority: Copy Animation Between Shapes](#19-copy-animation-between-shapes)
20. [LOW Priority: Slide Hidden Property](#20-slide-hidden-property)
21. [Summary Table](#summary-table)

---

## 1. Shape Alignment and Distribution

**Priority: HIGH**

### Why It's Useful

Shape alignment and distribution are among the most frequently used layout operations in PowerPoint. An AI agent building slide layouts needs to align shapes to each other or to the slide, and distribute them evenly. Without this, the agent must calculate positions manually -- error-prone and verbose.

### COM API

- `ShapeRange.Align(AlignCmd, RelativeTo)` -- align shapes
- `ShapeRange.Distribute(DistributeCmd, RelativeTo)` -- distribute shapes evenly

### Constants

```python
# MsoAlignCmd
msoAlignLefts = 0
msoAlignCenters = 1
msoAlignRights = 2
msoAlignTops = 3
msoAlignMiddles = 4
msoAlignBottoms = 5

# MsoDistributeCmd
msoDistributeHorizontally = 0
msoDistributeVertically = 1
```

### Python Code Example

```python
def _align_shapes_impl(slide_index, shape_names, align_cmd, relative_to_slide):
    """Align multiple shapes using ShapeRange.Align."""
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    # Create ShapeRange from names (pass as tuple for COM SAFEARRAY)
    shape_range = slide.Shapes.Range(tuple(shape_names))

    # msoTrue=-1 = relative to slide, msoFalse=0 = relative to each other
    relative = -1 if relative_to_slide else 0
    shape_range.Align(align_cmd, relative)

    return {
        "success": True,
        "aligned_count": shape_range.Count,
        "align_cmd": align_cmd,
        "relative_to_slide": relative_to_slide,
    }

def _distribute_shapes_impl(slide_index, shape_names, distribute_cmd, relative_to_slide):
    """Distribute shapes evenly."""
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    shape_range = slide.Shapes.Range(tuple(shape_names))
    relative = -1 if relative_to_slide else 0
    shape_range.Distribute(distribute_cmd, relative)

    return {
        "success": True,
        "distributed_count": shape_range.Count,
    }
```

### Proposed Tools (2 new tools)

| Tool Name | Description |
|---|---|
| `ppt_align_shapes` | Align multiple shapes (left, center, right, top, middle, bottom) relative to each other or the slide |
| `ppt_distribute_shapes` | Distribute shapes evenly (horizontal or vertical) |

### References

- [ShapeRange.Align method](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shaperange.align)
- [ShapeRange.Distribute method](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shaperange.distribute)
- [MsoAlignCmd enumeration](https://learn.microsoft.com/en-us/office/vba/api/office.msoaligncmd)

---

## 2. Slide Size / Page Setup

**Priority: HIGH**

### Why It's Useful

Setting the slide dimensions is essential when creating presentations for different contexts: widescreen (16:9), standard (4:3), A4 portrait for printed documents, custom poster sizes, etc. An AI agent must be able to get and set these values to create purpose-appropriate presentations.

### COM API

- `Presentation.PageSetup.SlideWidth` -- slide width in points (R/W)
- `Presentation.PageSetup.SlideHeight` -- slide height in points (R/W)
- `Presentation.PageSetup.SlideSize` -- preset size constant (R/W)
- `Presentation.PageSetup.SlideOrientation` -- portrait/landscape (R/W)
- `Presentation.PageSetup.NotesOrientation` -- notes page orientation (R/W)

### Constants

```python
# PpSlideSizeType
ppSlideSizeOnScreen = 1           # 10x7.5 inches (4:3)
ppSlideSizeLetterPaper = 2        # 10x7.5 inches
ppSlideSizeA4Paper = 3            # 10.83x7.5 inches
ppSlideSize35MM = 4               # 11.25x7.5 inches
ppSlideSizeOverhead = 5           # 10x7.5 inches
ppSlideSizeBanner = 6             # 8x1 inches
ppSlideSizeCustom = 7             # Custom dimensions
ppSlideSizeOnScreen16x9 = 8       # 13.33x7.5 inches (16:9)
ppSlideSizeOnScreen16x10 = 9      # 13.33x8.33 inches (16:10)
ppSlideSizeWidescreen = 10        # 13.33x7.5 inches (widescreen)

# MsoOrientation
msoOrientationHorizontal = 1      # Landscape
msoOrientationVertical = 2        # Portrait
msoOrientationMixed = -2
```

### Python Code Example

```python
def _get_slide_size_impl():
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    ps = pres.PageSetup
    return {
        "width": round(ps.SlideWidth, 2),
        "height": round(ps.SlideHeight, 2),
        "width_inches": round(ps.SlideWidth / 72, 2),
        "height_inches": round(ps.SlideHeight / 72, 2),
        "slide_size": ps.SlideSize,
        "orientation": ps.SlideOrientation,
    }

def _set_slide_size_impl(width=None, height=None, slide_size=None, orientation=None):
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    ps = pres.PageSetup

    if slide_size is not None:
        ps.SlideSize = slide_size
    if width is not None:
        ps.SlideWidth = width
    if height is not None:
        ps.SlideHeight = height
    if orientation is not None:
        ps.SlideOrientation = orientation

    return {
        "success": True,
        "width": round(ps.SlideWidth, 2),
        "height": round(ps.SlideHeight, 2),
    }
```

### Proposed Tools (2 new tools)

| Tool Name | Description |
|---|---|
| `ppt_get_slide_size` | Get the current slide dimensions, size preset, and orientation |
| `ppt_set_slide_size` | Set slide dimensions (width/height in points, preset, or orientation) |

### References

- [PageSetup.SlideWidth property](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.pagesetup.slidewidth)
- [PageSetup.SlideHeight property](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.pagesetup.slideheight)
- [PageSetup.SlideSize property](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.pagesetup.slidesize)
- [PageSetup.SlideOrientation property](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.PageSetup.SlideOrientation)

---

## 3. Copy Formatting (Format Painter)

**Priority: HIGH**

### Why It's Useful

Format painter is one of the most-used features in PowerPoint. An AI agent that styles presentations needs to copy formatting from a reference shape to target shapes efficiently, rather than manually reading and setting every property.

### COM API

- `Shape.PickUp()` -- copies shape-level formatting to internal buffer
- `Shape.Apply()` -- applies the picked-up formatting to the shape

**Note:** PickUp/Apply copies fill, line, shadow, and 3D effects. It does NOT copy text-level formatting (font name, size, bold, etc.). For text formatting transfer, individual font properties must be copied manually.

### Python Code Example

```python
def _copy_formatting_impl(slide_index, source_name, target_names):
    """Copy shape formatting from source to one or more targets."""
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    source = _get_shape(slide, source_name)
    source.PickUp()

    applied_to = []
    for name in target_names:
        target = _get_shape(slide, name)
        target.Apply()
        applied_to.append(target.Name)

    return {
        "success": True,
        "source": source.Name,
        "applied_to": applied_to,
    }
```

### Proposed Tools (1 new tool)

| Tool Name | Description |
|---|---|
| `ppt_copy_formatting` | Copy shape formatting (fill, line, shadow, effects) from one shape to one or more target shapes via PickUp/Apply |

### References

- [Shape.PickUp method](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.pickup)
- [Shape.Apply method](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.apply)

---

## 4. Clipboard Operations (Copy/Paste Shapes)

**Priority: HIGH**

### Why It's Useful

Copy/paste is fundamental for duplicating content between slides, between presentations, or for converting shapes (paste special). The existing `duplicate_shape` tool only duplicates within the same slide. An AI agent needs to:
- Copy a shape from one slide and paste it to another
- Copy shapes between presentations
- Paste with specific format options (e.g., as picture)

### COM API

- `Shape.Copy()` / `ShapeRange.Copy()` -- copy to clipboard
- `Shape.Cut()` / `ShapeRange.Cut()` -- cut to clipboard
- `Shapes.Paste()` -- paste from clipboard, returns ShapeRange
- `Shapes.PasteSpecial(DataType, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Link)` -- paste with format control
- `Slide.Copy()` -- copy a slide to clipboard
- `Slides.Paste(Index)` -- paste a slide at specified position

### PpPasteDataType Constants

```python
ppPasteDefault = 0
ppPasteBitmap = 1
ppPasteEnhancedMetafile = 2
ppPasteMetafilePicture = 3
ppPasteGIF = 4
ppPasteJPG = 5
ppPastePNG = 6
ppPasteText = 7
ppPasteHTML = 8
ppPasteRTF = 9
ppPasteOLEObject = 10
ppPasteShape = 11
```

### Python Code Example

```python
def _copy_shape_to_slide_impl(src_slide_index, shape_name, dst_slide_index):
    """Copy a shape from one slide to another."""
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    src_slide = pres.Slides(src_slide_index)
    dst_slide = pres.Slides(dst_slide_index)

    shape = _get_shape(src_slide, shape_name)
    shape.Copy()

    pasted = dst_slide.Shapes.Paste()
    new_shape = pasted(1)

    return {
        "success": True,
        "new_shape_name": new_shape.Name,
        "destination_slide": dst_slide_index,
    }

def _paste_special_impl(slide_index, data_type):
    """Paste clipboard content with specific format."""
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    pasted = slide.Shapes.PasteSpecial(DataType=data_type)
    names = [pasted(i).Name for i in range(1, pasted.Count + 1)]

    return {
        "success": True,
        "pasted_count": pasted.Count,
        "shape_names": names,
    }
```

### Proposed Tools (2 new tools)

| Tool Name | Description |
|---|---|
| `ppt_copy_shape_to_slide` | Copy a shape from one slide and paste it to another (same or different slide) |
| `ppt_paste_special` | Paste clipboard content with a specific data type (bitmap, metafile, text, etc.) |

### References

- [Shape.Copy method](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.copy)
- [Shapes.Paste method](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shapes.paste)
- [Shapes.PasteSpecial method](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shapes.pastespecial)

---

## 5. Undo / Redo

**Priority: HIGH**

### Why It's Useful

An AI agent making multiple changes needs to be able to undo mistakes. Undo/Redo are among the most critical safety operations. There is no `Application.Undo()` method in PowerPoint's COM model, so `ExecuteMso` is the only viable approach.

### COM API

- `Application.CommandBars.ExecuteMso("Undo")` -- undo last action
- `Application.CommandBars.ExecuteMso("Redo")` -- redo last undone action
- `Application.CommandBars.GetEnabledMso("Undo")` -- check if undo is available
- `Application.CommandBars.GetEnabledMso("Redo")` -- check if redo is available

### Python Code Example

```python
def _undo_impl(times=1):
    """Undo the last N actions."""
    app = ppt._get_app_impl()
    undone = 0
    for _ in range(times):
        if not app.CommandBars.GetEnabledMso("Undo"):
            break
        app.CommandBars.ExecuteMso("Undo")
        undone += 1
    return {"success": True, "actions_undone": undone}

def _redo_impl(times=1):
    """Redo the last N undone actions."""
    app = ppt._get_app_impl()
    redone = 0
    for _ in range(times):
        if not app.CommandBars.GetEnabledMso("Redo"):
            break
        app.CommandBars.ExecuteMso("Redo")
        redone += 1
    return {"success": True, "actions_redone": redone}
```

### Proposed Tools (2 new tools)

| Tool Name | Description |
|---|---|
| `ppt_undo` | Undo the last N actions (default 1) via ExecuteMso |
| `ppt_redo` | Redo the last N undone actions (default 1) via ExecuteMso |

### Caveats

- Each call to `ExecuteMso("Undo")` undoes one action.
- There is no way to inspect the undo stack or get a list of undoable operations.
- COM operations are placed on the undo stack, so they can be undone.

### References

- [CommandBars.ExecuteMso method](https://learn.microsoft.com/en-us/office/vba/api/office.commandbars.executemso)
- [CommandBars.GetEnabledMso method](https://learn.microsoft.com/en-us/office/vba/api/Office.CommandBars.GetEnabledMso)

---

## 6. Image Cropping

**Priority: HIGH**

### Why It's Useful

Image cropping is essential for adjusting photos and screenshots inserted into presentations. AI agents placing images frequently need to crop to show only the relevant portion. This is a core image manipulation operation not currently exposed.

### COM API

The `PictureFormat` object on a shape provides cropping properties:

- `Shape.PictureFormat.CropLeft` -- crop from left in points
- `Shape.PictureFormat.CropRight` -- crop from right in points
- `Shape.PictureFormat.CropTop` -- crop from top in points
- `Shape.PictureFormat.CropBottom` -- crop from bottom in points
- `Shape.PictureFormat.Crop` -- returns a `Crop` object (PowerPoint 2010+) with additional properties: `PictureWidth`, `PictureHeight`, `PictureOffsetX`, `PictureOffsetY`, `ShapeWidth`, `ShapeHeight`

**Important:** Cropping is relative to the ORIGINAL (unscaled) size. If you insert a 100pt-wide image, scale it to 200pt, and set CropLeft=50, it crops 100pt from the visible image.

### Python Code Example

```python
def _crop_picture_impl(slide_index, shape_name_or_index,
                        crop_left=None, crop_right=None,
                        crop_top=None, crop_bottom=None):
    """Set cropping on a picture shape."""
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    pf = shape.PictureFormat

    if crop_left is not None:
        pf.CropLeft = crop_left
    if crop_right is not None:
        pf.CropRight = crop_right
    if crop_top is not None:
        pf.CropTop = crop_top
    if crop_bottom is not None:
        pf.CropBottom = crop_bottom

    return {
        "success": True,
        "shape_name": shape.Name,
        "crop_left": round(pf.CropLeft, 2),
        "crop_right": round(pf.CropRight, 2),
        "crop_top": round(pf.CropTop, 2),
        "crop_bottom": round(pf.CropBottom, 2),
    }
```

### Proposed Tools (1 new tool)

| Tool Name | Description |
|---|---|
| `ppt_crop_picture` | Set or get cropping on a picture shape (crop from left, right, top, bottom in points) |

### References

- [PictureFormat.CropLeft property](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.pictureformat.cropleft)
- [PictureFormat.CropRight property](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.pictureformat.cropright)
- [PictureFormat.Crop property](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.pictureformat.crop)

---

## 7. Slide Background Formatting

**Priority: HIGH**

### Why It's Useful

Slide backgrounds define the visual foundation of every slide. An AI agent creating visually rich presentations needs to set solid colors, gradients, or image backgrounds per slide. This is a major design capability gap.

### COM API

- `Slide.FollowMasterBackground` -- True/False whether slide uses master background
- `Slide.Background.Fill` -- FillFormat object for background
- `Slide.Background.Fill.Solid()` -- set solid fill
- `Slide.Background.Fill.ForeColor.RGB` -- background color
- `Slide.Background.Fill.TwoColorGradient(Style, Variant)` -- gradient background
- `Slide.Background.Fill.UserPicture(PictureFile)` -- picture background
- `Slide.Background.Fill.Transparency` -- background transparency

### Python Code Example

```python
def _set_slide_background_impl(slide_index, fill_type, color=None,
                                 gradient_color1=None, gradient_color2=None,
                                 gradient_style=None, image_path=None,
                                 transparency=None):
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    # Must disable FollowMasterBackground to set custom background
    slide.FollowMasterBackground = 0  # msoFalse

    fill = slide.Background.Fill

    if fill_type == "solid":
        fill.Solid()
        if color is not None:
            fill.ForeColor.RGB = hex_to_int(color)
    elif fill_type == "gradient":
        style_val = GRADIENT_STYLE_MAP.get(gradient_style, 1)
        fill.TwoColorGradient(style_val, 1)
        if gradient_color1:
            fill.ForeColor.RGB = hex_to_int(gradient_color1)
        if gradient_color2:
            fill.BackColor.RGB = hex_to_int(gradient_color2)
    elif fill_type == "picture":
        abs_path = os.path.abspath(image_path)
        fill.UserPicture(abs_path)
    elif fill_type == "none":
        fill.Background()
    elif fill_type == "master":
        slide.FollowMasterBackground = -1  # msoTrue

    if transparency is not None:
        fill.Transparency = transparency

    return {"success": True, "slide_index": slide_index, "fill_type": fill_type}
```

### Proposed Tools (1 new tool)

| Tool Name | Description |
|---|---|
| `ppt_set_slide_background` | Set slide background (solid color, gradient, picture, or reset to master) |

### References

- [Slide.Background property](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.background)
- [Slide.BackgroundStyle property](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.backgroundstyle)

---

## 8. Shape Effects (Glow, Reflection, Soft Edge, 3D)

**Priority: MEDIUM**

### Why It's Useful

Professional presentations use shape effects like glow, reflection, soft edges, and 3D/bevel to create polished visuals. The current server only supports shadow. Adding the remaining effects completes the shape formatting story.

### COM API

#### Glow
- `Shape.Glow.Radius` -- glow radius in points (Float, R/W)
- `Shape.Glow.Color` -- glow color (ColorFormat object)
- `Shape.Glow.Transparency` -- glow transparency (0-1)

#### Reflection
- `Shape.Reflection.Type` -- preset type (MsoReflectionType)
- `Shape.Reflection.Blur` -- blur amount
- `Shape.Reflection.Offset` -- distance offset
- `Shape.Reflection.Size` -- percentage of shape height (0-100)
- `Shape.Reflection.Transparency` -- transparency (0-1)

#### Soft Edge
- `Shape.SoftEdge.Type` -- preset type (MsoSoftEdgeType, e.g., `msoSoftEdge6`)
- `Shape.SoftEdge.Radius` -- soft edge radius in points

#### 3D / Bevel
- `Shape.ThreeD.BevelTopType` -- top bevel type (MsoBevelType)
- `Shape.ThreeD.BevelTopInset` -- top bevel width in points
- `Shape.ThreeD.BevelTopDepth` -- top bevel height in points
- `Shape.ThreeD.BevelBottomType` -- bottom bevel type
- `Shape.ThreeD.BevelBottomInset` -- bottom bevel width
- `Shape.ThreeD.BevelBottomDepth` -- bottom bevel height
- `Shape.ThreeD.Depth` -- extrusion depth
- `Shape.ThreeD.ExtrusionColor` -- extrusion color
- `Shape.ThreeD.PresetMaterial` -- surface material
- `Shape.ThreeD.PresetLightingDirection` -- light direction
- `Shape.ThreeD.Visible` -- 3D effect visibility

### Python Code Example

```python
def _set_glow_impl(slide_index, shape_name_or_index, radius, color=None, transparency=None):
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    glow = shape.Glow
    glow.Radius = radius
    if color is not None:
        glow.Color.RGB = hex_to_int(color)
    if transparency is not None:
        glow.Transparency = transparency

    return {"success": True, "shape_name": shape.Name}

def _set_reflection_impl(slide_index, shape_name_or_index,
                           reflection_type=None, blur=None, offset=None,
                           size=None, transparency=None):
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    ref = shape.Reflection
    if reflection_type is not None:
        ref.Type = reflection_type
    if blur is not None:
        ref.Blur = blur
    if offset is not None:
        ref.Offset = offset
    if size is not None:
        ref.Size = size
    if transparency is not None:
        ref.Transparency = transparency

    return {"success": True, "shape_name": shape.Name}

def _set_soft_edge_impl(slide_index, shape_name_or_index, radius):
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    shape.SoftEdge.Radius = radius

    return {"success": True, "shape_name": shape.Name}
```

### Proposed Tools (3 new tools)

| Tool Name | Description |
|---|---|
| `ppt_set_glow` | Set glow effect on a shape (radius, color, transparency) |
| `ppt_set_reflection` | Set reflection effect (type, blur, offset, size, transparency) |
| `ppt_set_soft_edge` | Set soft edge effect (radius in points; 0 to remove) |

**Note:** 3D/Bevel is complex (many properties). Consider a single `ppt_set_3d_effect` tool with optional params.

| Tool Name | Description |
|---|---|
| `ppt_set_3d_effect` | Set 3D/bevel effect on a shape (bevel type, depth, material, lighting) |

### References

- [Shape.Glow property](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Shape.Glow)
- [Shape.Reflection property](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.reflection)
- [Shape.SoftEdge property](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.softedge)
- [Shape.ThreeD property](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.threed)
- [GlowFormat object](https://learn.microsoft.com/en-us/office/vba/api/office.glowformat)
- [Working with Glow and Reflection (Office 2010)](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/hh148188(v=office.14))

---

## 9. Merge Shapes

**Priority: MEDIUM**

### Why It's Useful

Shape merging (union, intersect, subtract, combine, fragment) enables creation of complex custom shapes from simple primitives. This is a powerful design capability for AI agents creating logos, icons, or custom graphics within presentations.

### COM API

**Method 1: ShapeRange.MergeShapes (PowerPoint 2013+)**

```
ShapeRange.MergeShapes(MergeCmd, PrimaryShape)
```

- `MergeCmd` -- Required, MsoMergeCmd constant
- `PrimaryShape` -- Optional, the shape whose formatting the result inherits

**MsoMergeCmd Constants:**

| Constant | Value | Description |
|---|---|---|
| `msoMergeUnion` | 1 | Union -- merge into one shape covering all area |
| `msoMergeCombine` | 2 | Combine -- merge but remove overlapping areas |
| `msoMergeIntersect` | 3 | Intersect -- keep only overlapping area |
| `msoMergeSubtract` | 4 | Subtract -- remove overlap from primary shape |
| `msoMergeFragment` | 5 | Fragment -- split into separate pieces at intersections |

**Method 2: ExecuteMso (PowerPoint 2010+)**

Requires shapes to be selected in the UI first:
- `ExecuteMso("ShapesUnion")`
- `ExecuteMso("ShapesCombine")`
- `ExecuteMso("ShapesIntersect")`
- `ExecuteMso("ShapesSubtract")`
- `ExecuteMso("ShapesFragment")`

### Python Code Example

```python
# MsoMergeCmd constants
msoMergeUnion = 1
msoMergeCombine = 2
msoMergeIntersect = 3
msoMergeSubtract = 4
msoMergeFragment = 5

MERGE_CMD_MAP = {
    "union": msoMergeUnion,
    "combine": msoMergeCombine,
    "intersect": msoMergeIntersect,
    "subtract": msoMergeSubtract,
    "fragment": msoMergeFragment,
}

def _merge_shapes_impl(slide_index, shape_names, merge_cmd, primary_shape_name=None):
    """Merge shapes using ShapeRange.MergeShapes."""
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    shape_range = slide.Shapes.Range(tuple(shape_names))

    primary = None
    if primary_shape_name:
        primary = _get_shape(slide, primary_shape_name)

    if primary:
        shape_range.MergeShapes(merge_cmd, primary)
    else:
        shape_range.MergeShapes(merge_cmd)

    # After merge, we need to find the resulting shape
    # It's typically the last shape or the one at the primary shape's position
    return {"success": True, "merge_type": merge_cmd}
```

### Proposed Tools (1 new tool)

| Tool Name | Description |
|---|---|
| `ppt_merge_shapes` | Merge multiple shapes (union, combine, intersect, subtract, fragment) |

### Caveats

- `ShapeRange.MergeShapes` is only available in PowerPoint 2013+.
- For PowerPoint 2010, `ExecuteMso` with selection is the only option.
- After merging, the original shapes are destroyed and replaced by the result.
- The `PrimaryShape` determines the formatting of the result.

### References

- [ShapeRange.MergeShapes method](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shaperange.mergeshapes)
- [MsoMergeCmd enumeration](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.core.msomergecmd?view=office-pia)

---

## 10. Shape Flip

**Priority: MEDIUM**

### Why It's Useful

Flipping shapes horizontally or vertically is a common transformation for creating mirrored layouts, arrow direction changes, or symmetric designs. While rotation is already in `update_shape`, flip is a separate operation.

### COM API

- `Shape.Flip(FlipCmd)` -- flip a shape
  - `msoFlipHorizontal = 0`
  - `msoFlipVertical = 1`
- `Shape.HorizontalFlip` -- read-only, True if horizontally flipped
- `Shape.VerticalFlip` -- read-only, True if vertically flipped

### Python Code Example

```python
def _flip_shape_impl(slide_index, shape_name_or_index, flip_direction):
    """Flip a shape horizontally or vertically."""
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    # msoFlipHorizontal=0, msoFlipVertical=1
    flip_cmd = 0 if flip_direction == "horizontal" else 1
    shape.Flip(flip_cmd)

    return {
        "success": True,
        "shape_name": shape.Name,
        "horizontal_flip": bool(shape.HorizontalFlip),
        "vertical_flip": bool(shape.VerticalFlip),
    }
```

### Proposed Tools (1 new tool)

| Tool Name | Description |
|---|---|
| `ppt_flip_shape` | Flip a shape horizontally or vertically |

**Alternative:** Could be added as a `flip` parameter to the existing `ppt_update_shape` tool.

### References

- [Shape.Flip method](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Shape.Flip)
- [Shape.HorizontalFlip property](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.horizontalflip)

---

## 11. Comments

**Priority: MEDIUM**

### Why It's Useful

Comments are essential for collaboration workflows. An AI agent reviewing or providing feedback on presentations needs to add comments to specific slides, read existing comments, and manage comment threads. This is critical for AI-assisted review workflows.

### COM API

- `Slide.Comments` -- Comments collection
- `Slide.Comments.Add2(Left, Top, Author, AuthorInitials, Text, ProviderID, UserID)` -- add a comment (replaces hidden `Add`)
- `Comment.Text` -- comment text
- `Comment.Author` -- author name
- `Comment.AuthorInitials` -- author initials
- `Comment.DateTime` -- timestamp
- `Comment.Left`, `Comment.Top` -- position
- `Comment.Delete()` -- delete a comment
- `Slide.Comments.Count` -- number of comments

### Python Code Example

```python
def _add_comment_impl(slide_index, text, author="AI Agent", author_initials="AI",
                        left=0, top=0):
    """Add a comment to a slide."""
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    # Use Add2 (recommended over deprecated Add)
    comment = slide.Comments.Add2(
        left, top,               # Position
        author, author_initials, # Author info
        text,                    # Comment text
        "AD", ""                 # ProviderID, UserID
    )

    return {
        "success": True,
        "comment_index": comment.Index if hasattr(comment, 'Index') else None,
        "text": text,
        "author": author,
    }

def _list_comments_impl(slide_index):
    """List all comments on a slide."""
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    comments = []
    for i in range(1, slide.Comments.Count + 1):
        c = slide.Comments(i)
        comments.append({
            "index": i,
            "author": c.Author,
            "author_initials": c.AuthorInitials,
            "text": c.Text,
            "datetime": str(c.DateTime),
            "left": c.Left,
            "top": c.Top,
        })

    return {
        "slide_index": slide_index,
        "comments_count": slide.Comments.Count,
        "comments": comments,
    }

def _delete_comment_impl(slide_index, comment_index):
    """Delete a specific comment on a slide."""
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    slide.Comments(comment_index).Delete()
    return {"success": True}
```

### Proposed Tools (3 new tools)

| Tool Name | Description |
|---|---|
| `ppt_add_comment` | Add a comment to a slide at a specific position |
| `ppt_list_comments` | List all comments on a slide |
| `ppt_delete_comment` | Delete a specific comment by index |

### Caveats

- `Comments.Add2` replaces the deprecated `Comments.Add` method.
- Modern comments (threaded comments) in Microsoft 365 may behave differently from classic comments.
- The `ProviderID` and `UserID` parameters in `Add2` are required but can be empty strings for non-enterprise scenarios.

### References

- [Comments.Add2 method](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.comments.add2)
- [Comment object](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Comment)
- [Slide.Comments property](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.comments)

---

## 12. Tags (Custom Metadata)

**Priority: MEDIUM**

### Why It's Useful

Tags allow storing arbitrary key-value metadata on presentations, slides, and shapes. For an AI agent, this is invaluable for:
- Marking shapes that have been processed or generated
- Storing template variable names for data-driven presentations
- Tracking workflow state across multiple tool invocations
- Adding semantic labels to shapes for later retrieval

### COM API

Available on Presentation, Slide, and Shape objects:

- `object.Tags.Add(Name, Value)` -- add a tag (both strings)
- `object.Tags(Name)` -- read a tag value by name
- `object.Tags.Count` -- number of tags
- `object.Tags.Name(Index)` -- get tag name by 1-based index
- `object.Tags.Value(Index)` -- get tag value by 1-based index
- `object.Tags.Delete(Name)` -- delete a tag

### Python Code Example

```python
def _set_tag_impl(slide_index, shape_name_or_index, tag_name, tag_value,
                    target_type="shape"):
    """Set a tag on a shape, slide, or presentation."""
    app = ppt._get_app_impl()
    pres = app.ActivePresentation

    if target_type == "presentation":
        target = pres
    elif target_type == "slide":
        target = pres.Slides(slide_index)
    else:
        slide = pres.Slides(slide_index)
        target = _get_shape(slide, shape_name_or_index)

    target.Tags.Add(tag_name, tag_value)

    return {"success": True, "tag_name": tag_name, "tag_value": tag_value}

def _get_tags_impl(slide_index, shape_name_or_index, target_type="shape"):
    """Get all tags from a shape, slide, or presentation."""
    app = ppt._get_app_impl()
    pres = app.ActivePresentation

    if target_type == "presentation":
        target = pres
    elif target_type == "slide":
        target = pres.Slides(slide_index)
    else:
        slide = pres.Slides(slide_index)
        target = _get_shape(slide, shape_name_or_index)

    tags = {}
    for i in range(1, target.Tags.Count + 1):
        tags[target.Tags.Name(i)] = target.Tags.Value(i)

    return {"tags": tags, "count": target.Tags.Count}
```

### Proposed Tools (2 new tools)

| Tool Name | Description |
|---|---|
| `ppt_set_tag` | Set a custom key-value tag on a shape, slide, or presentation |
| `ppt_get_tags` | Get all custom tags from a shape, slide, or presentation |

### References

- [Tags object](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.tags)
- [Shape.Tags property](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.tags)

---

## 13. Replace Fonts Presentation-Wide

**Priority: MEDIUM**

### Why It's Useful

When rebranding or ensuring consistency across a presentation, replacing all instances of one font with another is a critical operation. The COM API provides a built-in method that is much more efficient than iterating through every shape manually.

### COM API

- `Presentation.Fonts.Replace(Original, Replacement)` -- replace all occurrences of a font
- `Presentation.Fonts.Count` -- number of fonts used
- `Presentation.Fonts(Index).Name` -- font name at index

### Python Code Example

```python
def _replace_font_impl(original_font, replacement_font):
    """Replace all instances of one font with another."""
    app = ppt._get_app_impl()
    pres = app.ActivePresentation

    pres.Fonts.Replace(original_font, replacement_font)

    return {
        "success": True,
        "original": original_font,
        "replacement": replacement_font,
    }

def _list_fonts_impl():
    """List all fonts used in the presentation."""
    app = ppt._get_app_impl()
    pres = app.ActivePresentation

    fonts = []
    for i in range(1, pres.Fonts.Count + 1):
        fonts.append(pres.Fonts(i).Name)

    return {"fonts": fonts, "count": pres.Fonts.Count}
```

### Proposed Tools (2 new tools)

| Tool Name | Description |
|---|---|
| `ppt_replace_font` | Replace all instances of one font with another presentation-wide |
| `ppt_list_fonts` | List all fonts currently used in the presentation |

### Caveats

- `Fonts.Replace` does not affect fonts inside charts or embedded OLE objects.

### References

- [Fonts.Replace method](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.fonts.replace)

---

## 14. Lock/Unlock Aspect Ratio

**Priority: MEDIUM**

### Why It's Useful

When resizing shapes -- especially images -- controlling whether the aspect ratio is locked prevents accidental distortion. An AI agent resizing images needs to toggle this property before changing dimensions.

### COM API

- `Shape.LockAspectRatio` -- MsoTriState (R/W)
  - `msoTrue = -1` -- locked (proportional resizing)
  - `msoFalse = 0` -- unlocked (independent width/height)

### Python Code Example

```python
def _set_lock_aspect_ratio_impl(slide_index, shape_name_or_index, locked):
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    shape.LockAspectRatio = -1 if locked else 0  # msoTrue / msoFalse

    return {
        "success": True,
        "shape_name": shape.Name,
        "lock_aspect_ratio": locked,
    }
```

### Implementation Note

Rather than a separate tool, this could be added as a `lock_aspect_ratio` parameter to the existing `ppt_update_shape` tool. That way, the user can set it in the same call where they resize the shape.

### Proposed Change

Add `lock_aspect_ratio: Optional[bool]` to `UpdateShapeInput` in `shapes.py`.

### References

- [Shape.LockAspectRatio property](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.lockaspectratio)

---

## 15. Export Shape as Image

**Priority: MEDIUM**

### Why It's Useful

Exporting individual shapes as images is useful for:
- Generating thumbnails of specific elements
- Extracting diagrams or charts for use elsewhere
- Creating image assets from shape compositions

### COM API

- `Shape.Export(PathName, FilterType, ScaleWidth, ScaleHeight, ExportMode)` -- export shape as image

### PpShapeFormatType Constants

```python
ppShapeFormatGIF = 0
ppShapeFormatJPG = 1
ppShapeFormatPNG = 2
ppShapeFormatBMP = 3
ppShapeFormatWMF = 4
ppShapeFormatEMF = 5
ppShapeFormatSVG = 6    # Windows version 2302+
```

### Python Code Example

```python
def _export_shape_impl(slide_index, shape_name_or_index, file_path, format="png",
                         width=None, height=None):
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    abs_path = os.path.abspath(file_path)

    FORMAT_MAP = {
        "gif": 0, "jpg": 1, "png": 2, "bmp": 3,
        "wmf": 4, "emf": 5, "svg": 6,
    }
    fmt = FORMAT_MAP.get(format.lower(), 2)  # default PNG

    if width and height:
        shape.Export(abs_path, fmt, width, height)
    else:
        shape.Export(abs_path, fmt)

    return {
        "success": True,
        "file_path": abs_path,
        "format": format,
    }
```

### Proposed Tools (1 new tool)

| Tool Name | Description |
|---|---|
| `ppt_export_shape` | Export a single shape as an image file (PNG, JPG, SVG, etc.) |

### References

- [Shape.Export method](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.export)

---

## 16. Selection Operations

**Priority: LOW**

### Why It's Useful

Selection operations are primarily needed as prerequisites for ExecuteMso-based operations (merge shapes, alignment via ExecuteMso). COM-based alignment via ShapeRange is preferred, but selection is still useful for:
- Inspecting what the user has currently selected
- Setting up selections for operations that require it

### COM API

- `Shape.Select(Replace)` -- select a shape; Replace=True replaces selection, False adds to it
- `ActiveWindow.Selection.Type` -- what's selected (ppSelectionNone/Slides/Shapes/Text)
- `ActiveWindow.Selection.ShapeRange` -- the selected shapes
- `ActiveWindow.Selection.TextRange` -- the selected text

### Python Code Example

```python
def _select_shapes_impl(slide_index, shape_names):
    """Select specific shapes in the PowerPoint UI."""
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    window = app.ActiveWindow

    window.View.GotoSlide(slide_index)
    slide = pres.Slides(slide_index)

    # Select first shape (replace existing selection)
    _get_shape(slide, shape_names[0]).Select()  # Replace=True (default)

    # Add remaining shapes to selection
    for name in shape_names[1:]:
        _get_shape(slide, name).Select(0)  # msoFalse = add to selection

    return {"success": True, "selected_count": len(shape_names)}

def _get_selection_impl():
    """Get information about the current selection."""
    app = ppt._get_app_impl()
    window = app.ActiveWindow
    sel = window.Selection

    result = {"type": sel.Type}

    if sel.Type == 2:  # ppSelectionShapes
        shapes = []
        for i in range(1, sel.ShapeRange.Count + 1):
            shapes.append(sel.ShapeRange(i).Name)
        result["shapes"] = shapes
    elif sel.Type == 3:  # ppSelectionText
        result["text"] = sel.TextRange.Text

    return result
```

### Proposed Tools (2 new tools)

| Tool Name | Description |
|---|---|
| `ppt_select_shapes` | Select specific shapes in the UI (for ExecuteMso operations) |
| `ppt_get_selection` | Get information about the current UI selection |

### References

- [Shape.Select method](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.select)
- [Selection object](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.selection)

---

## 17. StartNewUndoEntry

**Priority: LOW**

### Why It's Useful

When an AI agent performs multiple COM operations that logically form one "action", grouping them into a single undo entry improves the user experience. Without this, each individual COM call is a separate undo step, requiring many Ctrl+Z presses to fully undo a complex operation.

### COM API

- `Application.StartNewUndoEntry()` -- starts a new undo boundary

When this method is called, all subsequent operations until the next call (or until macro completion) are grouped as a single undo entry.

### Python Code Example

```python
def _start_undo_entry_impl():
    """Start a new undo entry boundary.

    All subsequent COM operations will be grouped as a single undo step
    until the next StartNewUndoEntry call or until execution ends.
    """
    app = ppt._get_app_impl()
    app.StartNewUndoEntry()
    return {"success": True}
```

### Proposed Tools (1 new tool)

| Tool Name | Description |
|---|---|
| `ppt_start_undo_entry` | Start a new undo entry so subsequent operations can be undone as one step |

### Caveats

- Available since PowerPoint 2010.
- Normally, when macro execution ends, all operations are grouped as one undo entry. This method is useful for creating multiple undo boundaries within a single macro/session.

### References

- [Application.StartNewUndoEntry method](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Application.StartNewUndoEntry)

---

## 18. View Control

**Priority: LOW**

### Why It's Useful

Controlling the view state (zoom level, view type) is useful for automated workflows where the agent needs to set up the editor for the user, or switch between Normal, Slide Sorter, and other views.

### COM API

- `ActiveWindow.ViewType` -- get/set view type (PpViewType constant)
- `ActiveWindow.View.Zoom` -- get/set zoom percentage
- `Application.WindowState` -- get/set window state (normal, minimized, maximized)

### Python Code Example

```python
def _set_view_impl(view_type=None, zoom=None):
    app = ppt._get_app_impl()
    window = app.ActiveWindow

    if view_type is not None:
        window.ViewType = view_type
    if zoom is not None:
        window.View.Zoom = zoom

    return {
        "success": True,
        "view_type": window.ViewType,
        "zoom": window.View.Zoom,
    }
```

### Proposed Tools (1 new tool)

| Tool Name | Description |
|---|---|
| `ppt_set_view` | Set the editor view type (normal, slide sorter, etc.) and zoom level |

### References

- [DocumentWindow.ViewType property](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.documentwindow.viewtype)

---

## 19. Copy Animation Between Shapes

**Priority: LOW**

### Why It's Useful

When building consistent animations across a presentation, copying animations from one shape to another saves significant effort. PowerPoint 2010+ supports this via COM.

### COM API

- `Shape.PickupAnimation()` -- copy animation settings into buffer
- `Shape.ApplyAnimation()` -- apply buffered animation to shape

### Python Code Example

```python
def _copy_animation_impl(slide_index, source_name, target_name):
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    source = _get_shape(slide, source_name)
    target = _get_shape(slide, target_name)

    source.PickupAnimation()
    target.ApplyAnimation()

    return {"success": True, "source": source_name, "target": target_name}
```

### Proposed Tools (1 new tool)

| Tool Name | Description |
|---|---|
| `ppt_copy_animation` | Copy all animations from one shape to another via PickupAnimation/ApplyAnimation |

### References

- [Shape.PickupAnimation method](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.pickupanimation)
- [Shape.ApplyAnimation method](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.applyanimation)

---

## 20. Slide Hidden Property

**Priority: LOW**

### Why It's Useful

Hiding/showing slides is a common need when preparing different versions of a presentation (e.g., hiding backup slides). The `get_slide_info` tool already reads the hidden state, but there is no tool to SET it.

### COM API

- `Slide.SlideShowTransition.Hidden` -- MsoTriState (R/W), controls whether the slide is hidden in slideshow

### Python Code Example

```python
def _set_slide_hidden_impl(slide_index, hidden):
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    slide.SlideShowTransition.Hidden = -1 if hidden else 0  # msoTrue / msoFalse

    return {
        "success": True,
        "slide_index": slide_index,
        "hidden": hidden,
    }
```

### Implementation Note

Rather than a separate tool, this could be added as a parameter to an existing tool, or as a new lightweight tool.

### Proposed Tools (1 new tool)

| Tool Name | Description |
|---|---|
| `ppt_set_slide_hidden` | Hide or unhide a slide in the slideshow |

---

## Summary Table

| # | Feature | Priority | New Tools | COM Approach | Notes |
|---|---|---|---|---|---|
| 1 | Shape Alignment & Distribution | **HIGH** | 2 | `ShapeRange.Align` / `.Distribute` | Most-requested layout feature |
| 2 | Slide Size / Page Setup | **HIGH** | 2 | `PageSetup.SlideWidth` / `.SlideHeight` | Essential for presentation setup |
| 3 | Copy Formatting | **HIGH** | 1 | `Shape.PickUp` / `.Apply` | Format painter via COM |
| 4 | Clipboard (Copy/Paste) | **HIGH** | 2 | `Shape.Copy` / `Shapes.Paste` / `.PasteSpecial` | Cross-slide/cross-pres copy |
| 5 | Undo / Redo | **HIGH** | 2 | `ExecuteMso("Undo")` / `("Redo")` | Critical safety feature |
| 6 | Image Cropping | **HIGH** | 1 | `PictureFormat.CropLeft/Right/Top/Bottom` | Core image editing |
| 7 | Slide Background | **HIGH** | 1 | `Slide.Background.Fill` | Design essential |
| 8 | Shape Effects | **MEDIUM** | 4 | `Glow` / `Reflection` / `SoftEdge` / `ThreeD` | Professional visual effects |
| 9 | Merge Shapes | **MEDIUM** | 1 | `ShapeRange.MergeShapes` | Custom shape creation |
| 10 | Shape Flip | **MEDIUM** | 1 | `Shape.Flip(FlipCmd)` | Mirror transformation |
| 11 | Comments | **MEDIUM** | 3 | `Comments.Add2` / `.Delete` | Collaboration workflows |
| 12 | Tags (Metadata) | **MEDIUM** | 2 | `Tags.Add` / `Tags(Name)` | AI agent state tracking |
| 13 | Replace Fonts | **MEDIUM** | 2 | `Fonts.Replace` | Bulk font changes |
| 14 | Lock Aspect Ratio | **MEDIUM** | 0* | `Shape.LockAspectRatio` | *Add to update_shape |
| 15 | Export Shape as Image | **MEDIUM** | 1 | `Shape.Export` | Individual shape export |
| 16 | Selection Operations | **LOW** | 2 | `Shape.Select` / `Selection` | Prerequisite for ExecuteMso |
| 17 | StartNewUndoEntry | **LOW** | 1 | `Application.StartNewUndoEntry` | Undo grouping |
| 18 | View Control | **LOW** | 1 | `ActiveWindow.ViewType` / `.Zoom` | Editor state management |
| 19 | Copy Animation | **LOW** | 1 | `PickupAnimation` / `ApplyAnimation` | Animation reuse |
| 20 | Slide Hidden | **LOW** | 1 | `SlideShowTransition.Hidden` | Show/hide slides |

### Total New Tools: ~31

### Recommended Implementation Order

**Phase 4A (HIGH priority -- 11 tools):**
1. Shape Alignment & Distribution (2 tools)
2. Slide Size / Page Setup (2 tools)
3. Copy Formatting (1 tool)
4. Clipboard Operations (2 tools)
5. Undo / Redo (2 tools)
6. Image Cropping (1 tool)
7. Slide Background (1 tool)

**Phase 4B (MEDIUM priority -- 14 tools):**
8. Shape Effects (4 tools)
9. Merge Shapes (1 tool)
10. Shape Flip (1 tool)
11. Comments (3 tools)
12. Tags (2 tools)
13. Replace Fonts (2 tools)
14. Export Shape as Image (1 tool)

**Phase 4C (LOW priority -- 6 tools + 1 enhancement):**
15. Lock Aspect Ratio (enhance existing tool)
16. Selection Operations (2 tools)
17. StartNewUndoEntry (1 tool)
18. View Control (1 tool)
19. Copy Animation (1 tool)
20. Slide Hidden (1 tool)

---

## Additional ExecuteMso Commands Reference

For reference, these are useful ExecuteMso commands that could be exposed through a generic `ppt_execute_mso` tool:

```python
USEFUL_EXECUTEMSO_COMMANDS = {
    # Undo / Redo
    "Undo": "Undo the last action",
    "Redo": "Redo the last undone action",

    # Clipboard
    "Copy": "Copy selection to clipboard",
    "Cut": "Cut selection to clipboard",
    "Paste": "Paste from clipboard",
    "PasteSourceFormatting": "Paste keeping source formatting",
    "PasteDestinationFormatting": "Paste with destination formatting",
    "PastePicture": "Paste as picture",

    # Shape Merging (require shapes to be selected)
    "ShapesUnion": "Union of selected shapes",
    "ShapesCombine": "Combine selected shapes",
    "ShapesIntersect": "Intersect selected shapes",
    "ShapesSubtract": "Subtract selected shapes",
    "ShapesFragment": "Fragment selected shapes",

    # Alignment (require shapes to be selected)
    "ObjectsAlignLeftSmart": "Align left edges",
    "ObjectsAlignCenterHorizontalSmart": "Align centers horizontally",
    "ObjectsAlignRightSmart": "Align right edges",
    "ObjectsAlignTopSmart": "Align top edges",
    "ObjectsAlignMiddleVerticalSmart": "Align middles vertically",
    "ObjectsAlignBottomSmart": "Align bottom edges",
    "ObjectsDistributeHorizontalSmart": "Distribute horizontally",
    "ObjectsDistributeVerticalSmart": "Distribute vertically",

    # Selection
    "SelectAll": "Select all objects on slide",

    # Animation
    "AnimationPreview": "Preview animations on current slide",

    # Slideshow
    "SlideShowFromBeginning": "Start slideshow from slide 1",
    "SlideShowFromCurrent": "Start slideshow from current slide",
}
```

---

## References

### Microsoft Documentation
- [PowerPoint Shape object](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape)
- [ShapeRange object](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.ShapeRange)
- [PageSetup object](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.pagesetup)
- [PictureFormat object](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.pictureformat.cropleft)
- [Comments object](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.comments)
- [Tags object](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.tags)
- [Fonts object / Replace method](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.fonts.replace)
- [CommandBars.ExecuteMso](https://learn.microsoft.com/en-us/office/vba/api/office.commandbars.executemso)
- [Application.StartNewUndoEntry](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Application.StartNewUndoEntry)

### idMso Reference
- [idMso Control List for PowerPoint 2013/2010](http://youpresent.co.uk/idmso-control-list-powerpoint-2013-2010/)
- [Office Fluent UI Command Identifiers (GitHub)](https://github.com/OfficeDev/office-fluent-ui-command-identifiers)

### Internal Research Documents
- [08_hybrid_com_ui_patterns.md](./08_hybrid_com_ui_patterns.md) -- Format Painter, Alignment, ExecuteMso patterns
