"""Shape operations for PowerPoint COM automation.

Handles adding, listing, modifying, duplicating, deleting shapes,
and z-order management on PowerPoint slides.
"""

import json
import logging
from typing import Optional, Union

from pydantic import BaseModel, Field, ConfigDict

from utils.color import hex_to_int
from utils.com_wrapper import ppt
from utils.navigation import goto_slide
from ppt_com.constants import (
    SHAPE_TYPE_NAMES,
    msoTrue, msoFalse,
    msoGroup,
    msoTextOrientationHorizontal,
    msoBringToFront, msoSendToBack, msoBringForward, msoSendBackward,
    GRADIENT_STYLE_MAP,
)

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Friendly shape name -> MsoAutoShapeType mapping
# ---------------------------------------------------------------------------
SHAPE_NAME_MAP: dict[str, int] = {
    "rectangle": 1,
    "parallelogram": 2,
    "trapezoid": 3,
    "diamond": 4,
    "rounded_rectangle": 5,
    "octagon": 6,
    "triangle": 7,
    "right_triangle": 8,
    "oval": 9,
    "hexagon": 10,
    "cross": 11,
    "pentagon": 12,
    "can": 13,
    "cube": 14,
    "smiley_face": 17,
    "donut": 18,
    "no_symbol": 19,
    "heart": 21,
    "lightning_bolt": 22,
    "sun": 23,
    "moon": 24,
    "arc": 25,
    "right_arrow": 33,
    "left_arrow": 34,
    "up_arrow": 35,
    "down_arrow": 36,
    "left_right_arrow": 37,
    "up_down_arrow": 38,
    "quad_arrow": 39,
    "chevron": 52,
    "flowchart_process": 61,
    "flowchart_decision": 63,
    "flowchart_data": 64,
    "flowchart_document": 67,
    "flowchart_terminator": 69,
    "flowchart_connector": 73,
    "explosion": 89,
    "star_4point": 91,
    "star_5point": 92,
    "star_8point": 93,
    "star_16point": 94,
    "star_24point": 95,
    "star_32point": 96,
    "cloud": 179,
}

ZORDER_CMD_MAP: dict[str, int] = {
    "bring_to_front": msoBringToFront,
    "send_to_back": msoSendToBack,
    "bring_forward": msoBringForward,
    "send_backward": msoSendBackward,
}


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class AddShapeInput(BaseModel):
    """Input for adding an auto shape to a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_type: Union[int, str] = Field(
        ...,
        description=(
            "MsoAutoShapeType integer or friendly name "
            "(e.g. 'rectangle', 'oval', 'right_arrow', 'star_5point')"
        ),
    )
    left: float = Field(..., description="Left position in points")
    top: float = Field(..., description="Top position in points")
    width: float = Field(..., description="Width in points")
    height: float = Field(..., description="Height in points")
    text: Optional[str] = Field(default=None, description="Optional text content")
    # --- inline fill (optional — avoids a separate ppt_set_fill call) ---
    fill_color: Optional[str] = Field(
        default=None,
        description=(
            "Solid fill color '#RRGGBB'. Implies fill_type='solid' when fill_type is omitted. "
            "For gradient fills this is the start/fore color."
        ),
    )
    fill_type: Optional[str] = Field(
        default=None,
        description="Fill type: 'solid', 'none', or 'gradient'. Defaults to 'solid' when fill_color is given.",
    )
    fill_color2: Optional[str] = Field(
        default=None,
        description="Gradient end/back color '#RRGGBB'. Only used when fill_type='gradient'.",
    )
    fill_gradient_style: Optional[str] = Field(
        default=None,
        description=(
            "Gradient direction. One of: 'horizontal', 'vertical', 'diagonal_up', "
            "'diagonal_down', 'from_corner', 'from_center'. Only used when fill_type='gradient'."
        ),
    )
    fill_transparency: Optional[float] = Field(
        default=None,
        description="Fill transparency: 0.0 = opaque, 1.0 = fully transparent.",
    )
    # --- inline line/border (optional — avoids a separate ppt_set_line call) ---
    line_visible: Optional[bool] = Field(
        default=None,
        description="Border visibility. Set to false to remove the default border (recommended for most card/box shapes).",
    )
    line_color: Optional[str] = Field(
        default=None,
        description="Border color '#RRGGBB'. Implies line_visible=true if line_visible is not specified.",
    )
    line_weight: Optional[float] = Field(
        default=None,
        description="Border weight in points.",
    )


class AddTextboxInput(BaseModel):
    """Input for adding a text box to a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    left: float = Field(..., description="Left position in points")
    top: float = Field(..., description="Top position in points")
    width: float = Field(..., description="Width in points")
    height: float = Field(..., description="Height in points")
    text: Optional[str] = Field(default=None, description="Optional initial text content")
    # --- inline font (optional — avoids a separate ppt_format_text call) ---
    font_name: Optional[str] = Field(
        default=None,
        description=(
            "Font name applied to all text. Sets both the Latin font (Name) and the East Asian "
            "font (NameFarEast) — same behaviour as ppt_format_text."
        ),
    )
    font_size: Optional[float] = Field(default=None, description="Font size in points.")
    bold: Optional[bool] = Field(default=None, description="Bold on/off.")
    italic: Optional[bool] = Field(default=None, description="Italic on/off.")
    font_color: Optional[str] = Field(default=None, description="Text color '#RRGGBB'.")
    align: Optional[str] = Field(
        default=None,
        description="Paragraph alignment for all text: 'left', 'center', 'right', or 'justify'.",
    )


class AddPictureInput(BaseModel):
    """Input for adding an image to a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    file_path: str = Field(..., description="Path to image file")
    left: float = Field(..., description="Left position in points")
    top: float = Field(..., description="Top position in points")
    width: Optional[float] = Field(default=None, description="Width in points (auto-scale if not provided)")
    height: Optional[float] = Field(default=None, description="Height in points (auto-scale if not provided)")


class AddLineInput(BaseModel):
    """Input for adding a line to a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    begin_x: float = Field(..., description="Start X position in points")
    begin_y: float = Field(..., description="Start Y position in points")
    end_x: float = Field(..., description="End X position in points")
    end_y: float = Field(..., description="End Y position in points")


class ListShapesInput(BaseModel):
    """Input for listing shapes on a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")


class ShapeIdentifierInput(BaseModel):
    """Input for identifying a shape by name or index."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name: Optional[str] = Field(default=None, description="Shape name (preferred — indices shift when shapes are added/removed)")
    shape_index: Optional[int] = Field(default=None, ge=1, description="1-based shape index (unstable — prefer shape_name)")


class UpdateShapeInput(BaseModel):
    """Input for updating shape properties."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name: Optional[str] = Field(default=None, description="Shape name (preferred — indices shift when shapes are added/removed)")
    shape_index: Optional[int] = Field(default=None, ge=1, description="1-based shape index (unstable — prefer shape_name)")
    left: Optional[float] = Field(default=None, description="New left position in points")
    top: Optional[float] = Field(default=None, description="New top position in points")
    width: Optional[float] = Field(default=None, description="New width in points")
    height: Optional[float] = Field(default=None, description="New height in points")
    rotation: Optional[float] = Field(default=None, description="Rotation in degrees (0-360)")
    name: Optional[str] = Field(default=None, description="New name for the shape")


class SetZOrderInput(BaseModel):
    """Input for changing shape z-order."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name: Optional[str] = Field(default=None, description="Shape name (preferred — indices shift when shapes are added/removed)")
    shape_index: Optional[int] = Field(default=None, ge=1, description="1-based shape index (unstable — prefer shape_name)")
    command: str = Field(
        ...,
        description="Z-order command: 'bring_to_front', 'send_to_back', 'bring_forward', 'send_backward'",
    )


# ---------------------------------------------------------------------------
# Helper: find a shape by name or index
# ---------------------------------------------------------------------------
def _get_shape(slide, name_or_index: Union[str, int, None], shape_name: Optional[str] = None, shape_index: Optional[int] = None):
    """Find a shape on a slide by name or 1-based index.

    Accepts either a combined name_or_index parameter or separate
    shape_name/shape_index from Pydantic models.
    """
    if shape_name is not None:
        identifier = shape_name
    elif shape_index is not None:
        identifier = shape_index
    elif name_or_index is not None:
        identifier = name_or_index
    else:
        raise ValueError("Either shape_name or shape_index must be provided.")

    if isinstance(identifier, int):
        if identifier < 1 or identifier > slide.Shapes.Count:
            raise ValueError(
                f"Shape index {identifier} is out of range. "
                f"Slide has {slide.Shapes.Count} shapes (1-based)."
            )
        return slide.Shapes(identifier)

    # String name lookup
    for i in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes(i)
        if shape.Name == identifier:
            return shape
    raise ValueError(f"Shape '{identifier}' not found on this slide.")


def _resolve_shape_type(shape_type: Union[int, str]) -> int:
    """Resolve a shape type from int or friendly name string."""
    if isinstance(shape_type, int):
        return shape_type
    key = shape_type.strip().lower().replace(" ", "_").replace("-", "_")
    if key in SHAPE_NAME_MAP:
        return SHAPE_NAME_MAP[key]
    raise ValueError(
        f"Unknown shape type '{shape_type}'. "
        f"Use an integer MsoAutoShapeType or one of: {', '.join(sorted(SHAPE_NAME_MAP.keys()))}"
    )


# ---------------------------------------------------------------------------
# COM implementation functions (run on COM thread via ppt.execute)
# ---------------------------------------------------------------------------
def _add_shape_impl(
    slide_index, shape_type_int, left, top, width, height, text,
    fill_color, fill_type, fill_color2, fill_gradient_style, fill_transparency,
    line_visible, line_color, line_weight,
):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = slide.Shapes.AddShape(
        Type=shape_type_int, Left=left, Top=top, Width=width, Height=height
    )
    if text:
        shape.TextFrame.TextRange.Text = text

    # Inline fill — avoids a follow-up ppt_set_fill call
    _VALID_FILL_TYPES = {"solid", "none", "gradient"}
    if fill_type is not None and fill_type not in _VALID_FILL_TYPES:
        raise ValueError(f"Invalid fill_type '{fill_type}'. Must be one of: {sorted(_VALID_FILL_TYPES)}")
    if fill_color is not None or fill_type is not None or fill_transparency is not None:
        effective_type = fill_type or ("solid" if fill_color is not None else None)
        fill = shape.Fill
        if effective_type == "none":
            fill.Visible = msoFalse
        elif effective_type == "gradient":
            gstyle = GRADIENT_STYLE_MAP.get(fill_gradient_style or "horizontal", 1)
            fill.TwoColorGradient(Style=gstyle, Variant=1)
            if fill_color is not None:
                fill.ForeColor.RGB = hex_to_int(fill_color)
            if fill_color2 is not None:
                fill.BackColor.RGB = hex_to_int(fill_color2)
        elif effective_type == "solid":
            fill.Solid()
            if fill_color is not None:
                fill.ForeColor.RGB = hex_to_int(fill_color)
        if fill_transparency is not None and effective_type != "none":
            fill.Transparency = fill_transparency

    # Inline line/border — avoids a follow-up ppt_set_line call
    if line_visible is not None:
        shape.Line.Visible = msoTrue if line_visible else msoFalse
    if line_color is not None:
        shape.Line.ForeColor.RGB = hex_to_int(line_color)
    if line_weight is not None:
        shape.Line.Weight = line_weight

    return {
        "success": True,
        "shape_name": shape.Name,
        "shape_index": shape.ZOrderPosition,
        "shape_type": shape.AutoShapeType,
    }


def _add_textbox_impl(
    slide_index, left, top, width, height, text,
    font_name, font_size, bold, italic, font_color, align,
):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    textbox = slide.Shapes.AddTextbox(
        Orientation=msoTextOrientationHorizontal,
        Left=left, Top=top, Width=width, Height=height,
    )
    if text:
        textbox.TextFrame.TextRange.Text = text

    # Inline font — avoids a follow-up ppt_format_text call
    if any(x is not None for x in [font_name, font_size, bold, italic, font_color]):
        font = textbox.TextFrame.TextRange.Font
        if font_name is not None:
            font.Name = font_name
            font.NameFarEast = font_name  # East Asian characters (e.g. Japanese)
        if font_size is not None:
            font.Size = font_size
        if bold is not None:
            font.Bold = msoTrue if bold else msoFalse
        if italic is not None:
            font.Italic = msoTrue if italic else msoFalse
        if font_color is not None:
            font.Color.RGB = hex_to_int(font_color)

    # Inline alignment — avoids a follow-up ppt_set_paragraph_format call
    if align is not None:
        _ALIGN = {"left": 1, "center": 2, "right": 3, "justify": 4}
        align_val = _ALIGN.get(align.lower())
        if align_val is None:
            raise ValueError(f"Invalid align '{align}'. Must be one of: {sorted(_ALIGN)}")
        textbox.TextFrame.TextRange.ParagraphFormat.Alignment = align_val

    return {
        "success": True,
        "shape_name": textbox.Name,
        "shape_index": textbox.ZOrderPosition,
    }


def _add_picture_impl(slide_index, file_path, left, top, width, height):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    w = width if width is not None else -1
    h = height if height is not None else -1
    pic = slide.Shapes.AddPicture(
        FileName=file_path,
        LinkToFile=msoFalse,
        SaveWithDocument=msoTrue,
        Left=left, Top=top, Width=w, Height=h,
    )
    return {
        "success": True,
        "shape_name": pic.Name,
        "shape_index": pic.ZOrderPosition,
        "width": round(pic.Width, 2),
        "height": round(pic.Height, 2),
    }


def _add_line_impl(slide_index, begin_x, begin_y, end_x, end_y):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    line = slide.Shapes.AddLine(
        BeginX=begin_x, BeginY=begin_y, EndX=end_x, EndY=end_y,
    )
    return {
        "success": True,
        "shape_name": line.Name,
        "shape_index": line.ZOrderPosition,
    }


def _list_shapes_impl(slide_index):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shapes = []
    for i in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes(i)
        has_text = False
        text_preview = ""
        try:
            if shape.HasTextFrame:
                has_text = True
                if shape.TextFrame.HasText:
                    full_text = shape.TextFrame.TextRange.Text
                    text_preview = full_text[:50] + ("..." if len(full_text) > 50 else "")
        except Exception:
            pass

        shapes.append({
            "index": i,
            "name": shape.Name,
            "id": shape.Id,
            "type": shape.Type,
            "type_name": SHAPE_TYPE_NAMES.get(shape.Type, f"Unknown({shape.Type})"),
            "left": round(shape.Left, 2),
            "top": round(shape.Top, 2),
            "width": round(shape.Width, 2),
            "height": round(shape.Height, 2),
            "has_text": has_text,
            "text_preview": text_preview,
        })
    return {
        "slide_index": slide_index,
        "shapes_count": slide.Shapes.Count,
        "shapes": shapes,
    }


def _get_shape_info_impl(slide_index, shape_name, shape_index):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, None, shape_name=shape_name, shape_index=shape_index)

    info = {
        "name": shape.Name,
        "id": shape.Id,
        "type": shape.Type,
        "type_name": SHAPE_TYPE_NAMES.get(shape.Type, f"Unknown({shape.Type})"),
        "left": round(shape.Left, 2),
        "top": round(shape.Top, 2),
        "width": round(shape.Width, 2),
        "height": round(shape.Height, 2),
        "rotation": round(shape.Rotation, 2),
        "z_order": shape.ZOrderPosition,
        "is_group": shape.Type == msoGroup,
        "has_animation": False,
        "aspect_ratio_locked": False,
        "text": None,
        "fill": None,
        "line": None,
    }

    # Animation check
    try:
        seq = slide.TimeLine.MainSequence
        for i in range(1, seq.Count + 1):
            if seq(i).Shape.Id == shape.Id:
                info["has_animation"] = True
                break
    except Exception:
        pass

    # Aspect ratio lock
    try:
        info["aspect_ratio_locked"] = shape.LockAspectRatio == msoTrue
    except Exception:
        pass

    # Text content
    try:
        if shape.HasTextFrame and shape.TextFrame.HasText:
            info["text"] = shape.TextFrame.TextRange.Text
    except Exception:
        pass

    # Fill info
    try:
        fill = shape.Fill
        info["fill"] = {
            "type": fill.Type,
            "visible": bool(fill.Visible),
        }
        try:
            info["fill"]["fore_color_rgb"] = fill.ForeColor.RGB
        except Exception:
            pass
        try:
            info["fill"]["transparency"] = round(fill.Transparency, 2)
        except Exception:
            pass
    except Exception:
        pass

    # Line info
    try:
        line = shape.Line
        info["line"] = {
            "visible": bool(line.Visible),
        }
        try:
            info["line"]["weight"] = round(line.Weight, 2)
        except Exception:
            pass
        try:
            info["line"]["fore_color_rgb"] = line.ForeColor.RGB
        except Exception:
            pass
        try:
            info["line"]["dash_style"] = line.DashStyle
        except Exception:
            pass
    except Exception:
        pass

    return info


def _update_shape_impl(slide_index, shape_name, shape_index, left, top, width, height, rotation, name):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, None, shape_name=shape_name, shape_index=shape_index)

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
    if name is not None:
        shape.Name = name

    return {
        "success": True,
        "shape_name": shape.Name,
        "left": round(shape.Left, 2),
        "top": round(shape.Top, 2),
        "width": round(shape.Width, 2),
        "height": round(shape.Height, 2),
    }


def _delete_shape_impl(slide_index, shape_name, shape_index):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, None, shape_name=shape_name, shape_index=shape_index)
    deleted_name = shape.Name
    shape.Delete()
    return {"success": True, "deleted": deleted_name}


def _duplicate_shape_impl(slide_index, shape_name, shape_index):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, None, shape_name=shape_name, shape_index=shape_index)
    dup = shape.Duplicate()
    new_shape = dup(1)
    new_shape.Left = shape.Left + 20
    new_shape.Top = shape.Top + 20
    return {
        "success": True,
        "new_shape_name": new_shape.Name,
        "new_shape_index": new_shape.ZOrderPosition,
    }


def _set_zorder_impl(slide_index, shape_name, shape_index, z_order_cmd):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, None, shape_name=shape_name, shape_index=shape_index)
    shape.ZOrder(z_order_cmd)
    return {"success": True, "shape_name": shape.Name, "new_z_position": shape.ZOrderPosition}


# ---------------------------------------------------------------------------
# MCP tool functions
# ---------------------------------------------------------------------------
def add_shape(params: AddShapeInput) -> str:
    """Add an auto shape to a slide.

    Supports rectangles, ovals, arrows, stars, flowchart shapes, and more.
    Use a friendly name like 'rectangle' or an MsoAutoShapeType integer.

    Args:
        params: Shape parameters including type, position, and size in points.

    Returns:
        JSON with shape name, index, and type of the created shape.
    """
    try:
        shape_type_int = _resolve_shape_type(params.shape_type)
        result = ppt.execute(
            _add_shape_impl,
            params.slide_index, shape_type_int,
            params.left, params.top, params.width, params.height,
            params.text,
            params.fill_color, params.fill_type, params.fill_color2,
            params.fill_gradient_style, params.fill_transparency,
            params.line_visible, params.line_color, params.line_weight,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to add shape: {str(e)}"})


def add_textbox(params: AddTextboxInput) -> str:
    """Add a text box to a slide.

    Creates a horizontal text box at the specified position and size.

    Args:
        params: Textbox parameters including position, size, and optional text.

    Returns:
        JSON with shape name and index of the created text box.
    """
    try:
        result = ppt.execute(
            _add_textbox_impl,
            params.slide_index,
            params.left, params.top, params.width, params.height,
            params.text,
            params.font_name, params.font_size, params.bold,
            params.italic, params.font_color, params.align,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to add textbox: {str(e)}"})


def add_picture(params: AddPictureInput) -> str:
    """Add an image from a file path to a slide.

    The image is embedded in the presentation. If width/height are not
    provided, the original image dimensions are used.

    Args:
        params: Picture parameters including file path, position, and optional size.

    Returns:
        JSON with shape name, index, and actual dimensions of the inserted image.
    """
    try:
        result = ppt.execute(
            _add_picture_impl,
            params.slide_index, params.file_path,
            params.left, params.top, params.width, params.height,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to add picture: {str(e)}"})


def add_line(params: AddLineInput) -> str:
    """Add a line to a slide.

    Creates a straight line from the begin point to the end point.

    Args:
        params: Line parameters including start and end coordinates in points.

    Returns:
        JSON with shape name and index of the created line.
    """
    try:
        result = ppt.execute(
            _add_line_impl,
            params.slide_index,
            params.begin_x, params.begin_y, params.end_x, params.end_y,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to add line: {str(e)}"})


def list_shapes(params: ListShapesInput) -> str:
    """List all shapes on a slide.

    Returns an array of shapes with their name, id, type, position, size,
    and a text preview (first 50 characters) for shapes that contain text.
    The index field reflects z-order (stacking order): index 1 is the
    backmost shape, the highest index is the frontmost shape.

    Args:
        params: Slide index to list shapes from.

    Returns:
        JSON with shapes count and array of shape info objects.
    """
    try:
        result = ppt.execute(_list_shapes_impl, params.slide_index)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to list shapes: {str(e)}"})


def get_shape_info(params: ShapeIdentifierInput) -> str:
    """Get detailed information about a specific shape.

    Returns name, id, type, position, size, rotation, z-order, full text
    content, fill info, line info, and metadata: is_group (True if this
    shape is a group container), has_animation (True if the shape has any
    animation in the main sequence), aspect_ratio_locked.

    Args:
        params: Slide index and shape identifier (name or index).

    Returns:
        JSON with detailed shape properties.
    """
    try:
        result = ppt.execute(
            _get_shape_info_impl,
            params.slide_index, params.shape_name, params.shape_index,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to get shape info: {str(e)}"})


def update_shape(params: UpdateShapeInput) -> str:
    """Update properties of an existing shape.

    Only updates properties that are provided (not None). Can change
    position, size, rotation, and name.

    Args:
        params: Shape identifier and properties to update.

    Returns:
        JSON with updated shape name and current position/size.
    """
    try:
        result = ppt.execute(
            _update_shape_impl,
            params.slide_index, params.shape_name, params.shape_index,
            params.left, params.top, params.width, params.height,
            params.rotation, params.name,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to update shape: {str(e)}"})


def delete_shape(params: ShapeIdentifierInput) -> str:
    """Delete a shape from a slide.

    Args:
        params: Slide index and shape identifier (name or index).

    Returns:
        JSON confirming the deleted shape name.
    """
    try:
        result = ppt.execute(
            _delete_shape_impl,
            params.slide_index, params.shape_name, params.shape_index,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to delete shape: {str(e)}"})


def duplicate_shape(params: ShapeIdentifierInput) -> str:
    """Duplicate a shape on the same slide.

    The duplicate is offset 20 points right and down from the original.

    Args:
        params: Slide index and shape identifier (name or index).

    Returns:
        JSON with the new duplicated shape's name and index.
    """
    try:
        result = ppt.execute(
            _duplicate_shape_impl,
            params.slide_index, params.shape_name, params.shape_index,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to duplicate shape: {str(e)}"})


def set_shape_zorder(params: SetZOrderInput) -> str:
    """Change the z-order (stacking position) of a shape.

    Commands: 'bring_to_front', 'send_to_back', 'bring_forward', 'send_backward'.

    Args:
        params: Shape identifier and z-order command.

    Returns:
        JSON with shape name and new z-order position.
    """
    try:
        cmd = params.command.strip().lower().replace(" ", "_").replace("-", "_")
        if cmd not in ZORDER_CMD_MAP:
            return json.dumps({
                "error": f"Unknown z-order command '{params.command}'. "
                f"Use one of: {', '.join(ZORDER_CMD_MAP.keys())}"
            })
        result = ppt.execute(
            _set_zorder_impl,
            params.slide_index, params.shape_name, params.shape_index,
            ZORDER_CMD_MAP[cmd],
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to set z-order: {str(e)}"})


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register all shape tools with the MCP server."""

    @mcp.tool(
        name="ppt_add_shape",
        annotations={
            "title": "Add Shape",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_add_shape(params: AddShapeInput) -> str:
        """Add an auto shape to a slide (rectangle, oval, arrow, star, etc.).

        Specify shape_type as a friendly name ('rectangle', 'oval', 'right_arrow',
        'star_5point', 'cloud', etc.) or an MsoAutoShapeType integer.
        All positions and sizes are in points (72 points = 1 inch).

        Optionally apply fill and border in the same call via fill_color, fill_type,
        fill_transparency, line_visible, line_color, and line_weight — avoids separate
        ppt_set_fill / ppt_set_line calls for common cases.
        Example: fill_color='#1E3A5F', line_visible=false creates a styled shape in one step.
        """
        return add_shape(params)

    @mcp.tool(
        name="ppt_add_textbox",
        annotations={
            "title": "Add Text Box",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_add_textbox(params: AddTextboxInput) -> str:
        """Add a text box to a slide.

        Creates a horizontal text box. Optionally set initial text content.
        All positions and sizes are in points (72 points = 1 inch).

        Optionally apply font styling in the same call via font_name, font_size, bold,
        italic, font_color, and align — avoids a separate ppt_format_text call.
        Example: text='Title', font_name='Segoe UI', font_size=32, bold=true,
        font_color='#FFFFFF', align='center' creates a fully styled label in one step.
        """
        return add_textbox(params)

    @mcp.tool(
        name="ppt_add_picture",
        annotations={
            "title": "Add Picture",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_add_picture(params: AddPictureInput) -> str:
        """Add an image from a file path to a slide.

        The image is embedded in the presentation. If width and height are
        omitted, the original image dimensions are preserved.
        """
        return add_picture(params)

    @mcp.tool(
        name="ppt_add_line",
        annotations={
            "title": "Add Line",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_add_line(params: AddLineInput) -> str:
        """Add a straight line to a slide.

        Draws a line from (begin_x, begin_y) to (end_x, end_y).
        All coordinates are in points (72 points = 1 inch).
        """
        return add_line(params)

    @mcp.tool(
        name="ppt_list_shapes",
        annotations={
            "title": "List Shapes",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_list_shapes(params: ListShapesInput) -> str:
        """List all shapes on a slide.

        Returns name, id, type, position, size, and text preview for each shape.
        """
        return list_shapes(params)

    @mcp.tool(
        name="ppt_get_shape_info",
        annotations={
            "title": "Get Shape Info",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_get_shape_info(params: ShapeIdentifierInput) -> str:
        """Get detailed information about a specific shape.

        Identify the shape by name (shape_name) or 1-based index (shape_index).
        Returns full text, fill info, line info, rotation, and z-order.
        """
        return get_shape_info(params)

    @mcp.tool(
        name="ppt_update_shape",
        annotations={
            "title": "Update Shape",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_update_shape(params: UpdateShapeInput) -> str:
        """Update properties of an existing shape.

        Identify the shape by name or index. Only provided properties are updated.
        Can change position (left, top), size (width, height), rotation, and name.
        """
        return update_shape(params)

    @mcp.tool(
        name="ppt_delete_shape",
        annotations={
            "title": "Delete Shape",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_delete_shape(params: ShapeIdentifierInput) -> str:
        """Delete a shape from a slide.

        Identify the shape by name (shape_name) or 1-based index (shape_index).
        This action cannot be undone via MCP (use PowerPoint's Ctrl+Z).
        """
        return delete_shape(params)

    @mcp.tool(
        name="ppt_duplicate_shape",
        annotations={
            "title": "Duplicate Shape",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_duplicate_shape(params: ShapeIdentifierInput) -> str:
        """Duplicate a shape on the same slide.

        Creates a copy offset 20 points right and down from the original.
        Returns the new shape's name and index.
        """
        return duplicate_shape(params)

    @mcp.tool(
        name="ppt_set_shape_zorder",
        annotations={
            "title": "Set Shape Z-Order",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_set_shape_zorder(params: SetZOrderInput) -> str:
        """Change the z-order (stacking position) of a shape.

        Commands: 'bring_to_front', 'send_to_back', 'bring_forward', 'send_backward'.
        Identify the shape by name (shape_name) or 1-based index (shape_index).
        """
        return set_shape_zorder(params)
