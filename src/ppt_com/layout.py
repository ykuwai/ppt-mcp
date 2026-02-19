"""Layout tools for PowerPoint COM automation.

Handles shape alignment, distribution, flipping, merging,
slide size, and slide background configuration.
"""

import json
import logging
import os
from typing import Optional, Union

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from utils.navigation import goto_slide
from utils.color import hex_to_int
from ppt_com.constants import (
    msoTrue, msoFalse,
    ALIGN_CMD_MAP, DISTRIBUTE_CMD_MAP, FLIP_CMD_MAP,
    MERGE_CMD_MAP, SLIDE_SIZE_MAP, GRADIENT_STYLE_MAP,
    VIEW_TYPE_NAMES,
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class AlignShapesInput(BaseModel):
    """Input for aligning shapes on a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_names: list[str] = Field(
        ..., min_length=2,
        description="List of shape names to align (minimum 2)",
    )
    align_to: str = Field(
        ...,
        description="Alignment direction: 'left', 'center', 'right', 'top', 'middle', or 'bottom'",
    )
    relative_to_slide: bool = Field(
        default=False,
        description="If true, align relative to the slide; otherwise align relative to each other",
    )


class DistributeShapesInput(BaseModel):
    """Input for distributing shapes evenly on a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_names: list[str] = Field(
        ..., min_length=3,
        description="List of shape names to distribute (minimum 3)",
    )
    direction: str = Field(
        ...,
        description="Distribution direction: 'horizontal' or 'vertical'",
    )
    relative_to_slide: bool = Field(
        default=False,
        description="If true, distribute relative to the slide edges; otherwise between outermost shapes",
    )


class GetSlideSizeInput(BaseModel):
    """Input for getting slide size (no parameters needed)."""
    model_config = ConfigDict(str_strip_whitespace=True)


class SetSlideSizeInput(BaseModel):
    """Input for setting slide size."""
    model_config = ConfigDict(str_strip_whitespace=True)

    width: Optional[float] = Field(
        default=None,
        description="Slide width in points (72 points = 1 inch)",
    )
    height: Optional[float] = Field(
        default=None,
        description="Slide height in points (72 points = 1 inch)",
    )
    preset: Optional[str] = Field(
        default=None,
        description="Preset size: '16:9', '4:3', 'a4', 'letter', 'widescreen', '16:10', '35mm', 'overhead', 'banner'",
    )
    orientation: Optional[str] = Field(
        default=None,
        description="Orientation: 'landscape' or 'portrait'",
    )


class SetSlideBackgroundInput(BaseModel):
    """Input for setting a slide's background."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    fill_type: str = Field(
        ...,
        description="Fill type: 'solid', 'gradient', 'picture', 'none', or 'master'",
    )
    color: Optional[str] = Field(
        default=None,
        description="Solid fill color as '#RRGGBB'",
    )
    gradient_color1: Optional[str] = Field(
        default=None,
        description="First gradient color as '#RRGGBB'",
    )
    gradient_color2: Optional[str] = Field(
        default=None,
        description="Second gradient color as '#RRGGBB'",
    )
    gradient_style: Optional[str] = Field(
        default=None,
        description="Gradient style: 'horizontal', 'vertical', 'diagonal_up', 'diagonal_down', 'from_corner', 'from_title', 'from_center'",
    )
    image_path: Optional[str] = Field(
        default=None,
        description="Absolute path to image file for picture fill",
    )
    transparency: Optional[float] = Field(
        default=None, ge=0.0, le=1.0,
        description="Fill transparency (0 = opaque, 1 = fully transparent)",
    )


class FlipShapeInput(BaseModel):
    """Input for flipping a shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    direction: str = Field(
        ...,
        description="Flip direction: 'horizontal' or 'vertical'",
    )


class MergeShapesInput(BaseModel):
    """Input for merging shapes using Boolean operations."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_names: list[str] = Field(
        ..., min_length=2,
        description="List of shape names to merge (minimum 2)",
    )
    merge_type: str = Field(
        ...,
        description="Merge type: 'union', 'combine', 'intersect', 'subtract', or 'fragment'",
    )
    primary_shape: Optional[str] = Field(
        default=None,
        description="Name of the primary shape (determines formatting of result). If omitted, the first shape in the range is used.",
    )


# ---------------------------------------------------------------------------
# Helper: find a shape by name or index
# ---------------------------------------------------------------------------
def _get_shape(slide, name_or_index: Union[str, int]):
    """Find a shape on a slide by name (str) or 1-based index (int).

    Args:
        slide: Slide COM object
        name_or_index: Shape name (str) or 1-based index (int)

    Returns:
        Shape COM object

    Raises:
        ValueError: If shape not found or index out of range
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
            if slide.Shapes(i).Name == name_or_index:
                return slide.Shapes(i)
        raise ValueError(f"Shape '{name_or_index}' not found on slide")


# ---------------------------------------------------------------------------
# COM implementation functions (run on COM thread via ppt.execute)
# ---------------------------------------------------------------------------
def _align_shapes_impl(slide_index, shape_names, align_to, relative_to_slide):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    # Validate align_to
    align_key = align_to.strip().lower()
    align_cmd = ALIGN_CMD_MAP.get(align_key)
    if align_cmd is None:
        raise ValueError(
            f"Unknown align_to '{align_to}'. "
            f"Valid values: {list(ALIGN_CMD_MAP.keys())}"
        )

    # Validate all shape names exist
    for name in shape_names:
        found = False
        for i in range(1, slide.Shapes.Count + 1):
            if slide.Shapes(i).Name == name:
                found = True
                break
        if not found:
            raise ValueError(f"Shape '{name}' not found on slide {slide_index}")

    relative = msoTrue if relative_to_slide else msoFalse
    shape_range = slide.Shapes.Range(tuple(shape_names))
    shape_range.Align(align_cmd, relative)

    return {
        "success": True,
        "aligned_count": len(shape_names),
        "align_to": align_key,
        "relative_to_slide": relative_to_slide,
    }


def _distribute_shapes_impl(slide_index, shape_names, direction, relative_to_slide):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    # Validate direction
    dir_key = direction.strip().lower()
    dist_cmd = DISTRIBUTE_CMD_MAP.get(dir_key)
    if dist_cmd is None:
        raise ValueError(
            f"Unknown direction '{direction}'. "
            f"Valid values: {list(DISTRIBUTE_CMD_MAP.keys())}"
        )

    # Validate all shape names exist
    for name in shape_names:
        found = False
        for i in range(1, slide.Shapes.Count + 1):
            if slide.Shapes(i).Name == name:
                found = True
                break
        if not found:
            raise ValueError(f"Shape '{name}' not found on slide {slide_index}")

    relative = msoTrue if relative_to_slide else msoFalse
    shape_range = slide.Shapes.Range(tuple(shape_names))
    shape_range.Distribute(dist_cmd, relative)

    return {
        "success": True,
        "distributed_count": len(shape_names),
        "direction": dir_key,
        "relative_to_slide": relative_to_slide,
    }


def _get_slide_size_impl():
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    ps = pres.PageSetup

    width_pt = ps.SlideWidth
    height_pt = ps.SlideHeight
    slide_size = ps.SlideSize
    orientation = ps.SlideOrientation

    # Reverse-lookup the preset name
    preset_name = None
    for name, val in SLIDE_SIZE_MAP.items():
        if val == slide_size:
            preset_name = name
            break

    return {
        "success": True,
        "width_points": round(width_pt, 2),
        "height_points": round(height_pt, 2),
        "width_inches": round(width_pt / 72.0, 4),
        "height_inches": round(height_pt / 72.0, 4),
        "slide_size_type": slide_size,
        "slide_size_name": preset_name,
        "orientation": "landscape" if orientation == 1 else "portrait",
    }


def _set_slide_size_impl(width, height, preset, orientation):
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    ps = pres.PageSetup

    # Preset dimensions in points (width, height)
    PRESET_DIMENSIONS = {
        "16:9": (960, 540),
        "widescreen": (960, 540),
        "4:3": (960, 720),
        "16:10": (960, 600),
        "a4": (842, 595),
        "a3": (1191, 842),
        "letter": (792, 612),
        "35mm": (792, 528),
        "overhead": (720, 540),
        "banner": (576, 72),
    }

    # Set preset FIRST (may change width/height)
    if preset is not None:
        preset_key = preset.strip().lower()
        dims = PRESET_DIMENSIONS.get(preset_key)
        if dims is None:
            raise ValueError(
                f"Unknown preset '{preset}'. "
                f"Valid values: {list(PRESET_DIMENSIONS.keys())}"
            )
        ps.SlideWidth = dims[0]
        ps.SlideHeight = dims[1]

    # Then set explicit width/height (overrides preset dimensions)
    if width is not None:
        ps.SlideWidth = width
    if height is not None:
        ps.SlideHeight = height

    # Then set orientation
    if orientation is not None:
        orient_key = orientation.strip().lower()
        if orient_key == "landscape":
            ps.SlideOrientation = 1
        elif orient_key == "portrait":
            ps.SlideOrientation = 2
        else:
            raise ValueError(
                f"Unknown orientation '{orientation}'. "
                f"Valid values: 'landscape', 'portrait'"
            )

    # Read back final values
    return {
        "success": True,
        "width_points": round(ps.SlideWidth, 2),
        "height_points": round(ps.SlideHeight, 2),
        "width_inches": round(ps.SlideWidth / 72.0, 4),
        "height_inches": round(ps.SlideHeight / 72.0, 4),
    }


def _set_slide_background_impl(slide_index, fill_type, color,
                                gradient_color1, gradient_color2,
                                gradient_style, image_path, transparency):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    fill_key = fill_type.strip().lower()

    if fill_key == "master":
        slide.FollowMasterBackground = msoTrue
        return {
            "success": True,
            "slide_index": slide_index,
            "fill_type": "master",
        }

    # Detach from master background
    slide.FollowMasterBackground = msoFalse
    fill = slide.Background.Fill

    if fill_key == "solid":
        if color is None:
            raise ValueError("color is required for solid fill")
        fill.Solid()
        fill.ForeColor.RGB = hex_to_int(color)

    elif fill_key == "gradient":
        if gradient_color1 is None or gradient_color2 is None:
            raise ValueError(
                "gradient_color1 and gradient_color2 are required for gradient fill"
            )
        style_key = (gradient_style or "horizontal").strip().lower()
        style_val = GRADIENT_STYLE_MAP.get(style_key)
        if style_val is None:
            raise ValueError(
                f"Unknown gradient_style '{gradient_style}'. "
                f"Valid values: {list(GRADIENT_STYLE_MAP.keys())}"
            )
        fill.TwoColorGradient(style_val, 1)
        fill.ForeColor.RGB = hex_to_int(gradient_color1)
        fill.BackColor.RGB = hex_to_int(gradient_color2)

    elif fill_key == "picture":
        if image_path is None:
            raise ValueError("image_path is required for picture fill")
        abs_path = os.path.abspath(image_path)
        if not os.path.isfile(abs_path):
            raise ValueError(f"Image file not found: {abs_path}")
        fill.UserPicture(abs_path)

    elif fill_key == "none":
        fill.Background()

    else:
        raise ValueError(
            f"Unknown fill_type '{fill_type}'. "
            f"Valid values: 'solid', 'gradient', 'picture', 'none', 'master'"
        )

    # Apply transparency if specified
    if transparency is not None and fill_key not in ("none", "master"):
        fill.Transparency = transparency

    return {
        "success": True,
        "slide_index": slide_index,
        "fill_type": fill_key,
    }


def _flip_shape_impl(slide_index, shape_name_or_index, direction):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    dir_key = direction.strip().lower()
    flip_cmd = FLIP_CMD_MAP.get(dir_key)
    if flip_cmd is None:
        raise ValueError(
            f"Unknown direction '{direction}'. "
            f"Valid values: {list(FLIP_CMD_MAP.keys())}"
        )

    shape.Flip(flip_cmd)

    # Read back flip state
    h_flip = shape.HorizontalFlip
    v_flip = shape.VerticalFlip

    return {
        "success": True,
        "shape_name": shape.Name,
        "horizontal_flip": bool(h_flip),
        "vertical_flip": bool(v_flip),
    }


def _merge_shapes_impl(slide_index, shape_names, merge_type, primary_shape):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    # Validate merge type
    merge_key = merge_type.strip().lower()
    merge_cmd = MERGE_CMD_MAP.get(merge_key)
    if merge_cmd is None:
        raise ValueError(
            f"Unknown merge_type '{merge_type}'. "
            f"Valid values: {list(MERGE_CMD_MAP.keys())}"
        )

    # Validate all shape names exist
    for name in shape_names:
        found = False
        for i in range(1, slide.Shapes.Count + 1):
            if slide.Shapes(i).Name == name:
                found = True
                break
        if not found:
            raise ValueError(f"Shape '{name}' not found on slide {slide_index}")

    shape_range = slide.Shapes.Range(tuple(shape_names))

    # MergeShapes always requires a primary shape reference.
    # If not specified, use the first shape in the list.
    if primary_shape is not None:
        primary = _get_shape(slide, primary_shape)
    else:
        primary = _get_shape(slide, shape_names[0])
    shape_range.MergeShapes(merge_cmd, primary)

    return {
        "success": True,
        "merge_type": merge_key,
        "merged_count": len(shape_names),
    }


# ---------------------------------------------------------------------------
# MCP tool functions (sync wrappers that delegate to COM thread)
# ---------------------------------------------------------------------------
def align_shapes(params: AlignShapesInput) -> str:
    """Align multiple shapes on a slide.

    Args:
        params: Slide index, shape names, alignment direction, and relative flag.

    Returns:
        JSON confirming the alignment operation.
    """
    try:
        result = ppt.execute(
            _align_shapes_impl,
            params.slide_index, params.shape_names,
            params.align_to, params.relative_to_slide,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to align shapes: {str(e)}"})


def distribute_shapes(params: DistributeShapesInput) -> str:
    """Distribute shapes evenly on a slide.

    Args:
        params: Slide index, shape names, direction, and relative flag.

    Returns:
        JSON confirming the distribution operation.
    """
    try:
        result = ppt.execute(
            _distribute_shapes_impl,
            params.slide_index, params.shape_names,
            params.direction, params.relative_to_slide,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to distribute shapes: {str(e)}"})


def get_slide_size(params: GetSlideSizeInput) -> str:
    """Get the current slide size of the active presentation.

    Args:
        params: No parameters needed.

    Returns:
        JSON with slide dimensions in points and inches.
    """
    try:
        result = ppt.execute(_get_slide_size_impl)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to get slide size: {str(e)}"})


def set_slide_size(params: SetSlideSizeInput) -> str:
    """Set the slide size of the active presentation.

    Args:
        params: Optional preset, width, height, and orientation.

    Returns:
        JSON with the resulting slide dimensions.
    """
    try:
        result = ppt.execute(
            _set_slide_size_impl,
            params.width, params.height,
            params.preset, params.orientation,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to set slide size: {str(e)}"})


def set_slide_background(params: SetSlideBackgroundInput) -> str:
    """Set the background of a specific slide.

    Args:
        params: Slide index, fill type, and fill-specific options.

    Returns:
        JSON confirming the background change.
    """
    try:
        result = ppt.execute(
            _set_slide_background_impl,
            params.slide_index, params.fill_type,
            params.color, params.gradient_color1,
            params.gradient_color2, params.gradient_style,
            params.image_path, params.transparency,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to set slide background: {str(e)}"})


def flip_shape(params: FlipShapeInput) -> str:
    """Flip a shape horizontally or vertically.

    Args:
        params: Slide index, shape identifier, and flip direction.

    Returns:
        JSON with the resulting flip state.
    """
    try:
        result = ppt.execute(
            _flip_shape_impl,
            params.slide_index, params.shape_name_or_index,
            params.direction,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to flip shape: {str(e)}"})


def merge_shapes(params: MergeShapesInput) -> str:
    """Merge shapes using a Boolean operation.

    Args:
        params: Slide index, shape names, merge type, and optional primary shape.

    Returns:
        JSON confirming the merge operation.
    """
    try:
        result = ppt.execute(
            _merge_shapes_impl,
            params.slide_index, params.shape_names,
            params.merge_type, params.primary_shape,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to merge shapes: {str(e)}"})


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register all layout tools with the MCP server."""

    @mcp.tool(
        name="ppt_align_shapes",
        annotations={
            "title": "Align Shapes",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_align_shapes(params: AlignShapesInput) -> str:
        """Align multiple shapes on a slide.

        Aligns shapes to a common edge or center.
        Provide at least 2 shape names.
        Set relative_to_slide=true to align relative to the slide boundaries.
        Align options: left, center, right, top, middle, bottom.
        """
        return align_shapes(params)

    @mcp.tool(
        name="ppt_distribute_shapes",
        annotations={
            "title": "Distribute Shapes",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_distribute_shapes(params: DistributeShapesInput) -> str:
        """Distribute shapes evenly on a slide.

        Spaces shapes evenly either horizontally or vertically.
        Provide at least 3 shape names.
        Set relative_to_slide=true to distribute relative to the slide edges.
        """
        return distribute_shapes(params)

    @mcp.tool(
        name="ppt_get_slide_size",
        annotations={
            "title": "Get Slide Size",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_get_slide_size(params: GetSlideSizeInput) -> str:
        """Get the current slide size of the active presentation.

        Returns width and height in both points and inches,
        along with the slide size preset name and orientation.
        """
        return get_slide_size(params)

    @mcp.tool(
        name="ppt_set_slide_size",
        annotations={
            "title": "Set Slide Size",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_set_slide_size(params: SetSlideSizeInput) -> str:
        """Set the slide size of the active presentation.

        Use preset for standard sizes ('16:9', '4:3', 'a4', 'letter', etc.)
        or specify exact width/height in points (72 points = 1 inch).
        Preset is applied first, then width/height, then orientation.
        """
        return set_slide_size(params)

    @mcp.tool(
        name="ppt_set_slide_background",
        annotations={
            "title": "Set Slide Background",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_set_slide_background(params: SetSlideBackgroundInput) -> str:
        """Set the background of a specific slide.

        Fill types: 'solid' (requires color), 'gradient' (requires gradient_color1
        and gradient_color2), 'picture' (requires image_path), 'none' (transparent),
        or 'master' (follow master slide background).
        Colors use '#RRGGBB' format.
        """
        return set_slide_background(params)

    @mcp.tool(
        name="ppt_flip_shape",
        annotations={
            "title": "Flip Shape",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_flip_shape(params: FlipShapeInput) -> str:
        """Flip a shape horizontally or vertically.

        Identify the shape by name or 1-based index.
        Direction: 'horizontal' mirrors left-right, 'vertical' mirrors top-bottom.
        Returns the resulting flip state.
        """
        return flip_shape(params)

    @mcp.tool(
        name="ppt_merge_shapes",
        annotations={
            "title": "Merge Shapes",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_merge_shapes(params: MergeShapesInput) -> str:
        """Merge shapes using a Boolean operation.

        Provide at least 2 shape names.
        Merge types: 'union' (combine outlines), 'combine' (XOR), 'intersect'
        (keep overlap), 'subtract' (remove overlap from primary), 'fragment'
        (split at intersections).
        Optionally specify primary_shape to control which shape's formatting is kept.
        """
        return merge_shapes(params)
