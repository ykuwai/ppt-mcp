"""Shape operations for PowerPoint COM automation.

Handles adding, listing, modifying, duplicating, deleting shapes,
and z-order management on PowerPoint slides.
"""

import json
import logging
from typing import Literal, Optional, Union

from pydantic import BaseModel, Field, ConfigDict, model_validator

from utils.offload import run_offloaded
from utils.color import hex_to_int, int_to_hex
from backend import ppt
from utils.navigation import goto_slide
from utils.redraw import FrozenRedraw
from utils.validation import font_size_warning
from ppt_com.shape_lookup import resolve_shape, walk_group_children
from ppt_com.constants import (
    SHAPE_TYPE_NAMES,
    msoTrue, msoFalse, msoTriStateMixed,
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
    "block_arc": 20,
    "double_bracket": 26,
    "double_brace": 27,
    "left_bracket": 29,
    "right_bracket": 30,
    "left_brace": 31,
    "right_brace": 32,
    "striped_right_arrow": 49,
    "notched_right_arrow": 50,
    "rectangular_callout": 105,
    "rounded_rectangular_callout": 106,
    "oval_callout": 107,
    "cloud_callout": 108,
    "frame": 158,
    "half_frame": 159,
    "l_shape": 162,
    "cloud": 179,
}

# ---------------------------------------------------------------------------
# Semantic labels for adjustment handles by MsoAutoShapeType.
# Maps auto_shape_type int → {1-based index: descriptive label}.
# Included in ppt_get_shape_info so AI consumers know what each handle controls.
# ---------------------------------------------------------------------------
ADJUSTMENT_LABELS: dict[int, dict[int, str]] = {
    # Basic shapes
    2: {1: "slant"},                                    # parallelogram
    3: {1: "top_width"},                                # trapezoid
    5: {1: "corner_radius"},                            # rounded_rectangle
    7: {1: "apex_x"},                                   # triangle
    10: {1: "side_width"},                              # hexagon
    11: {1: "arm_thickness"},                           # cross
    13: {1: "lid_height"},                              # can (cylinder)
    14: {1: "depth"},                                   # cube
    17: {1: "mouth_arc"},                               # smiley_face
    18: {1: "ring_thickness"},                          # donut
    20: {1: "start_angle", 2: "end_angle", 3: "thickness"},  # block_arc
    23: {1: "ray_length"},                              # sun
    24: {1: "crescent_width"},                          # moon
    # Brackets and braces
    26: {1: "curve_depth"},                             # double_bracket
    27: {1: "notch_size", 2: "notch_position"},         # double_brace
    29: {1: "curve_depth"},                             # left_bracket
    30: {1: "curve_depth"},                             # right_bracket
    31: {1: "notch_size", 2: "notch_position"},         # left_brace
    32: {1: "notch_size", 2: "notch_position"},         # right_brace
    # Arrows
    33: {1: "head_width", 2: "head_length"},            # right_arrow
    34: {1: "head_width", 2: "head_length"},            # left_arrow
    35: {1: "head_width", 2: "head_length"},            # up_arrow
    36: {1: "head_width", 2: "head_length"},            # down_arrow
    37: {1: "head_width", 2: "head_length"},            # left_right_arrow
    38: {1: "head_width", 2: "head_length"},            # up_down_arrow
    49: {1: "shaft_width", 2: "head_length"},           # striped_right_arrow
    50: {1: "shaft_width", 2: "head_length"},           # notched_right_arrow
    52: {1: "arrow_depth"},                             # chevron
    # Stars
    91: {1: "inner_radius"},                            # star_4point
    92: {1: "inner_radius"},                            # star_5point
    93: {1: "inner_radius"},                            # star_8point
    94: {1: "inner_radius"},                            # star_16point
    95: {1: "inner_radius"},                            # star_24point
    96: {1: "inner_radius"},                            # star_32point
    # Callouts (pointer_y/pointer_x position relative to shape; can exceed 0–1)
    105: {1: "pointer_y", 2: "pointer_x"},                              # rectangular_callout
    106: {1: "pointer_y", 2: "pointer_x", 3: "corner_radius"},         # rounded_rectangular_callout
    107: {1: "pointer_y", 2: "pointer_x"},                              # oval_callout
    108: {1: "pointer_y", 2: "pointer_x"},                              # cloud_callout
    # Structural shapes
    158: {1: "border_thickness"},                                        # frame
    159: {1: "arm_thickness_x", 2: "base_thickness_y"},                  # half_frame
    162: {1: "notch_depth_x", 2: "notch_depth_y"},                       # corner (L-shape)
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
ZORDER_FIELD_DESCRIPTION = (
    "Where the new shape lands in the stack. 'front' (default) is what "
    "PowerPoint does. 'behind_text' puts it directly below the lowest shape "
    "carrying text, which is what art under a caption wants and what 'back' "
    "gets wrong on a deck with a full bleed background. 'back' is the very "
    "bottom."
)


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
    text: Optional[str] = Field(
        default=None,
        description="Optional text content. "
        "\\n = paragraph break (Enter), \\v = line break (Shift+Enter) within the same paragraph.",
    )
    # --- inline text formatting (optional — avoids a separate ppt_format_text call) ---
    font_name: Optional[str] = Field(
        default=None,
        description="Font name applied to shape text. Sets both the Latin font (Name) and the East Asian font (NameFarEast) — same behaviour as ppt_add_textbox.",
    )
    font_size: Optional[float] = Field(
        default=None,
        description="Font size in points.",
    )
    bold: Optional[bool] = Field(
        default=None,
        description="Bold on/off.",
    )
    italic: Optional[bool] = Field(
        default=None,
        description="Italic on/off.",
    )
    font_color: Optional[str] = Field(
        default=None,
        description="Text color '#RRGGBB'.",
    )
    align: Optional[str] = Field(
        default=None,
        description="Paragraph alignment for shape text: 'left', 'center', 'right', or 'justify'.",
    )
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
    corner_radius: Optional[float] = Field(
        default=None,
        ge=0.0,
        le=1.0,
        description="Corner radius for rounded_rectangle shapes as a ratio. "
        "Value range: 0.0 (square corners) to 1.0 (maximum rounding). "
        "Mutually exclusive with corner_radius_pt. Ignored for other shape types.",
    )
    corner_radius_pt: Optional[float] = Field(
        default=None,
        gt=0.0,
        description="Corner radius in points for rounded_rectangle shapes. "
        "Clamped to half the shorter side of the shape. "
        "Mutually exclusive with corner_radius. Ignored for other shape types.",
    )

    zorder: Literal["front", "back", "behind_text"] = Field(
        default="front", description=ZORDER_FIELD_DESCRIPTION
    )

    @model_validator(mode="after")
    def check_corner_radius_exclusivity(self):
        """Ensure corner_radius and corner_radius_pt are mutually exclusive."""
        if self.corner_radius is not None and self.corner_radius_pt is not None:
            raise ValueError(
                "corner_radius and corner_radius_pt are mutually exclusive — "
                "set one or the other, not both"
            )
        return self


class AddTextboxInput(BaseModel):
    """Input for adding a text box to a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    left: float = Field(..., description="Left position in points")
    top: float = Field(..., description="Top position in points")
    width: float = Field(..., description="Width in points")
    height: float = Field(..., description="Height in points")
    text: Optional[str] = Field(
        default=None,
        description="Optional initial text content. "
        "\\n = paragraph break (Enter), \\v = line break (Shift+Enter) within the same paragraph.",
    )
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
    vertical_anchor: Optional[str] = Field(
        default=None,
        description="Vertical text anchor: 'top', 'middle', or 'bottom'.",
    )
    zorder: Literal["front", "back", "behind_text"] = Field(
        default="front", description=ZORDER_FIELD_DESCRIPTION
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
    zorder: Literal["front", "back", "behind_text"] = Field(
        default="front", description=ZORDER_FIELD_DESCRIPTION
    )


class AddLineInput(BaseModel):
    """Input for adding a line to a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    begin_x: float = Field(..., description="Start X position in points")
    begin_y: float = Field(..., description="Start Y position in points")
    end_x: float = Field(..., description="End X position in points")
    end_y: float = Field(..., description="End Y position in points")
    zorder: Literal["front", "back", "behind_text"] = Field(
        default="front", description=ZORDER_FIELD_DESCRIPTION
    )


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
    shape_names: Optional[list[str]] = Field(
        default=None,
        description=(
            "Several shapes to update together, in one call. Use with the "
            "d* offsets to shift a group of shapes by a fixed amount. If any "
            "name does not resolve, nothing moves."
        ),
    )
    all: bool = Field(
        default=False,
        description=(
            "Update every shape on the slide. Pair with exclude to leave the "
            "full bleed background where it is."
        ),
    )
    exclude: Optional[list[str]] = Field(
        default=None,
        description="Names to leave alone. Only with all or shape_names.",
    )
    left: Optional[float] = Field(default=None, description="New left position in points")
    top: Optional[float] = Field(default=None, description="New top position in points")
    width: Optional[float] = Field(default=None, description="New width in points")
    height: Optional[float] = Field(default=None, description="New height in points")
    rotation: Optional[float] = Field(default=None, description="Rotation in degrees (0-360)")
    dleft: Optional[float] = Field(
        default=None,
        description="Move right by this many points, relative to where the shape is now. Negative moves left.",
    )
    dtop: Optional[float] = Field(
        default=None,
        description="Move down by this many points, relative to where the shape is now. Negative moves up, which is the usual one.",
    )
    dwidth: Optional[float] = Field(default=None, description="Widen by this many points, relative to the current width.")
    dheight: Optional[float] = Field(default=None, description="Heighten by this many points, relative to the current height.")
    name: Optional[str] = Field(default=None, description="New name for the shape")
    adjustments: Optional[dict[int, float]] = Field(
        default=None,
        description=(
            "Shape-specific adjustment handle values. Keys are 1-based indices, "
            "values are floats (typically 0.0–1.0, but range varies by shape). "
            "Use ppt_get_shape_info to discover current values, count, and "
            "semantic labels for each handle. "
            "Examples: triangle apex {1: 0.25}, arrow head {1: 0.4, 2: 0.6}, "
            "cross thickness {1: 0.3}, star depth {1: 0.4}, callout pointer {1: 0.1, 2: 0.8}."
        ),
    )

    @model_validator(mode="after")
    def validate_adjustment_keys(self):
        if self.adjustments:
            for k in self.adjustments:
                if k < 1:
                    raise ValueError(
                        f"Adjustment index {k} must be >= 1 (1-based indexing)"
                    )
        return self

    @model_validator(mode="after")
    def validate_selection(self):
        chosen = [
            name for name, value in (
                ("shape_name", self.shape_name),
                ("shape_index", self.shape_index),
                ("shape_names", self.shape_names),
                ("all", self.all or None),
            ) if value is not None
        ]
        if not chosen:
            raise ValueError(
                "Say which shapes to update: shape_name, shape_index, "
                "shape_names, or all=true"
            )
        if len(chosen) > 1:
            raise ValueError(
                f"Use one way of choosing shapes, not {' and '.join(chosen)}"
            )
        if self.shape_names is not None and not self.shape_names:
            raise ValueError("shape_names must not be empty if provided")
        if self.exclude is not None and not (self.all or self.shape_names):
            raise ValueError("exclude only means something with all or shape_names")
        return self

    @model_validator(mode="after")
    def validate_offsets(self):
        for absolute, relative in (("left", "dleft"), ("top", "dtop"),
                                   ("width", "dwidth"), ("height", "dheight")):
            if getattr(self, absolute) is not None and getattr(self, relative) is not None:
                raise ValueError(
                    f"{absolute} and {relative} are mutually exclusive — set "
                    "the position or the offset, not both"
                )
        return self

    @model_validator(mode="after")
    def validate_single_shape_only_fields(self):
        # A rename would make duplicates, and adjustment handles mean
        # different things on different shapes.
        if self.all or self.shape_names:
            for field in ("name", "adjustments"):
                if getattr(self, field) is not None:
                    raise ValueError(
                        f"{field} applies to one shape, so it cannot be used "
                        "with all or shape_names"
                    )
        return self


class SetZOrderInput(BaseModel):
    """Input for changing shape z-order."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name: Optional[str] = Field(default=None, description="Shape name (preferred — indices shift when shapes are added/removed)")
    shape_index: Optional[int] = Field(default=None, ge=1, description="1-based shape index (unstable — prefer shape_name)")
    command: str = Field(
        ...,
        description=(
            "Z-order command: 'bring_to_front', 'send_to_back', "
            "'bring_forward', 'send_backward', or 'send_behind_text' "
            "(directly below the lowest shape carrying text, which is where "
            "art under a caption belongs; 'send_to_back' hides it under a "
            "full bleed background)"
        ),
    )


# ---------------------------------------------------------------------------
# Helper: find a shape by name or index
# ---------------------------------------------------------------------------
def _get_shape(slide, name_or_index: Union[str, int, None], shape_name: Optional[str] = None, shape_index: Optional[int] = None):
    """Find a shape on a slide by name or 1-based index.

    Accepts either a combined name_or_index parameter or separate
    shape_name/shape_index from Pydantic models. The lookup itself is
    ppt_com.shape_lookup.resolve_shape, so this reaches into groups too.
    """
    if shape_name is not None:
        identifier = shape_name
    elif shape_index is not None:
        identifier = shape_index
    elif name_or_index is not None:
        identifier = name_or_index
    else:
        raise ValueError("Either shape_name or shape_index must be provided.")

    return resolve_shape(slide, identifier)


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
# ---------------------------------------------------------------------------
# Where a new shape lands in the stack
# ---------------------------------------------------------------------------
ZORDER_PLACEMENTS = ("front", "back", "behind_text")

# Not an MsoZOrderCmd. PowerPoint has four commands and none of them is this
# one, so it travels as a sentinel the impl branches on.
BEHIND_TEXT = "send_behind_text"


def _carries_text(shape):
    """True when this shape, or anything inside it, has text on it."""
    try:
        if shape.HasTextFrame and shape.TextFrame.HasText:
            return True
    except Exception:
        # A group has no HasTextFrame at all, so the question moves inward.
        pass
    for child, _ in walk_group_children(shape):
        try:
            if child.HasTextFrame and child.TextFrame.HasText:
                return True
        except Exception:
            continue
    return False


def _text_positions(slide, shape):
    positions = []
    for i in range(1, slide.Shapes.Count + 1):
        other = slide.Shapes(i)
        if other.Name == shape.Name:
            continue
        if _carries_text(other):
            positions.append(other.ZOrderPosition)
    return positions


def place_in_zorder(slide, shape, where):
    """Put a freshly added shape where the caller asked for it.

    PowerPoint adds every shape at the front, which is wrong for art that
    belongs under a caption. `back` is not the answer either, because these
    decks usually have a full bleed background at the bottom and sending the
    new picture there hides it completely.

    `behind_text` walks it up from the bottom until it sits directly below the
    lowest shape carrying text. One step at a time, re-reading the positions,
    because the arithmetic for "how many steps" is different depending on
    where the shape started and getting it wrong is silent.

    Returns a dict to merge into the tool's answer, or {} for the default.
    """
    if where in (None, "front"):
        return {}

    if where == "back":
        shape.ZOrder(msoSendToBack)
        return {"zorder": "back", "z_position": shape.ZOrderPosition}

    if where != "behind_text":
        raise ValueError(
            f"Unknown zorder '{where}'. Use one of: "
            f"{', '.join(ZORDER_PLACEMENTS)}"
        )

    if not _text_positions(slide, shape):
        return {
            "zorder": "front",
            "z_position": shape.ZOrderPosition,
            "note": (
                "zorder was behind_text and nothing on this slide has text, "
                "so the shape was left at the front rather than hidden under "
                "the background."
            ),
        }

    shape.ZOrder(msoSendToBack)
    for _ in range(slide.Shapes.Count):
        lowest = min(_text_positions(slide, shape))
        if shape.ZOrderPosition + 1 >= lowest:
            break
        shape.ZOrder(msoBringForward)

    return {"zorder": "behind_text", "z_position": shape.ZOrderPosition}


def _add_shape_impl(
    slide_index, shape_type_int, left, top, width, height, text,
    font_name, font_size, bold, italic, font_color, align,
    fill_color, fill_type, fill_color2, fill_gradient_style, fill_transparency,
    line_visible, line_color, line_weight,
    corner_radius, corner_radius_pt,
    zorder="front",
):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    # Freeze the PowerPoint frame window's painting for the whole operation so
    # the intermediate default (theme accent) fill and any scroll-to-selection
    # are never drawn — only the finished, fully-styled shape is painted, in a
    # single clean repaint on exit. PowerPoint has no working
    # Application.ScreenUpdating, so the freeze is done at the Win32 level
    # (see utils.redraw.FrozenRedraw).
    with FrozenRedraw():
        goto_slide(app, slide_index)
        slide = pres.Slides(slide_index)
        shape = slide.Shapes.AddShape(
            Type=shape_type_int, Left=left, Top=top, Width=width, Height=height,
        )
        result = _apply_shape_attrs(
            shape, text, font_name, font_size, bold, italic, font_color, align,
            fill_color, fill_type, fill_color2, fill_gradient_style, fill_transparency,
            line_visible, line_color, line_weight, corner_radius, corner_radius_pt,
            width, height,
        )
        # Here rather than inside _apply_shape_attrs, which knows about a
        # shape and not about the slide it sits on. shape_index was read in
        # there, before the move, so it is corrected rather than left saying
        # where the shape started.
        placed = place_in_zorder(slide, shape, zorder)
        if "z_position" in placed:
            result["shape_index"] = placed["z_position"]
        result.update(placed)
        return result


def _apply_shape_attrs(
    shape, text, font_name, font_size, bold, italic, font_color, align,
    fill_color, fill_type, fill_color2, fill_gradient_style, fill_transparency,
    line_visible, line_color, line_weight, corner_radius, corner_radius_pt,
    width, height,
):
    if text:
        text = text.replace('\n', '\r')  # \n -> paragraph break (Enter)
        # \v (vertical tab) -> line break (Shift+Enter) — passed through as-is
        shape.TextFrame.TextRange.Text = text

        # Inline text formatting (same pattern as _add_textbox_impl)
        if font_name is not None or font_size is not None or bold is not None \
                or italic is not None or font_color is not None:
            font = shape.TextFrame.TextRange.Font
            if font_name is not None:
                font.Name = font_name
                font.NameFarEast = font_name
            if font_size is not None:
                font.Size = font_size
            if bold is not None:
                font.Bold = msoTrue if bold else msoFalse
            if italic is not None:
                font.Italic = msoTrue if italic else msoFalse
            if font_color is not None:
                font.Color.RGB = hex_to_int(font_color)

        if align is not None:
            _ALIGN = {"left": 1, "center": 2, "right": 3, "justify": 4}
            align_val = _ALIGN.get(align.lower())
            if align_val is None:
                raise ValueError(
                    f"Invalid align '{align}'. Must be one of: {sorted(_ALIGN)}"
                )
            shape.TextFrame.TextRange.ParagraphFormat.Alignment = align_val

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

    # Corner radius for rounded rectangles
    if corner_radius is not None or corner_radius_pt is not None:
        try:
            if shape.AutoShapeType == SHAPE_NAME_MAP["rounded_rectangle"]:
                if corner_radius_pt is not None:
                    # Absolute: convert points to COM ratio, clamp to 0.5
                    short_side = min(width, height)
                    adj_value = min(0.5, corner_radius_pt / short_side)
                else:
                    # Ratio: map user-facing 0.0–1.0 to COM's 0.0–0.5
                    adj_value = corner_radius * 0.5
                shape.Adjustments[1] = adj_value
        except Exception:
            logger.warning("Failed to set corner_radius on shape '%s'", shape.Name)

    return {
        "success": True,
        "shape_name": shape.Name,
        "shape_index": shape.ZOrderPosition,
        "shape_type": shape.AutoShapeType,
    }


def _add_textbox_impl(
    slide_index, left, top, width, height, text,
    font_name, font_size, bold, italic, font_color, align,
    vertical_anchor, zorder="front",
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
        text = text.replace('\n', '\r')  # \n -> paragraph break (Enter)
        # \v (vertical tab) -> line break (Shift+Enter) — passed through as-is
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

    # Inline vertical anchor — avoids a follow-up ppt_set_textframe call
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
        textbox.TextFrame.VerticalAnchor = anchor_val

    placed = place_in_zorder(slide, textbox, zorder)
    return {
        "success": True,
        "shape_name": textbox.Name,
        # After the placement, not before: a dict literal evaluates its
        # entries in order, so reading the position first reports where the
        # shape used to be.
        "shape_index": textbox.ZOrderPosition,
        **placed,
    }


def _add_picture_impl(slide_index, file_path, left, top, width, height,
                      zorder="front"):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    # Insert at natural size first to obtain true aspect ratio.
    # Passing Width/Height directly to AddPicture when only one dimension is
    # specified causes COM to set that dimension but leave the other at its
    # natural value, resulting in a distorted (non-proportional) image.
    pic = slide.Shapes.AddPicture(
        FileName=file_path,
        LinkToFile=msoFalse,
        SaveWithDocument=msoTrue,
        Left=left, Top=top, Width=-1, Height=-1,
    )
    if width is not None and height is not None:
        # Both specified: user intentionally overrides aspect ratio.
        pic.LockAspectRatio = msoFalse
        pic.Width = width
        pic.Height = height
    elif width is not None:
        pic.LockAspectRatio = msoTrue
        pic.Width = width
    elif height is not None:
        pic.LockAspectRatio = msoTrue
        pic.Height = height
    placed = place_in_zorder(slide, pic, zorder)
    return {
        "success": True,
        "shape_name": pic.Name,
        "shape_index": pic.ZOrderPosition,
        "width": round(pic.Width, 2),
        "height": round(pic.Height, 2),
        **placed,
    }


def _add_line_impl(slide_index, begin_x, begin_y, end_x, end_y,
                   zorder="front"):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    line = slide.Shapes.AddLine(
        BeginX=begin_x, BeginY=begin_y, EndX=end_x, EndY=end_y,
    )
    placed = place_in_zorder(slide, line, zorder)
    return {
        "success": True,
        "shape_name": line.Name,
        "shape_index": line.ZOrderPosition,
        **placed,
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


def _text_frame_state(shape):
    """Report the text frame settings that decide how text is drawn.

    ppt_get_text answers with the size a run was set to, and when the frame is
    allowed to shrink text to fit, that is not always the size on the slide.
    Nothing else said the setting was on. The words are the ones
    ppt_set_textframe accepts.

    autofit is the configured mode, not a measurement. AutoSize is all COM
    offers, and a shrink_to_fit box whose text already fits is drawn at its
    full size. Whether the text is being shrunk right now is what
    ppt_check_typography measures, by turning the setting off, reading the
    natural height and putting it back.

    Returns None for a shape with no text frame at all.
    """
    from ppt_com.text import (
        AUTO_SIZE_NAMES, ORIENTATION_NAMES, VERTICAL_ANCHOR_NAMES,
    )

    try:
        if not shape.HasTextFrame:
            return None
    except Exception:
        return None

    state = {
        "autofit": None,
        "word_wrap": None,
        "vertical_anchor": None,
        "orientation": None,
        "margins": None,
    }

    # AutoSize lives on TextFrame2. TextFrame's own AutoSize cannot say
    # shrink_to_fit, which is the one state worth reading.
    try:
        state["autofit"] = AUTO_SIZE_NAMES.get(shape.TextFrame2.AutoSize)
    except Exception:
        pass

    try:
        tf = shape.TextFrame
    except Exception:
        return state

    try:
        wrap = tf.WordWrap
        # Mixed is what a group of paragraphs answers, and it is neither.
        state["word_wrap"] = None if wrap == msoTriStateMixed else wrap == msoTrue
    except Exception:
        pass

    try:
        state["vertical_anchor"] = VERTICAL_ANCHOR_NAMES.get(tf.VerticalAnchor)
    except Exception:
        pass

    try:
        state["orientation"] = ORIENTATION_NAMES.get(tf.Orientation)
    except Exception:
        pass

    try:
        state["margins"] = {
            "left": round(tf.MarginLeft, 2),
            "right": round(tf.MarginRight, 2),
            "top": round(tf.MarginTop, 2),
            "bottom": round(tf.MarginBottom, 2),
        }
    except Exception:
        pass

    return state


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
        "text_frame": _text_frame_state(shape),
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
            info["fill"]["color_hex"] = int_to_hex(fill.ForeColor.RGB)
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
            info["line"]["color_hex"] = int_to_hex(line.ForeColor.RGB)
        except Exception:
            pass
        try:
            info["line"]["dash_style"] = line.DashStyle
        except Exception:
            pass
    except Exception:
        pass

    # Connector info
    try:
        cf = shape.ConnectorFormat
        conn_info = {}
        try:
            if cf.BeginConnected:
                conn_info["begin_connected_shape"] = cf.BeginConnectedShape.Name
                conn_info["begin_connection_site"] = cf.BeginConnectionSite
        except Exception:
            pass
        try:
            if cf.EndConnected:
                conn_info["end_connected_shape"] = cf.EndConnectedShape.Name
                conn_info["end_connection_site"] = cf.EndConnectionSite
        except Exception:
            pass
        if conn_info:
            info["connector_format"] = conn_info
    except Exception:
        pass

    # Adjustment handles
    try:
        adj_count = shape.Adjustments.Count
        if adj_count > 0:
            adj_dict = {}
            for i in range(1, adj_count + 1):
                try:
                    adj_dict[i] = round(shape.Adjustments[i], 4)
                except Exception:
                    pass
            info["adjustments"] = adj_dict
            info["adjustments_count"] = len(adj_dict)
            # Include semantic labels when available
            try:
                auto_type = shape.AutoShapeType
                labels = ADJUSTMENT_LABELS.get(auto_type)
                if labels:
                    info["adjustment_labels"] = labels
            except Exception:
                pass
    except Exception:
        pass

    return info


# ---------------------------------------------------------------------------
# Choosing what an update applies to
# ---------------------------------------------------------------------------
def select_targets(available, shape_names, all_shapes, exclude):
    """Work out which shapes an update should touch.

    Pure arithmetic, apart from the slide, so the awkward part can be tested
    without PowerPoint. `available` is the names on the slide at the top
    level, in z order.

    Returns one list, in the order the answer should come back in. An int is a
    position in `available`, a str is a name that is not at the top level and
    has to be resolved another way before anything is called missing, because
    a shape inside a group answers to its own name everywhere else.

    Positions rather than names, because two shapes on a slide can share a
    name. Going back through the name would move the first of them twice and
    leave the second where it was, while reporting both as done.

    Order follows the slide for `all`, and the caller's list otherwise, so a
    result reads in the order the caller thinks in.
    """
    excluded = set(exclude or ())
    if all_shapes:
        return [i for i, name in enumerate(available) if name not in excluded]

    first_at = {}
    for i, name in enumerate(available):
        first_at.setdefault(name, i)

    picked = []
    for name in shape_names:
        if name in excluded:
            continue
        picked.append(first_at.get(name, name))
    return picked


def _apply_geometry(shape, left, top, width, height, rotation,
                    dleft, dtop, dwidth, dheight):
    """Absolute values first, then the offsets, on one shape."""
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
    if dleft is not None:
        shape.Left = shape.Left + dleft
    if dtop is not None:
        shape.Top = shape.Top + dtop
    if dwidth is not None:
        shape.Width = shape.Width + dwidth
    if dheight is not None:
        shape.Height = shape.Height + dheight


def _geometry_of(shape):
    return {
        "shape_name": shape.Name,
        "left": round(shape.Left, 2),
        "top": round(shape.Top, 2),
        "width": round(shape.Width, 2),
        "height": round(shape.Height, 2),
    }


def _update_many_impl(slide_index, shape_names, all_shapes, exclude,
                      left, top, width, height, rotation,
                      dleft, dtop, dwidth, dheight):
    """Move or resize a set of shapes in one call.

    Nothing is written until every name has been resolved, so a typo leaves
    the slide alone rather than half shifted. Seventeen shapes moving one at a
    time is also seventeen repaints, hence the freeze.
    """
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)

    shapes = [slide.Shapes(i) for i in range(1, slide.Shapes.Count + 1)]
    order = [shape.Name for shape in shapes]

    targets, missing = [], []
    for pick in select_targets(order, shape_names, all_shapes, exclude):
        if isinstance(pick, int):
            targets.append(shapes[pick])
            continue
        # Not at the top level. It may still be a group's child, which
        # answers to its own name, or a "Group 20/Rounded Rectangle 22" path,
        # the way shape_name does.
        try:
            targets.append(resolve_shape(slide, pick))
        except ValueError:
            missing.append(pick)

    if missing:
        raise ValueError(
            "Nothing was moved. Not found on slide "
            f"{slide_index}, at the top level or inside a group: "
            f"{', '.join(missing)}. On the slide: {', '.join(order)}"
        )

    with FrozenRedraw():
        updated = []
        for shape in targets:
            _apply_geometry(shape, left, top, width, height, rotation,
                            dleft, dtop, dwidth, dheight)
            updated.append(_geometry_of(shape))

    return {"success": True, "count": len(updated), "updated": updated}


def _update_shape_impl(slide_index, shape_name, shape_index, left, top, width, height,
                       rotation, name, adjustments,
                       dleft=None, dtop=None, dwidth=None, dheight=None):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, None, shape_name=shape_name, shape_index=shape_index)

    _apply_geometry(shape, left, top, width, height, rotation,
                    dleft, dtop, dwidth, dheight)
    if name is not None:
        shape.Name = name

    # Apply adjustment handle values.
    if adjustments:
        try:
            adj_count = shape.Adjustments.Count
        except Exception:
            raise ValueError(
                f"Shape '{shape.Name}' does not support adjustment handles"
            )
        for idx, value in adjustments.items():
            idx = int(idx)  # ensure int even if str comes through deserialization
            if idx < 1 or idx > adj_count:
                raise ValueError(
                    f"Adjustment index {idx} out of range for shape "
                    f"'{shape.Name}' (has {adj_count} adjustment(s))"
                )
            shape.Adjustments[idx] = value

    result = {
        "success": True,
        "shape_name": shape.Name,
        "left": round(shape.Left, 2),
        "top": round(shape.Top, 2),
        "width": round(shape.Width, 2),
        "height": round(shape.Height, 2),
    }

    # Include current adjustment values in response when adjustments were set.
    if adjustments:
        adj_dict = {}
        for i in range(1, adj_count + 1):
            try:
                adj_dict[i] = round(shape.Adjustments[i], 4)
            except Exception:
                pass
        result["adjustments"] = adj_dict
        # Include semantic labels when available
        try:
            auto_type = shape.AutoShapeType
            labels = ADJUSTMENT_LABELS.get(auto_type)
            if labels:
                result["adjustment_labels"] = labels
        except Exception:
            pass

    return result


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

    # send_behind_text is not one of PowerPoint's four commands, it is a walk
    # up from the bottom. Same helper the adding tools use.
    if z_order_cmd == BEHIND_TEXT:
        placed = place_in_zorder(slide, shape, "behind_text")
        result = {"success": True, "shape_name": shape.Name,
                  "new_z_position": shape.ZOrderPosition}
        if "note" in placed:
            result["note"] = placed["note"]
        return result

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
            params.font_name, params.font_size, params.bold,
            params.italic, params.font_color, params.align,
            params.fill_color, params.fill_type, params.fill_color2,
            params.fill_gradient_style, params.fill_transparency,
            params.line_visible, params.line_color, params.line_weight,
            params.corner_radius, params.corner_radius_pt,
            params.zorder,
        )
        warn = font_size_warning(params.font_size)
        if warn:
            result["warning"] = warn
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
            params.vertical_anchor, params.zorder,
        )
        warn = font_size_warning(params.font_size)
        if warn:
            result["warning"] = warn
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
            params.zorder,
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
            params.zorder,
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

    text_frame carries autofit, word_wrap, vertical_anchor, orientation and
    the four margins, in the words ppt_set_textframe accepts, or null for a
    shape with no text frame. autofit is the configured mode, so
    "shrink_to_fit" says the text may be drawn smaller than the size
    ppt_get_text reports, not that it is; ppt_check_typography measures which
    one it is. The margins matter when working out whether a line fits,
    because the usable width is the shape width less the left and right
    margin, around 14pt on a default box.

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
    """Update properties of an existing shape, or of several at once.

    Only updates properties that are provided (not None). Can change
    position, size, rotation, name, and shape-specific adjustment handles.

    dleft, dtop, dwidth and dheight are offsets against what the shape has
    now, so moving something up by 26pt does not need its current top read
    first. shape_names and all pick several shapes, and exclude leaves some
    out, so shifting a whole slide except its background is one call rather
    than one per shape plus the arithmetic. Nothing is written until every
    name has resolved, so a typo leaves the slide alone.

    Adjustment handles control shape-specific geometry — e.g., triangle apex
    position, arrow proportions, callout pointer, star depth, cross thickness.
    Use ppt_get_shape_info to discover available adjustments and current values.

    Args:
        params: Shape identifier and properties to update.

    Returns:
        JSON with the updated shape name, position and size, and adjustment
        values. For shape_names or all, a count and one entry per shape.
    """
    try:
        if params.all or params.shape_names:
            result = ppt.execute(
                _update_many_impl,
                params.slide_index, params.shape_names, params.all,
                params.exclude,
                params.left, params.top, params.width, params.height,
                params.rotation,
                params.dleft, params.dtop, params.dwidth, params.dheight,
            )
        else:
            result = ppt.execute(
                _update_shape_impl,
                params.slide_index, params.shape_name, params.shape_index,
                params.left, params.top, params.width, params.height,
                params.rotation, params.name, params.adjustments,
                params.dleft, params.dtop, params.dwidth, params.dheight,
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

    Commands: 'bring_to_front', 'send_to_back', 'bring_forward',
    'send_backward', 'send_behind_text'.

    'send_behind_text' puts the shape directly below the lowest shape carrying
    text, which is where art under a caption belongs. It is not one of
    PowerPoint's own commands; 'send_to_back' is usually wrong for this
    because a deck with a full bleed background hides the shape completely.

    Args:
        params: Shape identifier and z-order command.

    Returns:
        JSON with shape name and new z-order position.
    """
    try:
        cmd = params.command.strip().lower().replace(" ", "_").replace("-", "_")
        if cmd != BEHIND_TEXT and cmd not in ZORDER_CMD_MAP:
            return json.dumps({
                "error": f"Unknown z-order command '{params.command}'. "
                f"Use one of: {', '.join(list(ZORDER_CMD_MAP) + [BEHIND_TEXT])}"
            })
        result = ppt.execute(
            _set_zorder_impl,
            params.slide_index, params.shape_name, params.shape_index,
            BEHIND_TEXT if cmd == BEHIND_TEXT else ZORDER_CMD_MAP[cmd],
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

        Optionally apply text styling in the same call via font_name, font_size, bold,
        italic, font_color, and align — avoids a separate ppt_format_text call.

        Optionally apply fill and border in the same call via fill_color, fill_type,
        fill_transparency, line_visible, line_color, and line_weight — avoids separate
        ppt_set_fill / ppt_set_line calls for common cases.

        For rounded_rectangle shapes, control corner rounding with either:
        - corner_radius (0.0–1.0): ratio-based, 0.0 = square, 1.0 = max rounding
        - corner_radius_pt (points): absolute size, e.g. 10 = 10pt radius
        These are mutually exclusive. Ignored for other shape types.

        Example: text='Label', font_size=14, bold=true, fill_color='#1E3A5F',
        line_visible=false creates a fully styled shape in one step.
        """
        return await run_offloaded(add_shape, params)

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

        Text line-break behaviour (same as ppt_set_text):
        - \\n = paragraph break (Enter) — starts a new paragraph.
        - \\v = line break (Shift+Enter) — soft return within the same paragraph.

        Optionally apply font styling in the same call via font_name, font_size, bold,
        italic, font_color, and align — avoids a separate ppt_format_text call.
        Use vertical_anchor ('top', 'middle', 'bottom') to control vertical text
        alignment — avoids a separate ppt_set_textframe call.
        Example: text='Title', font_name='Segoe UI', font_size=32, bold=true,
        font_color='#FFFFFF', align='center', vertical_anchor='middle' creates a
        fully styled, vertically centered label in one step.
        """
        return await run_offloaded(add_textbox, params)

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
        return await run_offloaded(add_picture, params)

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
        return await run_offloaded(add_line, params)

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
        return await run_offloaded(list_shapes, params)

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
        Returns full text, fill info, line info, rotation, z-order, and
        text_frame (autofit, word_wrap, vertical_anchor, orientation,
        margins). Read text_frame before sizing text to fit a box. autofit is
        the configured mode, so "shrink_to_fit" means the drawn size may be
        smaller than the size that was set, and ppt_check_typography is what
        says whether it currently is. The usable width is the shape width
        less the side margins.
        """
        return await run_offloaded(get_shape_info, params)

    @mcp.tool(
        name="ppt_update_shape",
        annotations={
            "title": "Update Shape",
            "readOnlyHint": False,
            "destructiveHint": False,
            # The absolute fields are idempotent and the d* offsets are not:
            # a retried dtop=-26 moves the shape another 26 points. The hint
            # is one value for the whole tool, so it takes the honest one.
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_update_shape(params: UpdateShapeInput) -> str:
        """Update properties of an existing shape, or of several at once.

        Identify the shape by name or index, or several with shape_names, or
        every shape on the slide with all=true plus exclude for the ones to
        leave alone. Only provided properties are updated.

        Absolute: left, top, width, height, rotation, name.
        Relative: dleft, dtop, dwidth, dheight, applied to what the shape has
        now. Prefer these for moving things, no read and no arithmetic first.

        Shifting a whole slide up by 26pt except its background is one call:
        all=true, exclude=["Picture 2"], dtop=-26. Nothing is written until
        every name has resolved, so a typo leaves the slide alone rather than
        half shifted.

        The d* offsets are not idempotent: sending the same call twice moves
        the shape twice. all=true means the shapes at the top level, so a
        group moves as one; name a group's child in shape_names to reach
        inside it.
        """
        return await run_offloaded(update_shape, params)

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
        return await run_offloaded(delete_shape, params)

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
        return await run_offloaded(duplicate_shape, params)

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

        Commands: 'bring_to_front', 'send_to_back', 'bring_forward',
        'send_backward', 'send_behind_text' (directly below the lowest shape
        that has text, for art that belongs under a caption).
        Identify the shape by name (shape_name) or 1-based index (shape_index).
        """
        return await run_offloaded(set_shape_zorder, params)


# ---------------------------------------------------------------------------
# macOS
# ---------------------------------------------------------------------------
# The implementations above walk COM. Their Apple Event counterparts have the
# same names and signatures, so on macOS they simply take their place; nothing
# else in this module changes.
from backend import IS_MACOS, use_mac_impls  # noqa: E402

if IS_MACOS:  # pragma: no cover - platform specific
    from ppt_mac import shapes as _mac_shapes

    use_mac_impls(globals(), _mac_shapes)
