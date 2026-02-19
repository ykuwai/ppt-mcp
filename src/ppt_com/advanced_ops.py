"""Advanced operations for PowerPoint COM automation.

Handles tags, font management, picture cropping, shape export,
slide visibility, shape selection, view control, animation copying,
picture insertion from URL, aspect ratio locking, and icon search.
"""

import json
import logging
import os
import tempfile
import time
import urllib.error
import urllib.request
from typing import Optional, Union

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from utils.navigation import goto_slide
from ppt_com.constants import (
    msoTrue, msoFalse,
    SHAPE_FORMAT_MAP,
    VIEW_TYPE_MAP, VIEW_TYPE_NAMES,
    ppSelectionNone, ppSelectionSlides, ppSelectionShapes, ppSelectionText,
)

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Icon metadata cache (lazy-loaded on first search)
# ---------------------------------------------------------------------------
_icon_cache = None        # list of icon dicts
_icon_cache_time = 0.0    # timestamp of last fetch
_ICON_CACHE_TTL = 86400   # 24 hours

_ICON_METADATA_URL = "https://fonts.google.com/metadata/icons"


def _fetch_icon_metadata():
    """Fetch and cache the Material Symbols icon metadata from Google Fonts.

    The first line of the response is `)]}'` (XSS protection) and must be
    stripped before parsing as JSON.  The parsed icons list is cached for
    24 hours to avoid repeated network calls.
    """
    global _icon_cache, _icon_cache_time

    now = time.time()
    if _icon_cache is not None and (now - _icon_cache_time) < _ICON_CACHE_TTL:
        return _icon_cache

    resp = urllib.request.urlopen(_ICON_METADATA_URL)
    raw = resp.read().decode("utf-8")

    # Strip XSS protection prefix  )]}'
    first_nl = raw.index("\n")
    json_str = raw[first_nl + 1:]
    data = json.loads(json_str)

    _icon_cache = data.get("icons", [])
    _icon_cache_time = now
    logger.info("Fetched %d icons from Google Fonts metadata", len(_icon_cache))
    return _icon_cache


def _search_icons(query: str, max_results: int = 20):
    """Search Material Symbols icons by keyword.

    Scoring:
    - Exact icon name match: +100
    - Full query (multi-word) found in icon name: +50
    - All query words found in icon name: +40
    - Query word found in icon name: +30
    - Exact tag match: +20
    - Query word found in a tag: +10
    - Query word found in a category: +5
    - Popularity bonus (normalized to 0-10 range)

    Returns a sorted list of dicts with name, tags, categories, score.
    """
    icons = _fetch_icon_metadata()
    query_lower = query.lower().strip()
    query_words = query_lower.split()

    results = []
    for icon in icons:
        name = icon.get("name", "")
        tags = [t.lower() for t in icon.get("tags", [])]
        categories = [c.lower() for c in icon.get("categories", [])]
        popularity = icon.get("popularity", 0)

        score = 0

        # Exact name match (query == name)
        if name == query_lower:
            score += 100
        # Full query string in name (e.g. "arrow_forward" contains "arrow forward" as "arrow_forward")
        elif query_lower.replace(" ", "_") == name:
            score += 90
        # Full query in name
        elif query_lower in name:
            score += 50

        # Bonus: all query words found in name
        if len(query_words) > 1 and all(w in name for w in query_words):
            score += 40

        # Per-word scoring
        for word in query_words:
            if word in name:
                score += 30
            for tag in tags:
                if word == tag:
                    score += 20
                elif word in tag:
                    score += 10
            for cat in categories:
                if word in cat:
                    score += 5

        if score > 0:
            # Small popularity bonus (normalized)
            score += min(popularity / 1000, 10)
            results.append({
                "name": name,
                "categories": icon.get("categories", []),
                "tags": icon.get("tags", [])[:8],  # limit tags for readability
                "popularity": popularity,
                "score": round(score, 2),
            })

    results.sort(key=lambda x: x["score"], reverse=True)
    return results[:max_results]


# ---------------------------------------------------------------------------
# Helper: find a shape by name or index
# ---------------------------------------------------------------------------
def _get_shape(slide, name_or_index: Union[str, int]):
    """Find a shape on a slide by name or 1-based index.

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
            if slide.Shapes(i).Name == name_or_index:
                return slide.Shapes(i)
        raise ValueError(f"Shape '{name_or_index}' not found on slide")


# ===========================================================================
# Pydantic input models
# ===========================================================================

# --- Tags ---
class SetTagInput(BaseModel):
    """Input for setting a tag on a shape, slide, or presentation."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: Optional[int] = Field(
        default=None, ge=1, description="1-based slide index (required for slide/shape targets)"
    )
    shape_name_or_index: Optional[Union[str, int]] = Field(
        default=None, description="Shape name (str) or 1-based index (int) for shape target. Prefer name — indices shift when shapes are added/removed"
    )
    tag_name: str = Field(..., description="Tag name (key)")
    tag_value: str = Field(..., description="Tag value")
    target_type: str = Field(
        default="shape",
        description="Target type: 'shape', 'slide', or 'presentation'",
    )


class GetTagsInput(BaseModel):
    """Input for getting tags from a shape, slide, or presentation."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: Optional[int] = Field(
        default=None, ge=1, description="1-based slide index (required for slide/shape targets)"
    )
    shape_name_or_index: Optional[Union[str, int]] = Field(
        default=None, description="Shape name (str) or 1-based index (int) for shape target. Prefer name — indices shift when shapes are added/removed"
    )
    target_type: str = Field(
        default="shape",
        description="Target type: 'shape', 'slide', or 'presentation'",
    )


# --- Fonts ---
class ReplaceFontInput(BaseModel):
    """Input for replacing a font throughout the presentation."""
    model_config = ConfigDict(str_strip_whitespace=True)

    original_font: str = Field(..., description="Font name to replace")
    replacement_font: str = Field(..., description="New font name")


# --- Set Default Fonts ---
class SetDefaultFontsInput(BaseModel):
    """Input for setting default fonts for the entire presentation."""
    model_config = ConfigDict(str_strip_whitespace=True)

    latin: Optional[str] = Field(
        default=None,
        description="Latin (alphabet/number) font name (e.g. 'Segoe UI', 'Calibri')",
    )
    east_asian: Optional[str] = Field(
        default=None,
        description="East Asian (Japanese/Chinese/Korean) font name (e.g. 'Meiryo', 'Yu Gothic UI')",
    )
    apply_to_existing: bool = Field(
        default=True,
        description="If true (default), also apply fonts to all existing text in the presentation. If false, only update the theme fonts for new text.",
    )


# --- Picture Crop ---
class CropPictureInput(BaseModel):
    """Input for cropping a picture shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    crop_left: Optional[float] = Field(default=None, description="Crop from left in points")
    crop_right: Optional[float] = Field(default=None, description="Crop from right in points")
    crop_top: Optional[float] = Field(default=None, description="Crop from top in points")
    crop_bottom: Optional[float] = Field(default=None, description="Crop from bottom in points")


# --- Shape Export ---
class ExportShapeInput(BaseModel):
    """Input for exporting a shape as an image file."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    file_path: str = Field(..., description="Output file path")
    format: str = Field(
        default="png",
        description="Image format: 'png', 'jpg', 'gif', 'bmp', 'wmf', or 'emf'",
    )
    width: Optional[int] = Field(default=None, description="Export width in pixels")
    height: Optional[int] = Field(default=None, description="Export height in pixels")


# --- Slide Hidden ---
class SetSlideHiddenInput(BaseModel):
    """Input for setting slide visibility."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    hidden: bool = Field(..., description="True to hide, False to show")


# --- Select Shapes ---
class SelectShapesInput(BaseModel):
    """Input for selecting multiple shapes on a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_names: list[str] = Field(
        ..., description="List of shape names to select"
    )


# --- View ---
class SetViewInput(BaseModel):
    """Input for setting the PowerPoint view type and zoom."""
    model_config = ConfigDict(str_strip_whitespace=True)

    view_type: Optional[str] = Field(
        default=None,
        description=(
            "View type: 'normal', 'slide_master', 'notes_page', 'handout_master', "
            "'notes_master', 'outline', 'slide_sorter', 'title_master', 'reading'"
        ),
    )
    zoom: Optional[int] = Field(
        default=None, ge=10, le=400, description="Zoom level (10-400)"
    )


# --- Copy Animation ---
class CopyAnimationInput(BaseModel):
    """Input for copying animation from one shape to another."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    source_shape: Union[str, int] = Field(
        ..., description="Source shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    target_shape: Union[str, int] = Field(
        ..., description="Target shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )


# --- Add Picture from URL ---
class AddPictureFromUrlInput(BaseModel):
    """Input for adding a picture from a URL."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    url: str = Field(..., description="URL of the image to download")
    left: float = Field(default=100, description="Left position in points")
    top: float = Field(default=100, description="Top position in points")
    width: Optional[float] = Field(default=None, description="Width in points (auto if not set)")
    height: Optional[float] = Field(default=None, description="Height in points (auto if not set)")
    svg_color: Optional[str] = Field(
        default=None,
        description=(
            "Replace 'currentColor' in SVG files with this color (e.g. '#1A73E8'). "
            "Only applies to SVG files."
        ),
    )
    fit: bool = Field(
        default=False,
        description=(
            "If true, fit the image within the width×height area while preserving "
            "aspect ratio and centering. Requires both width and height."
        ),
    )


# --- Add SVG Icon ---
class AddSvgIconInput(BaseModel):
    """Input for adding a Material Symbols icon as SVG image."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    icon_name: str = Field(
        ...,
        description=(
            "Material Symbols icon name (e.g. 'bolt', 'description', 'extension', "
            "'settings', 'home', 'search', 'favorite'). "
            "See https://fonts.google.com/icons for available icons."
        ),
    )
    left: float = Field(default=100, description="Left position in points")
    top: float = Field(default=100, description="Top position in points")
    width: float = Field(default=72, description="Width of the area in points")
    height: float = Field(default=72, description="Height of the area in points")
    color: str = Field(
        default="accent1",
        description=(
            "Icon color. Use '#RRGGBB' hex string or a theme color name "
            "(e.g. 'accent1', 'accent2', 'dark1', 'light1'). "
            "Default: 'accent1' (the presentation's main accent color)."
        ),
    )
    style: str = Field(
        default="outlined",
        description="Icon style: 'outlined', 'rounded', or 'sharp'",
    )
    filled: bool = Field(
        default=False,
        description="If true, use the filled variant of the icon instead of outline.",
    )


# --- Lock Aspect Ratio ---
class LockAspectRatioInput(BaseModel):
    """Input for locking/unlocking shape aspect ratio."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    locked: bool = Field(..., description="True to lock, False to unlock")


# --- Search Icons ---
class SearchIconsInput(BaseModel):
    """Input for searching Material Symbols icons by keyword."""
    model_config = ConfigDict(str_strip_whitespace=True)

    query: str = Field(
        ...,
        description=(
            "Search keyword(s) for finding icons. Examples: 'home', 'arrow', "
            "'settings gear', 'chart graph'. Multiple words narrow the search."
        ),
    )
    max_results: int = Field(
        default=20, ge=1, le=100,
        description="Maximum number of results to return (default: 20)",
    )


# ===========================================================================
# COM implementation functions (run on COM thread via ppt.execute)
# ===========================================================================

# ---------------------------------------------------------------------------
# Tags
# ---------------------------------------------------------------------------
def _resolve_target(app, target_type, slide_index, shape_name_or_index):
    """Resolve the target COM object based on target_type."""
    pres = app.ActivePresentation
    target_type_lower = target_type.strip().lower()

    if target_type_lower == "presentation":
        return pres
    elif target_type_lower == "slide":
        if slide_index is None:
            raise ValueError("slide_index is required for target_type='slide'")
        return pres.Slides(slide_index)
    elif target_type_lower == "shape":
        if slide_index is None:
            raise ValueError("slide_index is required for target_type='shape'")
        if shape_name_or_index is None:
            raise ValueError("shape_name_or_index is required for target_type='shape'")
        slide = pres.Slides(slide_index)
        return _get_shape(slide, shape_name_or_index)
    else:
        raise ValueError(
            f"Unknown target_type '{target_type}'. Use 'shape', 'slide', or 'presentation'."
        )


def _set_tag_impl(slide_index, shape_name_or_index, tag_name, tag_value, target_type):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    target = _resolve_target(app, target_type, slide_index, shape_name_or_index)
    target.Tags.Add(tag_name, tag_value)
    return {
        "success": True,
        "target_type": target_type,
        "tag_name": tag_name,
        "tag_value": tag_value,
    }


def _get_tags_impl(slide_index, shape_name_or_index, target_type):
    app = ppt._get_app_impl()
    target = _resolve_target(app, target_type, slide_index, shape_name_or_index)
    tags = {}
    for i in range(1, target.Tags.Count + 1):
        tags[target.Tags.Name(i)] = target.Tags.Value(i)
    return {
        "success": True,
        "target_type": target_type,
        "tags_count": target.Tags.Count,
        "tags": tags,
    }


# ---------------------------------------------------------------------------
# Fonts
# ---------------------------------------------------------------------------
def _replace_font_impl(original_font, replacement_font):
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    pres.Fonts.Replace(original_font, replacement_font)
    return {
        "success": True,
        "original_font": original_font,
        "replacement_font": replacement_font,
    }


def _list_fonts_impl():
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    fonts = []
    for i in range(1, pres.Fonts.Count + 1):
        fonts.append(pres.Fonts(i).Name)
    return {
        "success": True,
        "fonts_count": len(fonts),
        "fonts": fonts,
    }


# ---------------------------------------------------------------------------
# Set Default Fonts
# ---------------------------------------------------------------------------
def _set_default_fonts_impl(latin, east_asian, apply_to_existing):
    app = ppt._get_app_impl()
    pres = app.ActivePresentation

    if not latin and not east_asian:
        raise ValueError("At least one of 'latin' or 'east_asian' must be provided")

    theme_updated = False
    # Step 1: Update theme fonts for all slide masters (affects new text)
    try:
        masters_updated = 0
        for m in range(1, pres.SlideMasters.Count + 1):
            try:
                font_scheme = pres.SlideMasters(m).Theme.ThemeFontScheme
                # msoThemeFontLatin = 1, msoThemeFontEastAsian = 2
                if latin:
                    font_scheme.MajorFont(1).Name = latin
                    font_scheme.MinorFont(1).Name = latin
                if east_asian:
                    font_scheme.MajorFont(2).Name = east_asian
                    font_scheme.MinorFont(2).Name = east_asian
                masters_updated += 1
            except Exception as e:
                logger.warning("Failed to update theme fonts for master %d: %s", m, e)
        theme_updated = masters_updated > 0
    except Exception as e:
        logger.warning("Failed to iterate slide masters: %s", e)

    # Step 2: Apply to existing text (including shapes inside groups)
    def _apply_to_shape(shape):
        """Recursively apply fonts to a shape and any grouped children."""
        try:
            if shape.HasTextFrame:
                font = shape.TextFrame.TextRange.Font
                if latin:
                    font.Name = latin
                if east_asian:
                    font.NameFarEast = east_asian
                return 1
        except Exception:
            pass
        # Recurse into group members
        updated = 0
        try:
            for k in range(1, shape.GroupItems.Count + 1):
                updated += _apply_to_shape(shape.GroupItems(k))
        except Exception:
            pass
        return updated

    slides_processed = 0
    shapes_updated = 0
    if apply_to_existing:
        for i in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(i)
            slides_processed += 1
            for j in range(1, slide.Shapes.Count + 1):
                try:
                    shapes_updated += _apply_to_shape(slide.Shapes(j))
                except Exception:
                    pass

    result = {"success": True, "theme_updated": theme_updated}
    if latin:
        result["latin"] = latin
    if east_asian:
        result["east_asian"] = east_asian
    if apply_to_existing:
        result["slides_processed"] = slides_processed
        result["shapes_updated"] = shapes_updated

    return result


# ---------------------------------------------------------------------------
# Picture Crop
# ---------------------------------------------------------------------------
def _crop_picture_impl(slide_index, shape_name_or_index, crop_left, crop_right, crop_top, crop_bottom):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    pic_fmt = shape.PictureFormat
    if crop_left is not None:
        pic_fmt.CropLeft = crop_left
    if crop_right is not None:
        pic_fmt.CropRight = crop_right
    if crop_top is not None:
        pic_fmt.CropTop = crop_top
    if crop_bottom is not None:
        pic_fmt.CropBottom = crop_bottom

    return {
        "success": True,
        "shape_name": shape.Name,
        "crop_left": round(pic_fmt.CropLeft, 2),
        "crop_right": round(pic_fmt.CropRight, 2),
        "crop_top": round(pic_fmt.CropTop, 2),
        "crop_bottom": round(pic_fmt.CropBottom, 2),
    }


# ---------------------------------------------------------------------------
# Shape Export
# ---------------------------------------------------------------------------
def _export_shape_impl(slide_index, shape_name_or_index, file_path, format_type, width, height):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    abs_path = os.path.abspath(file_path)

    # Convert string format name to integer if needed
    if isinstance(format_type, str):
        fmt_key = format_type.strip().lower()
        fmt_int = SHAPE_FORMAT_MAP.get(fmt_key)
        if fmt_int is None:
            raise ValueError(
                f"Unknown format '{format_type}'. "
                f"Valid values: {list(SHAPE_FORMAT_MAP.keys())}"
            )
        format_type = fmt_int

    # Ensure output directory exists
    out_dir = os.path.dirname(abs_path)
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    if width is not None and height is not None:
        shape.Export(abs_path, format_type, width, height)
    elif width is not None:
        shape.Export(abs_path, format_type, width)
    else:
        shape.Export(abs_path, format_type)

    return {
        "success": True,
        "shape_name": shape.Name,
        "file_path": abs_path,
    }


# ---------------------------------------------------------------------------
# Slide Hidden
# ---------------------------------------------------------------------------
def _set_slide_hidden_impl(slide_index, hidden):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    slide.SlideShowTransition.Hidden = msoTrue if hidden else msoFalse
    return {
        "success": True,
        "slide_index": slide_index,
        "hidden": hidden,
    }


# ---------------------------------------------------------------------------
# Select Shapes
# ---------------------------------------------------------------------------
def _select_shapes_impl(slide_index, shape_names):
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    # Navigate to the slide first
    app.ActiveWindow.View.GotoSlide(slide_index)

    # Select first shape (replace=True is default)
    first_shape = _get_shape(slide, shape_names[0])
    first_shape.Select()

    # Add remaining shapes to selection (msoFalse=0 means add to selection)
    for name in shape_names[1:]:
        shape = _get_shape(slide, name)
        shape.Select(msoFalse)

    return {
        "success": True,
        "slide_index": slide_index,
        "selected_shapes": shape_names,
        "count": len(shape_names),
    }


# ---------------------------------------------------------------------------
# Get Selection
# ---------------------------------------------------------------------------
def _get_selection_impl():
    app = ppt._get_app_impl()
    selection = app.ActiveWindow.Selection
    sel_type = selection.Type

    result = {
        "success": True,
        "type": sel_type,
    }

    if sel_type == ppSelectionNone:
        result["type_name"] = "none"
    elif sel_type == ppSelectionSlides:
        result["type_name"] = "slides"
        slide_indices = []
        for i in range(1, selection.SlideRange.Count + 1):
            slide_indices.append(selection.SlideRange(i).SlideIndex)
        result["slide_indices"] = slide_indices
    elif sel_type == ppSelectionShapes:
        result["type_name"] = "shapes"
        shape_names = []
        for i in range(1, selection.ShapeRange.Count + 1):
            shape_names.append(selection.ShapeRange(i).Name)
        result["shape_names"] = shape_names
        result["count"] = selection.ShapeRange.Count
    elif sel_type == ppSelectionText:
        result["type_name"] = "text"
        result["text"] = selection.TextRange.Text

    return result


# ---------------------------------------------------------------------------
# View
# ---------------------------------------------------------------------------
def _set_view_impl(view_type, zoom):
    app = ppt._get_app_impl()
    window = app.ActiveWindow

    if view_type is not None:
        vt_key = view_type.strip().lower().replace(" ", "_").replace("-", "_")
        if vt_key not in VIEW_TYPE_MAP:
            raise ValueError(
                f"Unknown view_type '{view_type}'. "
                f"Use one of: {', '.join(VIEW_TYPE_MAP.keys())}"
            )
        window.ViewType = VIEW_TYPE_MAP[vt_key]

    if zoom is not None:
        window.View.Zoom = zoom

    current_view_type = window.ViewType
    current_zoom = window.View.Zoom

    return {
        "success": True,
        "view_type": VIEW_TYPE_NAMES.get(current_view_type, f"Unknown({current_view_type})"),
        "view_type_id": current_view_type,
        "zoom": current_zoom,
    }


# ---------------------------------------------------------------------------
# Copy Animation
# ---------------------------------------------------------------------------
def _copy_animation_impl(slide_index, source_shape, target_shape):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    src = _get_shape(slide, source_shape)
    tgt = _get_shape(slide, target_shape)

    src.PickupAnimation()
    tgt.ApplyAnimation()

    return {
        "success": True,
        "source_shape": src.Name,
        "target_shape": tgt.Name,
    }


# ---------------------------------------------------------------------------
# Add Picture from URL
# ---------------------------------------------------------------------------
def _add_picture_from_url_impl(slide_index, url, left, top, width, height, svg_color, fit):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    # Download the image
    resp = urllib.request.urlopen(url)
    content_type = resp.headers.get("Content-Type", "")
    is_svg = url.lower().endswith(".svg") or "svg" in content_type

    if is_svg:
        svg_text = resp.read().decode("utf-8")
        if svg_color:
            svg_text = svg_text.replace("currentColor", svg_color)
        tmp_fd, tmp_path = tempfile.mkstemp(suffix=".svg")
        os.close(tmp_fd)
        with open(tmp_path, "w", encoding="utf-8") as f:
            f.write(svg_text)
    else:
        data = resp.read()
        suffix = os.path.splitext(url.split("?")[0])[-1] or ".png"
        tmp_fd, tmp_path = tempfile.mkstemp(suffix=suffix)
        os.close(tmp_fd)
        with open(tmp_path, "wb") as f:
            f.write(data)

    try:
        abs_tmp = os.path.abspath(tmp_path)

        if fit and width is not None and height is not None:
            # Auto-size first, then fit to area
            # AddPicture(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
            pic = slide.Shapes.AddPicture(abs_tmp, 0, -1, left, top, -1, -1)
            pic.LockAspectRatio = -1  # msoTrue
            scale = min(width / pic.Width, height / pic.Height)
            new_w = pic.Width * scale
            new_h = pic.Height * scale
            pic.Width = new_w
            pic.Left = left + (width - new_w) / 2
            pic.Top = top + (height - new_h) / 2
        else:
            w = width if width is not None else -1
            h = height if height is not None else -1
            pic = slide.Shapes.AddPicture(abs_tmp, 0, -1, left, top, w, h)

        return {
            "success": True,
            "shape_name": pic.Name,
            "shape_index": pic.ZOrderPosition,
            "width": round(pic.Width, 2),
            "height": round(pic.Height, 2),
            "source_url": url,
        }
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


# ---------------------------------------------------------------------------
# Add SVG Icon
# ---------------------------------------------------------------------------
def _resolve_color(pres, color_str):
    """Resolve a color string to a hex '#RRGGBB' value.

    Accepts '#RRGGBB' hex strings directly, or theme color names
    like 'accent1', 'dark1', 'light2', etc.
    """
    if color_str.startswith("#"):
        return color_str

    # Theme color name -> resolve from presentation
    theme_map = {
        "dark1": 1, "light1": 2, "dark2": 3, "light2": 4,
        "accent1": 5, "accent2": 6, "accent3": 7, "accent4": 8,
        "accent5": 9, "accent6": 10, "hyperlink": 11,
        "followed_hyperlink": 12,
    }
    idx = theme_map.get(color_str.lower())
    if idx is None:
        raise ValueError(
            f"Unknown color '{color_str}'. Use '#RRGGBB' or theme name: "
            f"{list(theme_map.keys())}"
        )
    # ThemeColorScheme is 1-based
    bgr = pres.SlideMaster.Theme.ThemeColorScheme(idx).RGB
    r = bgr & 0xFF
    g = (bgr >> 8) & 0xFF
    b = (bgr >> 16) & 0xFF
    return f"#{r:02X}{g:02X}{b:02X}"


def _add_svg_icon_impl(slide_index, icon_name, left, top, width, height, color, style, filled):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    # Resolve theme color name to hex
    hex_color = _resolve_color(pres, color)

    # Build CDN URL (append -fill suffix for filled variant)
    base = "https://cdn.jsdelivr.net/npm/@material-symbols/svg-400@0.31.3"
    file_name = f"{icon_name}-fill" if filled else icon_name
    svg_url = f"{base}/{style}/{file_name}.svg"

    # Download SVG
    try:
        resp = urllib.request.urlopen(svg_url)
    except urllib.error.HTTPError as e:
        if e.code == 404:
            raise ValueError(
                f"Icon '{icon_name}' not found (style='{style}', filled={filled}). "
                f"Check the name at https://fonts.google.com/icons . "
                f"URL: {svg_url}"
            ) from None
        raise
    svg_text = resp.read().decode("utf-8")

    # Apply color: replace currentColor and inject fill on <svg> tag
    svg_text = svg_text.replace("currentColor", hex_color)
    if f'fill="{hex_color}"' not in svg_text:
        svg_text = svg_text.replace("<svg ", f'<svg fill="{hex_color}" ', 1)

    # Write to temp file
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".svg")
    os.close(tmp_fd)

    try:
        with open(tmp_path, "w", encoding="utf-8") as f:
            f.write(svg_text)

        abs_tmp = os.path.abspath(tmp_path)

        # Insert with auto-size, then fit
        # AddPicture(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
        pic = slide.Shapes.AddPicture(abs_tmp, 0, -1, left, top, -1, -1)

        # Fit to area preserving aspect ratio
        pic.LockAspectRatio = -1  # msoTrue
        scale = min(width / pic.Width, height / pic.Height)
        new_w = pic.Width * scale
        new_h = pic.Height * scale
        pic.Width = new_w
        pic.Left = left + (width - new_w) / 2
        pic.Top = top + (height - new_h) / 2

        return {
            "success": True,
            "shape_name": pic.Name,
            "shape_index": pic.ZOrderPosition,
            "width": round(pic.Width, 2),
            "height": round(pic.Height, 2),
            "icon_name": icon_name,
            "source_url": svg_url,
        }
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


# ---------------------------------------------------------------------------
# Lock Aspect Ratio
# ---------------------------------------------------------------------------
def _lock_aspect_ratio_impl(slide_index, shape_name_or_index, locked):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)
    shape.LockAspectRatio = msoTrue if locked else msoFalse
    return {
        "success": True,
        "shape_name": shape.Name,
        "locked": locked,
    }


# ===========================================================================
# MCP tool functions (sync wrappers that delegate to COM thread)
# ===========================================================================

# --- Tags ---
def set_tag(params: SetTagInput) -> str:
    """Set a tag (key-value pair) on a shape, slide, or presentation.

    Args:
        params: Target identification and tag name/value.

    Returns:
        JSON confirming the tag was set.
    """
    try:
        result = ppt.execute(
            _set_tag_impl,
            params.slide_index, params.shape_name_or_index,
            params.tag_name, params.tag_value, params.target_type,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to set tag: {str(e)}"})


def get_tags(params: GetTagsInput) -> str:
    """Get all tags from a shape, slide, or presentation.

    Args:
        params: Target identification.

    Returns:
        JSON with tag count and name-value pairs.
    """
    try:
        result = ppt.execute(
            _get_tags_impl,
            params.slide_index, params.shape_name_or_index, params.target_type,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to get tags: {str(e)}"})


# --- Fonts ---
def replace_font(params: ReplaceFontInput) -> str:
    """Replace a font throughout the active presentation.

    Args:
        params: Original and replacement font names.

    Returns:
        JSON confirming the font replacement.
    """
    try:
        result = ppt.execute(
            _replace_font_impl,
            params.original_font, params.replacement_font,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to replace font: {str(e)}"})


def list_fonts() -> str:
    """List all fonts used in the active presentation.

    Returns:
        JSON with font count and names.
    """
    try:
        result = ppt.execute(_list_fonts_impl)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to list fonts: {str(e)}"})


# --- Set Default Fonts ---
def set_default_fonts(params: SetDefaultFontsInput) -> str:
    """Set default fonts for the entire presentation.

    Args:
        params: Font names and whether to apply to existing text.

    Returns:
        JSON with theme update status and number of shapes updated.
    """
    try:
        result = ppt.execute(
            _set_default_fonts_impl,
            params.latin, params.east_asian, params.apply_to_existing,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to set default fonts: {str(e)}"})


# --- Picture Crop ---
def crop_picture(params: CropPictureInput) -> str:
    """Crop a picture shape.

    Args:
        params: Shape identification and crop values in points.

    Returns:
        JSON with current crop values after setting.
    """
    try:
        result = ppt.execute(
            _crop_picture_impl,
            params.slide_index, params.shape_name_or_index,
            params.crop_left, params.crop_right,
            params.crop_top, params.crop_bottom,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to crop picture: {str(e)}"})


# --- Shape Export ---
def export_shape(params: ExportShapeInput) -> str:
    """Export a shape as an image file.

    Args:
        params: Shape identification, file path, format, and optional dimensions.

    Returns:
        JSON with shape name and output file path.
    """
    try:
        fmt_key = params.format.strip().lower()
        if fmt_key not in SHAPE_FORMAT_MAP:
            return json.dumps({
                "error": f"Unknown format '{params.format}'. "
                f"Use one of: {', '.join(SHAPE_FORMAT_MAP.keys())}"
            })
        result = ppt.execute(
            _export_shape_impl,
            params.slide_index, params.shape_name_or_index,
            params.file_path, SHAPE_FORMAT_MAP[fmt_key],
            params.width, params.height,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to export shape: {str(e)}"})


# --- Slide Hidden ---
def set_slide_hidden(params: SetSlideHiddenInput) -> str:
    """Set a slide as hidden or visible in the slideshow.

    Args:
        params: Slide index and hidden state.

    Returns:
        JSON confirming the hidden state.
    """
    try:
        result = ppt.execute(
            _set_slide_hidden_impl,
            params.slide_index, params.hidden,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to set slide hidden: {str(e)}"})


# --- Select Shapes ---
def select_shapes(params: SelectShapesInput) -> str:
    """Select multiple shapes on a slide.

    Args:
        params: Slide index and list of shape names.

    Returns:
        JSON with selected shape names and count.
    """
    try:
        if not params.shape_names:
            return json.dumps({"error": "shape_names list must not be empty"})
        result = ppt.execute(
            _select_shapes_impl,
            params.slide_index, params.shape_names,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to select shapes: {str(e)}"})


# --- Get Selection ---
def get_selection() -> str:
    """Get the current selection in the active window.

    Returns:
        JSON with selection type and details.
    """
    try:
        result = ppt.execute(_get_selection_impl)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to get selection: {str(e)}"})


# --- View ---
def set_view(params: SetViewInput) -> str:
    """Set the PowerPoint view type and/or zoom level.

    Args:
        params: View type and zoom level.

    Returns:
        JSON with current view type and zoom after setting.
    """
    try:
        result = ppt.execute(
            _set_view_impl,
            params.view_type, params.zoom,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to set view: {str(e)}"})


# --- Copy Animation ---
def copy_animation(params: CopyAnimationInput) -> str:
    """Copy animation from one shape to another on the same slide.

    Args:
        params: Slide index, source shape, and target shape.

    Returns:
        JSON confirming the animation was copied.
    """
    try:
        result = ppt.execute(
            _copy_animation_impl,
            params.slide_index, params.source_shape, params.target_shape,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to copy animation: {str(e)}"})


# --- Add Picture from URL ---
def add_picture_from_url(params: AddPictureFromUrlInput) -> str:
    """Add a picture to a slide by downloading from a URL.

    Args:
        params: Slide index, URL, position, and optional dimensions.

    Returns:
        JSON with shape name, dimensions, and source URL.
    """
    try:
        result = ppt.execute(
            _add_picture_from_url_impl,
            params.slide_index, params.url,
            params.left, params.top, params.width, params.height,
            params.svg_color, params.fit,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to add picture from URL: {str(e)}"})


# --- Add SVG Icon ---
def add_svg_icon(params: AddSvgIconInput) -> str:
    """Add a Material Symbols icon as SVG image to a slide.

    Args:
        params: Slide index, icon name, position, dimensions, color, and style.

    Returns:
        JSON with shape name, dimensions, icon name, and source URL.
    """
    try:
        result = ppt.execute(
            _add_svg_icon_impl,
            params.slide_index, params.icon_name,
            params.left, params.top, params.width, params.height,
            params.color, params.style, params.filled,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to add SVG icon: {str(e)}"})


# --- Lock Aspect Ratio ---
def lock_aspect_ratio(params: LockAspectRatioInput) -> str:
    """Lock or unlock the aspect ratio of a shape.

    Args:
        params: Shape identification and lock state.

    Returns:
        JSON confirming the aspect ratio lock state.
    """
    try:
        result = ppt.execute(
            _lock_aspect_ratio_impl,
            params.slide_index, params.shape_name_or_index, params.locked,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to lock aspect ratio: {str(e)}"})


# --- Search Icons ---
def search_icons(params: SearchIconsInput) -> str:
    """Search Material Symbols icons by keyword.

    Fetches icon metadata from Google Fonts on first call (cached for 24h).

    Args:
        params: Search query and max results.

    Returns:
        JSON with matching icons (name, categories, tags, popularity, score).
    """
    try:
        results = _search_icons(params.query, params.max_results)
        return json.dumps({
            "success": True,
            "query": params.query,
            "count": len(results),
            "icons": results,
        })
    except Exception as e:
        return json.dumps({"error": f"Failed to search icons: {str(e)}"})


# ===========================================================================
# Tool registration
# ===========================================================================
def register_tools(mcp):
    """Register all advanced operations tools with the MCP server."""

    # --- Tags ---
    @mcp.tool(
        name="ppt_set_tag",
        annotations={
            "title": "Set Tag",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_set_tag(params: SetTagInput) -> str:
        """Set a tag (key-value pair) on a shape, slide, or presentation.

        Tags are custom metadata stored as name-value string pairs.
        Set target_type to 'shape' (default), 'slide', or 'presentation'.
        For shape targets, provide slide_index and shape_name_or_index.
        For slide targets, provide slide_index.
        """
        return set_tag(params)

    @mcp.tool(
        name="ppt_get_tags",
        annotations={
            "title": "Get Tags",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_get_tags(params: GetTagsInput) -> str:
        """Get all tags from a shape, slide, or presentation.

        Returns a dictionary of tag name-value pairs.
        Set target_type to 'shape' (default), 'slide', or 'presentation'.
        """
        return get_tags(params)

    # --- Fonts ---
    @mcp.tool(
        name="ppt_replace_font",
        annotations={
            "title": "Replace Font",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_replace_font(params: ReplaceFontInput) -> str:
        """Replace all occurrences of a font throughout the active presentation.

        Replaces every instance of original_font with replacement_font
        across all slides, shapes, and text ranges.
        """
        return replace_font(params)

    @mcp.tool(
        name="ppt_list_fonts",
        annotations={
            "title": "List Fonts",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_list_fonts() -> str:
        """List all fonts used in the active presentation.

        Returns the names of all fonts embedded or referenced in the presentation.
        """
        return list_fonts()

    @mcp.tool(
        name="ppt_set_default_fonts",
        annotations={
            "title": "Set Default Fonts",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_set_default_fonts(params: SetDefaultFontsInput) -> str:
        """Set default fonts for the entire presentation (Latin and East Asian separately).

        Updates theme fonts so new text uses the specified fonts.
        If apply_to_existing is true (default), also updates all existing text.
        Use 'latin' for alphabet/number fonts (e.g. 'Segoe UI') and
        'east_asian' for Japanese/Chinese/Korean fonts (e.g. 'Meiryo').
        At least one of latin or east_asian must be provided.
        """
        return set_default_fonts(params)

    # --- Picture Crop ---
    @mcp.tool(
        name="ppt_crop_picture",
        annotations={
            "title": "Crop Picture",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_crop_picture(params: CropPictureInput) -> str:
        """Crop a picture shape by setting crop values in points.

        Only provided crop values are updated. Returns current crop values
        (crop_left, crop_right, crop_top, crop_bottom) after applying changes.
        """
        return crop_picture(params)

    # --- Shape Export ---
    @mcp.tool(
        name="ppt_export_shape",
        annotations={
            "title": "Export Shape",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        },
    )
    async def tool_export_shape(params: ExportShapeInput) -> str:
        """Export a shape as an image file.

        Supports formats: 'png', 'jpg', 'gif', 'bmp', 'wmf', 'emf'.
        Optionally specify width and height in pixels.
        """
        return export_shape(params)

    # --- Slide Hidden ---
    @mcp.tool(
        name="ppt_set_slide_hidden",
        annotations={
            "title": "Set Slide Hidden",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_set_slide_hidden(params: SetSlideHiddenInput) -> str:
        """Set a slide as hidden or visible in the slideshow.

        Hidden slides are skipped during slideshow playback but remain
        in the presentation. Set hidden=true to hide, hidden=false to show.
        """
        return set_slide_hidden(params)

    # --- Select Shapes ---
    @mcp.tool(
        name="ppt_select_shapes",
        annotations={
            "title": "Select Shapes",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_select_shapes(params: SelectShapesInput) -> str:
        """Select multiple shapes on a slide by name.

        Navigates to the specified slide and selects the listed shapes.
        The first shape replaces any existing selection; remaining shapes
        are added to the selection.
        """
        return select_shapes(params)

    # --- Get Selection ---
    @mcp.tool(
        name="ppt_get_selection",
        annotations={
            "title": "Get Selection",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_get_selection() -> str:
        """Get the current selection in the active PowerPoint window.

        Returns the selection type (none, slides, shapes, text).
        For shapes, returns the list of selected shape names.
        For text, returns the selected text content.
        """
        return get_selection()

    # --- View ---
    @mcp.tool(
        name="ppt_set_view",
        annotations={
            "title": "Set View",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_set_view(params: SetViewInput) -> str:
        """Set the PowerPoint view type and/or zoom level.

        View types: 'normal', 'slide_master', 'notes_page', 'handout_master',
        'notes_master', 'outline', 'slide_sorter', 'title_master', 'reading'.
        Zoom range: 10-400. Returns current view_type and zoom after setting.
        """
        return set_view(params)

    # --- Copy Animation ---
    @mcp.tool(
        name="ppt_copy_animation",
        annotations={
            "title": "Copy Animation",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_copy_animation(params: CopyAnimationInput) -> str:
        """Copy animation effects from one shape to another on the same slide.

        Uses PickupAnimation/ApplyAnimation to transfer all animation
        settings from the source shape to the target shape.
        """
        return copy_animation(params)

    # --- Add Picture from URL ---
    @mcp.tool(
        name="ppt_add_picture_from_url",
        annotations={
            "title": "Add Picture from URL",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        },
    )
    async def tool_add_picture_from_url(params: AddPictureFromUrlInput) -> str:
        """Add a picture to a slide by downloading from a URL.

        Downloads the image to a temporary file, inserts it into the slide,
        and cleans up the temp file. Supports SVG files with optional
        currentColor replacement via svg_color. If fit=true with both
        width and height, the image is fitted within the area preserving
        aspect ratio and centered. If width/height are not specified,
        the original image dimensions are used.
        """
        return add_picture_from_url(params)

    # --- Add SVG Icon ---
    @mcp.tool(
        name="ppt_add_svg_icon",
        annotations={
            "title": "Add SVG Icon",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        },
    )
    async def tool_add_svg_icon(params: AddSvgIconInput) -> str:
        """Add a Material Symbols icon as SVG image to a slide.

        Downloads the icon from the Google Material Symbols CDN and inserts
        it fitted within the given area preserving aspect ratio.
        Icon styles: 'outlined', 'rounded', 'sharp'. Set filled=true for
        the filled variant. Color accepts '#RRGGBB' or theme names like
        'accent1'. Use ppt_search_icons to find icon names by keyword.
        """
        return add_svg_icon(params)

    # --- Lock Aspect Ratio ---
    @mcp.tool(
        name="ppt_lock_aspect_ratio",
        annotations={
            "title": "Lock Aspect Ratio",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_lock_aspect_ratio(params: LockAspectRatioInput) -> str:
        """Lock or unlock the aspect ratio of a shape.

        When locked, resizing the shape maintains its proportions.
        Set locked=true to lock, locked=false to unlock.
        """
        return lock_aspect_ratio(params)

    # --- Search Icons ---
    @mcp.tool(
        name="ppt_search_icons",
        annotations={
            "title": "Search Material Icons",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": True,
        },
    )
    async def tool_search_icons(params: SearchIconsInput) -> str:
        """Search Google's Material Symbols icon library by keyword.

        Returns matching icon names sorted by relevance (name, tags, and
        category matching + popularity). Each result includes the icon
        name, categories, sample tags, and popularity score.
        Use the returned icon name with ppt_add_svg_icon to insert it
        into a slide. Supports multi-word queries (e.g. 'arrow forward',
        'chart graph'). The metadata is fetched on first call and cached
        for 24 hours.
        """
        return search_icons(params)
