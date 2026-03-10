"""SmartArt operations for PowerPoint COM automation.

Handles creating SmartArt graphics, modifying nodes (set text, add, delete,
format), applying color schemes and quick styles, and listing available
SmartArt layouts, color schemes, and quick styles.
"""

import json
import logging
from typing import Optional, Union

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from utils.color import hex_to_int
from utils.navigation import goto_slide
from utils.validation import font_size_warning
from ppt_com.constants import msoSmartArt, SHAPE_TYPE_NAMES, msoTrue, msoFalse

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# English alias mapping: SmartArt layout Id (GUID) -> English name
# ---------------------------------------------------------------------------
# SmartArt layout names depend on the OS locale (e.g. Japanese Windows shows
# "基本プロセス" instead of "Basic Process").  The layout Id is a stable URN
# that does not change across locales.  This dict maps well-known Ids to their
# canonical English names so that callers can use English names regardless of
# the Windows display language.
#
# The mapping covers the most commonly used layouts.  Layouts not listed here
# will still work via their locale-specific name or index.
SMARTART_ENGLISH_ALIASES: dict[str, str] = {
    # --- List ---
    "urn:microsoft.com/office/officeart/2005/8/layout/vList2": "Basic Block List",
    "urn:microsoft.com/office/officeart/2005/8/layout/list1": "Lined List",
    "urn:microsoft.com/office/officeart/2005/8/layout/vList5": "Vertical Box List",
    "urn:microsoft.com/office/officeart/2005/8/layout/hList1": "Horizontal Bullet List",
    "urn:microsoft.com/office/officeart/2005/8/layout/vList1": "Vertical Bullet List",
    "urn:microsoft.com/office/officeart/2005/8/layout/default": "Basic List",
    "urn:microsoft.com/office/officeart/2005/8/layout/vList4": "Stacked List",
    "urn:microsoft.com/office/officeart/2005/8/layout/hList3": "Table List",
    "urn:microsoft.com/office/officeart/2005/8/layout/lProcess": "Vertical Accent List",
    "urn:microsoft.com/office/officeart/2005/8/layout/chevron2": "Grouped List",
    "urn:microsoft.com/office/officeart/2005/8/layout/hList6": "Square Accent List",
    "urn:microsoft.com/office/officeart/2005/8/layout/hierarchy5": "Horizontal Picture List",
    "urn:microsoft.com/office/officeart/2005/8/layout/vList6": "Picture Accent List",
    "urn:microsoft.com/office/officeart/2009/3/layout/PictureAccentBlocks": "Picture Accent Blocks",
    "urn:microsoft.com/office/officeart/2005/8/layout/hList7": "Trapezoid List",
    "urn:microsoft.com/office/officeart/2005/8/layout/vList3": "Vertical Arrow List",
    "urn:microsoft.com/office/officeart/2005/8/layout/target3": "Target List",
    "urn:microsoft.com/office/officeart/2005/8/layout/hList4": "Tab List",
    "urn:microsoft.com/office/officeart/2005/8/layout/chevron1": "Vertical Chevron List",
    "urn:microsoft.com/office/officeart/2005/8/layout/hList2": "Horizontal Labeled Hierarchy",
    # --- Process ---
    "urn:microsoft.com/office/officeart/2005/8/layout/process1": "Basic Process",
    "urn:microsoft.com/office/officeart/2005/8/layout/process2": "Step Up Process",
    "urn:microsoft.com/office/officeart/2005/8/layout/process3": "Step Down Process",
    "urn:microsoft.com/office/officeart/2005/8/layout/process4": "Accent Process",
    "urn:microsoft.com/office/officeart/2005/8/layout/process5": "Alternating Flow",
    "urn:microsoft.com/office/officeart/2005/8/layout/chevron3": "Continuous Block Process",
    "urn:microsoft.com/office/officeart/2005/8/layout/arrow2": "Process Arrows",
    "urn:microsoft.com/office/officeart/2005/8/layout/arrow6": "Circular Bending Process",
    "urn:microsoft.com/office/officeart/2005/8/layout/arrow1": "Basic Bending Process",
    "urn:microsoft.com/office/officeart/2005/8/layout/arrow3": "Upward Arrow",
    "urn:microsoft.com/office/officeart/2005/8/layout/arrow5": "Converging Arrows",
    "urn:microsoft.com/office/officeart/2005/8/layout/arrow4": "Diverging Arrows",
    "urn:microsoft.com/office/officeart/2005/8/layout/chevron4": "Chevron List",
    "urn:microsoft.com/office/officeart/2005/8/layout/vProcess5": "Descending Process",
    "urn:microsoft.com/office/officeart/2005/8/layout/funnel1": "Funnel",
    "urn:microsoft.com/office/officeart/2005/8/layout/gear1": "Gear",
    "urn:microsoft.com/office/officeart/2005/8/layout/hChevron3": "Phased Process",
    "urn:microsoft.com/office/officeart/2005/8/layout/arrow7": "Repeating Bending Process",
    "urn:microsoft.com/office/officeart/2005/8/layout/equation1": "Basic Equation",
    "urn:microsoft.com/office/officeart/2005/8/layout/equation2": "Vertical Equation",
    # --- Cycle ---
    "urn:microsoft.com/office/officeart/2005/8/layout/cycle1": "Basic Cycle",
    "urn:microsoft.com/office/officeart/2005/8/layout/cycle2": "Text Cycle",
    "urn:microsoft.com/office/officeart/2005/8/layout/cycle3": "Block Cycle",
    "urn:microsoft.com/office/officeart/2005/8/layout/cycle4": "Nondirectional Cycle",
    "urn:microsoft.com/office/officeart/2005/8/layout/cycle5": "Continuous Cycle",
    "urn:microsoft.com/office/officeart/2005/8/layout/cycle6": "Multidirectional Cycle",
    "urn:microsoft.com/office/officeart/2005/8/layout/cycle7": "Segmented Cycle",
    "urn:microsoft.com/office/officeart/2005/8/layout/radial1": "Basic Radial",
    "urn:microsoft.com/office/officeart/2005/8/layout/radial2": "Diverging Radial",
    "urn:microsoft.com/office/officeart/2005/8/layout/radial3": "Radial Venn",
    "urn:microsoft.com/office/officeart/2005/8/layout/radial4": "Radial Cluster",
    # --- Hierarchy ---
    "urn:microsoft.com/office/officeart/2005/8/layout/orgChart1": "Organization Chart",
    "urn:microsoft.com/office/officeart/2005/8/layout/hierarchy1": "Hierarchy",
    "urn:microsoft.com/office/officeart/2005/8/layout/hierarchy2": "Labeled Hierarchy",
    "urn:microsoft.com/office/officeart/2005/8/layout/hierarchy4": "Horizontal Hierarchy",
    "urn:microsoft.com/office/officeart/2005/8/layout/hierarchy3": "Table Hierarchy",
    "urn:microsoft.com/office/officeart/2005/8/layout/hierarchy6": "Horizontal Organization Chart",
    "urn:microsoft.com/office/officeart/2005/8/layout/hierarchy7": "Horizontal Multi-Level Hierarchy",
    "urn:microsoft.com/office/officeart/2009/3/layout/HalfCircleOrganizationChart": "Half Circle Organization Chart",
    "urn:microsoft.com/office/officeart/2005/8/layout/lProcess2": "Lined Hierarchy",
    # --- Relationship ---
    "urn:microsoft.com/office/officeart/2005/8/layout/venn1": "Basic Venn",
    "urn:microsoft.com/office/officeart/2005/8/layout/venn2": "Linear Venn",
    "urn:microsoft.com/office/officeart/2005/8/layout/venn3": "Stacked Venn",
    "urn:microsoft.com/office/officeart/2005/8/layout/target1": "Basic Target",
    "urn:microsoft.com/office/officeart/2005/8/layout/target2": "Nested Target",
    "urn:microsoft.com/office/officeart/2005/8/layout/balance1": "Balance",
    "urn:microsoft.com/office/officeart/2005/8/layout/opposingArrows1": "Opposing Arrows",
    "urn:microsoft.com/office/officeart/2005/8/layout/opposingIdeas1": "Opposing Ideas",
    "urn:microsoft.com/office/officeart/2005/8/layout/arrow8": "Counterbalance Arrows",
    "urn:microsoft.com/office/officeart/2005/8/layout/equation3": "Formula",
    # --- Matrix ---
    "urn:microsoft.com/office/officeart/2005/8/layout/matrix1": "Basic Matrix",
    "urn:microsoft.com/office/officeart/2005/8/layout/matrix2": "Titled Matrix",
    "urn:microsoft.com/office/officeart/2005/8/layout/matrix3": "Grid Matrix",
    "urn:microsoft.com/office/officeart/2005/8/layout/cycle8": "Cycle Matrix",
    # --- Pyramid ---
    "urn:microsoft.com/office/officeart/2005/8/layout/pyramid1": "Basic Pyramid",
    "urn:microsoft.com/office/officeart/2005/8/layout/pyramid2": "Inverted Pyramid",
    "urn:microsoft.com/office/officeart/2005/8/layout/pyramid3": "Pyramid List",
    "urn:microsoft.com/office/officeart/2005/8/layout/pyramid4": "Segmented Pyramid",
    # --- Picture ---
    "urn:microsoft.com/office/officeart/2005/8/layout/hList5": "Picture Lineup",
    "urn:microsoft.com/office/officeart/2005/8/layout/lProcess3": "Continuous Picture List",
    "urn:microsoft.com/office/officeart/2005/8/layout/vList7": "Vertical Picture Accent List",
    "urn:microsoft.com/office/officeart/2005/8/layout/vList8": "Alternating Picture Blocks",
    "urn:microsoft.com/office/officeart/2009/3/layout/TitlePictureLineup": "Title Picture Lineup",
    "urn:microsoft.com/office/officeart/2005/8/layout/hProcess9": "Picture Accent Process",
    "urn:microsoft.com/office/officeart/2009/3/layout/SnapshotPictureList": "Snapshot Picture List",
    "urn:microsoft.com/office/officeart/2009/3/layout/BubblePictureList": "Bubble Picture List",
    "urn:microsoft.com/office/officeart/2009/3/layout/PictureStrips": "Picture Strips",
    "urn:microsoft.com/office/officeart/2009/3/layout/CirclePictureHierarchy": "Circle Picture Hierarchy",
    "urn:microsoft.com/office/officeart/2005/8/layout/radial6": "Radial Picture List",
    # --- Other common ones ---
    "urn:microsoft.com/office/officeart/2005/8/layout/hProcess4": "Staggered Process",
    "urn:microsoft.com/office/officeart/2005/8/layout/hProcess5": "Process List",
    "urn:microsoft.com/office/officeart/2005/8/layout/hProcess7": "Circle Accent Timeline",
    "urn:microsoft.com/office/officeart/2005/8/layout/hProcess8": "Basic Timeline",
    "urn:microsoft.com/office/officeart/2005/8/layout/hProcess11": "Interconnected Block Process",
    "urn:microsoft.com/office/officeart/2009/3/layout/ArchitectureLayout": "Architecture Layout",
}

# Build a reverse mapping: lowercase English name -> layout Id
_ENGLISH_NAME_TO_ID: dict[str, str] = {
    v.lower(): k for k, v in SMARTART_ENGLISH_ALIASES.items()
}


def _resolve_layout(app, layout_name: str):
    """Resolve a SmartArt layout by name with English alias support.

    Search order:
    1. Exact match on locale-specific Name (case-insensitive, contains)
    2. Match via English alias name -> layout Id lookup

    Returns the SmartArtLayout COM object, or raises ValueError if not found.
    """
    name_lower = layout_name.lower()
    layouts = app.SmartArtLayouts

    # 1) Try locale-specific name (contains match, case-insensitive)
    for j in range(1, layouts.Count + 1):
        if name_lower in layouts(j).Name.lower():
            return layouts(j)

    # 2) Try English alias: check if the search string matches any English name
    #    and resolve via layout Id
    matched_ids = set()
    for eng_name, layout_id in _ENGLISH_NAME_TO_ID.items():
        if name_lower in eng_name:
            matched_ids.add(layout_id)

    if matched_ids:
        for j in range(1, layouts.Count + 1):
            try:
                lid = layouts(j).Id
            except Exception:
                continue
            if lid in matched_ids:
                return layouts(j)

    raise ValueError(
        f"SmartArt layout '{layout_name}' not found. "
        "Use ppt_list_smartart_layouts to see available layouts with English names."
    )


def _get_english_name(layout_id: str) -> Optional[str]:
    """Return the English alias for a layout Id, or None if not mapped."""
    return SMARTART_ENGLISH_ALIASES.get(layout_id)


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class AddSmartArtInput(BaseModel):
    """Input for adding a SmartArt graphic to a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    layout_name: Optional[str] = Field(
        default=None,
        description=(
            "Partial or full SmartArt layout name to search for (case-insensitive match). "
            "Accepts English names (e.g. 'Basic Process') regardless of OS locale."
        ),
    )
    layout_index: Optional[int] = Field(
        default=None, ge=1,
        description="1-based SmartArt layout index (used if layout_name is not provided, defaults to 1)",
    )
    left: float = Field(default=50.0, description="Left position in points")
    top: float = Field(default=50.0, description="Top position in points")
    width: float = Field(default=600.0, description="Width in points")
    height: float = Field(default=400.0, description="Height in points")
    node_texts: Optional[list[str]] = Field(
        default=None,
        description="List of text strings to populate SmartArt nodes in order",
    )
    color_index: Optional[int] = Field(
        default=None, ge=1,
        description="1-based color scheme index (from Application.SmartArtColors). Use ppt_list_smartart_layouts with list_type='colors' to find indices.",
    )
    style_index: Optional[int] = Field(
        default=None, ge=1,
        description="1-based quick style index (from Application.SmartArtQuickStyles). Use ppt_list_smartart_layouts with list_type='styles' to find indices. Applied before color_index.",
    )
    font_name: Optional[str] = Field(
        default=None,
        description="Font name to apply to all nodes (sets both Latin and East Asian font). E.g. 'BIZ UDPゴシック'.",
    )
    font_size: Optional[float] = Field(
        default=None, gt=0,
        description="Font size in points to apply to all nodes.",
    )
    bold: Optional[bool] = Field(
        default=None,
        description="Bold on/off for all nodes.",
    )
    font_color: Optional[str] = Field(
        default=None,
        description="Text color for all nodes as '#RRGGBB'. E.g. '#FFFFFF' for white.",
    )


class ModifySmartArtInput(BaseModel):
    """Input for modifying a SmartArt node or the SmartArt as a whole."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="SmartArt shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    action: str = Field(
        ..., description=(
            "Action to perform: "
            "'set_text' (set node text), "
            "'add_node' (add a new node), "
            "'delete_node' (remove a node), "
            "'change_layout' (switch to a different SmartArt layout; requires layout_index or layout_name), "
            "'change_color' (apply color scheme; requires color_index), "
            "'change_style' (apply quick style; requires style_index; optionally also applies color_index), "
            "'format_node' (set fill/line/font on one node; requires node_index), "
            "'format_all_nodes' (apply fill/line/font to every node)"
        )
    )
    node_index: Optional[int] = Field(
        default=None, ge=1,
        description="1-based node index in AllNodes. Required for 'set_text', 'delete_node', and 'format_node'. For 'add_node', specifies the node to add after (omit to append).",
    )
    text: Optional[str] = Field(
        default=None,
        description="Text to set on the node. Required for 'set_text', optional for 'add_node'.",
    )
    # --- layout change fields ---
    layout_name: Optional[str] = Field(
        default=None,
        description=(
            "Partial/full layout name for 'change_layout' (case-insensitive). "
            "Accepts English names (e.g. 'Basic Process') regardless of OS locale. "
            "Use ppt_list_smartart_layouts to find names."
        ),
    )
    layout_index: Optional[int] = Field(
        default=None, ge=1,
        description="1-based layout index for 'change_layout'. Use ppt_list_smartart_layouts to find indices.",
    )
    # --- styling fields ---
    color_index: Optional[int] = Field(
        default=None, ge=1,
        description="1-based color scheme index. Required for 'change_color'. Also applied after style change when used with 'change_style'.",
    )
    style_index: Optional[int] = Field(
        default=None, ge=1,
        description="1-based quick style index. Required for 'change_style'.",
    )
    font_name: Optional[str] = Field(
        default=None,
        description="Font name (sets both Latin and East Asian). Used with 'format_node' and 'format_all_nodes'.",
    )
    font_size: Optional[float] = Field(
        default=None, gt=0,
        description="Font size in points. Used with 'format_node' and 'format_all_nodes'.",
    )
    bold: Optional[bool] = Field(
        default=None,
        description="Bold on/off. Used with 'format_node' and 'format_all_nodes'.",
    )
    fill_color: Optional[str] = Field(
        default=None,
        description="Node fill color as '#RRGGBB'. Used with 'format_node' and 'format_all_nodes'.",
    )
    line_color: Optional[str] = Field(
        default=None,
        description="Node border color as '#RRGGBB'. Used with 'format_node' and 'format_all_nodes'.",
    )
    line_width: Optional[float] = Field(
        default=None, gt=0,
        description="Node border width in points. Used with 'format_node' and 'format_all_nodes'.",
    )
    font_color: Optional[str] = Field(
        default=None,
        description="Text color as '#RRGGBB'. Used with 'format_node' and 'format_all_nodes'.",
    )


class ListSmartArtInput(BaseModel):
    """Input for listing SmartArt layouts, color schemes, or quick styles."""
    model_config = ConfigDict(str_strip_whitespace=True)

    list_type: str = Field(
        default="layouts",
        description=(
            "What to list: "
            "'layouts' (SmartArt diagram layouts), "
            "'colors' (color schemes), "
            "'styles' (quick style templates), "
            "'categories' (distinct layout category names — use these values with the category filter)."
        ),
    )
    category: Optional[str] = Field(
        default=None,
        description=(
            "Filter layouts by category (exact or partial match, case-insensitive). "
            "Only applies to list_type='layouts'. "
            "Use list_type='categories' first to discover available category names in the current locale."
        ),
    )
    keyword: Optional[str] = Field(
        default=None,
        description=(
            "Filter by keyword in name (partial match, case-insensitive). "
            "Also matches English alias names for layouts. Applies to all list_types."
        ),
    )
    include_description: bool = Field(
        default=False,
        description="Include the description field in each entry. Omitted by default to keep output compact.",
    )


# ---------------------------------------------------------------------------
# Helper: find a shape on a slide
# ---------------------------------------------------------------------------
def _get_shape(slide, name_or_index: Union[str, int]):
    """Find a shape on a slide by name or 1-based index."""
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


def _apply_node_format(node, fill_color, line_color, line_width, font_name, font_size, bold, font_color):
    """Apply fill/line/font formatting to a single SmartArtNode."""
    if fill_color is not None or line_color is not None or line_width is not None:
        if node.Shapes.Count > 0:
            sh = node.Shapes(1)
            if fill_color is not None:
                sh.Fill.ForeColor.RGB = hex_to_int(fill_color)
            if line_color is not None:
                sh.Line.ForeColor.RGB = hex_to_int(line_color)
            if line_width is not None:
                sh.Line.Weight = line_width

    if font_name is not None or font_size is not None or bold is not None or font_color is not None:
        f = node.TextFrame2.TextRange.Font
        if font_size is not None:
            f.Size = font_size
        if font_name is not None:
            f.Name = font_name
            f.NameFarEast = font_name
        if bold is not None:
            f.Bold = msoTrue if bold else msoFalse
        if font_color is not None:
            f.Fill.ForeColor.RGB = hex_to_int(font_color)


# ---------------------------------------------------------------------------
# COM implementation functions (run on COM thread via ppt.execute)
# ---------------------------------------------------------------------------
def _add_smartart_impl(slide_index, layout_name, layout_index, left, top, width, height,
                       node_texts, color_index, style_index, font_name, font_size, bold, font_color):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)

    # Resolve the SmartArt layout
    if layout_name:
        layout = _resolve_layout(app, layout_name)
    else:
        idx = layout_index if layout_index else 1
        layout = app.SmartArtLayouts(idx)

    resolved_layout_name = layout.Name
    shape = slide.Shapes.AddSmartArt(layout, left, top, width, height)

    # Set node texts if provided
    if node_texts:
        smart_art = shape.SmartArt
        for i, text in enumerate(node_texts):
            if i < smart_art.AllNodes.Count:
                smart_art.AllNodes(i + 1).TextFrame2.TextRange.Text = text
            else:
                try:
                    node = smart_art.AllNodes(smart_art.AllNodes.Count).AddNode()
                    node.TextFrame2.TextRange.Text = text
                except Exception:
                    break  # Some layouts have fixed node counts

    # Apply styling (QuickStyle must come before Color — setting QuickStyle resets Color)
    smart_art = shape.SmartArt
    if style_index is not None:
        smart_art.QuickStyle = app.SmartArtQuickStyles(style_index)
    if color_index is not None:
        smart_art.Color = app.SmartArtColors(color_index)

    # Apply font to all nodes
    if font_name is not None or font_size is not None or bold is not None or font_color is not None:
        for i in range(1, smart_art.AllNodes.Count + 1):
            _apply_node_format(
                smart_art.AllNodes(i),
                None, None, None,  # no fill/line at creation via these params
                font_name, font_size, bold, font_color,
            )

    return {
        "success": True,
        "shape_name": shape.Name,
        "node_count": smart_art.AllNodes.Count,
        "layout_name": resolved_layout_name,
    }


def _modify_smartart_impl(slide_index, shape_name_or_index, action,
                          node_index, text,
                          layout_name, layout_index,
                          color_index, style_index,
                          font_name, font_size, bold, font_color,
                          fill_color, line_color, line_width):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    if shape.Type != msoSmartArt:
        raise ValueError(
            f"Shape '{shape.Name}' is not a SmartArt graphic "
            f"(type={SHAPE_TYPE_NAMES.get(shape.Type, shape.Type)})"
        )

    smart_art = shape.SmartArt

    if action == "set_text":
        if node_index is None:
            raise ValueError("node_index is required for 'set_text' action")
        if text is None:
            raise ValueError("text is required for 'set_text' action")
        smart_art.AllNodes(node_index).TextFrame2.TextRange.Text = text

    elif action == "add_node":
        if node_index and node_index <= smart_art.AllNodes.Count:
            new_node = smart_art.AllNodes(node_index).AddNode()
        else:
            new_node = smart_art.AllNodes(smart_art.AllNodes.Count).AddNode()
        if text:
            new_node.TextFrame2.TextRange.Text = text

    elif action == "delete_node":
        if node_index is None:
            raise ValueError("node_index is required for 'delete_node' action")
        smart_art.AllNodes(node_index).Delete()

    elif action == "change_layout":
        if layout_name:
            layout = _resolve_layout(app, layout_name)
        elif layout_index:
            layout = app.SmartArtLayouts(layout_index)
        else:
            raise ValueError("layout_name or layout_index is required for 'change_layout' action")
        smart_art.Layout = layout

    elif action == "change_color":
        if color_index is None:
            raise ValueError("color_index is required for 'change_color' action")
        smart_art.Color = app.SmartArtColors(color_index)

    elif action == "change_style":
        if style_index is None:
            raise ValueError("style_index is required for 'change_style' action")
        # QuickStyle must be set before Color — setting QuickStyle resets Color to theme default
        smart_art.QuickStyle = app.SmartArtQuickStyles(style_index)
        if color_index is not None:
            smart_art.Color = app.SmartArtColors(color_index)

    elif action == "format_node":
        if node_index is None:
            raise ValueError("node_index is required for 'format_node' action")
        node = smart_art.AllNodes(node_index)
        _apply_node_format(node, fill_color, line_color, line_width, font_name, font_size, bold, font_color)

    elif action == "format_all_nodes":
        for i in range(1, smart_art.AllNodes.Count + 1):
            _apply_node_format(
                smart_art.AllNodes(i),
                fill_color, line_color, line_width,
                font_name, font_size, bold, font_color,
            )

    else:
        raise ValueError(
            f"Unknown action '{action}'. Supported: "
            "'set_text', 'add_node', 'delete_node', "
            "'change_color', 'change_style', "
            "'format_node', 'format_all_nodes'"
        )

    return {
        "success": True,
        "action": action,
        "total_nodes": smart_art.AllNodes.Count,
    }


def _list_smartart_options_impl(list_type, category, keyword, include_description):
    app = ppt._get_app_impl()

    # --- categories: return distinct category names from SmartArtLayouts ---
    if list_type == "categories":
        collection = app.SmartArtLayouts
        seen = {}
        for i in range(1, collection.Count + 1):
            try:
                cat = collection(i).Category or ""
            except Exception:
                cat = ""
            if cat and cat not in seen:
                seen[cat] = 0
            if cat:
                seen[cat] += 1
        categories = [{"category": k, "count": v} for k, v in seen.items()]
        return {
            "success": True,
            "list_type": "categories",
            "total_layouts": collection.Count,
            "categories": categories,
        }

    if list_type == "layouts":
        collection = app.SmartArtLayouts
        key = "layouts"
    elif list_type == "colors":
        collection = app.SmartArtColors
        key = "colors"
    elif list_type == "styles":
        collection = app.SmartArtQuickStyles
        key = "styles"
    else:
        raise ValueError(f"Unknown list_type '{list_type}'. Use: 'layouts', 'colors', 'styles', or 'categories'")

    total = collection.Count
    cat_lower = category.lower() if category else None
    kw_lower = keyword.lower() if keyword else None

    items = []
    for i in range(1, total + 1):
        item = collection(i)
        item_name = item.Name

        # For layouts, resolve the English alias name via layout Id
        english_name = None
        if list_type == "layouts":
            try:
                english_name = _get_english_name(item.Id)
            except Exception:
                pass

        # Category filter (layouts only)
        if cat_lower and list_type == "layouts":
            try:
                item_cat = item.Category or ""
            except Exception:
                item_cat = ""
            if cat_lower not in item_cat.lower():
                continue

        # Keyword filter on name — also check English alias name for layouts
        if kw_lower:
            name_match = kw_lower in item_name.lower()
            eng_match = english_name and kw_lower in english_name.lower()
            if not name_match and not eng_match:
                continue

        entry = {"index": i, "name": item_name}
        if list_type == "layouts":
            if english_name:
                entry["english_name"] = english_name
            try:
                entry["category"] = item.Category
            except Exception:
                pass
        if include_description:
            entry["description"] = item.Description

        items.append(entry)

    return {
        "success": True,
        "list_type": list_type,
        "total_count": total,
        "filtered_count": len(items),
        key: items,
    }


# ---------------------------------------------------------------------------
# MCP tool functions (sync wrappers that delegate to COM thread)
# ---------------------------------------------------------------------------
def add_smartart(params: AddSmartArtInput) -> str:
    try:
        result = ppt.execute(
            _add_smartart_impl,
            params.slide_index, params.layout_name, params.layout_index,
            params.left, params.top, params.width, params.height,
            params.node_texts,
            params.color_index, params.style_index,
            params.font_name, params.font_size, params.bold, params.font_color,
        )
        warn = font_size_warning(params.font_size)
        if warn:
            result["warning"] = warn
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to add SmartArt: {str(e)}"})


def modify_smartart(params: ModifySmartArtInput) -> str:
    try:
        result = ppt.execute(
            _modify_smartart_impl,
            params.slide_index, params.shape_name_or_index,
            params.action, params.node_index, params.text,
            params.layout_name, params.layout_index,
            params.color_index, params.style_index,
            params.font_name, params.font_size, params.bold, params.font_color,
            params.fill_color, params.line_color, params.line_width,
        )
        warn = font_size_warning(params.font_size)
        if warn:
            result["warning"] = warn
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to modify SmartArt: {str(e)}"})


def list_smartart_options(params: ListSmartArtInput) -> str:
    try:
        result = ppt.execute(
            _list_smartart_options_impl,
            params.list_type, params.category, params.keyword, params.include_description,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to list SmartArt options: {str(e)}"})


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register all SmartArt tools with the MCP server."""

    @mcp.tool(
        name="ppt_add_smartart",
        annotations={
            "title": "Add SmartArt",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_add_smartart(params: AddSmartArtInput) -> str:
        """Add a SmartArt graphic to a slide.

        Creates a SmartArt with the specified layout and optionally populates
        node text. Find layouts using ppt_list_smartart_layouts.

        layout_name accepts English names (e.g. 'Basic Process', 'Organization Chart')
        regardless of OS locale. On non-English Windows, the tool first tries the
        locale-specific name, then falls back to the English alias mapping.

        Styling at creation time:
        - color_index: apply a color scheme (use list_type='colors' to find indices)
        - style_index: apply a quick style template (use list_type='styles')
          NOTE: style is applied before color — set both to get both effects.
        - font_name, font_size, bold, font_color: apply to all nodes at once.

        All positions and sizes are in points (72 points = 1 inch).
        """
        return add_smartart(params)

    @mcp.tool(
        name="ppt_modify_smartart",
        annotations={
            "title": "Modify SmartArt",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_modify_smartart(params: ModifySmartArtInput) -> str:
        """Modify a SmartArt graphic.

        Actions:
        - 'set_text': update text of a node (requires node_index, text)
        - 'add_node': add a new node (optional node_index to insert after, optional text)
        - 'delete_node': remove a node (requires node_index)
        - 'change_layout': switch to a different SmartArt layout (requires layout_index or
          layout_name). Accepts English names regardless of OS locale.
          Use ppt_list_smartart_layouts to find layouts. Node texts and
          count are preserved where the new layout allows.
        - 'change_color': apply a color scheme (requires color_index)
        - 'change_style': apply a quick style (requires style_index; also applies
          color_index if provided — QuickStyle resets Color, so set both together)
        - 'format_node': set fill/line/font on one node (requires node_index;
          use fill_color, line_color, line_width, font_name, font_size, bold)
        - 'format_all_nodes': apply fill/line/font to every node (same fields as format_node)

        Colors: '#RRGGBB' hex strings. Use ppt_list_smartart_layouts with
        list_type='colors' or list_type='styles' to discover available indices.
        """
        return modify_smartart(params)

    @mcp.tool(
        name="ppt_list_smartart_layouts",
        annotations={
            "title": "List SmartArt Layouts / Colors / Styles",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_list_smartart_layouts(params: ListSmartArtInput) -> str:
        """List SmartArt layouts, color schemes, or quick styles.

        list_type:
        - 'layouts' (default): all diagram layout templates.
          Each layout includes an english_name field (when available) so you can
          identify layouts regardless of OS locale.
        - 'colors': color schemes — use color_index with ppt_add/modify_smartart
        - 'styles': quick style templates — use style_index with ppt_add/modify_smartart
        - 'categories': distinct category names for the current locale — use these
          values with the category filter to narrow 'layouts' results

        keyword: filter by keyword in name (partial match, case-insensitive).
          Also matches English alias names for layouts.

        category: filter layouts by category (partial match, case-insensitive).
        include_description: set True to add verbose description text (off by default).

        Output is compact by default (index + name + english_name + category).
        """
        return list_smartart_options(params)
