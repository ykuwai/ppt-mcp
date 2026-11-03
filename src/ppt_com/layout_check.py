"""Layout checks: what is wrong with a slide, measured rather than looked at.

A slide preview shows a model what it built, but only a model that can read
images, and only when it thinks to ask. Most of what goes wrong on a generated
slide is geometry, and geometry can be answered with numbers: text that runs
into another shape, text that does not fit its box, a shape hanging off the
slide, three cards whose gaps are 24, 24 and 31 points, a title that sits three
points lower than the one on the slide before it.

This module answers those questions in two places. ``ppt_check_layout`` checks
one slide or the whole deck on request. The tools that create, move, resize or
retext a shape run the per-shape checks on the shape they just touched and put
what they find in their own result, so a problem is reported by the call that
caused it, while the caller still knows why it made that call. Setting
``PPT_LAYOUT_CHECKS=off`` turns the second off, for a write-heavy session that
would rather pay for the check once a slide is done.

The arithmetic is kept apart from the reading, so it can be tested without
PowerPoint and so both platforms answer the same way. ``_slide_facts_impl``
reads a slide into plain dicts; everything after that is pure.
"""

import json
import logging
import math
import os
import re
import tempfile
import unicodedata
from itertools import pairwise
from typing import Optional

from pydantic import BaseModel, ConfigDict, Field

from backend import IS_MACOS, ppt
from utils.color import int_to_hex
from utils.navigation import goto_slide
from utils.offload import run_offloaded
from utils.package import package_facts, picture_aspects  # noqa: F401
from ppt_com.constants import (
    msoChart, msoEmbeddedOLEObject, msoFillGradient, msoFillSolid, msoGroup, msoLine,
    msoLinkedOLEObject, msoLinkedPicture, msoMedia, msoPicture,
    msoPlaceholder, msoSmartArt, msoTable, msoTextBox, msoTrue,
    ppAlignCenter, ppAlignRight, ppDirectionRightToLeft,
    ppPlaceholderCenterTitle, ppPlaceholderDate, ppPlaceholderFooter,
    ppPlaceholderBody, ppPlaceholderHeader, ppPlaceholderObject,
    ppPlaceholderSlideNumber, ppPlaceholderTitle, ppPlaceholderVerticalTitle,
    msoAnimTriggerAfterPrevious, msoAnimTriggerWithPrevious,
)

logger = logging.getLogger(__name__)


def _check_each_write(value):
    """Whether the write tools check what they touched, from PPT_LAYOUT_CHECKS."""
    value = value.strip().lower()
    if value in ("on", "true", "1", "yes", ""):
        return True
    if value in ("off", "false", "0", "no"):
        return False
    # Read at import, so a typo must not stop the server from starting.
    logger.warning(
        "PPT_LAYOUT_CHECKS=%r is not a setting; use 'on' (the default), or 'off' "
        "to leave the checks to ppt_check_layout. Using 'on'.", value)
    return True


# Each check reads the slide back, which costs a call per write, and on the Mac
# several. ppt_check_layout runs either way.
CHECK_EACH_WRITE = _check_each_write(os.getenv("PPT_LAYOUT_CHECKS", "on"))


# ---------------------------------------------------------------------------
# Thresholds, in points unless they say otherwise
# ---------------------------------------------------------------------------

# A shape may sit right on an edge; a fraction of a point past it is rounding.
# The same allowance ppt_check_typography gives an off-slide edge.
_EDGE_TOLERANCE = 0.5

# PpEntryEffect values of the Morph transition: by object, word and character,
# as PowerPoint for Windows reads them from a deck.
MORPH_EFFECTS = frozenset({3954, 3955, 3956})

# Edges this close are one edge. PowerPoint stores positions in EMU and reads
# them back in points, so two shapes placed at the same number can differ in
# the second decimal.
_ALIGNED = 0.5

# Centres this close are one centre. A text box that grows or shrinks to fit
# its text keeps its top, so its middle moves by half the change, and a point
# between two centres cannot be seen the way it can between two edges.
_ALIGNED_CENTRE = 1.5

# Text this close to a side of its frame's inset keeps to that side: it is set
# left, right or centred there, or fills the frame.
_SNUG = 1.5

# Edges this close, and not closer, were meant to line up and do not. Further
# apart than this is a decision rather than a slip.
_NEAR_MISS = 4.0

# Text has to cover this much of another shape, in both directions, before it
# counts as running into it. Less is a glyph's side bearing touching a border.
_MIN_OVERLAP = 2.0

# How far apart a fix sets two things that collided: enough to read as two.
_CLEARANCE = 6.0

# How far apart the gaps in a row of like shapes can be before it shows.
_GAP_TOLERANCE = 2.0

# Shapes count as the same kind of thing, for spacing, when their widths and
# heights are within this fraction of each other.
_SIMILAR_SIZE = 0.15

# Two shapes are in the same row (or column) when they share at least this
# fraction of the smaller one's height (or width).
_SHARED_SPAN = 0.5

# Like shapes in a row whose heights (or in a column, widths) are this close
# are the same size.
_SIZE_TOLERANCE = 2.0

# Text closer than this to an edge of the card it sits on is pressed against
# it, when it has twice as much room the other way. Equally little room all
# round is a badge or a pill, and is meant.
_MIN_ROOM = 12.0

# A slide's side margin is the distance most of its shapes keep from the left
# and right edges. Only distances within this share of the slide's width
# count; a shape further in is placed, not kept to a margin.
_MARGIN_ZONE = 0.2

# Margins this close are one margin, and a shape further inside it than this
# is not keeping it.
_MARGIN_TOLERANCE = 2.0

# How far inside the margin a shape can be and still have been meant to keep
# it: this much, or a quarter of the margin when that is more. Further in is a
# decision, a corner logo against a centred composition, not a slip.
_MARGIN_NEAR = 10.0

# A shape covering this much of the slide is a backdrop, not part of the
# layout. It is not lined up against anything and it does not make the shapes
# on it its children.
_BACKDROP_SHARE = 0.9

# A fill more see-through than this mixes with whatever is under it, so its
# own colour is not the colour behind the text.
_OPAQUE_ENOUGH = 0.3

# WCAG 2.x contrast minimums. Large text is 18 pt, or 14 pt bold.
_CONTRAST_TEXT = 4.5
_CONTRAST_LARGE = 3.0
_LARGE_TEXT = 18.0
_LARGE_BOLD_TEXT = 14.0

# A deck's titles are compared once at least this many slides have one.
_MIN_TITLED_SLIDES = 3

# Title sizes this close are the same size.
_TITLE_SIZE_TOLERANCE = 1.0

# Without a title placeholder, a slide's title is its largest text in the top
# quarter of the slide, at this size or more.
_TITLE_ZONE = 0.25
_MIN_TITLE_SIZE = 24.0

# A title placeholder starting this share of the slide's width away from where
# the other titles start was put there by its layout, a caption layout's side
# column, and is styled for that place.
_TITLE_ELSEWHERE = 0.25

# A row of like shapes can have one a little wider than the rest by accident.
# Wider than this, and it is a design: a wide card beside two narrow ones.
_ALONG_SIZE_SLIP = 0.25

# A picture drawn this much wider or narrower for its height than its image
# is stretched. Less is rounding, or a crop a point off.
_STRETCH = 0.03

# How much of a shape's text the reader keeps, for the checks that read it.
_CONTENT_KEPT = 500

# PowerPoint's default text frame insets, for a frame whose margins could not
# be read.
_DEFAULT_MARGINS = {"left": 7.2, "right": 7.2, "top": 3.6, "bottom": 3.6}

# Placeholders the layout owns. They are not part of the slide's design, so
# they are not lined up, spaced, or reported as left empty.
_LAYOUT_PLACEHOLDERS = frozenset({
    ppPlaceholderSlideNumber, ppPlaceholderHeader,
    ppPlaceholderFooter, ppPlaceholderDate,
})

_TITLE_PLACEHOLDERS = frozenset({
    ppPlaceholderTitle, ppPlaceholderCenterTitle, ppPlaceholderVerticalTitle,
})

# Shapes that draw their whole box whatever their fill says.
_OPAQUE_TYPES = frozenset({
    msoPicture, msoLinkedPicture, msoTable, msoChart, msoGroup, msoMedia,
    msoSmartArt, msoEmbeddedOLEObject, msoLinkedOLEObject,
})

# The ones whose colour cannot be known, so text on them is not given a
# contrast verdict.
_UNKNOWN_BACKDROP_TYPES = frozenset({
    msoPicture, msoLinkedPicture, msoChart, msoGroup, msoMedia, msoSmartArt,
    msoEmbeddedOLEObject, msoLinkedOLEObject, msoTable,
})

# The finding types the write tools report on the shape they touched. The rest
# describe a slide or a deck rather than one shape, and a slide being built is
# allowed to be unfinished, so those wait for ppt_check_layout.
SHAPE_FINDINGS = frozenset({
    "off_slide", "overflow", "overlap", "low_contrast", "cramped_text",
    "narrow_margin", "misaligned", "uneven_sizes", "sample_text", "number_order",
    "text_direction", "mixed_direction", "stretched_picture", "title_edge", "font_missing",
    "small_text", "blurry_picture",
})

# Reported in this order, worst first.
_ORDER = (
    "off_slide", "overflow", "overlap", "stretched_picture", "blurry_picture", "low_contrast",
    "small_text", "cramped_text", "chart_overlap",
    "narrow_margin", "misaligned", "title_edge", "uneven_spacing", "uneven_sizes",
    "number_order", "text_direction", "mixed_direction", "font_missing", "empty_placeholder",
    "sample_text",
    "chart_sample_data", "missing_alt_text",
    "missing_title", "duplicate_titles",
    "inconsistent_titles",
    "inconsistent_body_text",
)


# ---------------------------------------------------------------------------
# Input model
# ---------------------------------------------------------------------------
class CheckLayoutInput(BaseModel):
    """Input for checking the layout of one slide or the whole deck."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    slide_index: Optional[int] = Field(
        default=None, ge=1,
        description=(
            "1-based slide index to check. Omit to check every slide, which "
            "also compares the slides' titles with each other."
        ),
    )


# ---------------------------------------------------------------------------
# Boxes
# ---------------------------------------------------------------------------
def _pt(value) -> str:
    """A number of points the way a person writes it: 12 pt, 12.5 pt."""
    return f"{round(value, 1):g} pt"


def _box(left, top, width, height):
    return (left, top, left + width, top + height)


def _width(box):
    return box[2] - box[0]


def _height(box):
    return box[3] - box[1]


def _shared(a, b):
    """Width and height of the area two boxes share, zero when they do not."""
    return (
        max(0.0, min(a[2], b[2]) - max(a[0], b[0])),
        max(0.0, min(a[3], b[3]) - max(a[1], b[1])),
    )


def _inside(inner, outer, tolerance=_EDGE_TOLERANCE) -> bool:
    return (
        inner[0] >= outer[0] - tolerance and inner[1] >= outer[1] - tolerance
        and inner[2] <= outer[2] + tolerance and inner[3] <= outer[3] + tolerance
    )


def _turned(shape) -> bool:
    """Whether a shape is rotated off the axes.

    Half a turn leaves the box where it was, so 180 degrees is not turned.
    """
    rotation = (shape.get("rotation") or 0.0) % 180.0
    return min(rotation, 180.0 - rotation) > 0.5


def shape_box(shape):
    """The shape's own box, or the box around it when it is rotated."""
    left, top = shape["left"], shape["top"]
    width, height = shape["width"], shape["height"]
    if not _turned(shape):
        return _box(left, top, width, height)
    angle = math.radians(shape.get("rotation") or 0.0)
    cos, sin = abs(math.cos(angle)), abs(math.sin(angle))
    turned_w = width * cos + height * sin
    turned_h = width * sin + height * cos
    cx, cy = left + width / 2, top + height / 2
    return _box(cx - turned_w / 2, cy - turned_h / 2, turned_w, turned_h)


def text_box(shape):
    """Where the shape's text is drawn, or None when it has none.

    PowerPoint reports this directly, and that is what is used. When it would
    not, the position is worked out from the frame: the margins, the vertical
    anchor and the first paragraph's alignment place a block of the measured
    width and height inside the box. Overflowing text is placed the way
    PowerPoint draws it, past the bottom for a top anchor, past both edges for
    a middle one.

    Wrapped lines end at the frame's insets, whatever PowerPoint reports: for
    right-to-left text, Hebrew or Arabic, it reports lines up to 11 pt wider
    than it draws them, past the inset and even past the frame.
    """
    text = shape.get("text")
    if not text or not text.get("chars"):
        return None
    width, height = text.get("width"), text.get("height")
    if width is None or height is None:
        return None
    if text.get("left") is not None and text.get("top") is not None:
        left, right = text["left"], text["left"] + width
        margins = text.get("margins")
        if (margins and text.get("word_wrap") is not False and not _turned(shape)
                and text.get("orientation") in (None, "horizontal")):
            inner_left = shape["left"] + margins["left"]
            inner_right = shape["left"] + shape["width"] - margins["right"]
            if max(left, inner_left) < min(right, inner_right):
                left, right = max(left, inner_left), min(right, inner_right)
        return _box(left, text["top"], right - left, height)

    margins = text.get("margins") or _DEFAULT_MARGINS
    inner_left = shape["left"] + margins["left"]
    inner_top = shape["top"] + margins["top"]
    usable_w = shape["width"] - margins["left"] - margins["right"]
    usable_h = shape["height"] - margins["top"] - margins["bottom"]

    alignment = text.get("alignment")
    if alignment == ppAlignCenter:
        left = inner_left + (usable_w - width) / 2
    elif alignment == ppAlignRight:
        left = inner_left + usable_w - width
    else:
        left = inner_left

    anchor = text.get("vertical_anchor")
    if anchor == "middle":
        top = inner_top + (usable_h - height) / 2
    elif anchor == "bottom":
        top = inner_top + usable_h - height
    else:
        top = inner_top
    return _box(left, top, width, height)


def paints_box(shape) -> bool:
    """Whether the shape draws its whole box rather than only its text.

    A filled shape, an outlined one, a picture or a table can be seen edge to
    edge. A text box with neither fill nor line is only its text, so a box
    drawn larger than its text covers nothing.

    PowerPoint for Mac does not say whether a line is drawn, so there the
    shape's kind decides. A text box or placeholder has no line unless one is
    added, and an autoshape has one unless it is taken away.
    """
    if shape["type"] in _OPAQUE_TYPES:
        return True
    fill = shape.get("fill") or {}
    if fill.get("visible"):
        return True
    line = shape.get("line_visible")
    if line is None:
        return shape["type"] not in (msoTextBox, msoPlaceholder)
    return bool(line)


def _is_line(shape) -> bool:
    return shape["type"] == msoLine or bool(shape.get("connector"))


def ink_box(shape):
    """What the shape puts on the slide: its box, its text, or nothing."""
    if not shape.get("visible", True) or _is_line(shape):
        return None
    if paints_box(shape):
        return shape_box(shape)
    return text_box(shape)


def _is_empty_placeholder(shape) -> bool:
    text = shape.get("text")
    return (
        shape.get("placeholder_type") is not None
        and text is not None
        and not text.get("chars")
    )


def _is_backdrop(shape, slide_w, slide_h) -> bool:
    if not slide_w or not slide_h:
        return False
    box = shape_box(shape)
    width, height = _shared(box, (0.0, 0.0, slide_w, slide_h))
    return width * height >= _BACKDROP_SHARE * slide_w * slide_h


def _arranged(shapes, slide_w, slide_h):
    """The shapes that make up the layout: the ones lined up and spaced.

    Not lines, not rotated shapes, not backdrops, not the layout's own footer
    placeholders, and not empty placeholders, which are reported on their own
    and draw nothing once the slide is shown.
    """
    arranged = []
    for shape in shapes:
        if not shape.get("visible", True) or _is_line(shape) or _turned(shape):
            continue
        if shape.get("placeholder_type") in _LAYOUT_PLACEHOLDERS:
            continue
        if _is_empty_placeholder(shape):
            continue
        if shape["width"] <= 0 or shape["height"] <= 0:
            continue
        if _is_backdrop(shape, slide_w, slide_h):
            continue
        arranged.append(shape)
    return arranged


# ---------------------------------------------------------------------------
# How a slide is put together: what sits on what, rows, margins
# ---------------------------------------------------------------------------
def _num(value) -> str:
    """A number of points as a tool call takes it: 640, 219.2."""
    return f"{round(value, 2):g}"


def _belongs_to(s, other, boxes) -> bool:
    """Whether ``s`` sits on ``other``, the way text sits on a card.

    Mostly inside, rather than wholly: the text on a card still belongs to the
    card when it runs past its edge, which is its own finding. A text box set
    to grow with its text can end up with a larger frame than the card it is
    on, so a frame that only holds text belongs to a shape that paints, at any
    size.
    """
    if other["type"] == msoTable:
        # Its cells hold their own text; a text box over them is in the way.
        return False
    mine, theirs = boxes[s["name"]], boxes[other["name"]]
    width, height = _shared(mine, theirs)
    if width * height < 0.5 * _width(mine) * _height(mine):
        return False
    if paints_box(s) != paints_box(other):
        # Never the other way round: a card is not inside a text frame.
        return paints_box(other)
    return _width(theirs) * _height(theirs) > _width(mine) * _height(mine)


def _depths(arranged, boxes):
    """How many shapes each one sits on: none for a card, one for its text."""
    return {
        s["name"]: sum(
            1 for other in arranged
            if other is not s and _belongs_to(s, other, boxes)
        )
        for s in arranged
    }


def _similar(a_box, b_box) -> bool:
    """Whether two boxes are about one size, in both directions."""
    aw, ah, bw, bh = _width(a_box), _height(a_box), _width(b_box), _height(b_box)
    return (abs(aw - bw) <= _SIMILAR_SIZE * max(aw, bw)
            and abs(ah - bh) <= _SIMILAR_SIZE * max(ah, bh))


def _like_groups(members, boxes):
    """Rows and columns of like shapes among ``members``, two or more each.

    Shapes are alike when their widths and heights are within
    ``_SIMILAR_SIZE`` of each other. A row shares its height and runs left to
    right, a column shares its width and runs down, and each comes sorted that
    way, as ``("horizontal", row)`` or ``("vertical", column)``.
    """
    groups = []
    for axis, direction, start in (("y", "horizontal", 0), ("x", "vertical", 1)):
        def joined(a, b, axis=axis):
            a_box, b_box = boxes[a["name"]], boxes[b["name"]]
            return _similar(a_box, b_box) and _related(a_box, b_box, axis)

        for group in _connected(members, joined):
            if len(group) >= 2:
                group.sort(key=lambda s, start=start: boxes[s["name"]][start])
                groups.append((direction, group))
    return groups


def _text_anchor(shape):
    """Which sides of its frame a text box's text keeps to, as PowerPoint drew it.

    A pair, across and down: ``"left"``, ``"center"``, ``"right"`` or
    ``"full"`` when the text fills the frame's width, and ``"top"``,
    ``"middle"``, ``"bottom"`` or ``"full"`` down it. Either is None when the
    text sits nowhere in particular, and the pair is None when the text was not
    measured. Read off where the text was drawn rather than off the paragraph
    settings, which a reader does not always fetch and which do not say how a
    frame grew around its text.
    """
    words = text_box(shape)
    if words is None:
        return None
    margins = (shape.get("text") or {}).get("margins") or _DEFAULT_MARGINS
    box = shape_box(shape)
    inner = (box[0] + margins["left"], box[1] + margins["top"],
             box[2] - margins["right"], box[3] - margins["bottom"])

    def keeps(low_gap, high_gap, low, middle, high):
        if low_gap <= _SNUG and high_gap <= _SNUG:
            return "full"
        if low_gap <= _SNUG:
            return low
        if high_gap <= _SNUG:
            return high
        # Measured text can sit a few points off the centre it was set to, so
        # the two gaps only have to be close for their size.
        if abs(low_gap - high_gap) <= max(_SNUG, 0.2 * (low_gap + high_gap)):
            return middle
        return None

    return (keeps(words[0] - inner[0], inner[2] - words[2], "left", "center", "right"),
            keeps(words[1] - inner[1], inner[3] - words[3], "top", "middle", "bottom"))


def _seen_edges(shape):
    """The edges of a shape a reader can see, or None for all of them.

    All of a shape that paints. Of a frame that only holds text, the ones its
    text keeps to: the left of left-set text, the centre of centred text. The
    rest of the frame cannot be seen, and lining it up means nothing.
    """
    if paints_box(shape):
        return None
    anchor = _text_anchor(shape)
    if anchor is None:
        # A frame with no text drawn in it shows nothing to line up with,
        # and one a write did not read cannot be judged by its frame.
        return set()
    across, down = anchor
    seen = set()
    seen |= {"left", "center", "right"} if across in (None, "full") else {across}
    seen |= {"top", "middle", "bottom"} if down in (None, "full") else {down}
    return seen


def _margins(arranged, boxes, depth, slide_w, slide_h):
    """The margins the slide's shapes keep from its edges, or None.

    The side margin is a distance top level shapes keep from both the left and
    the right edge, the one the most of them keep. A shape touching an edge
    bleeds off it on purpose and keeps no margin, so it does not count. Of a
    frame holding only text, the sides its text keeps to count, and a wide one
    set to a side counts on both, the way a title frame is drawn across the
    slide. A centred one counts on neither: centred text keeps no margin.

    A title and a footer rarely keep the same distance from the top and the
    bottom, so the margin there is the one the topmost shape keeps, usually
    the title, and no more than the side margin.
    """
    if not slide_w or not slide_h:
        return None
    lefts, rights, tops = [], [], []
    for shape in arranged:
        if depth[shape["name"]]:
            continue
        box = boxes[shape["name"]]
        wide = _width(box) >= 0.5 * slide_w
        if paints_box(shape):
            sides = ("left", "right")
        elif shape.get("text") is None:
            # A frame a write did not read: a title's, across the slide,
            # still marks the margins; a narrow one cannot say.
            sides = ("left", "right") if wide else ()
        else:
            across = (_text_anchor(shape) or (None, None))[0]
            if across == "center" or across is None:
                sides = ()
            elif across == "full" or wide:
                sides = ("left", "right")
            else:
                sides = (across,)
        for side in sides:
            distance = box[0] if side == "left" else slide_w - box[2]
            if _EDGE_TOLERANCE < distance < _MARGIN_ZONE * slide_w:
                (lefts if side == "left" else rights).append(distance)
        if _EDGE_TOLERANCE < box[1] < _MARGIN_ZONE * slide_h:
            tops.append(box[1])
    kept = []
    for value in {round(d, 1) for d in lefts}:
        on_left = sum(1 for d in lefts if abs(d - value) <= _MARGIN_TOLERANCE)
        on_right = sum(1 for d in rights if abs(d - value) <= _MARGIN_TOLERANCE)
        if on_right:
            # The most shapes keep it; between equals, the narrower.
            kept.append((on_left + on_right, -value))
    if not kept:
        return None
    side = -max(kept)[1]
    return {"horizontal": side, "vertical": round(min([side] + tops), 1)}


def _near_margin(distance, margin) -> bool:
    """Whether a shape this far from the edge was meant to keep the margin."""
    return margin - max(_MARGIN_NEAR, 0.25 * margin) <= distance < margin - _MARGIN_TOLERANCE


# Narrower than this, a frame cannot hold the words of a label.
_MIN_LABEL_WIDTH = 60.0

# A label sits this close above or below the dot or icon it labels.
_LABEL_GAP = 30.0


def _centred_text(shape) -> bool:
    """Whether a shape is a frame of centred text, a label under an icon or a dot."""
    return not paints_box(shape) and (_text_anchor(shape) or (None, None))[0] == "center"


def _centred_shorten(shape, low, high):
    """The fix for centred text too near the edge for a narrower frame to help."""
    box = shape_box(shape)
    centre = (box[0] + box[2]) / 2
    room = max(0.0, 2 * min(centre - low, high - centre))
    return (f"Shorten the text of '{shape['name']}': centred where it is, at "
            f"{_pt(centre)}, it has {_pt(room)} across before the margin. Moving "
            "it would take it off what it labels; to make room, move that in "
            "from the edge as well.")


def _centred_fit(shape, low, high):
    """The fix that keeps centred text between ``low`` and ``high``, or None.

    Centred text labels whatever it is centred on, so moving it would take it
    off its mark. Narrowing its frame around the same centre makes the text
    wrap inside the limits and leaves it where it was.
    """
    box = shape_box(shape)
    centre = (box[0] + box[2]) / 2
    half = min(centre - low, high - centre)
    if 2 * half < _MIN_LABEL_WIDTH:
        return None
    return (f"Narrow '{shape['name']}' to {_pt(2 * half)} around its centre, so "
            f"its text wraps inside the margin and stays centred where it is "
            f"(ppt_update_shape left={_num(centre - half)} width={_num(2 * half)}), "
            "or shorten the text.")


def _layout(shapes, slide_w, slide_h):
    """What the checks on how a slide is put together share.

    The shapes that make up the layout and their boxes, how many shapes each
    sits on, the rows and columns of like shapes, and the slide's margins.
    ``groups`` are the top level ones, which are spaced and kept to the
    margins. ``rows`` and ``columns`` map each shape that paints to the like
    shapes beside it on its own level, which are kept one size.

    Text set beside an icon or under a dot is spaced by what it labels, and
    whether it lines up with it is the alignment check's question, so only
    text that labels nothing makes rows of its own: a column of big figures.
    """
    arranged = _arranged(shapes, slide_w, slide_h)
    boxes = {s["name"]: shape_box(s) for s in arranged}
    depth = _depths(arranged, boxes)
    rows, columns = {}, {}
    for level in sorted(set(depth.values())):
        painted = [s for s in arranged if depth[s["name"]] == level and paints_box(s)]
        for direction, group in _like_groups(painted, boxes):
            names = [s["name"] for s in group]
            for name in names:
                (rows if direction == "horizontal" else columns)[name] = names
    painted = [s for s in arranged if depth[s["name"]] == 0 and paints_box(s)]
    heading = _heading(arranged, slide_h)
    free_text = [s for s in arranged if depth[s["name"]] == 0 and not paints_box(s)
                 and not _labels_something(s, painted, boxes) and s["name"] not in heading]
    return {
        "arranged": arranged, "boxes": boxes, "depth": depth,
        "groups": _like_groups(painted, boxes) + _like_groups(free_text, boxes),
        "rows": rows, "columns": columns,
        "slide_width": slide_w, "slide_height": slide_h,
        "margins": _margins(arranged, boxes, depth, slide_w, slide_h),
    }


# A subtitle sits this close under its title.
_SUBTITLE_GAP = 30.0


def _heading(arranged, slide_h):
    """The names of the slide's title and the subtitle right under it.

    They head the slide rather than belong to its content, so a column of
    figures below is not spaced against them: the gap under a heading is
    meant to be larger.
    """
    title = find_title({"shapes": arranged}, slide_h)
    if title is None:
        return set()
    names = {title["name"]}
    above = text_box(title) or shape_box(title)
    for shape in arranged:
        words = text_box(shape)
        if shape is title or words is None or paints_box(shape):
            continue
        if 0 <= words[1] - above[3] <= _SUBTITLE_GAP and abs(words[0] - above[0]) <= _NEAR_MISS:
            names.add(shape["name"])
    return names


def _labels_something(shape, painted, boxes):
    """Whether a text frame sits right beside a shape that paints: a label."""
    words = text_box(shape) or boxes[shape["name"]]
    for other in painted:
        box = boxes[other["name"]]
        beside = _related(words, box, "y") and -_EDGE_TOLERANCE <= max(
            box[0] - words[2], words[0] - box[2]) <= _LABEL_GAP
        under = _related(words, box, "x") and -_EDGE_TOLERANCE <= max(
            box[1] - words[3], words[1] - box[3]) <= _LABEL_GAP
        if beside or under:
            return True
    return False


def _spans(row, boxes, slide_w, margin) -> bool:
    """Whether a row is meant to reach across between the margins.

    One end on a margin and the other near the far side, as a row of cards
    under a title is. A row set in from both sides is placed where it is.
    """
    zone = _MARGIN_ZONE * slide_w
    left, right = boxes[row[0]["name"]][0], slide_w - boxes[row[-1]["name"]][2]
    on_margin = min(abs(left - margin), abs(right - margin)) <= _MARGIN_TOLERANCE
    return on_margin and left < zone and right < zone


def _row_fix(row, boxes, slide_w, margin, layout=None):
    """The fix that spreads ``row`` evenly between the side margins, or None
    when it is already there or too wide to fit.

    A card moves with what is on it, its label or its figure, so the fix for
    a row of cards carries each card's contents along.
    """
    widths = [_width(boxes[s["name"]]) for s in row]
    room = slide_w - 2 * margin - sum(widths)
    if room <= 0:
        return None
    gap = room / (len(row) - 1)
    moves, x = [], margin
    for shape, width in zip(row, widths, strict=True):
        by = x - boxes[shape["name"]][0]
        if abs(by) > 0.05:
            carried = []
            if layout is not None:
                carried = [s["name"] for s in layout["arranged"]
                           if s is not shape and s["name"] in boxes
                           and _belongs_to(s, shape, boxes)]
            moves.append((shape["name"], x, by, carried))
        x += width + gap
    if not moves:
        return None
    tail = (f"Every gap is then {_pt(gap)}, and the row keeps the {_pt(margin)} side "
            f"margins the rest of the slide keeps.")
    if not any(carried for _, _, _, carried in moves):
        listed = " and ".join(f"of '{name}' to {_pt(left)}" for name, left, _, _ in moves)
        return f"Set left {listed} (ppt_update_shape). {tail}"
    calls = [f"'{name}'" + (" with what is on it" if carried else "") + f" to {_pt(left)} "
             f"(ppt_update_shape shape_names={[name] + carried} dleft={_num(by)})"
             for name, left, by, carried in moves]
    return f"Move {', and '.join(calls)}. {tail}"


_OPPOSITE = {"top": "bottom", "bottom": "top", "left": "right", "right": "left"}
_WHERE = {"top": "above", "bottom": "below", "left": "on the left", "right": "on the right"}


def _room(words, card):
    """How far text is from each edge of the card it sits on, negative past it."""
    return {
        "top": words[1] - card[1], "bottom": card[3] - words[3],
        "left": words[0] - card[0], "right": card[2] - words[2],
    }


def _wanted_room(room, side):
    """The room text should have on ``side`` of its card.

    What it has on the opposite side, the way a card is padded evenly, but no
    more than it has on the closer of the other two, so text set at the top
    of a tall card is not told to move to the middle of it. The closer one,
    because ragged lines leave more room on their own side than the card was
    padded with.
    """
    across = ("left", "right") if side in ("top", "bottom") else ("top", "bottom")
    return max(0.0, min(room[_OPPOSITE[side]], min(room[s] for s in across)))


def _padding(room, side, want) -> str:
    """Why the room asked for is that much: ', as it has above', when it is."""
    if abs(room[_OPPOSITE[side]] - want) <= 0.05:
        return f", as it has {_WHERE[_OPPOSITE[side]]}"
    return ""


def _card_of(shape, layout):
    """The shape that paints under this one's text, the innermost when they
    nest, or None."""
    boxes = layout["boxes"]
    if shape["name"] not in boxes:
        return None
    cards = [other for other in layout["arranged"]
             if other is not shape and paints_box(other)
             and _belongs_to(shape, other, boxes)]
    if not cards:
        return None
    return min(cards, key=lambda c: _width(boxes[c["name"]]) * _height(boxes[c["name"]]))


def _grow(card, side, amount, layout):
    """The fix that gives ``card`` ``amount`` more room on ``side``.

    The rest of its row grows with it when the room is height, and the rest
    of its column when it is width, so like cards stay one size.
    """
    name, box = card["name"], layout["boxes"][card["name"]]
    if side in ("top", "bottom"):
        word, group = "taller", "row"
        call = f"height={_num(_height(box) + amount)}"
        mates = layout["rows"].get(name, [])
    else:
        word, group = "wider", "column"
        call = f"width={_num(_width(box) + amount)}"
        mates = layout["columns"].get(name, [])
    if side == "top":
        call += f" top={_num(box[1] - amount)}"
    elif side == "left":
        call += f" left={_num(box[0] - amount)}"
    others = [m for m in mates if m != name]
    if not others:
        return f"make '{name}' {_pt(amount)} {word} (ppt_update_shape {call})"
    listed = ", ".join(f"'{m}'" for m in others)
    return (f"make '{name}' {_pt(amount)} {word}, and {listed} with it so the "
            f"{group} stays one size (ppt_update_shape shape_names="
            f"{[name] + others} {call})")


# ---------------------------------------------------------------------------
# Per shape findings
# ---------------------------------------------------------------------------
def _finding(slide_index, kind, shapes, message, fix, severity="error", **extra):
    finding = {
        "slide_index": slide_index,
        "type": kind,
        "severity": severity,
        "shapes": list(shapes),
        "message": message,
        "fix": fix,
    }
    finding.update(extra)
    return finding


def on_screen(animations):
    """When each animated shape is on the screen during the slide show.

    ``animations`` is the slide's main sequence in order, each effect a dict
    of the shape's name, its category (as ``list_animations`` reports it)
    and its trigger, None where it cannot be read. A click starts a step; an
    effect with or after the previous one shares its step. Where the trigger
    is unknown, as on the Mac, each effect is taken as its own click: that
    can only keep apart shapes that really are apart, never part two that are
    on the screen together.

    The answer maps a shape that enters or leaves to the steps it is on the
    screen for, as (from, to) pairs, ``to`` excluded. A shape with no entrance
    or exit is on the screen throughout, and is left out.
    """
    step, shown = 0, {}
    for effect in animations:
        if effect.get("trigger") not in (msoAnimTriggerWithPrevious, msoAnimTriggerAfterPrevious):
            step += 1
        kind = effect.get("category")
        if kind not in ("entrance", "exit"):
            continue
        spans = shown.get(effect["shape"])
        if spans is None:
            # Its first entrance or exit says whether it is there at the start.
            spans = shown[effect["shape"]] = [] if kind == "entrance" else [[0, None]]
        is_open = bool(spans) and spans[-1][1] is None
        if kind == "entrance" and not is_open:
            spans.append([step, None])
        elif kind == "exit" and is_open:
            spans[-1][1] = step
    return {name: [(a, math.inf if b is None else b) for a, b in spans]
            for name, spans in shown.items()}


def moving(animations):
    """The shapes a motion path moves during the slide show."""
    return {effect["shape"] for effect in animations
            if effect.get("category") == "motion_path"}


def _together(a, b) -> bool:
    """Whether two shapes are ever on the screen at once.

    A build-up lays one line where the last one was, each shown in turn:
    drawn on top of each other in the file, never seen so.
    """
    always = [(0, math.inf)]
    return any(s < e2 and s2 < e
               for s, e in a.get("shown") or always
               for s2, e2 in b.get("shown") or always)


def _parked(shape, box, slide_w, slide_h) -> bool:
    """Whether a shape is off the slide on purpose.

    One a motion path moves is judged by where it goes, not where it starts.
    A sound's icon dragged off the slide is how it is hidden, and it still
    plays. On a slide next to a Morph transition, a shape wholly off the slide
    is waiting to move on, or has just moved off.
    """
    if shape.get("moves"):
        return True
    outside = box[2] <= 0 or box[0] >= slide_w or box[3] <= 0 or box[1] >= slide_h
    return outside and (shape.get("type") == msoMedia or shape.get("morph"))


def _off_slide(slide_index, shapes, slide_w, slide_h, layout):
    """Shapes, or the text of shapes, that reach past an edge of the slide.

    Measured on what is drawn. A text box whose frame hangs over the edge with
    its text well inside is not wrong, and a box that fits on the slide with
    its text spilling off the bottom is.

    The fix brings the shape back inside the margin the rest of the slide
    keeps, not just onto it: moved by the overshoot alone, it would end up
    flush against the edge.
    """
    if not slide_w or not slide_h:
        return []
    margins = layout["margins"] or {}
    findings = []
    for shape in shapes:
        if not shape.get("visible", True):
            continue
        box = shape_box(shape) if _is_line(shape) else ink_box(shape)
        if box is None or _parked(shape, box, slide_w, slide_h):
            continue
        past = {
            "left": -box[0],
            "top": -box[1],
            "right": box[2] - slide_w,
            "bottom": box[3] - slide_h,
        }
        past = {edge: round(amount, 1) for edge, amount in past.items()
                if amount > _EDGE_TOLERANCE}
        if not past:
            continue
        parts = [f"{_pt(amount)} past the {edge} edge" for edge, amount in past.items()]
        name = shape["name"]
        moves, calls, kept, too_big = [], [], [], []
        for low, high, field, keep, words, bigger in (
            ("left", "right", "left", margins.get("horizontal"), ("right", "left"), "wider"),
            ("top", "bottom", "top", margins.get("vertical"), ("down", "up"), "taller"),
        ):
            if low in past and high in past:
                too_big.append(bigger)
                continue
            for edge, sign, word in ((low, 1, words[0]), (high, -1, words[1])):
                if edge not in past:
                    continue
                amount = past[edge] + (keep or 0.0)
                moves.append(f"{word} by {'' if keep else 'at least '}{_pt(amount)}")
                calls.append(f"{field}={_num(shape[field] + sign * amount)}")
                if keep:
                    kept.append(_pt(keep))
        keep = margins.get("horizontal") or 0.0
        centred = None
        if _centred_text(shape) and not {"top", "bottom"} & set(past):
            # Moved, centred text would leave what it is centred on.
            centred = _centred_fit(shape, keep, slide_w - keep) or _centred_shorten(
                shape, keep, slide_w - keep)
        if too_big:
            fix = f"'{name}' is {' and '.join(too_big)} than the slide: make it smaller."
        elif centred:
            fix = centred
        else:
            keeping = (f", to keep the {' and '.join(kept)} margin"
                       f"{'s' if len(kept) > 1 else ''} the rest of the slide keeps"
                       if kept else "")
            fix = (f"Move '{name}' {' and '.join(moves)}{keeping} "
                   f"(ppt_update_shape {' '.join(calls)}), or make it smaller.")
        findings.append(_finding(
            slide_index, "off_slide", [name],
            f"'{name}' runs {' and '.join(parts)} of the slide.",
            fix, edges=past,
        ))
    return findings


def _overflow(slide_index, shapes, layout):
    """Text that needs more room than its box gives it.

    The same verdict ppt_get_text(measure=True) gives, from the same
    arithmetic, so the two never disagree about a box.
    """
    from ppt_com.text import build_measurement

    findings = []
    for shape in shapes:
        text = shape.get("text")
        if not shape.get("visible", True) or not text or not text.get("chars"):
            continue
        measurement = build_measurement(
            [], text.get("width"), text.get("height"),
            shape["width"], shape["height"], text.get("margins"),
            text.get("word_wrap"), text.get("autofit"), text.get("orientation"),
        )
        if not measurement["overflows"]:
            continue
        name = shape["name"]
        needed = measurement["text_height_pt"]
        room = measurement["usable_height_pt"]
        if text.get("orientation") in ("vertical", "upward", "downward"):
            needed = measurement["text_width_pt"]
            room = measurement["usable_width_pt"]
        elif needed is not None and room is not None and needed <= room + 0.5:
            # Wrapping is off and the overflow is a line too long for the box.
            if text.get("autofit") == "shape_to_fit":
                # A box sized to its text is as wide as its longest line, so a
                # line cannot run out of its side. A Hebrew or Arabic line read
                # wider than the box is PowerPoint for Mac reporting it up to
                # 11 pt wider than it draws it (see text_box).
                continue
            needed = measurement["text_width_pt"]
            room = measurement["usable_width_pt"]
            findings.append(_finding(
                slide_index, "overflow", [name],
                f"A line of text in '{name}' is {_pt(needed)} long and the "
                f"box has {_pt(room)} for it, so it runs out of the side.",
                f"Make '{name}' {_pt(needed - room)} wider (ppt_update_shape "
                f"width={_num(shape['width'] + needed - room)}), turn word "
                "wrap on (ppt_set_textframe word_wrap=true), or shorten the line.",
                needed_pt=needed, available_pt=room,
            ))
            continue
        vertical = text.get("orientation") in ("vertical", "upward", "downward")
        field = "width" if vertical else "height"
        if not vertical and paints_box(shape) and layout["rows"].get(name):
            # One step of a process, one card of a row: growing it alone
            # would make it the odd one out.
            fix = (f"Cut the text or make the font smaller, or "
                   f"{_grow(shape, 'bottom', needed - room, layout)}.")
        else:
            grow = needed - room
            word = "wider" if vertical else "taller"
            blocker = _in_the_way(shape, grow, vertical, shapes, layout)
            if blocker:
                # A title box grown down into the body below it trades one
                # problem for another; the words are what has to give.
                fix = (f"Cut the text or make the font smaller. Making '{name}' "
                       f"{_pt(grow)} {word} would run it into {blocker}.")
            else:
                fix = (f"Make '{name}' {_pt(grow)} {word} "
                       f"(ppt_update_shape {field}={_num(shape[field] + grow)}), "
                       "cut the text, or make the font smaller.")
        findings.append(_finding(
            slide_index, "overflow", [name],
            f"The text in '{name}' needs {_pt(needed)} and the box has "
            f"{_pt(room)}, so it spills out of the box.",
            fix, needed_pt=needed, available_pt=room,
        ))
    return findings


def _in_the_way(shape, grow, vertical, shapes, layout):
    """What a box grown by ``grow`` would run into, as words, or None.

    A box grows from its top down, or for sideways text from its left across,
    so only the strip it gains is looked at. The card a text sits on, and what
    sits on a card, are not in its way.
    """
    box = shape_box(shape)
    strip = ((box[2], box[1], box[2] + grow, box[3]) if vertical
             else (box[0], box[3], box[2], box[3] + grow))
    slide_w, slide_h = layout.get("slide_width"), layout.get("slide_height")
    if slide_w and slide_h and (strip[2] > slide_w + _EDGE_TOLERANCE
                                or strip[3] > slide_h + _EDGE_TOLERANCE):
        return "the edge of the slide"
    boxes = layout["boxes"]
    for other in shapes:
        if other is shape or not other.get("visible", True) or _is_line(other):
            continue
        if slide_w and slide_h and _is_backdrop(other, slide_w, slide_h):
            continue
        if shape["name"] in boxes and other["name"] in boxes and (
                _belongs_to(shape, other, boxes) or _belongs_to(other, shape, boxes)):
            continue
        ink = ink_box(other)
        if ink is None:
            continue
        width, height = _shared(strip, ink)
        if width > _MIN_OVERLAP and height > _MIN_OVERLAP:
            return f"'{other['name']}'" if paints_box(other) else f"the text of '{other['name']}'"
    return None


def _overlap(slide_index, shapes, layout):
    """Text that runs into another shape, or into other text.

    Only text is judged. Two filled shapes crossing may be a Venn diagram or a
    badge on a card, and nothing about the slide says which, but text drawn
    over something, or under it, is never meant. Text that sits wholly inside
    another shape is on it, like the words on a card, and is fine, unless the
    shape is a table, whose cells hold text of their own, or a card laid in
    front of it, which hides it.

    Text running into several cards of one row is one finding, not one per
    card: a subtitle wrapped onto a second line is one problem.
    """
    findings, hits, seen = [], {}, set()
    for a in shapes:
        a_text = text_box(a) if a.get("visible", True) else None
        if a_text is None or _turned(a) or a.get("moves"):
            continue
        for b in shapes:
            # Text spilling out of its own box is the overflow check's.
            if b is a or _turned(b) or b.get("moves") or not _together(a, b):
                continue
            b_ink = ink_box(b)
            if b_ink is None:
                continue
            if (_inside(a_text, b_ink) and b["type"] != msoTable
                    and not (b.get("z", 0) > a.get("z", 0) and _solid_fill(b))):
                continue
            width, height = _shared(a_text, b_ink)
            if width <= _MIN_OVERLAP or height <= _MIN_OVERLAP:
                continue
            pair = frozenset((a["name"], b["name"]))
            if pair in seen:
                continue
            seen.add(pair)
            spill = _spill_finding(slide_index, a, a_text, b, b_ink, width, height, layout)
            if spill is not None:
                findings.append(spill)
            else:
                hits.setdefault(a["name"], (a, a_text, []))[2].append((b, b_ink, width, height))
    for a, a_text, collided in hits.values():
        rows = {}
        for hit in collided:
            row = tuple(layout["rows"].get(hit[0]["name"]) or [hit[0]["name"]])
            rows.setdefault(row, []).append(hit)
        for group in rows.values():
            findings.append(_collision_finding(slide_index, a, a_text, group, shapes, layout))
    return findings


def _spill_finding(slide_index, a, a_text, b, b_ink, width, height, layout):
    """Text that sits on a card and runs past its edge, or None."""
    a_name, b_name = a["name"], b["name"]
    # Text mostly on a shape underneath it is meant to be on it, like the
    # words on a card, and the problem is that it does not fit. Saying where
    # it runs past is more use than saying the two collide.
    area = _width(a_text) * _height(a_text)
    if not (paints_box(b) and b.get("z", 0) < a.get("z", 0) and b["type"] != msoTable
            and area and width * height >= 0.5 * area):
        return None
    room = _room(a_text, b_ink)
    past = {edge: round(-amount, 1) for edge, amount in room.items()
            if -amount > _EDGE_TOLERANCE}
    parts = " and ".join(f"{_pt(amount)} past its {edge} edge"
                         for edge, amount in past.items())
    # Grown by the overshoot alone, the card would end exactly where the
    # text does. It grows by as much again as the text has on the other
    # side, so the text is padded evenly.
    if b_name in layout["boxes"]:
        grows = []
        for edge, amount in past.items():
            want = _wanted_room(room, edge)
            grows.append(
                _grow(b, edge, amount + want, layout)
                + (f", which leaves the text {_pt(want)} {_WHERE[edge]}"
                   f"{_padding(room, edge, want)}" if want else "")
            )
        fix = f"{'; '.join(grows)}. Or make the text of '{a_name}' shorter or smaller."
        fix = fix[0].upper() + fix[1:]
    else:
        fix = (f"Make '{b_name}' larger by that much, or make the text of "
               f"'{a_name}' shorter or smaller.")
    return _finding(
        slide_index, "overlap", [a_name, b_name],
        f"The text of '{a_name}' sits on '{b_name}' and runs {parts}.",
        fix, overlap_pt={"width": round(width, 1), "height": round(height, 1)},
        past_edges=past,
    )


def _collision_finding(slide_index, a, a_text, group, shapes, layout):
    """Text running into one shape, or into several of one row."""
    a_name = a["name"]
    others = [b for b, _, _, _ in group]
    inks = [ink for _, ink, _, _ in group]
    width = max(w for _, _, w, _ in group)
    height = max(h for _, _, _, h in group)
    union = (min(i[0] for i in inks), min(i[1] for i in inks),
             max(i[2] for i in inks), max(i[3] for i in inks))
    # What a shape puts on the slide is its text when it does not paint its box.
    targets = ", ".join(f"'{b['name']}'" if paints_box(b) else f"the text of '{b['name']}'"
                        for b in others)
    overlap_pt = {"width": round(width, 1), "height": round(height, 1)}
    # Wholly inside: on a table's cells, or under a card laid in front of it.
    contained = _inside(a_text, union)
    on_table = contained and all(b["type"] == msoTable for b in others)
    in_front = [b for b in others if b.get("z", 0) > a.get("z", 0)]
    if contained and len(others) == 1 and in_front and not on_table:
        # Words meant for the card, added before it: the card hides them.
        b_name = others[0]["name"]
        return _finding(
            slide_index, "overlap", [a_name, b_name],
            f"The text of '{a_name}' is hidden under '{b_name}', which is in front of it.",
            f"Bring '{a_name}' to the front (ppt_set_shape_zorder shape_name='{a_name}' "
            f"command='bring_to_front'), so it sits on '{b_name}'.",
            overlap_pt=overlap_pt,
        )
    # The cheaper way apart: along whichever axis they share less of.
    vertical = height <= width
    low, high = (1, 3) if vertical else (0, 2)
    sign = 1 if a_text[low] + a_text[high] >= union[low] + union[high] else -1
    # Far enough to leave the gap past the far side of what it ran into, which
    # for text lying wholly on a table is all of the table.
    clear = (union[high] - a_text[low] if sign > 0 else a_text[high] - union[low]) + _CLEARANCE
    field, ways = ("top", ("down", "up")) if vertical else ("left", ("right", "left"))
    # Shorter words help text that crosses an edge, not text lying on a table.
    shorten = "" if on_table else f", or shorten the text of '{a_name}'"
    # The smaller of the two moves, a caption rather than the table it
    # captions, unless that would put it on something else: then the other.
    a_box = shape_box(a)
    theirs = sum(_width(shape_box(b)) * _height(shape_box(b)) for b in others)
    if _width(a_box) * _height(a_box) <= theirs and not _lands_on_something(
            a, a_text, vertical, sign * clear, shapes, others, layout):
        fix = (f"Move '{a_name}' {ways[0] if sign > 0 else ways[1]} by {_pt(clear)} "
               f"(ppt_update_shape d{field}={_num(sign * clear)}), which leaves "
               f"{_pt(_CLEARANCE)} between it and {targets}{shorten}.")
    else:
        boxes = layout["boxes"]
        carried = [s for s in layout["arranged"]
                   if s is not a and all(s is not b for b in others)
                   and any(b["name"] in boxes and _belongs_to(s, b, boxes) for b in others)]
        names = [b["name"] for b in others] + [c["name"] for c in carried]
        them = "it" if len(others) == 1 else "them"
        with_them = f" together with the shapes on {them}" if carried else ""
        fix = (f"Move {targets} {ways[1] if sign > 0 else ways[0]} by {_pt(clear)}"
               f"{with_them} (ppt_update_shape shape_names={names} "
               f"d{field}={_num(-sign * clear)}), which leaves {_pt(_CLEARANCE)} "
               f"between {them} and the text of '{a_name}'{shorten}.")
    if on_table:
        where = "is hidden under" if in_front else "lies on"
        message = f"The text of '{a_name}' {where} {targets}, over its cells."
    else:
        message = (f"The text of '{a_name}' runs into {targets}: they share an area "
                   f"{_pt(width)} wide and {_pt(height)} tall.")
    return _finding(
        slide_index, "overlap", [a_name] + [b["name"] for b in others],
        message, fix, overlap_pt=overlap_pt,
    )


def _solid_fill(shape) -> bool:
    """Whether the shape is filled with one colour that hides what is under it."""
    fill = shape.get("fill") or {}
    return (bool(fill.get("visible")) and bool(fill.get("solid"))
            and (fill.get("transparency") or 0.0) <= _OPAQUE_ENOUGH and not _is_line(shape))


def _lands_on_something(a, a_text, vertical, by, shapes, others, layout):
    """Whether text moved this far would run into something else, or off the slide."""
    moved = ((a_text[0], a_text[1] + by, a_text[2], a_text[3] + by) if vertical
             else (a_text[0] + by, a_text[1], a_text[2] + by, a_text[3]))
    slide_w, slide_h = layout.get("slide_width"), layout.get("slide_height")
    if slide_w and slide_h and (
            moved[0] < -_EDGE_TOLERANCE or moved[1] < -_EDGE_TOLERANCE
            or moved[2] > slide_w + _EDGE_TOLERANCE or moved[3] > slide_h + _EDGE_TOLERANCE):
        return True
    boxes = layout["boxes"]
    # The card it sits on is not in its way.
    holders = {s["name"] for s in layout["arranged"]
               if a["name"] in boxes and s is not a and _belongs_to(a, s, boxes)}
    skip = {a["name"]} | {b["name"] for b in others} | holders
    for shape in shapes:
        if shape["name"] in skip or not shape.get("visible", True) or _is_line(shape):
            continue
        if slide_w and slide_h and _is_backdrop(shape, slide_w, slide_h):
            continue
        ink = ink_box(shape)
        if ink is None:
            continue
        # Close enough to touch is as bad as touching: the fix would jam it
        # against the next thing, a subtitle up against the title.
        near = _CLEARANCE / 2
        width, height = _shared((moved[0] - near, moved[1] - near, moved[2] + near, moved[3] + near), ink)
        if width > _MIN_OVERLAP and height > _MIN_OVERLAP:
            return True
    return False


_FROM = {"master": "the slide master", "layout": "the slide's layout"}


# Placeholders that mark where a template means content to go.
_CONTENT_PLACEHOLDERS = frozenset({
    ppPlaceholderTitle, ppPlaceholderCenterTitle, ppPlaceholderBody, ppPlaceholderObject,
})


def _sparse(graphic) -> bool:
    """Whether a template graphic is drawn in lines rather than painted solid.

    A group of thin curves down the side of the slide, a fan of lines in a
    corner: their box is mostly empty, so it says less about where they are
    than a band's or a logo's does.
    """
    if graphic["type"] == msoGroup:
        return True
    fill = graphic.get("fill") or {}
    return (graphic["type"] not in _OPAQUE_TYPES and fill.get("visible") is False
            and not (graphic.get("text") or {}).get("chars"))


def _subtract(box, cut):
    """The parts of ``box`` outside ``cut``, as up to four boxes."""
    width, height = _shared(box, cut)
    if width <= 0 or height <= 0:
        return [box]
    top, bottom = max(box[1], cut[1]), min(box[3], cut[3])
    parts = [(box[0], box[1], box[2], cut[1]), (box[0], cut[3], box[2], box[3]),
             (box[0], top, cut[0], bottom), (cut[2], top, box[2], bottom)]
    return [p for p in parts if p[2] - p[0] > _EDGE_TOLERANCE and p[3] - p[1] > _EDGE_TOLERANCE]


def _graphic_parts(graphic, ink, area):
    """Where a template graphic stands in the way of the slide's own shapes.

    A solid graphic, wherever it is drawn. A sparse one, less the part of it
    the template's own title and body placeholders cover: a theme that sets
    its title over the tips of its decoration means content to go there.
    """
    parts = [ink]
    if _sparse(graphic):
        for cut in area or ():
            parts = [piece for part in parts for piece in _subtract(part, cut)]
    return parts


def _template_collisions(slide_index, shapes, template, layout, area=()):
    """The slide's own shapes running into the template's graphics.

    A logo, a band down one side, a footer line: the template draws them on
    every slide, and nothing a slide adds is meant to cross them. Text over
    one is hard to read, and a card over one cuts into the design.

    Placeholders are where their layout put them and are left alone. So is a
    shape wholly on a graphic, which is the contrast check's to judge. One
    covering a graphic whole is reported: a chart sized over the logo and the
    confidentiality line hides them as surely as one crossing their edge. A
    backdrop over the whole slide is a decision about the whole slide, and is
    not a shape here at all.
    """
    slide_w, slide_h = layout.get("slide_width"), layout.get("slide_height")
    graphics, obstacles = [], []
    for graphic in template or ():
        if not graphic.get("visible", True) or _is_line(graphic) or _turned(graphic):
            continue
        if _is_backdrop(graphic, slide_w, slide_h):
            continue
        ink = ink_box(graphic)
        if ink is None and graphic.get("text") is None:
            # A graphic a write left unread, being far from what it touched,
            # still stands where a fix might move something.
            ink = shape_box(graphic)
        elif ink is None:
            continue
        for part in _graphic_parts(graphic, ink, area):
            obstacles.append((graphic, part))
            if graphic.get("text") is not None or paints_box(graphic):
                graphics.append((graphic, part))
    if not graphics:
        return []
    findings = []
    for a in shapes:
        if (not a.get("visible", True) or _is_line(a) or _turned(a)
                or a.get("placeholder_type") is not None):
            continue
        a_ink = shape_box(a) if paints_box(a) else text_box(a)
        if a_ink is None:
            continue
        hits, seen = [], set()
        for graphic, ink in graphics:
            width, height = _shared(a_ink, ink)
            if width <= _MIN_OVERLAP or height <= _MIN_OVERLAP:
                continue
            if _inside(a_ink, ink) or graphic["name"] in seen:
                continue
            seen.add(graphic["name"])
            hits.append((graphic, ink, width, height))
        if hits:
            findings.append(_template_collision_finding(slide_index, a, a_ink, hits, shapes,
                                                        obstacles, layout, bool(area)))
    return findings


# A template graphic this thin is a rule: a line under the title, a line over
# the footer.
_RULE = 3.0


def _template_rules(slide_index, shapes, template, layout):
    """Text, or a card, drawn across one of the template's rules.

    A rule under the title is drawn on every slide; a title that wraps onto a
    second line runs its words across it, and a card set over it cuts it. A
    rule is too thin to share an area with anything, so it is judged by
    whether it crosses what the shape draws. A placeholder's frame often
    reaches over its layout's rule by design, so for one it is the words that
    count, and only where they cross.
    """
    rules = []
    for graphic in template or ():
        if not graphic.get("visible", True) or _turned(graphic):
            continue
        box = shape_box(graphic)
        # Thin one way: a slanting line's box says nothing about where it is.
        thin = min(_width(box), _height(box))
        if thin > _RULE or not (_is_line(graphic) or paints_box(graphic)):
            continue
        if max(_width(box), _height(box)) < 4 * _RULE:
            continue
        rules.append((graphic, box, _width(box) >= _height(box)))
    if not rules:
        return []
    slide_w, slide_h = layout.get("slide_width"), layout.get("slide_height")
    obstacles = [(g, shape_box(g)) for g in template if g.get("visible", True)
                 and not _turned(g) and not _is_backdrop(g, slide_w, slide_h)]
    findings = []
    for a in shapes:
        if not a.get("visible", True) or _is_line(a) or _turned(a):
            continue
        words_only = a.get("placeholder_type") is not None or not paints_box(a)
        ink = text_box(a) if words_only else shape_box(a)
        if ink is None:
            continue
        for graphic, box, across in rules:
            if across:
                line = (box[1] + box[3]) / 2
                crosses = (ink[1] + _MIN_OVERLAP < line < ink[3] - _MIN_OVERLAP
                           and _shared(ink, (box[0], ink[1], box[2], ink[3]))[0] > _MIN_OVERLAP)
            else:
                line = (box[0] + box[2]) / 2
                crosses = (ink[0] + _MIN_OVERLAP < line < ink[2] - _MIN_OVERLAP
                           and _shared(ink, (ink[0], box[1], ink[2], box[3]))[1] > _MIN_OVERLAP)
            if not crosses:
                continue
            name = a["name"]
            what = f"The text of '{name}'" if words_only else f"'{name}'"
            source = _FROM.get(graphic.get("source"))
            drawn = f", drawn by {source}" if source else ""
            named = f"'{graphic['name']}'"
            if words_only and across and (ink[1] + ink[3]) / 2 < line:
                # Words above the rule that grew down across it: a title that
                # wrapped.
                fix = (f"Keep '{name}' to the room above the line: shorten the words "
                       f"or make the font smaller until they end above "
                       f"{_pt(line - _MIN_OVERLAP)} (ppt_format_text font_size=...).")
            else:
                # A card takes back the side that crossed, or moves the
                # shorter way that leaves it on the slide and clear of the rest.
                fix = (_template_resize(a, box, named, layout)
                       or _template_shrink(a, box, named, keep=0.8)
                       or _template_move(a, ink, box, named, shapes, obstacles, layout)
                       or _template_shrink(a, box, named, keep=0.5)
                       or (f"Move '{name}' clear of {named}, or make it smaller: "
                           "nothing on the slide is meant to cross it."))
            findings.append(_finding(
                slide_index, "overlap", [name, graphic["name"]],
                f"{what} runs across '{graphic['name']}', a line the template draws"
                f"{drawn}.",
                fix, template_graphics=[graphic["name"]],
            ))
            break
    return findings


def _template_collision_finding(slide_index, a, a_ink, hits, shapes, graphics, layout,
                                carved=False):
    a_name = a["name"]
    width = max(w for _, _, w, _ in hits)
    height = max(h for _, _, _, h in hits)
    union = (min(i[0] for _, i, _, _ in hits), min(i[1] for _, i, _, _ in hits),
             max(i[2] for _, i, _, _ in hits), max(i[3] for _, i, _, _ in hits))
    named = " and ".join(f"'{g['name']}'" for g, _, _, _ in hits)
    sources = {g.get("source") for g, _, _, _ in hits}
    source = _FROM.get(sources.pop()) if len(sources) == 1 else None
    drawn = f", part of the template drawn by {source}" if source else ", part of the template"
    what = f"'{a_name}'" if paints_box(a) else f"The text of '{a_name}'"
    if all(_inside(ink, a_ink) for _, ink, _, _ in hits):
        message = f"{what} covers {named}{drawn}."
    elif carved and all(_sparse(g) for g, _, _, _ in hits):
        # Lines in a corner leave most of their box empty, so the shape may
        # touch none of them; what it did is leave the template's content area.
        kind = "a decoration" if len(hits) == 1 else "decorations"
        by = f"drawn by {source}" if source else "of the template"
        message = (f"{what} runs out of the template's content area into the space of "
                   f"{named}, {kind} {by}: an area {_pt(width)} wide and {_pt(height)} tall.")
    else:
        message = (f"{what} runs into {named}{drawn}: they share an area "
                   f"{_pt(width)} wide and {_pt(height)} tall.")
    # A chart trimmed a little at the bottom stays where it was, under its
    # title; moved, it would jam against the title. A small shape is better
    # moved than squeezed.
    fix = (_template_resize(a, union, named, layout)
           or _template_shrink(a, union, named, keep=0.8)
           or _template_move(a, a_ink, union, named, shapes, graphics, layout)
           or _template_shrink(a, union, named, keep=0.5)
           or (f"Move '{a_name}' clear of {named}, or make it smaller: the template keeps "
               "that part of the slide for its own graphics."))
    return _finding(
        slide_index, "overlap", [a_name] + [g["name"] for g, _, _, _ in hits],
        message, fix, overlap_pt={"width": round(width, 1), "height": round(height, 1)},
        template_graphics=[g["name"] for g, _, _, _ in hits],
    )


def _template_resize(a, union, named, layout):
    """Pull back the side of a card in a row that crosses a template graphic.

    Moving a card out of a row of like ones breaks the row's spacing; taking
    back the edge that crossed keeps the rest of it where it was. The size the
    rest of the row shares is the one to take when it is enough.
    """
    name = a["name"]
    boxes = layout["boxes"]
    if name not in boxes or not paints_box(a):
        return None
    for groups, low, field, word, kind in ((layout["rows"], 0, "width", "wide", "row"),
                                           (layout["columns"], 1, "height", "tall", "column")):
        mates = groups.get(name)
        if not mates or len(mates) < 2:
            continue
        box = boxes[name]
        high = low + 2
        size = box[high] - box[low]
        others = [boxes[m][high] - boxes[m][low] for m in mates if m != name and m in boxes]
        usual = _clusters(list(enumerate(others)), _SIZE_TOLERANCE)
        shared = sorted(v for _, v in usual[0])[len(usual[0]) // 2] if usual else None
        position = "left" if low == 0 else "top"
        if union[low] >= (box[low] + box[high]) / 2:
            # The graphic is past its far edge: the near edge stays.
            room = union[low] - _CLEARANCE - box[low]
            if room <= 0:
                continue
            if shared is None or shared > room:
                # Squeezed alone, the card would be the odd one out of its row
                # and its words might not fit; the whole row narrows instead.
                fit = _fit_row(mates, layout, low, None, union[low] - _CLEARANCE, kind, named)
                if fit:
                    return fit
            target = shared if shared is not None and shared <= room else room
            if target < 0.5 * size:
                continue
            like = f" like the rest of its {kind}" if target == shared else ""
            return (f"Make '{name}' {_pt(target)} {word}{like} (ppt_update_shape "
                    f"{field}={_num(target)}), which leaves it {_pt(union[low] - box[low] - target)} "
                    f"clear of {named}.")
        if union[high] <= (box[low] + box[high]) / 2:
            start = union[high] + _CLEARANCE
            room = box[high] - start
            if room <= 0:
                continue
            if shared is None or shared > room:
                fit = _fit_row(mates, layout, low, start, None, kind, named)
                if fit:
                    return fit
            target = shared if shared is not None and shared <= room else room
            if target < 0.5 * size:
                continue
            like = f" like the rest of its {kind}" if target == shared else ""
            new_start = box[high] - target
            return (f"Make '{name}' {_pt(target)} {word}{like}, keeping its far edge "
                    f"(ppt_update_shape {position}={_num(new_start)} {field}={_num(target)}), "
                    f"which leaves it {_pt(new_start - union[high])} clear of {named}.")
    return None


def _fit_row(mates, layout, low, start, end, kind, named, lead=None, grow=False):
    """Narrow a row (or column) of like shapes evenly to fit between ``start`` and ``end``.

    Either bound may be None, and then that end of the row stays where it is.
    The gaps stay as they are, every card takes one new size, and each card
    moves with what is on it. With ``grow`` the row may also widen to reach
    the bounds.
    """
    boxes = layout["boxes"]
    high = low + 2
    members = sorted((m for m in mates if m in boxes), key=lambda m: boxes[m][low])
    if len(members) < 2:
        return None
    sizes = [boxes[m][high] - boxes[m][low] for m in members]
    first, last = boxes[members[0]][low], boxes[members[-1]][high]
    gap = ((last - first) - sum(sizes)) / (len(members) - 1)
    begin = first if start is None else start
    finish = last if end is None else end
    size = (finish - begin - gap * (len(members) - 1)) / len(members)
    if gap < 0 or size < 0.5 * min(sizes) or (size >= min(sizes) and not grow):
        return None
    field, position = ("width", "left") if low == 0 else ("height", "top")
    moves = []
    for i, name in enumerate(members):
        by = begin + i * (size + gap) - boxes[name][low]
        if abs(by) <= 0.05:
            continue
        shape = next(s for s in layout["arranged"] if s["name"] == name)
        carried = [s["name"] for s in layout["arranged"]
                   if s is not shape and s["name"] in boxes and _belongs_to(s, shape, boxes)]
        with_it = " with what is on it" if carried else ""
        way = ("right" if by > 0 else "left") if low == 0 else ("down" if by > 0 else "up")
        moves.append(f"'{name}'{with_it} {way} by {_pt(abs(by))} (ppt_update_shape shape_names="
                     f"{[name] + carried} d{position}={_num(by)})")
    fix = (f"{lead or f'Make the {kind} fit before {named}'}: ppt_update_shape "
           f"shape_names={members} {field}={_num(size)}")
    if moves:
        fix += f", then move {', and '.join(moves)}"
    return fix + f". The gaps stay {_pt(gap)}."


def _template_shrink(a, union, named, keep):
    """Take back the side of a shape that crosses a template graphic.

    Only for a shape that paints, since a text frame made smaller keeps its
    text by wrapping it, and only when at least ``keep`` of it is left and its
    words still fit.
    """
    if not paints_box(a):
        return None
    box = shape_box(a)
    options = []
    for low, field, word, position in ((1, "height", "tall", "top"), (0, "width", "wide", "left")):
        high = low + 2
        size = box[high] - box[low]
        if size <= 0:
            continue
        def fits(target, low=low, size=size):
            return target >= keep * size and not _spills_at(
                a, target if low == 0 else _width(box), target if low == 1 else _height(box))

        if union[low] >= (box[low] + box[high]) / 2:
            target = union[low] - _CLEARANCE - box[low]
            if fits(target):
                options.append((size - target, (
                    f"Make '{a['name']}' {_pt(target)} {word} (ppt_update_shape "
                    f"{field}={_num(target)}), which leaves it {_pt(_CLEARANCE)} clear of {named}.")))
        elif union[high] <= (box[low] + box[high]) / 2:
            start = union[high] + _CLEARANCE
            target = box[high] - start
            if fits(target):
                options.append((size - target, (
                    f"Make '{a['name']}' {_pt(target)} {word}, keeping its far edge "
                    f"(ppt_update_shape {position}={_num(start)} {field}={_num(target)}), "
                    f"which leaves it {_pt(_CLEARANCE)} clear of {named}.")))
    return min(options)[1] if options else None


def _kept_edges(shape):
    """The sides of its frame a shape is seen to keep: what lining up with it means.

    A shape that paints its box keeps both. Text alone keeps only the side it
    is set to: a centred title's frame edges are nowhere a reader can see, and
    the far edge of left-set text is wherever its longest line ends.
    """
    if paints_box(shape):
        return {"left", "right"}
    anchor = _text_anchor(shape)
    if anchor is None:
        return {"left", "right"}
    return {"full": {"left", "right"}, "left": {"left"}, "right": {"right"}}.get(anchor[0], set())


def _template_move(a, a_ink, union, named, shapes, graphics, layout):
    """The shortest move that takes a shape clear of a template graphic.

    Whichever way is shortest and leaves it on the slide, away from the other
    shapes and the template's other graphics. The text on a card moves with it.
    Moved sideways, it goes on to the nearest edge something on the slide
    already keeps, the title's or the chart's, when that is not much further:
    clear of the graphic and in line with the content reads as placed.
    """
    slide_w, slide_h = layout.get("slide_width"), layout.get("slide_height")
    boxes = layout["boxes"]
    carried, holders = [], set()
    if a["name"] in boxes:
        for s in layout["arranged"]:
            if s is a or s["name"] not in boxes:
                continue
            if paints_box(a) and _belongs_to(s, a, boxes):
                carried.append(s)
            elif _belongs_to(a, s, boxes):
                holders.add(s["name"])
    skip = {a["name"]} | {s["name"] for s in carried} | holders
    # A band is often more than one shape: a thin line on top of it, a strip
    # beside it. Clear of the band means clear of what touches it too.
    grown = True
    while grown:
        grown = False
        for graphic, ink in graphics:
            if f"'{graphic['name']}'" in named:
                # Another piece of a graphic already being cleared: the part
                # of a decoration the template's own content covers stays out.
                continue
            reach = (union[0] - _CLEARANCE, union[1] - _CLEARANCE,
                     union[2] + _CLEARANCE, union[3] + _CLEARANCE)
            w, h = _shared(reach, ink)
            if w > 0 and h > 0 and not _inside(ink, union):
                union = (min(union[0], ink[0]), min(union[1], ink[1]),
                         max(union[2], ink[2]), max(union[3], ink[3]))
                if f"'{graphic['name']}'" not in named:
                    named = f"{named} and '{graphic['name']}'"
                grown = True

    def lands_clear(field, by):
        dx, dy = (by, 0.0) if field == "left" else (0.0, by)
        moved = (a_ink[0] + dx, a_ink[1] + dy, a_ink[2] + dx, a_ink[3] + dy)
        if slide_w and slide_h and (
                moved[0] < -_EDGE_TOLERANCE or moved[1] < -_EDGE_TOLERANCE
                or moved[2] > slide_w + _EDGE_TOLERANCE or moved[3] > slide_h + _EDGE_TOLERANCE):
            return False
        near = _CLEARANCE / 2
        reach = (moved[0] - near, moved[1] - near, moved[2] + near, moved[3] + near)
        inks = [ink for _, ink in graphics]
        for other in shapes:
            if other["name"] in skip or not other.get("visible", True) or _is_line(other):
                continue
            if slide_w and slide_h and _is_backdrop(other, slide_w, slide_h):
                continue
            inks.append(ink_box(other))
        for ink in inks:
            if ink is None:
                continue
            w, h = _shared(reach, ink)
            if w > _MIN_OVERLAP and h > _MIN_OVERLAP:
                return False
        return True

    shortest = sorted((abs(by), field, by, words) for field, by, words in (
        ("left", union[2] + _CLEARANCE - a_ink[0], "right"),
        ("left", union[0] - _CLEARANCE - a_ink[2], "left"),
        ("top", union[3] + _CLEARANCE - a_ink[1], "down"),
        ("top", union[1] - _CLEARANCE - a_ink[3], "up"),
    ))
    chosen = next(((field, by, words) for _, field, by, words in shortest
                   if lands_clear(field, by)), None)
    if chosen is None:
        return None
    field, by, words = chosen
    where = ""
    if field == "left":
        lines = []
        for s in layout["arranged"]:
            if s["name"] in skip:
                continue
            kept = _kept_edges(s)
            lines.append((s["left"] if "left" in kept else None,
                          s["left"] + s["width"] if "right" in kept else None,
                          f"in line with '{s['name']}'"))
        margin = (layout.get("margins") or {}).get("horizontal")
        if margin is not None and slide_w:
            lines.append((margin, slide_w - margin, "on the slide's margin"))
        right_edge = a["left"] + a["width"]
        if by > 0:
            options = sorted((left - a["left"], f", {label}") for left, _, label in lines
                             if left is not None and left - a["left"] >= by - _EDGE_TOLERANCE)
            if _reads_right_to_left(shapes):
                # A right-to-left slide keeps its right edges; the left
                # ones are wherever its lines happen to end.
                options = sorted((right - right_edge, f", {label}") for _, right, label in lines
                                 if right is not None
                                 and right - right_edge >= by - _EDGE_TOLERANCE) + options
        else:
            options = sorted((right - right_edge, f", {label}") for _, right, label in lines
                             if right is not None
                             and right - right_edge <= by + _EDGE_TOLERANCE)[::-1]
        for aligned, label in options:
            if abs(aligned - by) > max(40.0, abs(by)):
                continue
            if lands_clear(field, aligned):
                by, where = aligned, label
                break
    if carried:
        names = [a["name"]] + [s["name"] for s in carried]
        call = f"ppt_update_shape shape_names={names} d{field}={_num(by)}"
        along = " together with what is on it"
    else:
        call = f"ppt_update_shape d{field}={_num(by)}"
        along = ""
    if where:
        return (f"Move '{a['name']}' {words} by {_pt(abs(by))}{where}{along} ({call}), "
                f"clear of {named}.")
    return (f"Move '{a['name']}' {words} by {_pt(abs(by))}{along} ({call}), which "
            f"leaves {_pt(_CLEARANCE)} between it and {named}.")


def _backdrop(shape, shapes, background_hex, template=(), background_gradient=None,
              slide_box=None):
    """What the shape's text is drawn on, and what paints it.

    The shape's own fill when it has one, otherwise the nearest shape
    underneath that holds the whole of the text and paints something, then a
    graphic of the template's that does, and otherwise the slide's background.
    What it is drawn on is one colour, or a gradient across a box as
    ``{"gradient": ..., "box": ...}``, since where on it the text sits decides
    its colour there. A picture or a see-through fill in the way means it
    cannot be told, and it is None. What paints it is "fill", "shape",
    "template" or "background", and the shape or graphic that does, when one
    does.
    """
    painted, colour = _paint(shape)
    if painted:
        return colour, "fill", shape
    words = text_box(shape)
    if words is None:
        return None, None, None
    below = [s for s in shapes if s.get("z", 0) < shape.get("z", 0)]
    for source, candidates in (("shape", below), ("template", list(template or ()))):
        for other in sorted(candidates, key=lambda s: s.get("z", 0), reverse=True):
            if not other.get("visible", True) or _is_line(other):
                continue
            if not _inside(words, shape_box(other)):
                continue
            if other["type"] in _UNKNOWN_BACKDROP_TYPES:
                return None, source, other
            painted, colour = _paint(other)
            if painted:
                return colour, source, other
    if background_gradient and slide_box and _opaque(background_gradient):
        return {"gradient": background_gradient, "box": slide_box}, "background", None
    return background_hex, "background", None


def _paint(shape):
    """Whether the shape's fill paints, and with what: one colour, a gradient
    across the shape's box, or None when that cannot be told."""
    fill = shape.get("fill")
    if not fill or not fill.get("visible"):
        return False, None
    if (fill.get("transparency") or 0.0) > _OPAQUE_ENOUGH:
        return True, None
    if fill.get("solid"):
        return True, fill.get("color_hex")
    gradient = fill.get("gradient")
    if gradient and _opaque(gradient):
        if _turned(shape):
            # Turned with its shape, the gradient runs at an angle of its own.
            gradient = {**gradient, "angle": None}
        return True, {"gradient": gradient, "box": shape_box(shape)}
    return True, None


def _opaque(gradient) -> bool:
    """Whether every stop of a gradient hides what is under it."""
    stops = gradient.get("stops") or []
    return bool(stops) and all((stop[2] or 0.0) <= _OPAQUE_ENOUGH for stop in stops)


def _backdrop_colour(shape, shapes, background_hex, template=()):
    """The one colour the shape's text is drawn on, or None when it is not one."""
    colour = _backdrop(shape, shapes, background_hex, template)[0]
    return colour if isinstance(colour, str) else None


def _as_designed(shape, source, follows_background) -> bool:
    """Whether a placeholder's text is still the template's own design.

    A placeholder keeps the colours its layout gave it, and the layout chose
    them for what it draws behind it: a section layout's dark panel, a caption
    column. PowerPoint for Mac cannot read a layout's own graphics, so that
    panel may not be in what was read, and judging the text against the slide's
    background then reports a light title on a dark panel as unreadable.

    So text in theme colours, on a placeholder, over what the template paints,
    is taken as meant. Text the caller coloured, text on a shape the caller
    added, and text on a background the caller changed are the caller's, and
    are checked. So is every placeholder when the caller changes the theme's
    colours, which ``trust_template`` turns this off for.
    """
    if shape.get("placeholder_type") is None:
        return False
    if source == "background":
        if follows_background is False:
            return False
    elif source != "template":
        return False
    runs = (shape.get("text") or {}).get("runs") or []
    return bool(runs) and all(run.get("theme") for run in runs)


def _low_contrast(slide_index, shapes, background_hex, template=(), follows_background=None,
                  layout=None, trust_template=True, background_gradient=None, theme_colors=None):
    from ppt_com.themes import _contrast_ratio

    # A theme whose own text colour does not read on its own background was
    # changed by hand, and is no longer the template's design either.
    trouble = _theme_trouble(theme_colors)
    trust = trust_template and trouble is None
    slide_box = None
    if layout and layout.get("slide_width") and layout.get("slide_height"):
        slide_box = (0.0, 0.0, layout["slide_width"], layout["slide_height"])
    findings = []
    for shape in shapes:
        text = shape.get("text")
        if not shape.get("visible", True) or not text or not text.get("chars"):
            continue
        backdrop, painted_by, painter = _backdrop(shape, shapes, background_hex, template,
                                                  background_gradient, slide_box)
        if backdrop is None and painter is not None and painter.get("photo"):
            found = _photo_contrast(slide_index, shape, painter, theme_colors)
            if found is not None:
                findings.append(found)
            continue
        if not backdrop or (trust and _as_designed(shape, painted_by, follows_background)):
            continue
        if isinstance(backdrop, dict):
            found = _gradient_contrast(slide_index, shape, shapes, backdrop, painted_by, painter,
                                       theme_colors)
            if found is not None:
                findings.append(found)
            continue
        worst = None
        for run in text.get("runs") or []:
            colour = run.get("color_hex")
            if not colour:
                continue
            size = run.get("size") or 0.0
            large = size >= _LARGE_TEXT or (run.get("bold") and size >= _LARGE_BOLD_TEXT)
            needed = _CONTRAST_LARGE if large else _CONTRAST_TEXT
            ratio = _contrast_ratio(colour, backdrop)
            if ratio < needed and (worst is None or ratio / needed < worst[0] / worst[1]):
                worst = (ratio, needed, colour, size)
        if worst is None:
            continue
        ratio, needed, colour, size = worst
        name = shape["name"]
        # The colour the slide's other text already uses keeps it looking like
        # one slide; black or white is the answer when none of them is enough.
        ours = [c for c in _text_colours(shapes, shape, size)
                if _contrast_ratio(c, backdrop) >= needed]
        themed = _theme_text_colours(theme_colors, lambda c, b=backdrop: _contrast_ratio(c, b), needed)
        if ours:
            suggestion, source = ours[0], ", the colour of the other text on the slide"
        elif themed:
            suggestion, source = themed[0], ", the theme's text colour"
        else:
            dark, light = _contrast_ratio("#000000", backdrop), _contrast_ratio("#FFFFFF", backdrop)
            suggestion, source = ("#000000" if dark >= light else "#FFFFFF"), ""
        recolour = (f"make the text {suggestion}{source} (contrast "
                    f"{_contrast_ratio(suggestion, backdrop):.1f}:1): ppt_format_text "
                    f"color='{suggestion}'")
        on = ""
        fix = f"{recolour[0].upper()}{recolour[1:]}. Or change what is behind it."
        runs = text.get("runs") or []
        if not trust and runs and all(run.get("theme") for run in runs):
            # The theme's colours were changed, and this text takes its colour
            # from the theme: putting the theme right fixes every slide.
            why = f", and {trouble}" if trouble else ""
            fix = (f"Its colour is the theme's{why}: choose theme colours that read on each "
                   f"other (ppt_set_theme_colors), or {recolour}.")
        elif painted_by == "template" and painter is not None:
            # Text the caller put on a band or panel of the template's: the
            # template keeps that place for itself, so off it is the better fix.
            on = f" It sits on '{painter['name']}', part of the template."
            move = None
            if layout is not None and shape.get("placeholder_type") is None:
                obstacles = [(g, ink_box(g) or shape_box(g)) for g in template or ()
                             if g.get("visible", True) and not _is_line(g)
                             and not _is_backdrop(g, layout.get("slide_width"), layout.get("slide_height"))]
                move = _template_move(shape, text_box(shape), shape_box(painter),
                                      f"'{painter['name']}'", shapes, obstacles, layout)
            if move:
                fix = f"{move[:-1]}. Or {recolour}."
        findings.append(_finding(
            slide_index, "low_contrast", [name],
            f"Text in '{name}' is {colour} on {backdrop}, a contrast of "
            f"{ratio:.1f}:1 where {needed:g}:1 is the minimum for its size.{on}",
            fix, text_color=colour, backdrop_color=backdrop,
            contrast=round(ratio, 2), minimum=needed,
        ))
    return findings


# Text on a photo is hard to read when this share of what is behind it, or
# more, falls short of the contrast its size needs. A colour that reads on all
# but this share of it is good enough to suggest.
_PHOTO_SHORT = 0.2
_PHOTO_GOOD = 0.05


def _photo_contrast(slide_index, shape, picture, theme_colors=None):
    """Text drawn on a photo, judged against the photo's own pixels under it.

    A photo is light in one place and dark in another, so no one colour
    stands for it: the text is held to every sampled pixel behind it, and
    reported when a fifth of them or more fall short.
    """
    from ppt_com.themes import _contrast_ratio

    words = text_box(shape)
    photo = picture["photo"]
    box = shape_box(picture)
    if words is None or _width(box) <= 0 or _height(box) <= 0:
        return None
    columns, rows = photo["columns"], photo["rows"]
    behind = []
    for r in range(rows):
        y = box[1] + (r + 0.5) * _height(box) / rows
        if not words[1] <= y <= words[3]:
            continue
        for c in range(columns):
            x = box[0] + (c + 0.5) * _width(box) / columns
            # A see-through pixel shows what is under the picture, which the
            # sample does not hold.
            if words[0] <= x <= words[2] and photo["colours"][r * columns + c]:
                behind.append(photo["colours"][r * columns + c])
    if len(behind) < 4:
        return None
    worst = None
    for run in (shape.get("text") or {}).get("runs") or []:
        colour = run.get("color_hex")
        if not colour:
            continue
        size = run.get("size") or 0.0
        large = size >= _LARGE_TEXT or (run.get("bold") and size >= _LARGE_BOLD_TEXT)
        needed = _CONTRAST_LARGE if large else _CONTRAST_TEXT
        ratios = sorted(_contrast_ratio(colour, b) for b in behind)
        short = sum(1 for r in ratios if r < needed) / len(ratios)
        if short >= _PHOTO_SHORT and (worst is None or short > worst[0]):
            worst = (short, ratios[0], needed, colour)
    if worst is None:
        return None
    short, least, needed, colour = worst
    name = shape["name"]

    def fails(candidate):
        return sum(1 for b in behind if _contrast_ratio(candidate, b) < needed) / len(behind)

    candidates = [c for key in ("dark1", "light1") if (c := (theme_colors or {}).get(key))]
    candidates += ["#000000", "#FFFFFF"]
    best = min(dict.fromkeys(candidates), key=fails)
    if fails(best) <= _PHOTO_GOOD:
        fix = (f"Make the text {best} (ppt_format_text color='{best}'), which reads on "
               f"{round(100 * (1 - fails(best)))}% of the photo behind it")
    else:
        light = sum(_luminance(b) for b in behind) / len(behind) > 0.5
        scrim = "#FFFFFF" if light else "#000000"
        pad = 8.0
        # A side that lands within a few points of the text box's own edge
        # goes on it: a card just off the box reads as a slip.
        frame = shape_box(shape)
        card = [words[0] - pad, words[1] - pad, words[2] + pad, words[3] + pad]
        for side in range(4):
            if abs(card[side] - frame[side]) <= pad:
                card[side] = frame[side]
        fix = (f"Put a see-through card behind the text: ppt_add_shape shape_type='rectangle' "
               f"left={_num(card[0])} top={_num(card[1])} "
               f"width={_num(card[2] - card[0])} height={_num(card[3] - card[1])} "
               f"fill_color='{scrim}' fill_transparency=0.35 line_visible=false, "
               f"bring the text back in front of it (ppt_set_shape_zorder shape_name='{name}' "
               f"command='bring_to_front'), and make the text {'#000000' if light else '#FFFFFF'}")
    return _finding(
        slide_index, "low_contrast", [name, picture["name"]],
        f"Text in '{name}' is {colour} on the photo '{picture['name']}': "
        f"{round(100 * short)}% of the photo behind it falls short of {needed:g}:1 "
        f"against it, as little as {least:.1f}:1.",
        f"{fix}. Or move the text to a calmer part of the photo.",
        contrast_ratio=round(least, 2), required_ratio=needed, on_photo=True,
    )


def _luminance(colour):
    red, green, blue = (int(colour[i:i + 2], 16) / 255 for i in (1, 3, 5))
    return 0.2126 * red + 0.7152 * green + 0.0722 * blue


def _theme_text_colours(theme_colors, ratio, needed):
    """The theme's text colours that read on a backdrop, best first.

    A write reads only the shape it touched, so the slide's other text
    colours are not known to it; the theme's are, and keep the slide in its
    palette where black or white would not.
    """
    found = [colour for key in ("dark1", "light1", "dark2", "light2")
             if (colour := (theme_colors or {}).get(key)) and ratio(colour) >= needed]
    return sorted(dict.fromkeys(found), key=ratio, reverse=True)


def _text_colours(shapes, exclude, size=None):
    """The colours the slide's other text is set in.

    The colour of text about ``size`` first, body text for body text rather
    than the colour of the headings, and then the most used.
    """
    counts = {}
    for shape in shapes:
        if shape is exclude or not shape.get("visible", True):
            continue
        for run in (shape.get("text") or {}).get("runs") or []:
            colour = (run.get("color_hex") or "").upper()
            if not colour:
                continue
            alike = bool(size) and abs((run.get("size") or 0.0) - size) <= 2.0
            seen = counts.setdefault(colour, [0, 0])
            seen[0] += alike
            seen[1] += 1
    return sorted(counts, key=lambda c: (-counts[c][0], -counts[c][1]))


# A theme's text colour on its background colour, and its second pair, read
# at least this well in every theme Office ships.
_THEME_PAIRS = (("dark1", "light1", _CONTRAST_TEXT), ("dark2", "light2", _CONTRAST_LARGE))


def _theme_trouble(colours):
    """The pair of a theme's own colours that does not read, in words, or None."""
    from ppt_com.themes import _contrast_ratio

    for text, back, needed in _THEME_PAIRS:
        a, b = (colours or {}).get(text), (colours or {}).get(back)
        if a and b and _contrast_ratio(a, b) < needed:
            return f"the theme's {text}, {a}, is {_contrast_ratio(a, b):.1f}:1 on its {back}, {b}"
    return None


def _gradient_contrast(slide_index, shape, shapes, paint, painted_by, painter, theme_colors=None):
    """Text too faint on some of the gradient behind it, or None.

    Where a linear gradient's angle is known, the colours behind the text are
    those across the part of it the text covers. Otherwise the text could be
    anywhere on it, and only text too faint on all of it is reported.
    PowerPoint blends a theme's gradient straight across in sRGB, and one made
    with a two-colour fill along a curve in linear light. Which a fill uses
    cannot be read, so a verdict has to hold for both.
    """
    from ppt_com.themes import _contrast_ratio

    stops = sorted(paint["gradient"]["stops"], key=lambda stop: stop[0])
    words = text_box(shape)
    span = _gradient_span(paint["box"], paint["gradient"].get("angle"), words) if words else None
    blends = [_gradient_colours(stops, span, blend) for blend in (_blend_srgb, _blend_curve)]
    worst = None
    for run in shape["text"].get("runs") or []:
        colour = run.get("color_hex")
        if not colour:
            continue
        size = run.get("size") or 0.0
        large = size >= _LARGE_TEXT or (run.get("bold") and size >= _LARGE_BOLD_TEXT)
        needed = _CONTRAST_LARGE if large else _CONTRAST_TEXT
        pick = min if span else max
        # The kinder of the two blends: a finding has to hold for either.
        ratio = max(pick(_contrast_ratio(colour, c) for c in colours) for colours in blends)
        if ratio < needed and (worst is None or ratio / needed < worst[0] / worst[1]):
            worst = (ratio, needed, colour, size)
    if worst is None:
        return None
    ratio, needed, colour, size = worst
    name = shape["name"]
    ends = (_colour_at(stops, span[0], _blend_srgb), _colour_at(stops, span[1], _blend_srgb)) \
        if span else (stops[0][1], stops[-1][1])
    if painted_by == "background" or painter is None:
        what = "the background's gradient"
    elif painted_by == "fill":
        what = "its own gradient fill"
    elif painted_by == "template":
        what = f"'{painter['name']}', a gradient of the template's"
    else:
        what = f"'{painter['name']}', a gradient"
    if span:
        on = f"on {what}, {ends[0]} to {ends[1]} behind it, a contrast of as little as {ratio:.1f}:1"
    else:
        on = f"on {what}, {ends[0]} to {ends[1]}, a contrast of at most {ratio:.1f}:1 anywhere on it"
    # A colour that reads on every part of the gradient the text could be on.
    everywhere = [c for colours in blends for c in colours]
    ours = [c for c in _text_colours(shapes, shape, size)
            if min(_contrast_ratio(c, b) for b in everywhere) >= needed]
    themed = _theme_text_colours(theme_colors, lambda c: min(_contrast_ratio(c, b) for b in everywhere),
                                 needed)
    choices = ([(c, ", the colour of the other text on the slide") for c in ours[:1]]
               + [(c, ", the theme's text colour") for c in themed[:1]] + [
        (c, "") for c in ("#000000", "#FFFFFF")
        if min(_contrast_ratio(c, b) for b in everywhere) >= needed])
    if choices:
        suggestion, source = choices[0]
        least = min(_contrast_ratio(suggestion, b) for b in everywhere)
        there = "there" if span else "anywhere on it"
        fix = (f"Make the text {suggestion}{source} (contrast {least:.1f}:1 or more {there}): "
               f"ppt_format_text color='{suggestion}'. Or put it on a card of its own colour.")
    else:
        fix = (f"No one colour reads on all of the gradient behind '{name}': put it on a card "
               "of its own colour, or move it where the gradient is all dark or all light.")
    return _finding(
        slide_index, "low_contrast", [name],
        f"Text in '{name}' is {colour} {on}, where {needed:g}:1 is the minimum for its size.",
        fix, text_color=colour, backdrop_gradient=list(ends),
        contrast=round(ratio, 2), minimum=needed,
    )


def _gradient_span(box, angle, words):
    """The lowest and highest positions along a linear gradient that the words cover.

    PowerPoint lays a linear gradient's angle across its box as if the box were
    square, so its ends fall on the box's corners whatever the box's shape.
    None when the gradient does not run in a line or its angle is unknown.
    """
    if angle is None:
        return None
    x0, y0, x1, y1 = box
    width, height = x1 - x0, y1 - y0
    if width <= 0 or height <= 0:
        return None
    c, s = math.cos(math.radians(angle)), math.sin(math.radians(angle))
    across = abs(c) + abs(s)
    left, right = max(words[0], x0), min(words[2], x1)
    top, bottom = max(words[1], y0), min(words[3], y1)
    if right < left or bottom < top:
        return None
    ts = [((x - x0) / width - 0.5) * c / across + ((y - y0) / height - 0.5) * s / across + 0.5
          for x in (left, right) for y in (top, bottom)]
    return max(0.0, min(ts)), min(1.0, max(ts))


def _gradient_colours(stops, span, blend):
    """The colours a gradient takes over ``span``, or all of it when that is None."""
    low, high = span if span else (0.0, 1.0)
    positions = {low, high} | {stop[0] for stop in stops if low < stop[0] < high}
    positions |= {low + (high - low) * i / 8 for i in range(1, 8)}
    return [_colour_at(stops, t, blend) for t in sorted(positions)]


def _colour_at(stops, t, blend):
    """The colour of a gradient at position ``t``, its stops blended by ``blend``."""
    if t <= stops[0][0]:
        return stops[0][1]
    for (p0, c0, _), (p1, c1, _) in pairwise(stops):
        if t <= p1:
            return c0 if p1 - p0 < 1e-9 else blend(c0, c1, (t - p0) / (p1 - p0))
    return stops[-1][1]


def _channels(colour):
    colour = colour.lstrip("#")
    return [int(colour[i:i + 2], 16) for i in (0, 2, 4)]


def _blend_srgb(a, b, f):
    """Straight across in sRGB, the way a theme's gradient is drawn."""
    return "#" + "".join(f"{round(x + (y - x) * f):02X}"
                         for x, y in zip(_channels(a), _channels(b)))


def _blend_curve(a, b, f):
    """Along an S-curve in linear light, the way a two-colour fill is drawn."""
    def linear(v):
        v /= 255.0
        return v / 12.92 if v <= 0.04045 else ((v + 0.055) / 1.055) ** 2.4

    def encoded(v):
        v = 12.92 * v if v <= 0.0031308 else 1.055 * v ** (1 / 2.4) - 0.055
        return round(min(max(v, 0.0), 1.0) * 255)

    f = f * f * (3 - 2 * f)
    return "#" + "".join(f"{encoded(linear(x) + (linear(y) - linear(x)) * f):02X}"
                         for x, y in zip(_channels(a), _channels(b)))


# Edge names per axis, in the order they are preferred when two tie.
_X_EDGES = ("left", "center", "right")
_Y_EDGES = ("top", "middle", "bottom")


def _edges(box, axis):
    if axis == "x":
        return dict(zip(_X_EDGES, (box[0], (box[0] + box[2]) / 2, box[2]), strict=True))
    return dict(zip(_Y_EDGES, (box[1], (box[1] + box[3]) / 2, box[3]), strict=True))


def _related(a_box, b_box, axis) -> bool:
    """Shapes stacked in a column are lined up left and right; shapes in a row,
    top and bottom. Shapes that are neither have nothing to line up."""
    if axis == "x":
        shared = _shared(a_box, b_box)[0]
        return shared >= _SHARED_SPAN * min(_width(a_box), _width(b_box))
    shared = _shared(a_box, b_box)[1]
    return shared >= _SHARED_SPAN * min(_height(a_box), _height(b_box))


def _within(edge, gap, centre=_ALIGNED_CENTRE) -> bool:
    """Whether two lines this far apart are one line.

    ``centre`` is how far apart two centres can be and still be one; None
    holds them to an edge's allowance. Between two shapes of one size, a
    centre off is every edge off by as much.
    """
    if centre is not None and edge in ("center", "middle"):
        return gap <= centre
    return gap <= _ALIGNED


def _misaligned(slide_index, layout, only=None):
    """Edges a few points apart that were meant to be one edge.

    Each pair of shapes in a row is compared top, middle and bottom, and each
    pair in a column left, centre and right. A pair already sharing one of
    those lines is aligned. A pair whose nearest lines are less than
    ``_NEAR_MISS`` apart, but not together, is the slip this reports.

    Of the two, the one to move is the one lined up with fewer other shapes,
    and between equals the one added later. One finding names it with the
    shapes on the line it nearly meets.

    Only shapes on the same level are compared: cards with cards, and the text
    on one card with the text on the next. Text is placed on its card, inset
    from its edges, rather than lined up with it or with the cards beside it.
    Of a text frame only the edges its text keeps to are compared: the centre
    of a centred name, not the sides of a frame no one can see.

    Levels are judged from the cards down. The icon and text on a card that
    slipped slipped with it, so they are judged where the card's fix will put
    them, and the card's fix moves them along.
    """
    arranged, boxes, depth = layout["arranged"], layout["boxes"], layout["depth"]
    seen = {s["name"]: _seen_edges(s) for s in arranged}
    # A photo the full height of the slide has its middle at the slide's
    # middle, which no list beside it is lined up with; likewise across.
    slide_w, slide_h = layout.get("slide_width"), layout.get("slide_height")
    for s in arranged:
        box, dropped = boxes[s["name"]], set()
        if slide_h and box[1] <= _EDGE_TOLERANCE and box[3] >= slide_h - _EDGE_TOLERANCE:
            dropped.add("middle")
        if slide_w and box[0] <= _EDGE_TOLERANCE and box[2] >= slide_w - _EDGE_TOLERANCE:
            dropped.add("center")
        if dropped:
            edges = seen[s["name"]]
            every = {"left", "center", "right", "top", "middle", "bottom"}
            seen[s["name"]] = (every if edges is None else edges) - dropped
    placed = dict(boxes)
    slips = []
    for level in sorted(set(depth.values())):
        members = [s for s in arranged if depth[s["name"]] == level]
        found = []
        for axis in ("x", "y"):
            found += _slips_along(axis, members, placed, seen, _card_agreement(members, axis, layout))
        for slip in found:
            mover = slip["mover"]
            slip["carried"] = [
                s for s in arranged
                if depth[s["name"]] > level and paints_box(mover)
                and _belongs_to(s, mover, boxes)
            ]
            for shape in slip["carried"]:
                placed[shape["name"]] = _moved(placed[shape["name"]], slip["axis"], -slip["delta"])
        slips += found

    findings = []
    for slip in slips:
        finding = _misaligned_finding(slide_index, slip, placed)
        if only is None or set(finding["shapes"]) & set(only):
            findings.append(finding)
    return findings


def _off_centre_labels(slide_index, layout):
    """Centred labels set well off the centre of the mark they label.

    A label under a dot or an icon is centred on it. One a few points off is
    an ordinary slip and the alignment check reports it; this finds the label
    whose frame was drawn wider than the rest from the same left offset, so
    its text is centred a long way to one side, up to a quarter of its own
    width, beyond which it is not that mark's label at all.
    """
    boxes, depth = layout["boxes"], layout["depth"]
    findings = []
    for label in layout["arranged"]:
        words = text_box(label) if _centred_text(label) else None
        if words is None:
            continue
        mark, nearest = None, None
        for shape in layout["arranged"]:
            if shape is label or not paints_box(shape) or depth[shape["name"]] != depth[label["name"]]:
                continue
            box = boxes[shape["name"]]
            gap = max(box[1] - words[3], words[1] - box[3])
            if not _related(words, box, "x") or not -_EDGE_TOLERANCE <= gap <= _LABEL_GAP:
                continue
            if nearest is None or gap < nearest:
                mark, nearest = shape, gap
        if mark is None:
            continue
        frame, mark_box = boxes[label["name"]], boxes[mark["name"]]
        off = (frame[0] + frame[2]) / 2 - (mark_box[0] + mark_box[2]) / 2
        if not _NEAR_MISS < abs(off) <= 0.25 * _width(words):
            continue
        name, left = label["name"], round(frame[0] - off, 2)
        findings.append(_finding(
            slide_index, "misaligned", [name, mark["name"]],
            f"The center of '{name}' is {_pt(abs(off))} "
            f"{'right' if off > 0 else 'left'} of the center of '{mark['name']}', "
            "the mark it labels.",
            f"Set left of '{name}' to {_pt(left)} so it is centred on "
            f"'{mark['name']}' (ppt_update_shape left={left:g}).",
            severity="warning", edge="center", off_by_pt=round(abs(off), 2),
        ))
    return findings


def _moved(box, axis, by):
    """A box moved along one axis."""
    if axis == "x":
        return (box[0] + by, box[1], box[2] + by, box[3])
    return (box[0], box[1] + by, box[2], box[3] + by)


def _card_agreement(members, axis, layout):
    """For each shape on a card, how many others sit at the same place on theirs.

    The price on one pricing card, set 2 pt lower on its card than the prices
    on the others, is the one that slipped, whichever card is the odd one.
    """
    boxes, start = layout["boxes"], 0 if axis == "x" else 1
    offsets = {}
    for shape in members:
        card = _card_of(shape, layout)
        if card is not None:
            offsets[shape["name"]] = boxes[shape["name"]][start] - boxes[card["name"]][start]
    return {
        name: sum(1 for other, value in offsets.items()
                  if other != name and abs(value - mine) <= _ALIGNED)
        for name, mine in offsets.items()
    }


def _slips_along(axis, members, boxes, seen, agreement=None):
    """The slips among one level's shapes along one axis."""
    agreement = agreement or {}
    aligned = {s["name"]: agreement.get(s["name"], 0) for s in members}
    near = []
    for i, a in enumerate(members):
        for b in members[i + 1:]:
            a_box, b_box = boxes[a["name"]], boxes[b["name"]]
            if not _related(a_box, b_box, axis):
                continue
            a_edges, b_edges = _edges(a_box, axis), _edges(b_box, axis)
            gaps = {
                edge: abs(a_edges[edge] - b_edges[edge]) for edge in a_edges
                if all(seen[s["name"]] is None or edge in seen[s["name"]] for s in (a, b))
            }
            if not gaps:
                continue
            span = _width if axis == "x" else _height
            allowance = None
            if abs(span(a_box) - span(b_box)) > 2 * _ALIGNED:
                # The middle of a line box is not where its letters look
                # centred, and the gap between the two grows with the type: a
                # big figure and its label are not held to a point and a half.
                text_only = not paints_box(a) and not paints_box(b)
                allowance = (max(_ALIGNED_CENTRE, 0.05 * (span(a_box) + span(b_box)))
                             if text_only else _ALIGNED_CENTRE)
            if any(_within(edge, gap, allowance) for edge, gap in gaps.items()):
                aligned[a["name"]] += 1
                aligned[b["name"]] += 1
                continue
            closest = min(gaps, key=gaps.get)
            if gaps[closest] <= _NEAR_MISS:
                near.append((a, b, closest))

    grouped = {}
    for a, b, edge in near:
        # The one that stays is lined up with more shapes, or was there
        # first. Stacking order is creation order unless something moved a
        # shape in it, which is as close as a slide comes to a history.
        if (aligned[a["name"]], -a.get("z", 0)) > (aligned[b["name"]], -b.get("z", 0)):
            anchor, mover = a, b
        else:
            anchor, mover = b, a
        grouped.setdefault((mover["name"], edge), (mover, edge, []))[2].append(anchor)
    return [_slip(axis, mover, edge, anchors, boxes) for mover, edge, anchors in grouped.values()]


def _slip(axis, mover, edge, anchors, boxes):
    """How far ``mover``'s edge is from the line most of ``anchors`` share."""
    values = [_edges(boxes[a["name"]], axis)[edge] for a in anchors]
    # Where most of the others are; between equals, the nearest.
    mine = _edges(boxes[mover["name"]], axis)[edge]
    target = max(
        {round(v, 2) for v in values},
        key=lambda v: (sum(1 for x in values if _within(edge, abs(x - v))), -abs(v - mine)),
    )
    # Only the shapes on that line are named. One that is off it as well is
    # off against them too, and has its own finding.
    on_line = [a for a, v in zip(anchors, values, strict=True) if _within(edge, abs(v - target))]
    return {"axis": axis, "mover": mover, "edge": edge, "anchors": on_line,
            "delta": mine - target}


def _misaligned_finding(slide_index, slip, placed):
    mover, axis, edge, delta = slip["mover"], slip["axis"], slip["edge"], slip["delta"]
    carried = slip.get("carried") or []
    name = mover["name"]
    field = "left" if axis == "x" else "top"
    # Where it goes is worked out from where it will be once the card it sits
    # on has been put back, which is where it was judged.
    value = round(placed[name][0 if axis == "x" else 1] - delta, 2)
    if axis == "x":
        direction, way = ("right of", "left") if delta > 0 else ("left of", "right")
    else:
        direction, way = ("below", "up") if delta > 0 else ("above", "down")
    others = ", ".join(f"'{a['name']}'" for a in slip["anchors"])
    if carried:
        names = [name] + [c["name"] for c in carried]
        listed = ", ".join(f"'{c['name']}'" for c in carried)
        fix = (f"Move '{name}' {way} by {_pt(abs(delta))} together with the shapes "
               f"on it, {listed} (ppt_update_shape shape_names={names} "
               f"d{field}={_num(-delta)}).")
    else:
        names = [name]
        fix = (f"Set {field} of '{name}' to {_pt(value)} (ppt_update_shape "
               f"{field}={value:g}), or ppt_align_shapes with align_to='{edge}'.")
    return _finding(
        slide_index, "misaligned", names[:1] + [a["name"] for a in slip["anchors"]] + names[1:],
        f"The {edge} of '{name}' is {_pt(abs(delta))} {direction} the "
        f"{edge} of {others}, close enough to look like a slip.",
        fix, severity="warning", edge=edge, off_by_pt=round(abs(delta), 2),
    )


# A title whose letters start this far from the edge the slide's body keeps,
# and no further, was meant to line up with it: further is a layout that sets
# the title apart on purpose, an indented band or a side column.
_TITLE_BODY_MIN = 2.0
_TITLE_BODY_MAX = 12.0


def _title_body_edge(slide_index, shapes, layout):
    """A title whose letters do not start where the body below it does.

    A title box set at the same left as the cards under it starts its words
    one inset further in, 7.2 pt by default, and the eye reads the title and
    the body as two edges a little apart. A right-set title, a Hebrew deck's,
    is held to the body's right edge. The template's own title over its own
    body placeholder is its design and left alone.
    """
    slide_w, slide_h = layout["slide_width"], layout["slide_height"]
    title = find_title({"shapes": shapes}, slide_h) if slide_h else None
    if title is None:
        return []
    # A write reads only what it touched; a shape above whose words were not
    # read may be the real title, and the one found only the first line under it.
    if any(s.get("unread") and s["top"] <= title["top"] + _EDGE_TOLERANCE for s in shapes):
        return []
    side = _title_side(title)
    words = text_box(title)
    if side == "center" or words is None:
        return []
    right = side == "right"
    body = []
    for shape in layout["arranged"]:
        if (shape is title or layout["depth"].get(shape["name"]) != 0
                or shape.get("placeholder_type") in _LAYOUT_PLACEHOLDERS):
            continue
        ink = ink_box(shape)
        if ink is None or ink[1] < words[3] - _EDGE_TOLERANCE:
            continue
        # A shape run off an edge is a bleed, not part of the column.
        if ink[0] <= _EDGE_TOLERANCE or (slide_w and ink[2] >= slide_w - _EDGE_TOLERANCE):
            continue
        # Centred text, or text set to the other side, keeps no edge on this
        # one: where its words start moves with their length, and with half
        # of any change made to its frame.
        if ("right" if right else "left") not in _kept_edges(shape):
            continue
        body.append((shape, ink))
    if not body:
        return []
    edge = max(ink[2] for _, ink in body) if right else min(ink[0] for _, ink in body)
    keepers = [s for s, ink in body if abs((ink[2] if right else ink[0]) - edge) <= _ALIGNED]
    delta = (words[2] if right else words[0]) - edge
    if not _TITLE_BODY_MIN < abs(delta) <= _TITLE_BODY_MAX:
        return []
    if title.get("placeholder_type") is not None and all(
            k.get("placeholder_type") is not None for k in keepers):
        return []
    name = title["name"]
    others = ", ".join(f"'{k['name']}'" for k in keepers)
    where = "right" if right else "left"
    if delta > 0:
        direction = "left of" if right else "right of"
    else:
        direction = "right of" if right else "left of"
    return [_finding(
        slide_index, "title_edge", [name] + [k["name"] for k in keepers],
        f"The {where} end of the title's words in '{name}' is {_pt(abs(delta))} "
        f"{direction} the {where} edge of {others} below it, so the title and "
        "the body read as two edges a little apart.",
        _body_to_title(keepers, delta, right, layout),
        severity="warning", edge=where, off_by_pt=round(abs(delta), 2),
    )]


def _body_to_title(keepers, delta, right, layout):
    """Bring the body's edge to the title's words, keeping its other edge.

    The body moves rather than the title, since where a deck's titles sit is
    held from slide to slide. A row of cards is fitted again with its gaps
    kept, and each card carries what is on it.
    """
    boxes = layout["boxes"]
    side = "right" if right else "left"
    lead = f"Bring the body's {side} edge to the title's words"
    row = layout["rows"].get(keepers[0]["name"])
    if row and len(row) > 1:
        edge = boxes[keepers[0]["name"]][2 if right else 0] + delta
        # The whole row widens with its first card, so that it stays one
        # width: one card made wider is the odd one out.
        fix = _fit_row(row, layout, 0, None if right else edge, edge if right else None,
                       "row", "", lead=lead, grow=True)
        if fix:
            return fix
    calls = []
    for keeper in keepers:
        carried = [s["name"] for s in layout["arranged"]
                   if s is not keeper and s["name"] in boxes and _belongs_to(s, keeper, boxes)]
        if right:
            calls.append(f"ppt_update_shape shape_name='{keeper['name']}' dwidth={_num(delta)}")
        elif carried:
            calls.append(f"ppt_update_shape shape_names={[keeper['name']] + carried} "
                         f"dleft={_num(delta)}, then ppt_update_shape shape_name="
                         f"'{keeper['name']}' dwidth={_num(-delta)}")
        else:
            calls.append(f"ppt_update_shape shape_name='{keeper['name']}' "
                         f"dleft={_num(delta)} dwidth={_num(-delta)}")
    return f"{lead}, keeping its {'left' if right else 'right'} edge: {'; '.join(calls)}."


def _connected(members, joined):
    """Split members into groups, two sharing a group when ``joined`` says so,
    directly or through others."""
    parent = list(range(len(members)))

    def root(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    for i, a in enumerate(members):
        for j in range(i + 1, len(members)):
            if joined(a, members[j]):
                parent[root(i)] = root(j)
    groups = {}
    for i, member in enumerate(members):
        groups.setdefault(root(i), []).append(member)
    return list(groups.values())


def _uneven_spacing(slide_index, layout, slide_w):
    """Rows and columns of like shapes whose gaps differ.

    Three or more shapes of about the same size, side by side in a row or
    stacked in a column, are one group: cards, steps, columns of icons. Their
    gaps are meant to be equal. Shapes sitting inside another shape belong to
    it, so the text on three cards is not a fourth row to space.

    A row that reaches across the slide is spread between the side margins
    rather than between its own two ends. When an end card is the one out of
    place, evening out the gaps around it leaves the row off centre.
    """
    boxes = layout["boxes"]
    margin = (layout["margins"] or {}).get("horizontal")
    findings = []
    for direction, members in layout["groups"]:
        if len(members) < 3:
            continue
        # A row is spaced along x, a column along y.
        start = 0 if direction == "horizontal" else 1
        gaps = [
            boxes[b["name"]][start] - boxes[a["name"]][start + 2]
            for a, b in zip(members, members[1:], strict=False)
        ]
        if min(gaps) <= 0 or max(gaps) - min(gaps) <= _GAP_TOLERANCE:
            continue
        names = [s["name"] for s in members]
        fix = None
        if direction == "horizontal" and margin is not None and _spans(members, boxes, slide_w, margin):
            fix = _row_fix(members, boxes, slide_w, margin, layout)
        if fix is None:
            even = sum(gaps) / len(gaps)
            fix = (f"ppt_distribute_shapes with these shape names and "
                   f"direction='{direction}', or make every gap {_pt(even)}.")
        listed = ", ".join(f"'{n}'" for n in names)
        findings.append(_finding(
            slide_index, "uneven_spacing", names,
            f"The gaps between {listed} are "
            f"{', '.join(_pt(g) for g in gaps)}; a "
            f"{'row' if direction == 'horizontal' else 'column'} of like "
            "shapes reads as one group when they are equal.",
            fix, severity="warning", gaps_pt=[round(g, 1) for g in gaps],
        ))
    return findings


def _featured(names, sizes, layout):
    """Whether the shapes of each size share a colour no other size has.

    A pricing table's middle plan, a chart's top box: made bigger and painted
    differently to stand out. A card that grew by accident kept its colour.
    """
    fills = {}
    for shape in layout["arranged"]:
        if shape["name"] in sizes:
            fills[shape["name"]] = ((shape.get("fill") or {}).get("color_hex") or "").upper()
    if not all(fills.get(n) for n in names):
        return False
    colours = []
    for cluster in _clusters(list(sizes.items()), _SIZE_TOLERANCE):
        shades = {fills[name] for name, _ in cluster}
        if len(shades) != 1:
            return False
        colours.append(shades.pop())
    return len(set(colours)) == len(colours)


def _uneven_sizes(slide_index, layout):
    """Like shapes side by side that are not quite one size.

    Cards in a row are meant to be one height, and in a column one width.
    Shapes far enough apart in size are different kinds of thing and are not
    grouped at all, so what this finds is a card grown to hold its text while
    the rest of its row was left as it was. Only shapes that paint are
    judged, since a text box's frame cannot be seen.
    """
    boxes = layout["boxes"]
    findings, seen = [], set()
    for groups, measure, field, word, kind in (
        (layout["rows"], _height, "height", "tall", "row"),
        (layout["columns"], _width, "width", "wide", "column"),
    ):
        for names in groups.values():
            if tuple(names) in seen:
                continue
            seen.add(tuple(names))
            sizes = {name: measure(boxes[name]) for name in names}
            biggest = max(sizes.values())
            if biggest - min(sizes.values()) <= _SIZE_TOLERANCE:
                continue
            if _featured(names, sizes, layout):
                continue
            if _graded(names, sizes):
                continue
            clusters = sorted(
                _clusters(list(sizes.items()), _SIZE_TOLERANCE),
                key=lambda g: -max(v for _, v in g),
            )
            described = "; ".join(
                f"{_pt(max(v for _, v in g))} for "
                f"{', '.join(repr(n) for n in names if n in dict(g))}"
                for g in clusters
            )
            findings.append(_finding(
                slide_index, "uneven_sizes", names,
                f"A {kind} of like shapes is not one {field}: {described}.",
                f"Make them all {_pt(biggest)} {word}, unless one is meant to "
                f"stand out: ppt_update_shape shape_names={names} "
                f"{field}={_num(biggest)}.",
                severity="warning",
                sizes_pt={name: round(size, 1) for name, size in sizes.items()},
            ))
    return findings + _uneven_along(slide_index, layout, seen)


def _graded(names, sizes) -> bool:
    """Whether like shapes step through their sizes on purpose.

    A slip is one card off the size the rest share. The bars of a chart drawn
    from shapes, a ranking or a funnel share no size at all: three or more of
    them, every one a different size, growing or shrinking in order from the
    first to the last. Their sizes are what they show.
    """
    if len(names) < 3:
        return False
    ordered = [sizes[name] for name in names]
    steps = [b - a for a, b in zip(ordered, ordered[1:])]
    return (all(step > _SIZE_TOLERANCE for step in steps)
            or all(step < -_SIZE_TOLERANCE for step in steps))


def _uneven_along(slide_index, layout, seen):
    """A row whose cards are not one width, or a column not one height.

    Along the row a size can be a design, a wide card beside two narrow ones,
    so only the clear case is reported: three or more like shapes, most of
    them one size, and the rest within a quarter of it. That is a card
    stretched by accident, often into whatever is beside the row.
    """
    boxes = layout["boxes"]
    findings = []
    for groups, measure, field, word, kind, direction in (
        (layout["rows"], _width, "width", "wide", "row", "horizontal"),
        (layout["columns"], _height, "height", "tall", "column", "vertical"),
    ):
        for names in groups.values():
            key = ("along", tuple(names))
            if key in seen or len(names) < 3:
                continue
            seen.add(key)
            sizes = {name: measure(boxes[name]) for name in names}
            clusters = _clusters(list(sizes.items()), _SIZE_TOLERANCE)
            usual = clusters[0]
            if len(clusters) < 2 or 2 * len(usual) <= len(names):
                continue
            size = sorted(v for _, v in usual)[len(usual) // 2]
            odd = [name for name in names if name not in dict(usual)]
            if any(abs(sizes[name] - size) > _ALONG_SIZE_SLIP * size for name in odd):
                continue
            if _featured(names, sizes, layout):
                continue
            described = "; ".join(
                f"{_pt(sizes[name])} for '{name}'" for name in odd
            ) + f"; {_pt(size)} for the rest"
            fix = (f"Make {', '.join(repr(n) for n in odd)} {_pt(size)} {word} like the rest "
                   f"(ppt_update_shape shape_names={odd} {field}={_num(size)})")
            if any(name != names[-1] for name in odd):
                fix += (f", then ppt_distribute_shapes with shape_names={names} and "
                        f"direction='{direction}' to even the gaps again")
            findings.append(_finding(
                slide_index, "uneven_sizes", names,
                f"A {kind} of like shapes is not one {field}: {described}.",
                fix + ".", severity="warning",
                sizes_pt={name: round(v, 1) for name, v in sizes.items()},
            ))
    return findings


def _cramped_text(slide_index, layout):
    """Text pressed against an edge of the card it sits on.

    The text does not cross the edge, so nothing overlaps, but it has next to
    no room on one side and plenty on the other, and reads as cut off. A card
    grown by exactly as much as its text overflowed leaves it just like this.

    The text on a card is judged as one block: a quote with its attribution
    under it, centred together, has its room above the quote and below the
    attribution, not above each line.
    """
    boxes = layout["boxes"]
    on_card = {}
    for shape in layout["arranged"]:
        if paints_box(shape):
            # Text in a shape's own frame is inset by the frame's margins, and
            # whether it fits there is the overflow check's question.
            continue
        words = text_box(shape)
        card = _card_of(shape, layout) if words else None
        if card is not None:
            on_card.setdefault(card["name"], (card, []))[1].append((shape, words))
    findings = []
    for card_name, (card, texts) in on_card.items():
        block = (min(w[0] for _, w in texts), min(w[1] for _, w in texts),
                 max(w[2] for _, w in texts), max(w[3] for _, w in texts))
        room = _room(block, boxes[card_name])
        if min(room.values()) < -_EDGE_TOLERANCE:
            continue  # It runs past the edge, which overlap reports.
        for side in ("bottom", "top", "right", "left"):
            want = _wanted_room(room, side)
            have = max(room[side], 0.0)
            if have >= _MIN_ROOM or have >= 0.5 * want:
                continue
            # Named after the text that sets the block's edge on that side.
            name = next(s["name"] for s, w in texts
                        if abs(_room(w, boxes[card_name])[side] - room[side]) <= 0.05)
            names, short = [s["name"] for s, _ in texts], want - have
            moved = (f"'{name}'" if len(names) == 1
                     else f"the text on '{card_name}' ({', '.join(repr(n) for n in names)})")
            call = "" if len(names) == 1 else f"shape_names={names} "
            if side == "bottom":
                fix = _grow(card, "bottom", short, layout)
            elif side == "top":
                fix = f"move {moved} down by {_pt(short)} (ppt_update_shape {call}dtop={_num(short)})"
                if room["bottom"] - short < want:
                    fix += f", and {_grow(card, 'bottom', want - room['bottom'] + short, layout)}"
            elif side == "left":
                fix = f"move {moved} right by {_pt(short)} (ppt_update_shape {call}dleft={_num(short)})"
            else:
                fix = (f"make '{name}' {_pt(short)} narrower so its lines wrap "
                       f"sooner (ppt_update_shape dwidth={_num(-short)}), or "
                       f"{_grow(card, 'right', short, layout)}")
            opposite = _OPPOSITE[side]
            others = [n for n in names if n != name]
            text = (f"The text of '{name}' is" if not others else
                    f"The text of '{name}', set with {', '.join(repr(n) for n in others)} "
                    "as one block, is")
            findings.append(_finding(
                slide_index, "cramped_text", [name, *others, card_name],
                f"{text} {_pt(have)} from the {side} edge of "
                f"'{card_name}', with {_pt(room[opposite])} {_WHERE[opposite]}, "
                "so it looks pressed against the edge.",
                f"{fix[0].upper()}{fix[1:]}, which leaves it {_pt(want)} "
                f"{_WHERE[side]}{_padding(room, side, want)}.",
                severity="warning", room_pt={k: round(v, 1) for k, v in room.items()},
            ))
            break
    return findings


def _narrow_margins(slide_index, layout, slide_w):
    """Shapes closer to the side of the slide than the rest of it keeps.

    A shape touching the edge bleeds off it on purpose and is left alone, and
    one past it is off the slide, which is its own finding. What is left is a
    shape that stops short of the edge but inside the margin the other shapes
    keep: a row one card too far to the right, or a box pushed back onto the
    slide only just far enough. Only a near miss counts. A logo in the corner
    of a slide whose text is centred well in from the sides is placed there.
    """
    margins = layout["margins"]
    if not margins:
        return []
    side = margins["horizontal"]
    boxes = layout["boxes"]
    rows = {s["name"]: members for direction, members in layout["groups"]
            if direction == "horizontal" for s in members}
    findings = []
    for shape in layout["arranged"]:
        # Centred text is placed by what it labels, not by the margin. Off the
        # slide altogether, it is reported as such.
        if layout["depth"][shape["name"]] or _centred_text(shape):
            continue
        ink = ink_box(shape)
        if ink is None:
            continue
        for edge, distance in (("left", ink[0]), ("right", slide_w - ink[2])):
            if distance <= _EDGE_TOLERANCE or not _near_margin(distance, side):
                continue
            name, box = shape["name"], boxes[shape["name"]]
            row = rows.get(name)
            fix = (_row_fix(row, boxes, slide_w, side, layout)
                   if row and _spans(row, boxes, slide_w, side) else None)
            if fix is None and _width(box) > slide_w - 2 * side:
                fix = (f"Make '{name}' {_pt(_width(box) - slide_w + 2 * side)} "
                       f"narrower and move it to the margin (ppt_update_shape "
                       f"left={_num(side)} width={_num(slide_w - 2 * side)}).")
            elif fix is None:
                # What came too near is the ink, which for text is not always
                # the frame: the frame moves by what its text needs.
                by = min(max(ink[0], side), slide_w - side - _width(ink)) - ink[0]
                fix = (f"Move '{name}' {'right' if by > 0 else 'left'} by "
                       f"{_pt(abs(by))} (ppt_update_shape left={_num(box[0] + by)}).")
            findings.append(_finding(
                slide_index, "narrow_margin", [name],
                f"'{name}' comes within {_pt(distance)} of the {edge} edge of "
                f"the slide, where the rest of the slide keeps {_pt(side)}.",
                fix, severity="warning", edge=edge,
                distance_pt=round(distance, 1), margin_pt=side,
            ))
            break
    return findings


# A picture smaller than this share of the slide is a logo or an icon, set in
# a corner on purpose. Text less tall than this share of the slide is a
# footnote or a page number, set low on purpose.
_LOGO_SHARE = 0.03
_BLOCK_SHARE = 0.10

# The most room the top and bottom are held to: the edge is crowded within
# half of it.
_VERTICAL_ROOM = 40.0


def _vertical_margins(slide_index, layout, slide_h):
    """Content run almost to the top or the bottom edge of the slide.

    The side margins say how much room a slide keeps; a card, a chart or a
    block of text that stops a few points short of the bottom edge looks
    pushed against it. Touching the edge is a bleed and is left alone, and so
    are a footnote or a logo set low, which are placed there. A row of cards
    is one finding, moved or shortened together.
    """
    margins = layout["margins"]
    slide_w = layout["slide_width"]
    if not margins or not slide_h or not slide_w:
        return []
    # A slide with nothing near its top, a quote set in the middle, says
    # little about how close to the edge it means things to go.
    want = min(margins["vertical"], _VERTICAL_ROOM)
    crowded = {}
    for shape in layout["arranged"]:
        if layout["depth"][shape["name"]] or shape.get("placeholder_type") in _LAYOUT_PLACEHOLDERS:
            continue
        ink = ink_box(shape)
        if ink is None:
            continue
        if paints_box(shape):
            if (shape.get("type") in _PICTURES
                    and _width(ink) * _height(ink) < _LOGO_SHARE * slide_w * slide_h):
                continue
        elif _height(ink) < _BLOCK_SHARE * slide_h:
            continue
        for edge, distance in (("top", ink[1]), ("bottom", slide_h - ink[3])):
            if _EDGE_TOLERANCE < distance < want / 2:
                crowded[shape["name"]] = (shape, ink, edge, distance)
                break
    findings, done = [], set()
    for name, (_, ink, edge, distance) in crowded.items():
        if name in done:
            continue
        row = [m for m in layout["rows"].get(name) or [name] if m in crowded
               and crowded[m][2] == edge] or [name]
        done.update(row)
        by = want - min(crowded[m][3] for m in row)
        inks = [crowded[m][1] for m in row]
        # Moving away from the edge runs into whatever is on the other side.
        if edge == "bottom":
            near = min(ink[1] for ink in inks)
            others = [ink_box(s) for s in layout["arranged"] if s["name"] not in row
                      and not layout["depth"][s["name"]]]
            room = min([near - o[3] for o in others if o and o[3] <= near + _EDGE_TOLERANCE
                        and _shared((o[0], 0, o[2], 1), (min(i[0] for i in inks), 0,
                                                         max(i[2] for i in inks), 1))[0] > 0],
                       default=near)
        else:
            far = max(ink[3] for ink in inks)
            others = [ink_box(s) for s in layout["arranged"] if s["name"] not in row
                      and not layout["depth"][s["name"]]]
            room = min([o[1] - far for o in others if o and o[1] >= far - _EDGE_TOLERANCE
                        and _shared((o[0], 0, o[2], 1), (min(i[0] for i in inks), 0,
                                                         max(i[2] for i in inks), 1))[0] > 0],
                       default=slide_h - far)
        names = row
        target = f"shape_names={names}" if len(names) > 1 else f"shape_name='{names[0]}'"
        quoted = [f"'{n}'" for n in names]
        which = quoted[0] if len(names) == 1 else ", ".join(quoted[:-1]) + " and " + quoted[-1]
        if room - _CLEARANCE >= by:
            fix = (f"Move {which} {'down' if edge == 'top' else 'up'} by {_pt(by)} "
                   f"(ppt_update_shape {target} dtop={_num(by if edge == 'top' else -by)}).")
        elif not all(_shortens(crowded[m][0], by, edge) for m in row):
            # The words are what is near the edge, or a shorter box would
            # spill them: a smaller frame only moves the problem.
            fix = (f"Cut the text in {which} or make the font smaller, so "
                   f"{'they' if len(names) > 1 else 'it'} "
                   f"{'start' if edge == 'top' else 'end'}{'' if len(names) > 1 else 's'} "
                   f"{_pt(by)} further from the edge: there is no room to move "
                   f"{'them' if len(names) > 1 else 'it'} {'down' if edge == 'top' else 'up'}, "
                   "and a shorter box would leave the words where they are or spill them.")
        else:
            keep = "bottom" if edge == "top" else "top"
            extra = f" dtop={_num(by)}" if edge == "top" else ""
            fix = (f"Make {which} {_pt(by)} shorter, keeping {'their' if len(names) > 1 else 'its'} "
                   f"{keep} (ppt_update_shape {target} dheight={_num(-by)}{extra}): there is no "
                   f"room to move {'them' if len(names) > 1 else 'it'} "
                   f"{'down' if edge == 'top' else 'up'}.")
        findings.append(_finding(
            slide_index, "narrow_margin", names,
            f"{which[0].upper()}{which[1:]} {'come' if len(names) > 1 else 'comes'} within "
            f"{_pt(distance)} of the {edge} edge of the slide, where the slide keeps "
            f"{_pt(want)}, so {'they look' if len(names) > 1 else 'it looks'} pushed "
            "against the edge.",
            fix, severity="warning", edge=edge, distance_pt=round(distance, 1),
            margin_pt=want,
        ))
    return findings


def _shortens(shape, by, edge):
    """Whether making ``shape`` ``by`` shorter, from ``edge``, moves the edge
    the check sees.

    A box that paints its frame is measured by the frame. For bare text the
    words are, and they move only when they are set to that edge of the box:
    words set to the top stay where they are when the bottom comes up. A box
    whose text would no longer fit trades the margin for an overflow.
    """
    if not paints_box(shape) and (_text_anchor(shape) or (None, None))[1] != edge:
        return False
    return not _spills_at(shape, shape["width"], shape["height"] - by)


def _spills_at(shape, width, height):
    """Whether the text in ``shape`` would spill out of a box this size.

    A box narrower than its words wraps them onto lines it was never measured
    with, so it counts as spilling: how many more lines is not known here.
    """
    text = shape.get("text")
    if not text or not text.get("chars"):
        return False
    margins = text.get("margins") or _DEFAULT_MARGINS
    if text.get("width") and text["width"] > width - margins["left"] - margins["right"] + 0.5:
        return True
    from ppt_com.text import build_measurement

    return build_measurement(
        [], text.get("width"), text.get("height"), width, height, text.get("margins"),
        text.get("word_wrap"), text.get("autofit"), text.get("orientation"),
    )["overflows"]


# Text a template or a draft leaves for someone to replace. The words in any
# case. "Your title" and its kind only as a line of their own, and a bracketed
# stand-in only when it is nothing but the stand-in, so that "Grow your
# company" and "[Number of users]" are text. Stand-in figures only in capitals and only with a currency sign, a
# decimal point or a unit, so that neither a word nor a Roman numeral is one.
_SAMPLE_WORDS = re.compile(
    r"lorem\s+ipsum|click\s+to\s+(?:add|edit)\b"
    r"|(?<![^\n])(?:your|insert)\s+(?:text|title|name|company|logo|content|subtitle|tagline"
    r"|caption|image|photo|picture)\b(?=(?:\s+here)?[ \t]*(?:\n|$))"
    r"|(?:^|\n)\s*(?:add\s+)?(?:text|title|subtitle|caption|content)\s+(?:goes\s+)?here\b"
    r"|\[\s*(?:(?:insert|your)\s+)?(?:(?:company|client|project|presentation|document)\s+)?"
    r"(?:name|date|title|subtitle|tagline|logo|image|text|amount|number)\s*\]",
    re.IGNORECASE,
)
_SAMPLE_FIGURES = re.compile(
    r"(?<![A-Za-z0-9])(?:[$\u20ac\u00a3]XX+(?:[.,]X+)*[%KMB]?|XX+(?:[.,]X+)+[%KMB]?|XX+[%KMB])"
    r"(?![A-Za-z0-9])"
)


def _sample_text(slide_index, shapes):
    """Text a template or a draft left behind to be replaced.

    Lorem ipsum, "Your title here", "[Company name]", "XX%". A model fills a
    slide in passes and loses track of one, and on a finished slide it reads
    as unfinished work.
    """
    findings = []
    for shape in shapes:
        content = (shape.get("text") or {}).get("content") or ""
        if not content or not shape.get("visible", True):
            continue
        found = _SAMPLE_WORDS.search(content) or _SAMPLE_FIGURES.search(content)
        if not found:
            continue
        name = shape["name"]
        findings.append(_finding(
            slide_index, "sample_text", [name],
            f"'{name}' still has stand-in text: \"{found.group(0)}\".",
            f"Replace it with the real content (ppt_find_replace_text, or "
            f"ppt_set_text for all of '{name}'). If '{name}' is not needed, say "
            "so before deleting it: it may be text the slide does want.",
            severity="warning", sample=found.group(0),
        ))
    return findings


# The bidi classes of letters: a paragraph takes its direction from its first
# one. Digits have none of their own.
_LEFT_TO_RIGHT = frozenset({"L"})
_RIGHT_TO_LEFT = frozenset({"R", "AL"})
_DIGITS = frozenset({"EN", "AN"})
# A single one of these between two digits keeps them one number, "1948-1967",
# "19.10.2026" or "1,000", in either direction.
_JOINERS = frozenset({"ES", "CS"})

# The word that makes a range read right to left, by script.
_RANGE_WORDS = {"Hebrew": "{a} עד {b}", "Arabic": "من {a} إلى {b}"}
# What stands between the two ends of a range.
_DASHES = "-–—−~"


def _as_range(line, numbers, gaps):
    """The two ends of a range, "10% – 20%", with their signs, or None.

    A range is two numbers with a dash between them and nothing else but
    spaces and the signs of the numbers, rising: "3 - 1" is a score, and
    "2020: 5" is a label and its figure.
    """
    if len(numbers) != 2:
        return None
    gap = gaps[0]
    dashes = [i for i, ch in enumerate(gap) if ch in _DASHES]
    if len(dashes) != 1 or not all(ch.isspace() or ch in _DASHES or unicodedata.bidirectional(ch) == "ET"
                                   for ch in gap):
        return None
    start = line.index(numbers[0])
    # A minus right before the first number is its sign; one before the
    # second would be a second dash, and the line is not a range.
    sign = -1 if line[:start].rstrip()[-1:] in ("-", "−") else 1
    try:
        if sign * float(numbers[0].replace(",", "")) >= float(numbers[1].replace(",", "")):
            return None
    except ValueError:
        pass
    at = start + len(numbers[0]) + dashes[0]
    return line[:at].strip(), line[at + 1:].strip()


def _numbers_only(line):
    """The separate numbers in a line with no letter in it, and the text
    between them, or None when a letter sets the line's direction.

    Numbers are told apart as the bidi algorithm tells them apart: digits
    joined by one sign or separator, "1948-1967", are one number.
    """
    numbers, gaps, gap = [], [], ""
    in_number = False
    for ch in line:
        kind = unicodedata.bidirectional(ch)
        if kind in _LEFT_TO_RIGHT or kind in _RIGHT_TO_LEFT:
            return None
        if kind in _DIGITS:
            if in_number:
                numbers[-1] += ch
            elif numbers and len(gap) == 1 and unicodedata.bidirectional(gap) in _JOINERS:
                numbers[-1] += gap + ch
            else:
                if numbers:
                    gaps.append(gap)
                numbers.append(ch)
            in_number, gap = True, ""
        else:
            in_number = False
            gap += ch
    return numbers, gaps


def _script(text):
    """'Hebrew' or 'Arabic' when the text reads right to left, 'Latin' when it
    reads left to right, None with no letters."""
    rtl = {"Hebrew": 0, "Arabic": 0}
    ltr = 0
    for ch in text:
        kind = unicodedata.bidirectional(ch)
        if kind in _RIGHT_TO_LEFT:
            rtl["Hebrew" if "֐" <= ch <= "׿" or "יִ" <= ch <= "ﭏ"
                else "Arabic"] += 1
        elif kind in _LEFT_TO_RIGHT:
            ltr += 1
    if not ltr and not any(rtl.values()):
        return None
    script = max(rtl, key=rtl.get)
    return script if rtl[script] > ltr else "Latin"


def _reads_right_to_left(shapes):
    """Whether most of the letters on a slide read right to left."""
    content = " ".join((s.get("text") or {}).get("content") or "" for s in shapes
                       if s.get("visible", True))
    return _script(content) in _RANGE_WORDS


def _paragraphs(content):
    """A shape's paragraphs, in the order PowerPoint numbers them."""
    return re.split(r"\r\n|[\r\n]", content)


def _direction_call(paragraphs, indices):
    """The ppt_set_paragraph_format call that sets these paragraphs right to left."""
    if len(indices) == 1 and len([p for p in paragraphs if p.strip()]) > 1:
        return f"ppt_set_paragraph_format paragraph_index={indices[0]} direction='rtl'"
    if len(indices) > 1 and len(indices) < len([p for p in paragraphs if p.strip()]):
        return ("ppt_set_paragraph_format direction='rtl' with paragraph_index "
                + ", ".join(str(i) for i in indices))
    return "ppt_set_paragraph_format direction='rtl'"


def _number_order(slide_index, shapes):
    """A line of numbers only, on a slide that reads right to left.

    "1516 – 1917" on a Hebrew slide is drawn left to right, so the start year
    is on the left and a reader starting on the right reads the range
    backwards; a list of years comes out the same. Digits are not letters, so
    nothing in the line says otherwise, and PowerPoint draws it that way on
    Windows and on the Mac even with the paragraph set right to left. A word
    in the line, in a paragraph set right to left, reads right: "1516 עד 1917".
    """
    findings = []
    slide_text = " ".join((s.get("text") or {}).get("content") or "" for s in shapes)
    for shape in shapes:
        text = shape.get("text") or {}
        content = text.get("content") or ""
        if not content or not shape.get("visible", True):
            continue
        script = _script(content) or _script(slide_text)
        if script not in _RANGE_WORDS:
            continue
        paragraphs = _paragraphs(content)
        directions = text.get("directions") or []
        for index, paragraph in enumerate(paragraphs, start=1):
            line = paragraph.strip()
            parts = _numbers_only(line)
            if not parts:
                continue
            numbers, gaps = parts
            # Numbers set apart by spaces and brackets alone, a phone number,
            # are read left to right in any language.
            if len(numbers) < 2 or not any(gap.strip(" ()[]") for gap in gaps):
                continue
            # Arabic-Indic digits are drawn from the right by themselves.
            if any(unicodedata.bidirectional(ch) == "AN" for number in numbers for ch in number):
                continue
            name = shape["name"]
            where = f"slide_indices=[{slide_index}] shape_name='{name}' "
            set_rtl = (index <= len(directions) and directions[index - 1] == "rtl")
            then = "" if set_rtl else (
                f", then set the line right to left "
                f"({_direction_call(paragraphs, [index])}), which PowerPoint for "
                "Windows needs to draw it from the right")
            ends = _as_range(line, numbers, gaps)
            if ends:
                example = _RANGE_WORDS[script].format(a=ends[0], b=ends[1])
                how = (f"replace it with \"{example}\" (ppt_find_replace_text {where}"
                       f"find_text=\"{line}\" replace_text=\"{example}\"){then}")
            else:
                how = (f"start it with a {script} word, the label it belongs to "
                       f"(ppt_find_replace_text {where}find_text=\"{line}\"){then}")
            findings.append(_finding(
                slide_index, "number_order", [name],
                f"'{name}' has a line of numbers only, \"{line}\", on a slide "
                f"that reads right to left. With no letter in it, PowerPoint "
                f"draws the line left to right, so \"{numbers[0]}\" is on the "
                f"left and a {script} reader reads \"{numbers[-1]}\" first.",
                f"Put a word in the line: {how}.",
                line=line,
            ))
    return findings


def _drawn_order(line, right_to_left, digits_as_letters=False):
    """The characters of one paragraph in the order they are drawn, left to
    right, in a paragraph of the given direction.

    The Unicode bidirectional algorithm for a line with no embeddings: weak
    types resolved (W1-W7), neutrals (N1-N2), levels (I1-I2), then each run
    from the highest level down reversed (L2). Enough to tell whether a
    paragraph draws the same both ways, which is all it is asked.

    PowerPoint for Mac follows it. PowerPoint for Windows was seen drawing
    digits as if they were Latin letters: "1516 עד 1917", set left to right,
    as 1516, the word, 1917, where the algorithm puts both numbers on the
    left. ``digits_as_letters`` draws them that way.
    """
    base = "R" if right_to_left else "L"
    kinds = []
    for ch in line:
        kind = unicodedata.bidirectional(ch) or "L"
        if digits_as_letters and kind == "EN":
            kind = "L"
        if kind in ("LRE", "RLE", "LRO", "RLO", "PDF", "LRI", "RLI", "FSI", "PDI", "BN"):
            kind = "ON"
        kinds.append(kind)
    # W1: a combining mark takes the type of what it marks.
    for i, kind in enumerate(kinds):
        if kind == "NSM":
            kinds[i] = kinds[i - 1] if i else base
    # W2, W3: a European number after Arabic letters is an Arabic number.
    strong = base
    for i, kind in enumerate(kinds):
        if kind in ("L", "R", "AL"):
            strong = kind
        elif kind == "EN" and strong == "AL":
            kinds[i] = "AN"
    kinds = ["R" if kind == "AL" else kind for kind in kinds]
    # W4: one separator between two numbers of a kind joins them.
    for i in range(1, len(kinds) - 1):
        before, after = kinds[i - 1], kinds[i + 1]
        if kinds[i] == "ES" and before == after == "EN":
            kinds[i] = "EN"
        elif kinds[i] == "CS" and before == after and before in ("EN", "AN"):
            kinds[i] = "EN" if before == "EN" else "AN"
    # W5: currency and percent signs beside a European number belong to it.
    for i, kind in enumerate(kinds):
        if kind != "ET":
            continue
        j = i
        while j < len(kinds) and kinds[j] == "ET":
            j += 1
        k = i
        while k > 0 and kinds[k - 1] == "ET":
            k -= 1
        if (j < len(kinds) and kinds[j] == "EN") or (k > 0 and kinds[k - 1] == "EN"):
            for m in range(k, j):
                kinds[m] = "EN"
    # W6, W7.
    kinds = ["ON" if kind in ("ES", "ET", "CS") else kind for kind in kinds]
    strong = base
    for i, kind in enumerate(kinds):
        if kind in ("L", "R"):
            strong = kind
        elif kind == "EN" and strong == "L":
            kinds[i] = "L"
    # N1, N2: neutrals between two of a direction take it, numbers counting
    # as right to left; the rest take the paragraph's.
    def direction(kind):
        return "L" if kind == "L" else "R" if kind in ("R", "EN", "AN") else None
    i = 0
    while i < len(kinds):
        if direction(kinds[i]) is not None:
            i += 1
            continue
        j = i
        while j < len(kinds) and direction(kinds[j]) is None:
            j += 1
        before = direction(kinds[i - 1]) if i else base
        after = direction(kinds[j]) if j < len(kinds) else base
        for m in range(i, j):
            kinds[m] = before if before == after else base
        i = j
    # I1, I2.
    level = 1 if right_to_left else 0
    levels = []
    for kind in kinds:
        if level == 0:
            levels.append(1 if kind == "R" else 2 if kind in ("EN", "AN") else 0)
        else:
            levels.append(2 if kind in ("L", "EN", "AN") else 1)
    # L2.
    order = list(range(len(line)))
    for top in range(max(levels, default=0), 0, -1):
        i = 0
        while i < len(order):
            if levels[order[i]] < top:
                i += 1
                continue
            j = i
            while j < len(order) and levels[order[j]] >= top:
                j += 1
            order[i:j] = order[i:j][::-1]
            i = j
    return "".join(line[i] for i in order)


def _hebrew_and_arabic(line):
    """Whether a line has both Hebrew and Arabic letters in it.

    By the bidirectional algorithm such a line draws the same in either
    direction, both scripts reading right to left. PowerPoint for Mac does not
    draw it so: it splits the line into a Hebrew run and an Arabic run, and in
    a paragraph set left to right the spaces and the dot between them,
    "תודה · شكرا", land against the Arabic word on the far side.
    """
    hebrew = arabic = False
    for ch in line:
        if unicodedata.bidirectional(ch) not in _RIGHT_TO_LEFT:
            continue
        if "\u0590" <= ch <= "\u05ff" or "\ufb1d" <= ch <= "\ufb4f":
            hebrew = True
        else:
            arabic = True
    return hebrew and arabic


def _text_direction(slide_index, shapes):
    """Hebrew or Arabic paragraphs left set left to right.

    PowerPoint for Windows draws a paragraph in the direction it is set,
    whatever its letters, and a paragraph is left to right until it is set
    otherwise: the text keeps to the left of the paragraph, a full stop or a
    question mark ends up on the left of the line, and a number in it is
    read in the wrong place, "1516 עד 1917" as 1917 to 1516. PowerPoint for
    Mac draws by the first letter instead, so the slide looks right there and
    reads wrong once it is opened on Windows. A paragraph that draws the same
    either way, words alone, is not reported.
    """
    findings = []
    for shape in shapes:
        text = shape.get("text") or {}
        directions = text.get("directions")
        content = text.get("content") or ""
        if not directions or not content or not shape.get("visible", True):
            continue
        paragraphs = _paragraphs(content)
        wrong, mixed = [], []
        for index, (paragraph, direction) in enumerate(zip(paragraphs, directions), start=1):
            # The Mac answers None for a paragraph not set right to left, and
            # draws it by its first letter, which reads right but for one case:
            # Hebrew and Arabic in one line, drawn there as on Windows.
            if direction == "rtl" or _script(paragraph) not in _RANGE_WORDS:
                continue
            if direction is None and not _hebrew_and_arabic(paragraph):
                continue
            first = next((ch for ch in paragraph
                          if unicodedata.bidirectional(ch) in _LEFT_TO_RIGHT | _RIGHT_TO_LEFT), "")
            if unicodedata.bidirectional(first) not in _RIGHT_TO_LEFT:
                continue
            wrong.append(index)
            line = paragraph.strip()
            if any(_drawn_order(line, False, digits) != _drawn_order(line, True, digits)
                   for digits in (False, True)) or _hebrew_and_arabic(line):
                mixed.append(line)
        if not mixed:
            continue
        name = shape["name"]
        sample = mixed[0] if len(mixed[0]) <= 40 else mixed[0][:40] + "…"
        lines = "a paragraph" if len(wrong) == 1 else f"{len(wrong)} paragraphs"
        findings.append(_finding(
            slide_index, "text_direction", [name],
            f"'{name}' has {lines} of {_script(content)} set left to right and not "
            f"tagged as {_script(content)}, such as \"{sample}\". PowerPoint for "
            "Windows draws them left to right: the punctuation and numbers in them "
            "land on the wrong side. PowerPoint for Mac goes by the first letter "
            "instead, so the slide reads right there and wrong on Windows.",
            f"Set them right to left ({_direction_call(paragraphs, wrong)}).",
            paragraphs=wrong,
        ))
    return findings


# Latin words, run together by spaces or hyphens: "F1", "PowerPoint",
# "Claude Code". Then what follows them up to the next letter.
_LATIN_RUN = re.compile(r"[A-Za-z][A-Za-z0-9]*(?:[ \-][A-Za-z0-9]+)*")
_AFTER_LATIN = re.compile(r"[^\sA-Za-z0-9\u0590-\u08FF\uFB1D-\uFEFF\u200f]*\s?")
_RLM = "\u200f"


def _mixed_direction(slide_index, shapes):
    """Latin words in a right-to-left paragraph, with a space or a mark after.

    PowerPoint splits the text into runs by script as it is written, and the
    space or punctuation after a Latin word goes into the Latin run.
    In a paragraph drawn right to left that run is drawn as one block, so
    "בדור F1 ההורים" comes out with two spaces on one side of F1 and none
    on the other, and the colon of "F1: שי טצו" is drawn before F1. A
    right-to-left mark (U+200F) after the word, and after its punctuation,
    puts them in the Hebrew run, where they are drawn in their place. Seen
    the same on Windows and on the Mac.
    """
    findings = []
    for shape in shapes:
        text = shape.get("text") or {}
        content = text.get("content") or ""
        if not content or not shape.get("visible", True):
            continue
        directions = text.get("directions") or []
        paragraphs = _paragraphs(content)
        hits = []
        for index, paragraph in enumerate(paragraphs, start=1):
            if _script(paragraph) not in _RANGE_WORDS:
                continue
            first = next((ch for ch in paragraph
                          if unicodedata.bidirectional(ch) in _LEFT_TO_RIGHT | _RIGHT_TO_LEFT), "")
            set_rtl = index <= len(directions) and directions[index - 1] == "rtl"
            if not set_rtl and unicodedata.bidirectional(first) not in _RIGHT_TO_LEFT:
                continue
            for match in _LATIN_RUN.finditer(paragraph):
                after = _AFTER_LATIN.match(paragraph, match.end()).group(0)
                if not after:
                    continue
                marks = after.rstrip()
                rest = paragraph[match.end() + len(after):]
                if not marks and not rest:
                    continue
                hits.append((match.group(0), after, marks))
        if not hits:
            continue
        name = shape["name"]
        word, after, marks = hits[0]
        find = word + after
        fixed = word + _RLM + (marks + _RLM if marks else "") + after[len(marks):]
        shown = fixed.replace(_RLM, "\\u200f")
        more = "" if len(hits) == 1 else f", and {len(hits) - 1} more like it"
        findings.append(_finding(
            slide_index, "mixed_direction", [name],
            f"'{name}' has Latin text in a right-to-left paragraph, \"{find.strip()}\"{more}. "
            "PowerPoint puts the space or punctuation after it in the Latin run and "
            "draws it on the wrong side of the word.",
            f"Put a right-to-left mark (U+200F) after the word"
            f"{' and after its punctuation' if marks else ''}: ppt_find_replace_text "
            f"slide_indices=[{slide_index}] shape_name='{name}' find_text=\"{find}\" "
            f"replace_text=\"{shown}\". The mark is invisible.",
            severity="warning", examples=[h[0] for h in hits],
        ))
    return findings


_PICTURES = frozenset({msoPicture, msoLinkedPicture})


def _stretched(slide_index, shapes, slide_area=None):
    """A picture drawn in other proportions than its image's.

    A photo or a logo given both a width and a height that do not match it is
    squeezed or stretched, and faces and lettering show it at once. The fix
    keeps the side that was meant and fits the other to it, about the same
    centre, so the picture stays where it was put.
    """
    findings = []
    for shape in shapes:
        natural = shape.get("natural_aspect")
        if (not natural or shape.get("type") not in _PICTURES
                or not shape.get("visible", True) or shape["height"] <= 0):
            continue
        drawn = shape["width"] / shape["height"]
        ratio = drawn / natural
        if abs(ratio - 1) <= _STRETCH:
            continue
        name = shape["name"]
        if ratio > 1:
            width = shape["height"] * natural
            left = shape["left"] + (shape["width"] - width) / 2
            how = "wider"
            fix = (f"Make it {_pt(shape['width'] - width)} narrower to fit its image, "
                   f"keeping its height and centre (ppt_update_shape width={_num(width)} "
                   f"left={_num(left)})")
        else:
            height = shape["width"] / natural
            top = shape["top"] + (shape["height"] - height) / 2
            how = "taller"
            fix = (f"Make it {_pt(shape['height'] - height)} shorter to fit its image, "
                   f"keeping its width and centre (ppt_update_shape height={_num(height)} "
                   f"top={_num(top)})")
        if shape.get("lock_aspect"):
            # A locked picture keeps whatever proportions it has, the wrong
            # ones too: given a new width, its height follows.
            fix = (f"Unlock its proportions first (ppt_lock_aspect_ratio "
                   f"locked=false), then {fix[0].lower()}{fix[1:]}")
        crop = (f"Crop the image to the proportions it is drawn at, keeping its size "
                f"(ppt_crop_picture shape_name_or_index='{name}' crop_fit='shape')")
        stretch = round(abs(ratio - 1) * 100) if ratio > 1 else round(abs(1 / ratio - 1) * 100)
        # A photo was given its box on purpose, and cropping keeps the box and
        # the slide around it; a logo cropped loses part of its mark.
        small = slide_area and shape["width"] * shape["height"] < _LOGO_SHARE * slide_area
        if small:
            fix = f"{fix}, or {crop[0].lower()}{crop[1:]}"
        else:
            fix = f"{crop}, or {fix[0].lower()}{fix[1:]}"
        findings.append(_finding(
            slide_index, "stretched_picture", [name],
            f"'{name}' is stretched: it is drawn {_pt(shape['width'])} by "
            f"{_pt(shape['height'])}, {stretch}% {how} for its size than its "
            f"image, so what is in it looks squeezed.",
            f"{fix}.",
            stretch_pct=stretch,
        ))
    return findings


# Text smaller than this is hard to read when the slide is projected. A
# designed deck sets its labels and source lines in 10 or 11 pt; below that is
# text squeezed to fit.
_SMALLEST_TEXT = 10.0

# What PowerPoint puts in a slide's corners itself, set small on purpose.
_SMALL_BY_DESIGN = frozenset({ppPlaceholderDate, ppPlaceholderFooter, ppPlaceholderSlideNumber})


def _small_text(slide_index, shapes):
    """Text set too small to read when the slide is projected.

    Judged by the smallest size in the shape: a line in 9 pt under a heading
    in 24 is the line the back of the room cannot read. The date, footer and
    slide number are left alone; they are small by design.
    """
    findings = []
    for shape in shapes:
        text = shape.get("text") or {}
        if (not shape.get("visible", True) or not text.get("chars")
                or shape.get("placeholder_type") in _SMALL_BY_DESIGN):
            continue
        sizes = [run["size"] for run in text.get("runs") or [] if run.get("size")]
        # Text set to shrink to fit is judged by the saved file: how far
        # PowerPoint shrank it, and the size it is set in when the file says.
        # Read live, both platforms give the size already shrunk (18 pt at
        # 70% reads as 13), and PowerPoint for Mac cannot say it shrinks.
        shrunk = shape.get("shrunk") or {}
        scale = shrunk.get("scale") or 1.0
        set_in = shrunk.get("smallest")
        if set_in is None and sizes:
            set_in = min(sizes) / scale
        if set_in is None or set_in * scale >= _SMALLEST_TEXT - 0.05:
            continue
        name, smallest = shape["name"], round(set_in * scale, 1)
        if scale < 1.0:
            drawn = (f"set in {_pt(round(set_in, 1))} and shrunk by PowerPoint to "
                     f"{_pt(smallest)} to fit its box")
            fix = (f"Give it the room its words need at {_num(_SMALLEST_TEXT)} pt or more: a "
                   f"bigger box (ppt_update_shape), fewer words, or both")
        else:
            drawn = f"set in {_pt(smallest)}"
            fix = (f"Set it to at least {_num(_SMALLEST_TEXT)} pt (ppt_format_text "
                   f"shape_name_or_index='{name}' font_size={_num(_SMALLEST_TEXT)})")
            if text.get("autofit") == "shrink_to_fit":
                fix += (". PowerPoint shrinks it to fit its box, so give it the room as well "
                        "(ppt_update_shape), or say less")
        findings.append(_finding(
            slide_index, "small_text", [name],
            f"Text in '{name}' is {drawn}, too small to read when the slide is "
            f"projected; {_pt(_SMALLEST_TEXT)} is about the smallest that is.",
            f"{fix}.", severity="warning", size=smallest,
        ))
    return findings


# A picture drawn larger than one of its image's pixels to a point is spread
# thin, and soft on a projector.
_PIXELS_PER_POINT = 1.0


def _blurry(slide_index, shapes):
    """A picture drawn larger than its image has pixels for."""
    findings = []
    for shape in shapes:
        pixels = shape.get("pixels")
        if (not pixels or shape.get("type") not in _PICTURES or not shape.get("visible", True)
                or shape["width"] <= 0 or shape["height"] <= 0):
            continue
        across, down = pixels
        density = min(across / shape["width"], down / shape["height"])
        if density >= _PIXELS_PER_POINT:
            continue
        name = shape["name"]
        width = shape["width"] * density / _PIXELS_PER_POINT
        height = shape["height"] * density / _PIXELS_PER_POINT
        left = shape["left"] + (shape["width"] - width) / 2
        top = shape["top"] + (shape["height"] - height) / 2
        findings.append(_finding(
            slide_index, "blurry_picture", [name],
            f"'{name}' is drawn {_pt(shape['width'])} by {_pt(shape['height'])} from an "
            f"image {round(across)} by {round(down)} pixels, so it looks soft when the "
            f"slide is projected.",
            f"Use a larger image, at least {math.ceil(shape['width'] * _PIXELS_PER_POINT)} "
            f"pixels across, or draw it at most {_pt(width)} by {_pt(height)}, keeping "
            f"its centre (ppt_update_shape width={_num(width)} height={_num(height)} "
            f"left={_num(left)} top={_num(top)}).",
            severity="warning", pixels_per_point=round(density, 2),
        ))
    return findings


def _missing_alt_text(slide_index, shapes):
    """A picture, chart or SmartArt with no alt text for a screen reader.

    Known only from the saved file; one whose file says nothing is left
    alone. A picture marked as decoration needs none, and alt text that is
    only the image's file name says nothing.
    """
    findings = []
    for shape in shapes:
        if shape.get("described") is not False or not shape.get("visible", True):
            continue
        name = shape["name"]
        if IS_MACOS:
            fix = (f"Describe what it shows in PowerPoint: right-click '{name}', View Alt "
                   "Text. PowerPoint for Mac does not let a tool set it; on Windows, "
                   f"ppt_update_shape shape_name='{name}' alt_text='...' does.")
        else:
            fix = (f"Describe what it shows, in a sentence: ppt_update_shape "
                   f"shape_name='{name}' alt_text='...'.")
        findings.append(_finding(
            slide_index, "missing_alt_text", [name],
            f"'{name}' has no alt text, so a screen reader cannot say what it shows.",
            fix, severity="warning",
        ))
    return findings


# Chart parts sharing less than this, either way, only touch.
_CHART_TOUCH = 2.0


def _chart_overlaps(slide_index, shapes):
    """Parts of a chart drawn over each other: data labels, the legend, the
    title and the plot.

    PowerPoint lays a chart out itself, and with small slices or many points
    its labels land on each other; a legend or title moved by hand can sit
    on the bars. Read on Windows only: PowerPoint for Mac has no chart
    object to ask.
    """
    def meet(a, b):
        width, height = _shared(a, b)
        return width > _CHART_TOUCH and height > _CHART_TOUCH

    findings = []
    for shape in shapes:
        parts = shape.get("chart_parts")
        if not parts or not shape.get("visible", True):
            continue
        name, problems, fixes = shape["name"], [], []
        labels = parts.get("labels") or []
        crowded = {i for i, a in enumerate(labels) for j, b in enumerate(labels)
                   if i < j and meet(a, b)} | {j for i, a in enumerate(labels)
                                               for j, b in enumerate(labels)
                                               if i < j and meet(a, b)}
        if crowded:
            problems.append(f"{len(crowded)} of its {len(labels)} data labels run into each other")
            fixes.append("make the chart larger, show labels on fewer series "
                         "(ppt_set_chart_series show_data_labels=false), or put the small "
                         "values together")
        legend, plot, title = parts.get("legend"), parts.get("plot"), parts.get("title")
        if legend and plot and meet(legend, plot):
            problems.append("its legend lies on the plot")
            fixes.append(f"put the legend beside the plot (ppt_format_chart "
                         f"shape_name_or_index='{name}' legend_position='bottom')")
        if title and (plot and meet(title, plot) or legend and meet(title, legend)):
            problems.append("its title lies on the plot or the legend")
            fixes.append(f"put the title back above the plot (ppt_format_chart "
                         f"shape_name_or_index='{name}' title_position='top')")
        if not problems:
            continue
        findings.append(_finding(
            slide_index, "chart_overlap", [name],
            f"In '{name}', {'; '.join(problems)}.",
            f"{'; '.join(fixes)[0].upper()}{'; '.join(fixes)[1:]}.",
            severity="warning",
        ))
    return findings


# The names PowerPoint gives a new chart's categories and series.
_SAMPLE_CATEGORY = re.compile(r"^Category \d+$")
# The pie's sample quarters are also what a real quarterly chart is called, so
# they count only beside sample series or numbers.
_SAMPLE_QUARTER = re.compile(r"^[1-4](?:st|nd|rd|th) Qtr$")
_SAMPLE_SERIES = re.compile(r"^(?:Series \d+|Y Values)$")


def _sample_values():
    from gvml.charts import (
        DEFAULT_BUBBLE_SERIES, DEFAULT_PIE_SERIES, DEFAULT_SERIES, DEFAULT_XY_SERIES,
    )
    return [tuple(s["values"]) for s in
            DEFAULT_SERIES + DEFAULT_PIE_SERIES + DEFAULT_XY_SERIES + DEFAULT_BUBBLE_SERIES]


def _chart_sample_data(slide_index, shapes):
    """A chart still showing the data PowerPoint put in it.

    A new chart comes with four categories named "Category 1" to 4, series
    named "Series 1" to 3, and made-up numbers. A model fills in the series
    and forgets the names, or adds a chart and moves on; either way the
    chart reads as a placeholder.
    """
    findings = []
    samples = None
    for shape in shapes:
        chart = shape.get("chart")
        if not chart or not shape.get("visible", True):
            continue
        if samples is None:
            samples = _sample_values()
        left = []
        series = chart.get("series") or []
        names = [s.get("name") for s in series if _SAMPLE_SERIES.match(s.get("name") or "")]
        numbers = [s.get("name") for s in series
                   if tuple(round(v, 4) for v in s.get("values") or [] if v is not None) in samples]
        categories = [c for c in chart.get("categories") or [] if _SAMPLE_CATEGORY.match(c or "")
                      or ((names or numbers) and _SAMPLE_QUARTER.match(c or ""))]
        if categories:
            left.append("categories named " + ", ".join(f"'{c}'" for c in categories[:4]))
        if names:
            left.append("series named " + ", ".join(f"'{n}'" for n in names[:3]))
        if numbers:
            left.append("PowerPoint's sample numbers")
        if not left:
            continue
        name = shape["name"]
        findings.append(_finding(
            slide_index, "chart_sample_data", [name],
            f"The chart '{name}' still has the data PowerPoint puts in a new "
            f"chart: {'; '.join(left)}.",
            f"Put the real data in it (ppt_set_chart_data on '{name}', with "
            "the categories and every series named), or delete the chart.",
        ))
    return findings


def _fonts(slide_index, shapes, theme_fonts=None):
    """Text set in a font this computer does not have, or drawn in another.

    PowerPoint draws a font it cannot find in a stand-in, and says nothing:
    the deck looks right where it was made and different here. Hebrew and
    Arabic are drawn in the run's right-to-left font, which setting a font
    by name left as it was, so a title asked for in David showed in the
    theme's font. That one is known only on Windows, where it can be read.
    """
    findings = []
    usual = {}
    for shape in shapes:
        text = shape.get("text") or {}
        missing = set(text.get("missing_fonts") or [])
        for run in text.get("runs") or []:
            if run.get("font") and run["font"] not in missing and not run["font"].startswith("+"):
                usual[run["font"]] = usual.get(run["font"], 0) + 1
    for shape in shapes:
        text = shape.get("text") or {}
        if not shape.get("visible", True) or not text.get("chars"):
            continue
        name = shape["name"]
        missing = text.get("missing_fonts") or []
        if missing:
            listed = ", ".join(f"'{f}'" for f in missing)
            instead = max(usual, key=usual.get) if usual else None
            use = (f"ppt_format_text font_name='{instead}', the slide's own font"
                   if instead else "ppt_format_text font_name= one the rest of the deck uses")
            findings.append(_finding(
                slide_index, "font_missing", [name],
                f"'{name}' asks for {listed}, which this computer does not have, "
                "so PowerPoint draws it in a stand-in font here, as it will on "
                "any computer without it.",
                f"Set it in a font that is installed ({use}), or install {listed} "
                "on every computer that will show the deck.",
                severity="warning", fonts=missing,
            ))
            continue
        fonts = {run.get("font") for run in text.get("runs") or [] if run.get("font")}
        drawn = text.get("script_font")
        themed = theme_fonts or {}
        asked = next(iter(fonts)) if len(fonts) == 1 else None
        # The theme's Latin font was not asked for by name, so a right-to-left
        # font beside it is no sign of one set only halfway. A right-to-left
        # font other than the theme's was chosen; the one left behind when a
        # font is set by name is the theme's, or with a theme that names none,
        # PowerPoint's default, which cannot be told from a choice.
        by_name = asked and not asked.startswith("+") and asked not in set(themed.get("latin") or ())
        left_behind = not themed.get("script") or drawn in set(themed["script"])
        if (by_name and left_behind and drawn and not drawn.startswith("+")
                and drawn.lower() != asked.lower()):
            script = _script(text.get("content") or "") or "right-to-left"
            findings.append(_finding(
                slide_index, "font_missing", [name],
                f"The {script} in '{name}' is drawn in '{drawn}', not '{asked}': "
                f"PowerPoint draws right-to-left letters in a font of their own, "
                f"and only the Latin one was set to '{asked}'.",
                (f"Set the font again (ppt_format_text font_name='{asked}'), which "
                 "sets it for right-to-left letters too."
                 if text.get("script_font_settable", True) else
                 f"Set the right-to-left font to '{asked}' in PowerPoint's Font dialog: "
                 "PowerPoint for Mac does not let it be set through its scripting, and "
                 "PowerPoint for Windows sets it with ppt_format_text font_name."),
                severity="warning", fonts=[asked],
            ))
    return findings


def _empty_placeholders(slide_index, shapes):
    findings = []
    for shape in shapes:
        if not _is_empty_placeholder(shape):
            continue
        if shape.get("placeholder_type") in _LAYOUT_PLACEHOLDERS:
            continue
        name = shape["name"]
        findings.append(_finding(
            slide_index, "empty_placeholder", [name],
            f"'{name}' is an empty placeholder. It draws nothing in the slide "
            "show, but in the editor it shows its prompt text, and it still "
            "takes the space the layout gave it.",
            f"Put text in '{name}', or delete it with ppt_delete_shape.",
            severity="warning",
        ))
    return findings


def slide_findings(slide, slide_w, slide_h, only=None, trust_template=True):
    """Every finding for one slide, worst first.

    ``slide`` is one entry of what ``_slide_facts_impl`` returns. With
    ``only``, a set of shape names, the answer is the per shape findings that
    involve one of those shapes, which is what a write tool reports. Without
    ``trust_template``, placeholder text in theme colours is judged like any
    other: see ``_as_designed``.
    """
    index = slide["slide_index"]
    shapes = _animated(slide, _settled(slide["shapes"]))
    template = slide.get("template") or []
    layout = _layout(shapes, slide_w, slide_h)
    findings = (
        _off_slide(index, shapes, slide_w, slide_h, layout)
        + _overflow(index, shapes, layout)
        + _overlap(index, shapes, layout)
        + _template_collisions(index, shapes, template, layout, slide.get("template_area"))
        + _template_rules(index, shapes, template, layout)
        + _low_contrast(index, shapes, slide.get("background_hex"), template,
                        slide.get("follows_background"), layout, trust_template,
                        slide.get("background_gradient"), slide.get("theme_colors"))
        + _cramped_text(index, layout)
        + _narrow_margins(index, layout, slide_w)
        + _vertical_margins(index, layout, slide_h)
        + _misaligned(index, layout, only=only)
        + _title_body_edge(index, shapes, layout)
        + _off_centre_labels(index, layout)
        + _uneven_sizes(index, layout)
        + _sample_text(index, shapes)
        + _chart_sample_data(index, shapes)
        + _fonts(index, shapes, slide.get("theme_fonts"))
        + _stretched(index, shapes, slide_w * slide_h)
        + _blurry(index, shapes)
        + _small_text(index, shapes)
        + _missing_alt_text(index, shapes)
        + _chart_overlaps(index, shapes)
        + _number_order(index, shapes)
        + _text_direction(index, shapes)
        + _mixed_direction(index, shapes)
    )
    findings += _inside_groups(index, slide, slide_w, slide_h, trust_template)
    if only is not None:
        # Text on a touched card is drawn on its fill, so whether it can still
        # be read is that write's to report as well: a card filled dark under
        # dark words.
        drawn_on = set()
        for shape in shapes:
            _, source, painter = _backdrop(shape, shapes, slide.get("background_hex"), template)
            if source == "shape" and painter["name"] in only:
                drawn_on.add(shape["name"])
        # Two like shapes of two sizes may be a row or a chart a model is still
        # drawing, the third bar not yet added; the slide's own check judges
        # the finished set.
        return _on_hidden(slide, _sorted([
            f for f in findings
            if f["type"] in SHAPE_FINDINGS and (
                any(name in only for name in f["shapes"])
                or (f["type"] == "low_contrast" and f["shapes"][0] in drawn_on))
            and not (f["type"] == "uneven_sizes" and len(f["shapes"]) == 2)
        ]))
    spacing = _uneven_spacing(index, layout, slide_w)
    # A row whose gaps are reported is put back between the margins by that
    # finding's fix, so its cards are not told about the margin as well.
    spaced = {name for f in spacing for name in f["shapes"]}
    findings = [f for f in findings
                if not (f["type"] == "narrow_margin" and f["shapes"][0] in spaced)]
    findings += spacing + _empty_placeholders(index, shapes)
    return _on_hidden(slide, _sorted(findings))


def _animated(slide, shapes):
    """The shapes, each with when it is on the screen and whether it moves.

    See ``on_screen``. A slide with no animations and no Morph next to it
    keeps its shapes as they are.
    """
    animations = slide.get("animations") or []
    shown, moves = on_screen(animations), moving(animations)
    if not (shown or moves or slide.get("morph")):
        return shapes
    return [{**shape, "shown": shown.get(shape["name"]),
             "moves": shape["name"] in moves, "morph": bool(slide.get("morph"))}
            for shape in shapes]


def _on_hidden(slide, findings):
    """Say so in each finding on a hidden slide.

    It is still part of the deck, and shown once unhidden, so it is checked
    like any other; the reader decides whether it matters now.
    """
    if not slide.get("hidden"):
        return findings
    for finding in findings:
        finding["hidden_slide"] = True
        finding["message"] += f" Slide {slide['slide_index']} is hidden in the slide show."
    return findings


def _inside_groups(index, slide, slide_w, slide_h, trust_template=True):
    """The checks on a shape's own text, for the shapes inside groups.

    A group is laid out as one box; what is inside it is placed by whoever
    drew it. Its text is still text: it can spill out of its box, be left as
    stand-in words, read backwards or be set in a font that is not there. Its
    colour is judged against what is under it inside the group, or the slide.
    """
    inner = _settled(slide.get("inner") or [])
    if not inner:
        return []
    layout = _layout(inner, slide_w, slide_h)
    # What is under a group shows through it, so the slide's own shapes are
    # backdrops too; only the findings on what is inside are this check's.
    names = {s["name"] for s in inner}
    around = [s for s in _settled(slide["shapes"]) if s.get("type") not in (msoGroup, msoSmartArt)]
    contrast = [f for f in _low_contrast(
                    index, around + inner, slide.get("background_hex"), slide.get("template") or (),
                    slide.get("follows_background"), _layout(around + inner, slide_w, slide_h),
                    trust_template, slide.get("background_gradient"), slide.get("theme_colors"))
                if f["shapes"][0] in names]
    findings = (_overflow(index, inner, layout)
                + contrast
                + _sample_text(index, inner)
                + _number_order(index, inner)
                + _text_direction(index, inner)
                + _mixed_direction(index, inner)
                + _fonts(index, inner, slide.get("theme_fonts"))
                + _small_text(index, inner))
    unreachable = {s["name"]: s for s in inner if s.get("reachable") is False}
    for finding in findings:
        member = unreachable.get(finding["shapes"][0])
        if member is None:
            continue
        groups, own = member.get("groups") or [member["in_group"]], member["name"].rsplit("/", 1)[-1]
        # A group in a group is still one after the outer is ungrouped.
        calls = ", then ".join(f"ppt_ungroup_shapes shape_name_or_index='{g}'" for g in groups)
        named = " and then ".join(f"'{g}'" for g in groups)
        finding["fix"] = (f"Ungroup {named} first ({calls}): PowerPoint for Mac does not "
                          f"reach inside a group. Then, with '{own}' as the shape: "
                          + finding["fix"].replace(member["name"], own))
    return findings


def _settled(shapes):
    """The shapes as PowerPoint leaves them once it has laid out their text.

    A box set to fit its text grows to it, but not always at once: after a
    change to the fonts of the whole deck, PowerPoint resizes it when it next
    draws the slide, and a read before that finds the text taller than the
    box. Such a box is taken at the height it is about to have, from its top
    down, so it is not reported as overflowing, and what it grows into is.
    """
    settled = []
    for shape in shapes:
        text = shape.get("text") or {}
        margins = text.get("margins")
        needed = text.get("height")
        if (text.get("autofit") == "shape_to_fit" and margins and needed is not None
                and text.get("orientation") in (None, "horizontal")):
            height = round(needed + margins["top"] + margins["bottom"], 2)
            if height > shape["height"] + _EDGE_TOLERANCE:
                shape = {**shape, "height": height}
        settled.append(shape)
    return settled


def _sorted(findings):
    rank = {kind: i for i, kind in enumerate(_ORDER)}
    return sorted(findings, key=lambda f: (f["slide_index"] or 0, rank.get(f["type"], len(rank))))


# ---------------------------------------------------------------------------
# Deck findings
# ---------------------------------------------------------------------------
def _largest_size(shape):
    sizes = [run.get("size") for run in (shape.get("text") or {}).get("runs") or []
             if run.get("size")]
    return max(sizes) if sizes else None


def find_title(slide, slide_h):
    """The shape that is this slide's title, or None.

    A title placeholder with text when there is one. Generated decks often
    have no placeholders at all, so otherwise it is the largest text in the
    top quarter of the slide, at a title's size. A cover slide's title sits
    lower than that and is left out, which is right: it is meant to differ.
    """
    for shape in slide["shapes"]:
        text = shape.get("text")
        if shape.get("placeholder_type") in _TITLE_PLACEHOLDERS and text and text.get("chars"):
            # The title slide's own title, and a section header's, set lower,
            # are meant to differ from the rest. Where the text is drawn says
            # how low: a title placeholder's frame is often far taller than
            # the line in it.
            words = text_box(shape) or shape_box(shape)
            if shape.get("placeholder_type") == ppPlaceholderCenterTitle or (
                slide_h and words[1] > _TITLE_ZONE * slide_h
            ):
                return None
            return shape
    if not slide_h:
        return None
    best = None
    for shape in slide["shapes"]:
        text = shape.get("text")
        if not shape.get("visible", True) or not text or not text.get("chars"):
            continue
        if shape["top"] > _TITLE_ZONE * slide_h:
            continue
        size = _largest_size(shape)
        if size is None or size < _MIN_TITLE_SIZE:
            continue
        # The topmost: a big figure further down is content, not the title.
        if best is None or (-shape["top"], size) > (-best["top"], _largest_size(best)):
            best = shape
    return best


def _clusters(values, tolerance):
    """Group (slide, value) pairs whose values are within tolerance of the
    group's first value, largest group first."""
    groups = []
    for slide_index, value in sorted(values, key=lambda pair: pair[1]):
        for group in groups:
            if abs(group[0][1] - value) <= tolerance:
                group.append((slide_index, value))
                break
        else:
            groups.append([(slide_index, value)])
    return sorted(groups, key=len, reverse=True)


def _slides(numbers):
    numbers = sorted(numbers)
    return ("slide " if len(numbers) == 1 else "slides ") + ", ".join(str(n) for n in numbers)


def _bookends(slides):
    """Slides dressed like the cover, which are not held to the content titles.

    A closing slide or a section divider often wears the cover's backdrop,
    one most slides do not have, with its words set large near the top: a
    "Thank you" in 44 pt was reported as a title out of size among 32 pt
    ones. Slides with the first slide's background, when that is not the
    deck's usual one, are left out of the comparison like the cover itself.
    """
    backgrounds = [slide.get("background_hex") for slide in slides]
    if len(slides) < 3 or not backgrounds[0]:
        return set()
    usual = max(set(backgrounds), key=backgrounds.count)
    if backgrounds[0] == usual:
        return set()
    return {slide["slide_index"] for slide in slides[1:]
            if slide.get("background_hex") == backgrounds[0]}


def _title_findings(slides, slide_h, slide_w=None):
    """Titles that differ between slides.

    Flipping from one slide to the next, a title that is a size larger, a few
    points lower, or set in another weight is the thing that jumps, whatever
    layout each slide is on: a slide on the layout's own title placeholder
    among slides with a title box of their own jumps just the same. Only
    titles at the top of a slide are compared, so a cover's or a section
    header's title, set lower and meant to differ, is not held to the rest.
    Colours are compared only between titles on the same backdrop, since a
    light title on a dark slide is meant.
    """
    by_index = {slide["slide_index"]: slide for slide in slides}
    bookends = _bookends(slides)
    titled = []
    for slide in slides:
        if slide["slide_index"] in bookends:
            continue
        title = find_title(slide, slide_h)
        if title is not None:
            titled.append((slide["slide_index"], title))
    if len(titled) < _MIN_TITLED_SLIDES:
        return []

    # A title is placed by the side its text keeps to. A right-aligned title,
    # a Hebrew or Arabic deck's, keeps its right edge from slide to slide while
    # its left edge moves with the length of the words; a centred one keeps its
    # middle. Text that fills its frame keeps to both sides.
    side = {i: _title_side(t) for i, t in titled}
    usual = {"top": _usual([(i, _title_edge(t, "top")) for i, t in titled])}
    for field in ("left", "right", "center"):
        values = [(i, _title_edge(t, field)) for i, t in titled if side[i] in (field, "full")]
        if values:
            usual[field] = _usual(values)

    def home(i, t):
        """The side this title keeps to, and how far its text is from the usual place there."""
        field = "left" if side[i] == "full" else side[i]
        return field, _title_edge(t, field) - usual[field]

    # A title placeholder far from where the others sit is where its layout
    # put it, in the size and colour the layout gave it for that place: a
    # caption layout's title in its side column. It is the template's design,
    # so it is held to none of the others. A title box of the caller's own is.
    if slide_w:
        titled = [(i, t) for i, t in titled
                  if t.get("placeholder_type") not in _TITLE_PLACEHOLDERS
                  or abs(home(i, t)[1]) <= _TITLE_ELSEWHERE * slide_w]
        if len(titled) < _MIN_TITLED_SLIDES:
            return []
    # A slide with a photo, or anything else that paints, where the other
    # titles sit has its title elsewhere because it has to: a different
    # layout, not a slip. Its title is not held to their position.
    elsewhere = set()
    for i, t in titled:
        words = text_box(t) or shape_box(t)
        shift = home(i, t)[1]
        spot = (words[0] - shift, usual["top"], words[2] - shift,
                usual["top"] + _height(words))
        if _title_spot_taken(t, by_index[i], spot, slide_w, slide_h):
            elsewhere.add(i)
    # Each layout of a template sets its own title: a caption layout's smaller,
    # a statement layout's larger. When most titles are the template's, the
    # template is what the deck follows, and a title it styled differs from
    # those on other layouts by design; it is held only to titles on its own
    # layout. A deck on two designs has a "Title and Content" in each, and
    # each styles its title its own way, so a layout is told by its design as
    # well as its name. A title the caller drew or recoloured is held to all
    # of them. When most titles are the caller's own, the deck follows those,
    # and every title is held to them.
    styled = {i: _template_title(t) for i, t in titled}
    layout_of = {i: (by_index[i].get("design"), by_index[i].get("layout")) for i, _ in titled}

    def held(i, usual):
        if not styled[i]:
            return True
        theirs = [j for j in usual if styled[j]]
        if 2 * len(theirs) <= len(usual):
            return True
        return any(layout_of[j] == layout_of[i] for j in theirs)

    findings = []
    for label, read, tolerance, field in (
        ("size", _largest_size, _TITLE_SIZE_TOLERANCE, None),
        ("left edge", lambda s: _title_edge(s, "left"), _NEAR_MISS, "left"),
        ("right edge", lambda s: _title_edge(s, "right"), _NEAR_MISS, "right"),
        ("centre", lambda s: _title_edge(s, "center"), _NEAR_MISS, "center"),
        ("top edge", lambda s: _title_edge(s, "top"), _NEAR_MISS, "top"),
    ):
        values = [(index, read(shape)) for index, shape in titled
                  if read(shape) is not None and not (field and index in elsewhere)
                  and (field in (None, "top") or side[index] in (field, "full"))]
        usual, rest = _usual_group(values, tolerance)
        # A value no two titles share is no usual value to hold the rest to:
        # a left-set title beside one placeholder title is two odd ones out.
        if len(usual) < 2 or not rest:
            continue
        groups = [usual] + _clusters(rest, tolerance)
        members = [j for j, _ in usual]
        groups = [usual] + [kept for kept in (
            [(i, v) for i, v in g if held(i, members)] for g in groups[1:]) if kept]
        if len(groups) < 2:
            continue
        usual_value = sorted(v for _, v in usual)[len(usual) // 2]
        described = "; ".join(
            f"{round(sorted(v for _, v in g)[len(g) // 2], 1):g} pt on "
            f"{_slides(i for i, _ in g)}"
            for g in groups
        )
        odd = [i for g in groups[1:] for i, _ in g]
        names = [shape["name"] for index, shape in titled if index in odd]
        if field is None:
            fix = (f"Set the titles on {_slides(odd)} to "
                   f"{round(usual_value, 1):g} pt (ppt_format_text "
                   f"font_size={round(usual_value, 1):g}).")
        else:
            moves = _title_moves(field, usual_value, odd, titled, by_index)
            fix = (f"Set the {label} of the titles on {_slides(odd)} to "
                   f"{round(usual_value, 1):g} pt: {moves}.")
        findings.append(_finding(
            None, "inconsistent_titles", names,
            f"Slide titles differ in {label}: {described}.",
            fix, severity="warning", slides=sorted(odd),
        ))

    findings += _title_differences(
        "alignment", [(index, title, side[index]) for index, title in titled
                      if side[index] != "full"],
        lambda value: f"to {value} alignment",
        lambda value: f"alignment='{value}'", held, tool="ppt_set_paragraph_format",
    )
    # Capitals the template sets by formatting are its style, and the words
    # under them are typed as they please, so only typed capitals are held.
    if not any((title.get("text") or {}).get("all_caps") for _, title in titled):
        findings += _title_differences(
            "capitals", [(index, title, _title_case(title)) for index, title in titled],
            lambda value: "in capitals" if value == "capitals" else "in normal case",
            lambda value: "with the title retyped " + (
                "in capitals" if value == "capitals" else "in normal case"),
            held, tool="ppt_set_text",
        )
    findings += _title_differences(
        "weight", [(index, title, _title_weight(title)) for index, title in titled],
        lambda value: f"in {value}" if value == "bold" else "in regular weight",
        lambda value: f"bold={'true' if value == 'bold' else 'false'}", held,
    )
    # A layout's title placeholder is set in the theme's heading font, and a
    # title box of one's own in the body font, so a deck mixing the two has
    # titles in two typefaces.
    findings += _title_differences(
        "font", [(index, title, _title_font(title)) for index, title in titled],
        lambda value: f"in {value}", lambda value: f"font_name='{value}'", held,
    )
    by_backdrop = {}
    for index, title in titled:
        slide = by_index[index]
        backdrop = _backdrop_colour(title, slide["shapes"], slide.get("background_hex"),
                                    slide.get("template"))
        if backdrop:
            by_backdrop.setdefault(backdrop, []).append((index, title, _title_colour(title)))
    for entries in by_backdrop.values():
        findings += _title_differences(
            "colour", entries, lambda value: f"to {value}",
            lambda value: f"color='{value}'", held,
        )
    return findings


def deck_findings(slides, slide_h, slide_w=None):
    """What differs between slides that should not: titles and body text."""
    findings = (_title_findings(slides, slide_h, slide_w) + _body_sizes(slides, slide_w, slide_h)
                + _titles_off_their_layout(slides) + _missing_titles(slides, slide_h)
                + _duplicate_titles(slides, slide_h))
    hidden = {slide["slide_index"] for slide in slides if slide.get("hidden")}
    for finding in findings:
        odd = sorted(hidden & set(finding.get("slides") or [finding.get("slide_index")]))
        if odd:
            finding["hidden_slides"] = odd
            finding["message"] += (f" {_slides(odd).capitalize()} "
                                   f"{'is' if len(odd) == 1 else 'are'} hidden in the slide show.")
    return findings


def _slide_heading(slide, slide_h):
    """The text that heads a slide, wherever it is: its title placeholder,
    the title ``find_title`` finds, or a cover's big words lower down."""
    for shape in slide["shapes"]:
        if (shape.get("placeholder_type") in _TITLE_PLACEHOLDERS
                and (shape.get("text") or {}).get("chars")):
            return shape
    title = find_title(slide, slide_h)
    if title is not None:
        return title
    big = [s for s in slide["shapes"] if s.get("visible", True)
           and (s.get("text") or {}).get("chars") and (_largest_size(s) or 0) >= _MIN_TITLE_SIZE]
    return max(big, key=_largest_size, default=None)


def _missing_titles(slides, slide_h):
    """Slides with nothing to head them.

    A screen reader, the outline and the list of slides PowerPoint shows
    when presenting all go by a slide's title; a slide without one is a
    number. Only slides whose text was read are judged.
    """
    findings = []
    for slide in slides:
        if not slide["shapes"] or any(s.get("unread") for s in slide["shapes"]):
            continue
        if _slide_heading(slide, slide_h) is not None:
            continue
        index = slide["slide_index"]
        empty = next((s["name"] for s in slide["shapes"]
                      if s.get("placeholder_type") in _TITLE_PLACEHOLDERS), None)
        if empty:
            fix = (f"Type the slide's title in '{empty}' (ppt_set_text slide_index={index} "
                   f"shape_name_or_index='{empty}').")
        else:
            fix = ("Add a title that says what the slide shows, placed and set like the "
                   "deck's other titles (ppt_add_textbox).")
        findings.append(_finding(
            index, "missing_title", [empty] if empty else [],
            f"Slide {index} has no title, so a screen reader, the outline and the list "
            "of slides show it only by its number.",
            fix, severity="warning",
        ))
    return findings


def _duplicate_titles(slides, slide_h):
    """Slides with the same title, which read as one slide in any list of them."""
    by_words = {}
    for slide in slides:
        title = _slide_heading(slide, slide_h)
        words = " ".join(((title or {}).get("text") or {}).get("content", "").split())
        if words:
            by_words.setdefault(words.casefold(), []).append((slide["slide_index"], title, words))
    findings = []
    for same in by_words.values():
        if len(same) < 2:
            continue
        numbers = [index for index, _, _ in same]
        words = same[0][2]
        findings.append(_finding(
            None, "duplicate_titles", [title["name"] for _, title, _ in same],
            f"{_slides(numbers).capitalize()} have the same title, '{words}', so they "
            "read as one slide in the outline and to a screen reader.",
            f"Make each say what its slide adds, such as '{words}: costs' and "
            f"'{words}: sales' (ppt_set_text).",
            severity="warning", slides=numbers,
        ))
    return findings


# A title placeholder this far from where its layout puts it, and no further,
# was nudged or resized by hand: further is a decision, a cover title moved
# up for a photo.
_LAYOUT_SLIP = 12.0


def _titles_off_their_layout(slides):
    """A title placeholder a little off where its layout puts it, on a layout
    no other slide uses.

    The titles of slides on one layout are compared with each other. One on
    a layout used once has nothing to be compared with but the layout, and a
    title dragged a few points or resized by hand shows as a jump when the
    deck is flipped through.
    """
    uses = {}
    for slide in slides:
        uses[(slide.get("design"), slide.get("layout"))] = uses.get(
            (slide.get("design"), slide.get("layout")), 0) + 1
    findings = []
    for slide in slides:
        home = slide.get("layout_title")
        if not home or uses[(slide.get("design"), slide.get("layout"))] != 1:
            continue
        title = next((s for s in slide["shapes"]
                      if s.get("placeholder_type") in _TITLE_PLACEHOLDERS
                      and s.get("visible", True)), None)
        if title is None:
            continue
        now = [title["left"], title["top"], title["width"], title["height"]]
        home = list(home)
        if title["height"] > home[3] and _spills_at(title, home[2], home[3]):
            # Grown to hold its words: the layout's box would spill them, and
            # putting it back undoes the overflow fix.
            home[3] = title["height"]
        off = [abs(a - b) for a, b in zip(now, home, strict=True)]
        if max(off) <= _ALIGNED or max(off) > _LAYOUT_SLIP:
            continue
        name = title["name"]
        words = []
        for field, a, b in zip(("left", "top", "width", "height"), now, home, strict=True):
            if abs(a - b) > _ALIGNED:
                words.append(f"{field} {_pt(a)} where the layout has {_pt(b)}")
        index = slide["slide_index"]
        findings.append(_finding(
            index, "inconsistent_titles", [name],
            f"The title on slide {index} is not where its layout '{slide.get('layout')}' "
            f"puts it: {'; '.join(words)}. No other slide is on that layout, so "
            "nothing else shows it, but it jumps when the deck is flipped through.",
            f"Put it back (ppt_update_shape shape_name='{name}' left={_num(home[0])} "
            f"top={_num(home[1])} width={_num(home[2])} height={_num(home[3])}).",
            severity="warning", slides=[index],
        ))
    return findings


# A text box is a slide's copy, rather than a label, a card's words or a
# source line, when it has at least this many characters in two paragraphs or
# more, across at least this share of the slide's width.
_BODY_CHARS = 80
_BODY_WIDTH = 0.6

# Body text this far from the usual size, and no further, was made smaller or
# larger to fit: further is a different kind of text, a quote set large or a
# dense appendix.
_BODY_SIZE_TOLERANCE = 1.0
_BODY_SIZE_SLIP = 6.0

_BODY_PLACEHOLDERS = frozenset({ppPlaceholderBody, ppPlaceholderObject})


def _body_text(slide, slide_w, slide_h):
    """The shape that carries a slide's copy, its size, and what it is
    compared with: its layout for a body placeholder, the deck's other
    copy for a text box. None when the slide has no copy that is clearly
    that, in one size.
    """
    title = find_title(slide, slide_h)
    if title is None:
        return None
    best = None
    for shape in slide["shapes"]:
        text = shape.get("text") or {}
        if (shape is title or not shape.get("visible", True)
                or text.get("autofit") == "shrink_to_fit"):
            continue
        placeholder = shape.get("placeholder_type")
        if placeholder in _BODY_PLACEHOLDERS:
            kind = ("layout", slide.get("design"), slide.get("layout"))
        elif placeholder is None and shape.get("type") == msoTextBox:
            paragraphs = [p for p in _paragraphs(text.get("content") or "") if p.strip()]
            if ((text.get("chars") or 0) < _BODY_CHARS or len(paragraphs) < 2
                    or not slide_w or shape["width"] < _BODY_WIDTH * slide_w):
                continue
            kind = ("copy",)
        else:
            continue
        sizes = {run.get("size") for run in text.get("runs") or [] if run.get("size")}
        if len(sizes) != 1:
            continue
        if best is None or text["chars"] > best[0]["text"]["chars"]:
            best = (shape, sizes.pop(), kind)
    return best


def _body_sizes(slides, slide_w, slide_h):
    """Body text a size or two off the rest of the deck's.

    A slide that held too much had its text made smaller to fit, and flipping
    from one slide to the next the text jumps in size. Only like is compared
    with like: a body placeholder with those on the same layout, which gives
    them all one size, and a slide's main copy in a text box with the others.
    """
    by_kind = {}
    for slide in slides:
        found = _body_text(slide, slide_w, slide_h)
        if found:
            shape, size, kind = found
            by_kind.setdefault(kind, {})[slide["slide_index"]] = (shape, size)
    findings = []
    for bodies in by_kind.values():
        if len(bodies) < _MIN_TITLED_SLIDES:
            continue
        usual, rest = _usual_group([(i, size) for i, (_, size) in bodies.items()],
                                   _BODY_SIZE_TOLERANCE)
        if len(usual) < 2:
            continue
        usual_size = sorted(v for _, v in usual)[len(usual) // 2]
        odd = [(i, v) for i, v in rest if abs(v - usual_size) <= _BODY_SIZE_SLIP]
        if not odd:
            continue
        groups = [usual] + _clusters(odd, _BODY_SIZE_TOLERANCE)
        described = "; ".join(
            f"{round(sorted(v for _, v in g)[len(g) // 2], 1):g} pt on {_slides(i for i, _ in g)}"
            for g in groups)
        names = [bodies[i][0]["name"] for i, _ in odd]
        calls = "; ".join(f"'{bodies[i][0]['name']}' on slide {i}" for i, _ in sorted(odd))
        size = f"{round(usual_size, 1):g}"
        smaller = any(v < usual_size for _, v in odd)
        findings.append(_finding(
            None, "inconsistent_body_text", names,
            f"Body text differs in size between slides: {described}.",
            f"Set the body text to {size} pt (ppt_format_text font_size={size}) "
            f"on {calls}."
            + (" Where it then no longer fits, cut the text or split the slide "
               "rather than make it smaller." if smaller else ""),
            severity="warning", slides=sorted(i for i, _ in odd),
        ))
    return findings


def design_key(area):
    """A name for a slide's master, for PowerPoint for Mac, which names none.

    It answers no name for a design, a master or a layout. Where a master
    puts its title and body is what decides how its titles are set, so that
    stands for it, and two designs in one deck are told apart. ``area`` is the
    master's title and body placeholders, as ``template_area`` holds them.
    """
    if not area:
        return None
    return "master " + " ".join(",".join(f"{v:g}" for v in box) for box in sorted(area))


def _template_title(title) -> bool:
    """Whether a title is a placeholder still in the template's own colours."""
    runs = (title.get("text") or {}).get("runs") or []
    return (title.get("placeholder_type") in _TITLE_PLACEHOLDERS and bool(runs)
            and all(run.get("theme") for run in runs))


def _usual(values):
    """The middle of the largest group of (slide, value) pairs that agree."""
    group = _usual_group(values, _NEAR_MISS)[0]
    return sorted(v for _, v in group)[len(group) // 2]


def _usual_group(values, tolerance):
    """The (slide, value) pairs within ``tolerance`` of the value the most of
    them are near, and the rest.

    Gathered around a middle rather than from the lowest value up, the titles
    split the same way whichever of them happens to be lowest: titles at
    33.4, 33.6 and 37.6 pt are one group around 33.6, where counting from
    33.4 would leave 37.6 out.
    """
    def near(centre):
        return [pair for pair in values if abs(pair[1] - centre) <= tolerance]

    if not values:
        return [], []
    centre = max(sorted({value for _, value in values}),
                 key=lambda c: (len(near(c)), -sum(abs(v - c) for _, v in near(c))))
    return near(centre), [pair for pair in values if abs(pair[1] - centre) > tolerance]


def _title_side(title):
    """The side of its frame a title's text keeps to: left, right, center or full."""
    across = (_text_anchor(title) or (None, None))[0]
    return across if across in ("right", "center", "full") else "left"


def _title_spot_taken(title, slide, spot, slide_w, slide_h):
    """Whether something on this slide paints where the deck's titles go."""
    for shape in slide["shapes"]:
        if shape is title or not shape.get("visible", True) or _is_line(shape):
            continue
        if not paints_box(shape) or (slide_w and _is_backdrop(shape, slide_w, slide_h)):
            continue
        width, height = _shared(spot, shape_box(shape))
        if width > _MIN_OVERLAP and height > _MIN_OVERLAP:
            return True
    return False


def _title_edge(title, field):
    """Where a title's text starts, which is what jumps from slide to slide.

    The frame can say otherwise: a layout's title placeholder is often far
    taller than its line, with the line anchored at its middle or bottom.
    """
    box = text_box(title) or shape_box(title)
    return {"left": box[0], "top": box[1], "right": box[2],
            "center": (box[0] + box[2]) / 2}[field]


def _title_case(title):
    """'capitals' for a title typed all in capitals, 'normal' for one with
    small letters, None for one with too few letters to tell, or none with
    a case at all, Hebrew or Arabic."""
    letters = [ch for ch in (title.get("text") or {}).get("content") or "" if ch.isalpha()]
    cased = [ch for ch in letters if ch.isupper() or ch.islower()]
    if len(cased) < 4:
        return None
    return "capitals" if all(ch.isupper() for ch in cased) else "normal"


def _title_weight(title):
    flags = [bool(run.get("bold")) for run in (title.get("text") or {}).get("runs") or []
             if run.get("bold") is not None]
    if not flags:
        return None
    return "bold" if 2 * sum(flags) >= len(flags) else "regular"


def _title_font(title):
    fonts = [run["font"] for run in (title.get("text") or {}).get("runs") or []
             if run.get("font")]
    return _commonest(fonts)


def _commonest(values):
    """The value that comes up most; the first of those that tie, so that the
    same deck gives the same answer every time."""
    return max(dict.fromkeys(values), key=values.count) if values else None


def _title_colour(title):
    colours = [run["color_hex"].upper() for run in (title.get("text") or {}).get("runs") or []
               if run.get("color_hex")]
    return _commonest(colours)


def _title_differences(label, entries, as_words, as_call, held=None, tool="ppt_format_text"):
    """Titles that differ in a setting with no in-between, a weight or a colour.

    ``held(i, usual)`` says whether slide i's title is held to the titles of
    the slides in ``usual``; one that is not is not reported.
    """
    groups = {}
    for index, title, value in entries:
        if value is not None:
            groups.setdefault(value, []).append((index, title))
    if len(groups) < 2 or sum(len(g) for g in groups.values()) < _MIN_TITLED_SLIDES:
        return []
    ordered = sorted(groups.items(), key=lambda item: (-len(item[1]), min(i for i, _ in item[1])))
    if held is not None:
        usual_members = [j for j, _ in ordered[0][1]]
        ordered = [ordered[0]] + [(value, kept) for value, kept in (
            (value, [(i, t) for i, t in members if held(i, usual_members)])
            for value, members in ordered[1:]) if kept]
        if len(ordered) < 2:
            return []
    usual = ordered[0][0]
    odd = sorted(i for _, members in ordered[1:] for i, _ in members)
    names = [t["name"] for _, members in ordered[1:] for _, t in members]
    described = "; ".join(f"{value} on {_slides(i for i, _ in members)}" for value, members in ordered)
    return [_finding(
        None, "inconsistent_titles", names,
        f"Slide titles differ in {label}: {described}.",
        f"Set the titles on {_slides(odd)} {as_words(usual)} ({tool} "
        f"{as_call(usual)}).",
        severity="warning", slides=odd,
    )]


def _title_moves(field, value, odd, titled, by_index):
    """The calls that put each odd title on ``value``, with what lines up with it.

    A slide's body is often set to its title's left edge. Moving the title
    alone would leave the body out of line with it, so whatever shares the
    edge moves with the title.
    """
    parts = []
    for index, title in titled:
        if index not in odd:
            continue
        by = value - _title_edge(title, field)
        if field == "top":
            # A title placeholder in the right place with its line anchored at
            # the bottom of a tall frame: anchoring it at the top puts the text
            # where the others are, without moving the frame off the slide.
            down = (_text_anchor(title) or (None, None))[1]
            margins = (title.get("text") or {}).get("margins") or _DEFAULT_MARGINS
            if down not in ("top", "full") and abs(title["top"] + margins["top"] - value) <= _NEAR_MISS:
                parts.append(f"on slide {index}, ppt_set_textframe vertical_anchor='top' "
                             f"on '{title['name']}', which sets its text at the top of "
                             "its frame, where the other titles' text is")
                continue
            if down not in ("top", "full") and title["top"] + by < -_EDGE_TOLERANCE:
                # Its line sits low in a tall frame, and moving the frame far
                # enough to lift it would take the frame off the slide.
                top = max(0.0, value - margins["top"])
                parts.append(f"on slide {index}, ppt_set_textframe vertical_anchor='top' "
                             f"on '{title['name']}' and then ppt_update_shape top={_num(top)} "
                             "on it, which sets its text where the other titles' text is")
                continue
        along = [shape["name"] for shape in by_index[index]["shapes"]
                 if shape is not title
                 and abs(_frame_edge(shape, field) - _frame_edge(title, field)) <= _ALIGNED]
        moved = "top" if field == "top" else "left"
        if along:
            names = [title["name"]] + along
            listed = ", ".join(f"'{n}'" for n in along)
            parts.append(
                f"on slide {index}, ppt_update_shape shape_names={names} "
                f"d{moved}={_num(by)}, which moves {listed}, lined up with the "
                "title, along with it")
        else:
            parts.append(f"on slide {index}, ppt_update_shape {moved}="
                         f"{_num(title[moved] + by)} on '{title['name']}'")
    return "; ".join(parts)


def _frame_edge(shape, field):
    return {"left": shape["left"], "top": shape["top"],
            "right": shape["left"] + shape["width"],
            "center": shape["left"] + shape["width"] / 2}[field]


# ---------------------------------------------------------------------------
# What a reader has to read
# ---------------------------------------------------------------------------
def within_reach(shape, touched) -> bool:
    """Whether a write that touched ``touched`` needs this shape's fill and text.

    Geometry is all the alignment check needs, and it comes with every shape.
    Fill and text decide only overlap and contrast, and a shape can only be in
    the way of a touched shape, or behind its text, where the two meet. So the
    rest are left unread, which on macOS is most of the cost of a check.

    Placeholders are always read, since whether one is empty decides whether it
    is lined up against anything. So are like shapes in a touched shape's row
    or column, since whether they paint decides whether the row has to be one
    size. Text overflowing its own frame far enough to reach a touched shape is
    the one meeting this misses, and that shape's own write already reported
    the overflow.
    """
    if shape.get("placeholder_type") is not None:
        return True
    box = shape_box(shape)
    for other in touched:
        other_box = shape_box(other)
        if _similar(box, other_box) and (
            _related(box, other_box, "x") or _related(box, other_box, "y")
        ):
            return True
        for reach in (other_box, text_box(other)):
            if reach is None:
                continue
            width, height = _shared(box, reach)
            if width > 0 and height > 0:
                return True
    return False


# ---------------------------------------------------------------------------
# Reading a slide (Windows)
# ---------------------------------------------------------------------------
def _fill_facts(shape):
    try:
        fill = shape.Fill
    except Exception:
        return None
    facts = {"visible": None, "solid": None, "color_hex": None, "transparency": None}
    try:
        facts["visible"] = fill.Visible == msoTrue
    except Exception:
        pass
    try:
        facts["solid"] = fill.Type == msoFillSolid
    except Exception:
        pass
    try:
        facts["color_hex"] = int_to_hex(fill.ForeColor.RGB)
    except Exception:
        pass
    try:
        facts["transparency"] = round(fill.Transparency, 2)
    except Exception:
        pass
    try:
        if fill.Type == msoFillGradient:
            facts["gradient"] = _gradient_facts(fill)
    except Exception:
        pass
    return facts


def _gradient_facts(fill):
    """A gradient's stops, as [position 0..1, colour, transparency], and its
    angle when it runs in a line: a gradient from the centre or a corner has
    none, and asking for one fails."""
    try:
        stops = fill.GradientStops
        listed = [[round(stops(i).Position, 4), int_to_hex(stops(i).Color.RGB),
                   round(stops(i).Transparency, 2)] for i in range(1, stops.Count + 1)]
    except Exception:
        logger.debug("Cannot read a gradient's stops", exc_info=True)
        return None
    try:
        angle = round(fill.GradientAngle, 2)
    except Exception:
        angle = None
    return {"angle": angle, "stops": listed} if listed else None


def _text_facts(shape, full):
    """What a layout check needs to know about a shape's text.

    Every read stands on its own, so a frame that will not report one thing
    still reports the rest, and a missing number stays None rather than
    becoming a zero that would read as a measurement.

    Where the text is drawn is always read, because any shape's text can be
    in another's way. The frame's settings and the text's colours and sizes
    are read only when ``full``: they answer overflow, contrast and titles,
    which a write tool asks only about the shape it touched.
    """
    from ppt_com.shapes import _text_frame_state

    tr = shape.TextFrame.TextRange
    try:
        content = tr.Text or ""
    except Exception:
        content = ""
    facts = {
        "chars": len(content.strip()),
        "content": content[:_CONTENT_KEPT],
        "left": None, "top": None, "width": None, "height": None,
        "alignment": None, "runs": [],
        "margins": None, "word_wrap": None, "autofit": None,
        "orientation": None, "vertical_anchor": None,
    }
    if full:
        state = _text_frame_state(shape) or {}
        for key in ("margins", "word_wrap", "autofit", "orientation", "vertical_anchor"):
            facts[key] = state.get(key)
    if not facts["chars"]:
        return facts

    # An empty frame has no bounds, which is why this waits for text.
    for key, attribute in (("left", "BoundLeft"), ("top", "BoundTop"),
                           ("width", "BoundWidth"), ("height", "BoundHeight")):
        try:
            facts[key] = round(getattr(tr, attribute), 2)
        except Exception:
            pass
    # Alignment only places text that PowerPoint would not place itself.
    if facts["left"] is None or facts["top"] is None:
        try:
            facts["alignment"] = tr.Paragraphs(1).ParagraphFormat.Alignment
        except Exception:
            pass
    if not full:
        return facts

    styles = set()
    try:
        run_count = tr.Runs().Count
    except Exception:
        run_count = 0
    for i in range(1, run_count + 1):
        try:
            run = tr.Runs(i)
            if not run.Text.strip():
                continue
            font = run.Font
            try:
                # A theme colour is the template's choice, an RGB one the caller's.
                theme = font.Color.ObjectThemeColor != 0
            except Exception:
                theme = None
            styles.add((int_to_hex(font.Color.RGB), round(font.Size, 1),
                        font.Bold == msoTrue, getattr(font, "Name", None) or None, theme))
        except Exception:
            continue
    facts["runs"] = [
        {"color_hex": colour, "size": size, "bold": bold, "font": name, "theme": theme}
        for colour, size, bold, name, theme in sorted(styles, key=lambda s: (s[0] or "", s[1] or 0))
    ]
    try:
        caps = shape.TextFrame2.TextRange.Font.Allcaps
        facts["all_caps"] = {msoTrue: True, 0: False}.get(caps)
    except Exception:
        pass
    if has_right_to_left(content):
        facts["directions"] = _directions(tr)
        try:
            # The font Hebrew and Arabic are drawn in; "" when the runs differ.
            facts["script_font"] = tr.Font.NameComplexScript or None
        except Exception:
            pass
    return facts


def has_right_to_left(content) -> bool:
    """Whether the text has a Hebrew or Arabic letter, so its paragraphs'
    direction is worth reading."""
    return any(unicodedata.bidirectional(ch) in _RIGHT_TO_LEFT for ch in content)


# Languages written right to left, by their primary language id: Arabic,
# Hebrew, Urdu, Farsi, Syriac, Divehi, Yiddish's in Windows' numbering.
_RTL_LANGUAGES = frozenset({0x01, 0x0D, 0x20, 0x29, 0x5A, 0x65, 0x3D})


def _directions(tr):
    """Each paragraph's direction as PowerPoint draws it, 'rtl' or 'ltr',
    None where unread.

    A paragraph is drawn right to left when it is set so, and also when its
    text is in a right-to-left language: PowerPoint tags Hebrew typed or set
    through it as Hebrew, and draws it from the right without being told. A
    file written by another program leaves the language out, and then only
    the setting counts.
    """
    try:
        count = tr.Paragraphs().Count
    except Exception:
        return None
    directions = []
    for i in range(1, count + 1):
        try:
            paragraph = tr.Paragraphs(i)
            if paragraph.ParagraphFormat.TextDirection == ppDirectionRightToLeft:
                directions.append("rtl")
                continue
            language = paragraph.LanguageID
            if language < 0:
                language = paragraph.Runs(1).LanguageID
            directions.append("rtl" if (language & 0x3FF) in _RTL_LANGUAGES else "ltr")
        except Exception:
            directions.append(None)
    return directions


def _geometry_facts(shape, z):
    """Where a shape is and what kind of shape it is."""
    facts = {
        "name": shape.Name,
        "z": z,
        "type": shape.Type,
        "left": round(shape.Left, 2),
        "top": round(shape.Top, 2),
        "width": round(shape.Width, 2),
        "height": round(shape.Height, 2),
        "rotation": None,
        "visible": True,
        "connector": False,
        "placeholder_type": None,
        "fill": None,
        "line_visible": None,
        "text": None,
    }
    try:
        facts["rotation"] = round(shape.Rotation, 2)
    except Exception:
        pass
    try:
        facts["visible"] = shape.Visible == msoTrue
    except Exception:
        pass
    try:
        facts["connector"] = shape.Connector == msoTrue
    except Exception:
        pass
    if facts["type"] == msoPlaceholder:
        try:
            facts["placeholder_type"] = shape.PlaceholderFormat.Type
        except Exception:
            pass
    if facts["type"] in (msoPicture, msoLinkedPicture):
        try:
            facts["lock_aspect"] = shape.LockAspectRatio == msoTrue
        except Exception:
            pass
    return facts


def _read_appearance(facts, shape, full):
    """Add what the shape draws: its fill and line, and its text."""
    if facts["type"] != msoGroup:
        facts["fill"] = _fill_facts(shape)
        try:
            facts["line_visible"] = shape.Line.Visible == msoTrue
        except Exception:
            pass
    try:
        has_frame = shape.HasTextFrame == msoTrue
    except Exception:
        has_frame = False
    if has_frame:
        try:
            facts["text"] = _text_facts(shape, full)
        except Exception:
            logger.debug("Cannot read the text of '%s'", facts["name"], exc_info=True)
    return facts


def _shape_facts(shape, z, detail):
    """One shape read in full, for a caller that has no slide around it."""
    return _read_appearance(_geometry_facts(shape, z), shape,
                            detail is None or shape.Name in detail)


def needs_background(shapes, detail) -> bool:
    """Whether the slide's background has to be read.

    It is the backdrop of last resort for the contrast check, and nothing else
    uses it, so a write that touched no text does not pay for reading it.
    """
    if detail is None:
        return True
    return any(
        (shape.get("text") or {}).get("chars")
        for shape in shapes if shape["name"] in detail
    )


def read_shapes(refs, geometry, appearance, detail):
    """Read a slide's shapes, as much of each as the check needs.

    ``geometry(ref, z)`` and ``appearance(facts, ref, full)`` are the
    platform's own readers. With no ``detail`` every shape is read in full.
    With names, those shapes are read in full first, and then the others
    only where they meet one of them; the rest stay geometry.
    """
    # A shape PowerPoint will not describe is left out, not allowed to fail the
    # whole slide.
    shapes, readable = [], []
    for z, ref in enumerate(refs, start=1):
        try:
            shapes.append(geometry(ref, z))
            readable.append(ref)
        except Exception:
            logger.debug("Cannot read shape %d", z, exc_info=True)
    refs = readable
    if detail is None:
        for facts, ref in zip(shapes, refs, strict=True):
            appearance(facts, ref, True)
        return shapes
    touched = []
    for facts, ref in zip(shapes, refs, strict=True):
        if facts["name"] in detail:
            touched.append(appearance(facts, ref, True))
    for facts, ref in zip(shapes, refs, strict=True):
        if facts["name"] not in detail and within_reach(facts, touched):
            # The colours of text on a touched card are read too, since they
            # are drawn on its fill.
            appearance(facts, ref, _on_touched(facts, touched))
        elif facts["name"] not in detail:
            # Its words are not known: no text is not the same as none read.
            facts["unread"] = True
    return shapes


def _on_touched(shape, touched) -> bool:
    """Whether the shape lies on a touched shape that paints, as words on a card."""
    box = shape_box(shape)
    for other in touched:
        if other.get("z", 0) >= shape.get("z", 0) or not paints_box(other):
            continue
        width, height = _shared(box, shape_box(other))
        if width * height >= 0.5 * _width(box) * _height(box):
            return True
    return False


def _background(slide):
    """The slide's background: its colour when it is one solid colour, and its
    gradient when it is one, each None otherwise.

    A slide that follows its layout's background, which follows the master's,
    is painted by whichever of them stopped following.
    """
    try:
        source = slide
        if slide.FollowMasterBackground == msoTrue:
            source = slide.CustomLayout
            if source.FollowMasterBackground == msoTrue:
                source = slide.Master
        fill = source.Background.Fill
        if fill.Type == msoFillSolid:
            return int_to_hex(fill.ForeColor.RGB), None
        if fill.Type == msoFillGradient:
            return None, _gradient_facts(fill)
    except Exception:
        logger.debug("Cannot read the slide background", exc_info=True)
    return None, None


def _theme_colours(slide, cache=None):
    """The text and background colours of the slide's theme, by name.

    Read once per design for a whole check, from ``cache``.
    """
    try:
        key = ("theme colours", slide.Design.Index)
    except Exception:
        key = None
    if cache is not None and key is not None and key in cache:
        return cache[key]
    try:
        scheme = slide.Design.SlideMaster.Theme.ThemeColorScheme
        colours = {name: int_to_hex(scheme(i).RGB)
                   for i, name in ((1, "dark1"), (2, "light1"), (3, "dark2"), (4, "light2"))}
    except Exception:
        logger.debug("Cannot read the theme's colours", exc_info=True)
        colours = None
    if cache is not None and key is not None:
        cache[key] = colours
    return colours


# Below anything on the slide itself, which counts from 1.
_TEMPLATE_Z = -10_000


def _template_facts(slide, touched=None, cache=None):
    """The graphics the slide's master and layout draw on it.

    A logo, a band down one side, a footer line: every slide made from the
    template has them, and the slide's own shapes have to keep clear of them.
    Placeholders are left out, since the slide has its own copies of those. The
    master's come first, being drawn first, and all of them below the slide's.

    A write reads what it needs and no more: where every graphic is, and the
    fill and text of those within reach of the shapes it touched, which are
    ``touched``. A whole check reads a layout's graphics once: ``cache`` keeps
    them by design and layout for the slides after.
    """
    try:
        if slide.DisplayMasterShapes != msoTrue:
            return []
    except Exception:
        logger.debug("Cannot tell whether the slide shows its template graphics", exc_info=True)
    key = None
    if cache is not None:
        try:
            key = (slide.Design.Index, slide.CustomLayout.Index)
        except Exception:
            key = None
        if key in cache:
            return cache[key]
    graphics = []
    sources = []
    try:
        if slide.CustomLayout.DisplayMasterShapes == msoTrue:
            sources.append(("master", slide.Master.Shapes))
    except Exception:
        logger.debug("Cannot read the master's graphics", exc_info=True)
    try:
        sources.append(("layout", slide.CustomLayout.Shapes))
    except Exception:
        logger.debug("Cannot read the layout's graphics", exc_info=True)
    for source, collection in sources:
        try:
            count = collection.Count
        except Exception:
            continue
        for i in range(1, count + 1):
            try:
                shape = collection(i)
                if shape.Type == msoPlaceholder:
                    continue
                facts = _geometry_facts(shape, _TEMPLATE_Z + len(graphics))
                if touched is None or within_reach(facts, touched):
                    _read_appearance(facts, shape, False)
            except Exception:
                logger.debug("Cannot read a %s graphic", source, exc_info=True)
                continue
            facts["source"] = source
            graphics.append(facts)
    if key is not None:
        cache[key] = graphics
    return graphics


def _template_area(slide, graphics, cache=None):
    """Where the template means content to go: its master's title and body frames.

    Read only when a sparse graphic needs it, since nothing else does.
    """
    if not any(_sparse(g) for g in graphics):
        return []
    key = None
    if cache is not None:
        try:
            key = ("area", slide.Design.Index)
        except Exception:
            key = None
        if key in cache:
            return cache[key]
    area = []
    try:
        shapes = slide.Master.Shapes
        for i in range(1, shapes.Count + 1):
            shape = shapes(i)
            if shape.Type != msoPlaceholder:
                continue
            try:
                if shape.PlaceholderFormat.Type not in _CONTENT_PLACEHOLDERS:
                    continue
            except Exception:
                continue
            area.append([round(shape.Left, 2), round(shape.Top, 2),
                         round(shape.Left + shape.Width, 2), round(shape.Top + shape.Height, 2)])
    except Exception:
        logger.debug("Cannot read the master's placeholders", exc_info=True)
    if key is not None:
        cache[key] = area
    return area


def _members_of(shape):
    """The shapes inside a group or a SmartArt graphic, with their paths."""
    from ppt_com.shape_lookup import PATH_SEPARATOR, walk_group_children

    prefix = shape.Name + PATH_SEPARATOR
    if shape.Type == msoGroup:
        yield from walk_group_children(shape, prefix)
        return
    try:
        items = shape.GroupItems
        count = items.Count
    except Exception:
        return
    for i in range(1, count + 1):
        try:
            child = items(i)
            yield child, prefix + child.Name
        except Exception:
            continue


def _inner_facts(shapes, refs):
    """The shapes inside the slide's groups and SmartArt graphics, read in full
    and named by the path the tools take: "Group 4/TextBox 2".

    The checks judge a group by its box; what is inside it has text that can
    still spill, read wrong or be left as stand-in words.
    """
    inner = []
    for facts, ref in zip(shapes, refs, strict=True):
        if facts["type"] not in (msoGroup, msoSmartArt):
            continue
        try:
            members = list(_members_of(ref))
        except Exception:
            logger.debug("Cannot read inside '%s'", facts["name"], exc_info=True)
            continue
        for order, (child, path) in enumerate(members, start=1):
            try:
                # Drawn in order over whatever is under the group.
                child_facts = _geometry_facts(child, facts["z"] + order / 1000)
                _read_appearance(child_facts, child, True)
            except Exception:
                logger.debug("Cannot read '%s'", path, exc_info=True)
                continue
            if child_facts["type"] in (msoGroup, msoSmartArt):
                continue
            child_facts["name"] = path
            child_facts["in_group"] = facts["name"]
            inner.append(child_facts)
    return inner


def _theme_fonts(slide, cache=None):
    """The theme's heading and body fonts, Latin and right to left."""
    key = None
    if cache is not None:
        try:
            key = ("theme fonts", slide.Design.Index)
        except Exception:
            key = None
        if key in cache:
            return cache[key]
    fonts = {"latin": [], "script": []}
    try:
        scheme = slide.Design.SlideMaster.Theme.ThemeFontScheme
    except Exception:
        scheme = None
    if scheme is not None:
        for part in ("MajorFont", "MinorFont"):
            for kind, index in (("latin", 1), ("script", 3)):   # msoThemeLatin, ComplexScript
                try:
                    name = getattr(scheme, part)(index).Name
                except Exception:
                    continue
                if name:
                    fonts[kind].append(name)
    if key is not None:
        cache[key] = fonts
    return fonts


def _layout_title(slide, cache=None):
    """Where the slide's layout puts its title, as a box, or None."""
    key = None
    if cache is not None:
        try:
            key = ("layout title", slide.Design.Index, slide.CustomLayout.Index)
        except Exception:
            key = None
        if key in cache:
            return cache[key]
    box = None
    try:
        shapes = slide.CustomLayout.Shapes
        for i in range(1, shapes.Count + 1):
            shape = shapes(i)
            if shape.Type != msoPlaceholder:
                continue
            try:
                if shape.PlaceholderFormat.Type not in _TITLE_PLACEHOLDERS:
                    continue
            except Exception:
                continue
            box = [round(shape.Left, 2), round(shape.Top, 2),
                   round(shape.Width, 2), round(shape.Height, 2)]
            break
    except Exception:
        logger.debug("Cannot read the layout's title", exc_info=True)
    if key is not None:
        cache[key] = box
    return box


def _slide_facts_impl(slide_indices, navigate, detail, template_cache=None):
    """Read the slides into the plain dicts the checks work on.

    ``detail`` is None to read every shape in full, or the names of the
    shapes a write touched; see ``read_shapes``. ``template_cache`` carries the
    template graphics already read from one call to the next, when a deck is
    read a slide at a time.
    """
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    total = pres.Slides.Count
    facts = {
        "slide_width": pres.PageSetup.SlideWidth,
        "slide_height": pres.PageSetup.SlideHeight,
        "slides": [],
    }
    if detail is not None:
        template_cache = None
    elif template_cache is None:
        template_cache = {}
    for slide_index in slide_indices:
        if slide_index < 1 or slide_index > total:
            raise ValueError(
                f"Slide index {slide_index} is out of range. "
                f"The presentation has {total} slides (1-based)."
            )
        if navigate:
            goto_slide(app, slide_index)
        slide = pres.Slides(slide_index)
        try:
            layout = slide.CustomLayout.Name
        except Exception:
            layout = None
        try:
            design = slide.Design.Name
        except Exception:
            design = None
        refs = [slide.Shapes(z) for z in range(1, slide.Shapes.Count + 1)]
        shapes = read_shapes(refs, _geometry_facts, _read_appearance, detail)
        if detail is None:
            _with_chart_parts(shapes, refs)
        touched = None if detail is None else [s for s in shapes if s["name"] in detail]
        try:
            follows = slide.FollowMasterBackground == msoTrue
        except Exception:
            follows = None
        template = _template_facts(slide, touched, template_cache)
        background_hex, background_gradient = (
            _background(slide) if needs_background(shapes, detail) else (None, None))
        try:
            hidden = slide.SlideShowTransition.Hidden == msoTrue
        except Exception:
            hidden = False
        facts["slides"].append({
            "slide_index": slide_index,
            "hidden": hidden,
            "animations": _animations(slide),
            # A Morph moves shapes onto this slide, or off it onto the next.
            "morph": _morph(slide) or (slide_index < total and _morph(pres.Slides(slide_index + 1))),
            "design": design,
            "layout": layout,
            "background_hex": background_hex,
            "background_gradient": background_gradient,
            "theme_colors": _theme_colours(slide, template_cache),
            "follows_background": follows,
            "template": template,
            "template_area": _template_area(slide, template, template_cache),
            "layout_title": _layout_title(slide, template_cache) if detail is None else None,
            "theme_fonts": _theme_fonts(slide, template_cache),
            "shapes": shapes,
            "inner": _inner_facts(shapes, refs) if detail is None else [],
        })
    return facts


# The most data labels of one chart that are read; a chart with more is not
# one whose labels can be read anyway, and each costs a call.
_MOST_LABELS = 120


def _with_chart_parts(shapes, refs):
    """Give each chart where its title, legend, plot and data labels are,
    in points within the chart. See ``_chart_overlaps``."""
    charts = {s["name"] for s in shapes if s.get("type") == msoChart}
    if not charts:
        return
    for ref in refs:
        try:
            if ref.Name not in charts or not ref.HasChart:
                continue
            parts = _chart_parts(ref.Chart)
        except Exception:
            logger.debug("Cannot read the parts of a chart", exc_info=True)
            continue
        for shape in shapes:
            if shape["name"] == ref.Name:
                shape["chart_parts"] = parts


def _chart_parts(chart):
    def box(part, inside=False):
        if inside:
            return _box(part.InsideLeft, part.InsideTop, part.InsideWidth, part.InsideHeight)
        return _box(part.Left, part.Top, part.Width, part.Height)

    parts = {"plot": box(chart.PlotArea, inside=True), "labels": []}
    if chart.HasTitle:
        parts["title"] = box(chart.ChartTitle)
    if chart.HasLegend:
        parts["legend"] = box(chart.Legend)
    collection = chart.SeriesCollection()
    for i in range(1, collection.Count + 1):
        points = collection.Item(i).Points()
        for j in range(1, points.Count + 1):
            if len(parts["labels"]) >= _MOST_LABELS:
                return parts
            point = points.Item(j)
            if not point.HasDataLabel:
                continue
            label = point.DataLabel
            if label.Width > 0 and label.Height > 0 and label.Text:
                parts["labels"].append(box(label))
    return parts


def _animations(slide):
    """The slide's main sequence, as ``on_screen`` takes it.

    None of it rather than part: a step left out would put the effects after
    it on the wrong click. Without it the slide is checked as it is drawn.
    """
    from ppt_com.animation import _get_animation_category

    effects = []
    try:
        sequence = slide.TimeLine.MainSequence
        for i in range(1, sequence.Count + 1):
            effect = sequence(i)
            effects.append({
                "shape": effect.Shape.Name,
                "category": _get_animation_category(effect.EffectType, bool(effect.Exit)),
                "trigger": effect.Timing.TriggerType,
            })
    except Exception:
        logger.debug("Cannot read the animations of a slide", exc_info=True)
        return []
    return effects


def _morph(slide) -> bool:
    """Whether the slide is entered with a Morph transition."""
    try:
        return slide.SlideShowTransition.EntryEffect in MORPH_EFFECTS
    except Exception:
        return False


def _package_facts_impl(slide_indices, charts=False):
    """``package_facts`` for the slides, from a copy of the deck saved aside.

    ``SaveCopyAs`` leaves the open deck as it was: not marked as changed,
    nothing to undo. ``charts`` is for the macOS reader, which reads them
    another way when the file cannot be used.
    """
    pres = ppt._get_pres_impl()
    try:
        suffix = os.path.splitext(pres.FullName)[1] or ".pptx"
    except Exception:
        suffix = ".pptx"
    if suffix.lower() not in (".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm"):
        suffix = ".pptx"
    handle, path = tempfile.mkstemp(suffix=suffix)
    os.close(handle)
    os.remove(path)
    try:
        pres.SaveCopyAs(path)
        return package_facts(path, slide_indices)
    finally:
        if os.path.exists(path):
            os.remove(path)


_FONTS = {}


def _installed_fonts_impl():
    """The names of the fonts PowerPoint can draw with on this computer.

    The fonts Windows has, by family, and the Office cloud fonts already
    downloaded. Read once: fonts installed later are seen after a restart.
    """
    if "installed" in _FONTS:
        return _FONTS["installed"]
    import win32gui
    from utils.fonts import names_in_folder

    names = set()

    def add(font, *_):
        names.add(font.lfFaceName)
        return 1

    hdc = win32gui.GetDC(0)
    try:
        win32gui.EnumFontFamilies(hdc, None, add)
    finally:
        win32gui.ReleaseDC(0, hdc)
    cloud = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "FontCache",
                         "4", "CloudFonts")
    if os.path.isdir(cloud):
        for family in os.listdir(cloud):
            names.add(family)
            names |= names_in_folder(os.path.join(cloud, family))
    # Windows lists a family by the name of its own language only: 游ゴシック is
    # "Yu Gothic" on an English Windows, and a deck made in Japan asks for the
    # first. The files hold both.
    _FONTS["installed"] = FontNames(names, (
        os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")))
    return _FONTS["installed"]


class FontNames(frozenset):
    """The installed fonts' names, lower case. One not among them is looked
    for in the font files of ``folders``, read once and only then."""

    def __new__(cls, names, folders=()):
        found = super().__new__(cls, (n.lower() for n in names if n))
        found.folders = tuple(folders)
        found.other_names = None
        return found

    def __contains__(self, name):
        if frozenset.__contains__(self, name):
            return True
        if self.other_names is None:
            from utils import fonts

            read = set()
            for folder in self.folders:
                read |= fonts.names_in_folder(folder)
            self.other_names = frozenset(n.lower() for n in read if n)
        return name in self.other_names


def _with_fonts(slides, installed):
    """Name the fonts each shape asks for that this computer does not have.

    ``installed`` is called for the list only when some text names a font.
    """
    if not any(run.get("font") for slide in slides
               for shape in slide["shapes"] + (slide.get("inner") or [])
               for run in (shape.get("text") or {}).get("runs") or []):
        return
    installed = installed()
    if not installed:
        return
    for slide in slides:
        for shape in slide["shapes"] + (slide.get("inner") or []):
            text = shape.get("text") or {}
            asked = {run.get("font") for run in text.get("runs") or [] if run.get("font")}
            missing = sorted(f for f in asked
                             if not f.startswith("+") and f.lower() not in installed)
            if missing:
                text["missing_fonts"] = missing


def _read_fonts():
    try:
        return ppt.execute(_installed_fonts_impl)
    except Exception:
        logger.debug("Cannot list the installed fonts", exc_info=True)
        return None


# The most pixels across a picture is exported at to judge text on it.
_PHOTO_PIXELS = 480


def _photo_samples_impl(slide_index, names):
    """The colours of each named picture, sampled from an export of it alone.

    The picture is written as a BMP the size it is drawn, at most
    _PHOTO_PIXELS across, without the text over it: the text's own pixels
    would otherwise count as its background. Exporting a shape changes
    nothing on the slide.
    """
    from ppt_com.constants import ppShapeFormatBMP
    from utils.bitmap import sample_bmp

    slide = ppt._get_pres_impl().Slides(slide_index)
    found = {}
    for name in names:
        handle, path = tempfile.mkstemp(suffix=".bmp")
        os.close(handle)
        try:
            shape = slide.Shapes(name)
            scale = min(1.0, _PHOTO_PIXELS / max(shape.Width, 1))
            shape.Export(path, ppShapeFormatBMP, round(shape.Width * scale),
                         round(shape.Height * scale))
            with open(path, "rb") as image:
                sampled = sample_bmp(image.read())
            if sampled:
                found[name] = sampled
        except Exception:
            logger.debug("Cannot export '%s'", name, exc_info=True)
        finally:
            if os.path.exists(path):
                os.remove(path)
    return found


def _with_photos(slides):
    """Give each picture that has text drawn on it the colours it has.

    Only for a whole check: exporting a picture is the slowest read there is.
    """
    for slide in slides:
        shapes = slide["shapes"]
        wanted = set()
        for shape in shapes:
            if not (shape.get("text") or {}).get("chars"):
                continue
            _, source, painter = _backdrop(shape, shapes, slide.get("background_hex"))
            if (source == "shape" and painter is not None and painter.get("type") in _PICTURES
                    and not _turned(painter)):
                wanted.add(painter["name"])
        if not wanted:
            continue
        try:
            found = ppt.execute(_photo_samples_impl, slide["slide_index"], sorted(wanted))
        except Exception:
            logger.debug("Cannot sample the pictures", exc_info=True)
            continue
        for shape in shapes:
            if shape["name"] in found:
                shape["photo"] = found[shape["name"]]


def _has_pictures(slides, names=None):
    return any(s.get("type") in _PICTURES and (names is None or s["name"] in names)
               for slide in slides for s in slide["shapes"])


def _needs_package(slides):
    """Whether the saved file has anything to add: pictures, charts and
    SmartArt, and text PowerPoint shrinks to fit. PowerPoint for Mac cannot
    say which text shrinks, and its file is read in place rather than saved
    aside, so there any text is reason enough."""
    return any(s.get("type") in (*_PICTURES, msoChart, msoSmartArt)
               or (s.get("text") or {}).get("autofit") == "shrink_to_fit"
               or (IS_MACOS and (s.get("text") or {}).get("chars"))
               for slide in slides for s in slide["shapes"])


def _has_charts(slides):
    return any(s.get("type") == msoChart for slide in slides for s in slide["shapes"])


def _with_package(slides, facts):
    """Give each picture its image's proportions and each chart its data,
    where they are known."""
    for slide in slides:
        index = slide["slide_index"]
        aspects = (facts.get("pictures") or {}).get(index) or {}
        pixels = (facts.get("pixels") or {}).get(index) or {}
        charts = (facts.get("charts") or {}).get(index) or {}
        described = (facts.get("described") or {}).get(index) or {}
        scales = (facts.get("text_scale") or {}).get(index) or {}
        for shape in slide["shapes"]:
            if shape["name"] in described:
                shape["described"] = described[shape["name"]]
            if shape["name"] in scales and shape.get("text"):
                shape["shrunk"] = scales[shape["name"]]
            if shape.get("type") in _PICTURES and shape["name"] in pixels:
                shape["pixels"] = pixels[shape["name"]]
            if shape.get("type") in _PICTURES and shape["name"] in aspects:
                shape["natural_aspect"] = aspects[shape["name"]]
            elif shape.get("type") == msoChart and shape["name"] in charts:
                shape["chart"] = charts[shape["name"]]


def _read_package(indices, charts=False):
    """``_package_facts_impl`` for the slides, or nothing when it fails."""
    try:
        return ppt.execute(_package_facts_impl, indices, charts) or {}
    except Exception:
        logger.debug("Cannot read the saved deck", exc_info=True)
        return {}


# ---------------------------------------------------------------------------
# Public functions
# ---------------------------------------------------------------------------
def check_layout(params: CheckLayoutInput) -> str:
    """Check one slide, or every slide and the deck as a whole."""
    # The text module's, which is already the Apple Event one on macOS.
    from ppt_com.text import _slide_count_impl

    try:
        if params.slide_index is not None:
            indices = [params.slide_index]
        else:
            indices = list(range(1, ppt.execute(_slide_count_impl) + 1))
        # A slide per call. PowerPoint is given a fixed time for each call, and
        # a long deck read in one, on a slow machine, can run past it; a slide
        # never does. The template graphics read on one slide serve the next.
        cache, slides, facts = {}, [], None
        for index in indices:
            facts = ppt.execute(_slide_facts_impl, [index], len(indices) == 1, None, cache)
            slides += facts["slides"]
        if facts is None:  # a deck with no slides
            return json.dumps({"issues": [], "total": 0, "slides_checked": 0})
        slide_w, slide_h = facts["slide_width"], facts["slide_height"]
        if _needs_package(slides):
            _with_package(slides, _read_package(indices, charts=_has_charts(slides)))
        _with_fonts(slides, _read_fonts)
        _with_photos(slides)
        findings = []
        for slide in slides:
            findings += slide_findings(slide, slide_w, slide_h)
        if params.slide_index is None:
            findings += deck_findings(slides, slide_h, slide_w)
        result = {
            "issues": findings,
            "total": len(findings),
            "slides_checked": len(slides),
        }
        return json.dumps(result, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


def _issues_for(slide_index, shapes, aspects=None, pixels=None):
    """The per-shape findings on one slide for the shapes a write touched.

    ``shapes`` are names, or 1-based positions for a tool that was given one;
    a position is turned into its name from a read of the slide's geometry.
    ``aspects`` are pictures' own proportions the write already knows, and
    ``pixels`` their images' sizes.
    """
    names = {s for s in shapes if isinstance(s, str) and s}
    positions = {s for s in shapes if isinstance(s, int) and not isinstance(s, bool)}
    if positions:
        facts = ppt.execute(_slide_facts_impl, [slide_index], False, [], None)
        names |= {s["name"] for s in facts["slides"][0]["shapes"] if s["z"] in positions}
    if not names:
        return []
    facts = ppt.execute(_slide_facts_impl, [slide_index], False, sorted(names), None)
    slides = facts["slides"]
    known, sizes = dict(aspects or {}), dict(pixels or {})
    unknown = {n for n in names if n not in known}
    if _has_pictures(slides, unknown):
        read = _read_package([slide_index])
        known = {**(read.get("pictures") or {}).get(slide_index, {}), **known}
        sizes = {**(read.get("pixels") or {}).get(slide_index, {}), **sizes}
    _with_package(slides, {"pictures": {slide_index: known}, "pixels": {slide_index: sizes}})
    _with_fonts(slides, _read_fonts)
    return slide_findings(slides[0], facts["slide_width"], facts["slide_height"],
                          only=names)


def report_layout(result: dict, slide_index: int, shape_names) -> dict:
    """Add what is now wrong with the shapes a write tool just touched.

    Called by those tools after their write has landed, with the names out of
    their own result, or the name or position they were given. Nothing is
    added when nothing is wrong. When the check itself could not run, the
    result says that instead, so an absent ``layout_issues`` always means the
    shapes were checked and are fine.
    """
    return report_layouts(result, {slide_index: shape_names})


def report_layouts(result: dict, touched) -> dict:
    """report_layout for a write that touched shapes on several slides.

    ``touched`` maps each slide index to the shapes touched on it.
    """
    # What a write hands the check is not for its caller, checked or not.
    natural = result.pop("_natural_aspect", None) if isinstance(result, dict) else None
    size = result.pop("_pixels", None) if isinstance(result, dict) else None
    if not CHECK_EACH_WRITE:
        return result
    touched = {index: [s for s in shapes if s not in (None, "")]
               for index, shapes in (touched or {}).items()}
    touched = {index: shapes for index, shapes in touched.items() if shapes}
    if not touched or not isinstance(result, dict) or "error" in result:
        return result
    aspects = {result["shape_name"]: natural} if natural and result.get("shape_name") else None
    pixels = {result["shape_name"]: size} if size and result.get("shape_name") else None
    issues = []
    try:
        for slide_index in sorted(touched):
            issues += _issues_for(slide_index, touched[slide_index], aspects, pixels)
    except Exception as e:  # noqa: BLE001 - the write itself already succeeded
        logger.debug("Layout check after a write failed", exc_info=True)
        result["layout_check_failed"] = str(e)
        return result
    if issues:
        result["layout_issues"] = issues
    return result


def _deck_findings_now(trust_template=True):
    """Every finding on every slide, and across the deck, read a slide per call."""
    from ppt_com.text import _slide_count_impl

    count = ppt.execute(_slide_count_impl)
    cache, slides, facts = {}, [], None
    for index in range(1, count + 1):
        facts = ppt.execute(_slide_facts_impl, [index], False, None, cache)
        slides += facts["slides"]
    if facts is None:
        return []
    found = []
    for slide in slides:
        found += slide_findings(slide, facts["slide_width"], facts["slide_height"],
                                trust_template=trust_template)
    return found + deck_findings(slides, facts["slide_height"], facts["slide_width"])


def _issue_key(finding):
    return finding.get("slide_index"), finding["type"], tuple(finding["shapes"])


def deck_snapshot(trust_template=True):
    """What is wrong with the deck before a change to the whole of it.

    None when it cannot be read, and then the change reports nothing rather
    than blaming itself for problems that were already there. A change to the
    theme's colours is the caller's choice, not the template's, so it is read
    without ``trust_template``; see ``_as_designed``. None too while
    ``CHECK_EACH_WRITE`` is off, without reading anything.
    """
    if not CHECK_EACH_WRITE:
        return None
    try:
        return {_issue_key(f) for f in _deck_findings_now(trust_template)}
    except Exception:  # noqa: BLE001 - the change itself still goes ahead
        logger.debug("Reading the deck before a deck-wide change failed", exc_info=True)
        return None


def report_deck_change(result: dict, before, trust_template=True) -> dict:
    """Add what a change to the whole deck left wrong that was not wrong before.

    A new theme, theme colours, fonts or slide size can change every slide at
    once: text grows past its box, a colour stops reading on its card, a shape
    lands off a smaller slide. What was already wrong is not this call's doing,
    so only what is new is reported.
    """
    if not isinstance(result, dict) or "error" in result or before is None:
        return result
    try:
        after = _deck_findings_now(trust_template)
    except Exception as e:  # noqa: BLE001 - the change itself already succeeded
        logger.debug("Layout check after a deck-wide change failed", exc_info=True)
        result["layout_check_failed"] = str(e)
        return result
    new = [f for f in after if _issue_key(f) not in before]
    if new:
        result["layout_issues"] = new
    return result


def report_contrast(result: dict, slide_indices) -> dict:
    """Add the text a background change left hard to read, on each slide it changed.

    A new background touches no shape, but every line of text on the slide is
    now drawn on it, so only the contrast check is run, over the whole slide.
    """
    if (not CHECK_EACH_WRITE or not isinstance(result, dict) or "error" in result
            or not slide_indices):
        return result
    issues = []
    try:
        cache = {}
        for slide_index in slide_indices:
            facts = ppt.execute(_slide_facts_impl, [slide_index], False, None, cache)
            slide = facts["slides"][0]
            issues += [f for f in slide_findings(slide, facts["slide_width"], facts["slide_height"])
                       if f["type"] == "low_contrast"]
    except Exception as e:  # noqa: BLE001 - the write itself already succeeded
        logger.debug("Contrast check after a background change failed", exc_info=True)
        result["layout_check_failed"] = str(e)
        return result
    if issues:
        result["layout_issues"] = issues
    return result


def touched_names(result: dict) -> list:
    """The shape names a write tool's result reports, one or several."""
    if not isinstance(result, dict):
        return []
    names = []
    for key in ("shape_name", "new_shape_name", "placeholder_name"):
        if result.get(key):
            names.append(result[key])
    for entry in result.get("updated") or []:
        if isinstance(entry, dict) and entry.get("shape_name"):
            names.append(entry["shape_name"])
    return names


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register the layout check tool with the MCP server."""

    @mcp.tool(
        name="ppt_check_layout",
        annotations={
            "title": "Check Slide Layout",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_check_layout(params: CheckLayoutInput) -> str:
        """Find layout problems on a slide, or across the deck, by measuring.

        Reports what is wrong as numbers rather than as a picture, so it works
        whether or not the caller can read an image, and it does not miss a
        three point slip the way an eye does. Each issue names the shapes and
        says how to fix it, worst first:

        - off_slide: a shape, or its text, reaches past a slide edge
        - overflow: text needs more room than its box gives it
        - overlap: text runs into another shape or other text, lies on a
          table's cells or is hidden under a card, or something the slide
          added runs into the template's own logo, band or footer
        - low_contrast: text colour against what is behind it is below the
          WCAG minimum for its size; on a gradient, against the part of it
          the text is on
        - cramped_text: text on a card pressed against one of its edges
        - narrow_margin: a shape closer to the side of the slide than the
          margin the rest of the slide keeps
        - misaligned: edges a few points apart that were meant to line up
        - uneven_spacing: a row or column of like shapes with unequal gaps
        - uneven_sizes: like shapes in a row (or column) that are not one
          height (or width), or one card in a row a little wider than the rest
        - empty_placeholder: a placeholder left without text
        - sample_text: text a template or a draft left to be replaced, such
          as lorem ipsum, "Your title here", "[Company name]" or "XX%"
        - small_text: text under 10 pt, at the size it is drawn
        - blurry_picture: a picture drawn larger than its image's pixels
        - chart_overlap: data labels on each other, or a legend or title on
          the plot (Windows only)
        - missing_alt_text: a picture, chart or SmartArt with no alt text
        - missing_title, duplicate_titles (whole deck only): a slide with no
          title, or slides with the same one
        A slide is judged as the slide show draws it: lines shown one at a
        time in one place, shapes a motion path moves, a sound's icon off the
        slide and shapes parked for a Morph are not reported.
        - inconsistent_titles (whole deck only): titles that differ in size,
          position, alignment, weight, typeface or colour from slide to slide;
          a right-set title is placed by its right edge, and the title each
          layout of a template styles for itself is left as it is

        Run it when a slide is finished, and on the whole deck before calling
        it done. A slide with no issues is finished as far as geometry goes;
        ppt_get_slide_preview is still the way to judge how it looks, and
        ppt_check_typography covers line breaks.

        The tools that add, copy, move, resize or retext a shape, or fill a
        placeholder, run the per shape checks on that shape themselves and
        report what they find in ``layout_issues``; a new slide background
        reports the text it left hard to read, and a new theme, theme colours,
        fonts or slide size what they broke across the deck. No
        ``layout_issues`` means it was checked and is fine;
        ``layout_check_failed`` means the check could not run.
        """
        return await run_offloaded(check_layout, params)


# On macOS the slide is read over Apple Events. Only _slide_facts_impl is
# swapped; the checks, and what decides how much of a slide to read, are shared.
from backend import use_mac_impls  # noqa: E402

if IS_MACOS:  # pragma: no cover - platform specific
    from ppt_mac import layout_check as _mac_layout_check
    use_mac_impls(globals(), _mac_layout_check)
