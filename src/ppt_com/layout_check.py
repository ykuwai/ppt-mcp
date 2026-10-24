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
from itertools import pairwise
from typing import Optional

from pydantic import BaseModel, ConfigDict, Field

from backend import ppt
from utils.color import int_to_hex
from utils.navigation import goto_slide
from utils.offload import run_offloaded
from ppt_com.constants import (
    msoChart, msoEmbeddedOLEObject, msoFillGradient, msoFillSolid, msoGroup, msoLine,
    msoLinkedOLEObject, msoLinkedPicture, msoMedia, msoPicture,
    msoPlaceholder, msoSmartArt, msoTable, msoTextBox, msoTrue,
    ppAlignCenter, ppAlignRight,
    ppPlaceholderCenterTitle, ppPlaceholderDate, ppPlaceholderFooter,
    ppPlaceholderBody, ppPlaceholderHeader, ppPlaceholderObject,
    ppPlaceholderSlideNumber, ppPlaceholderTitle, ppPlaceholderVerticalTitle,
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
    "narrow_margin", "misaligned", "uneven_sizes", "sample_text",
})

# Reported in this order, worst first.
_ORDER = (
    "off_slide", "overflow", "overlap", "low_contrast", "cramped_text",
    "narrow_margin", "misaligned", "uneven_spacing", "uneven_sizes",
    "empty_placeholder", "sample_text", "inconsistent_titles",
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
    of its column when it is width, so like cards stay one size. The call is
    relative, so each card grows from where it is: a row set out in steps
    keeps its steps.
    """
    name = card["name"]
    if side in ("top", "bottom"):
        word, group = "taller", "row"
        call = f"dheight={_num(amount)}"
        mates = layout["rows"].get(name, [])
    else:
        word, group = "wider", "column"
        call = f"dwidth={_num(amount)}"
        mates = layout["columns"].get(name, [])
    if side == "top":
        call += f" dtop={_num(-amount)}"
    elif side == "left":
        call += f" dleft={_num(-amount)}"
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


def _bleeds(shape, box, low, high, start, end) -> bool:
    """Whether a shape runs edge to edge along one axis on purpose.

    A background or a band drawn across the whole slide often hangs a point
    or two over, or well past both edges: nothing of it is lost, since what
    is past the edge was never meant to show. Its words, if any, still have
    to be on the slide.
    """
    if not paints_box(shape) or box[low] > start + _EDGE_TOLERANCE or box[high] < end - _EDGE_TOLERANCE:
        return False
    words = text_box(shape)
    return words is None or (words[low] >= start - _EDGE_TOLERANCE
                             and words[high] <= end + _EDGE_TOLERANCE)


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
        if box is None:
            continue
        past = {
            "left": -box[0],
            "top": -box[1],
            "right": box[2] - slide_w,
            "bottom": box[3] - slide_h,
        }
        past = {edge: round(amount, 1) for edge, amount in past.items()
                if amount > _EDGE_TOLERANCE}
        for low, high, start, end in ((0, 2, 0.0, slide_w), (1, 3, 0.0, slide_h)):
            if _bleeds(shape, box, low, high, start, end):
                for edge in (("left", "right") if low == 0 else ("top", "bottom")):
                    past.pop(edge, None)
        if not past:
            continue
        parts = [f"{_pt(amount)} past the {edge} edge" for edge, amount in past.items()]
        name = shape["name"]
        moves, calls, kept, too_big, fits = [], [], [], [], []
        for low, high, field, size, keep, words, bigger, span in (
            ("left", "right", "left", "width", margins.get("horizontal"), ("right", "left"),
             "wider", slide_w),
            ("top", "bottom", "top", "height", margins.get("vertical"), ("down", "up"),
             "taller", slide_h),
        ):
            if low not in past and high not in past:
                continue
            room = span - 2 * (keep or 0.0)
            if (low in past and high in past) or shape[size] > room + _EDGE_TOLERANCE:
                # No move puts it inside the margins: moved off one edge, it
                # runs off the other. It is brought to the margin and given
                # the room between the margins.
                (too_big if low in past and high in past else fits).append(
                    (bigger, field, size, keep or 0.0, room))
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
        if too_big or fits:
            sized = too_big + fits
            what = " and ".join(
                f"{bigger} than the slide" if (bigger, field, size, keep, room) in too_big
                else f"{bigger} than the {_pt(room)} between the margins the rest of the "
                     f"slide keeps"
                for bigger, field, size, keep, room in sized)
            calls += [f"{field}={_num(keep)} {size}={_num(room)}"
                      for _, field, size, keep, room in sized]
            why = ("no move puts it on the slide" if too_big
                   else "moving it back would run it past the other edge")
            fix = (f"'{name}' is {what}, so {why}: bring it to the margin and make it "
                   f"fit (ppt_update_shape {' '.join(calls)}).")
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
        if a_text is None or _turned(a):
            continue
        for b in shapes:
            # Text spilling out of its own box is the overflow check's.
            if b is a or _turned(b):
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
    # captions, unless that would put it on something else: then the other,
    # and then each the other way. A move that lands on something, or off
    # the slide, is not offered: a model would make it and be told again.
    a_box = shape_box(a)
    theirs = sum(_width(shape_box(b)) * _height(shape_box(b)) for b in others)
    boxes = layout["boxes"]
    carried = [s for s in layout["arranged"]
               if s is not a and all(s is not b for b in others)
               and any(b["name"] in boxes and _belongs_to(s, b, boxes) for b in others)]
    names = [b["name"] for b in others] + [c["name"] for c in carried]
    them = "it" if len(others) == 1 else "them"
    with_them = f" together with the shapes on {them}" if carried else ""
    # How far each would go the other way: the text past the near side of
    # what it ran into, or that past the near side of the text.
    back = (a_text[high] - union[low] if sign > 0 else union[high] - a_text[low]) + _CLEARANCE

    def move_text(way, by):
        if _lands_on_something(a, a_text, vertical, way * by, shapes, others, layout):
            return None
        return (f"Move '{a_name}' {ways[0] if way > 0 else ways[1]} by {_pt(by)} "
                f"(ppt_update_shape d{field}={_num(way * by)}), which leaves "
                f"{_pt(_CLEARANCE)} between it and {targets}{shorten}.")

    def move_them(way, by):
        moving = [(s, ink_box(s)) for s in others + carried]
        if _lands(moving, vertical, way * by, shapes, {a_name} | set(names), layout):
            return None
        return (f"Move {targets} {ways[0] if way > 0 else ways[1]} by {_pt(by)}"
                f"{with_them} (ppt_update_shape shape_names={names} "
                f"d{field}={_num(way * by)}), which leaves {_pt(_CLEARANCE)} "
                f"between {them} and the text of '{a_name}'{shorten}.")

    tries = [lambda: move_text(sign, clear), lambda: move_them(-sign, clear),
             lambda: move_text(-sign, back), lambda: move_them(sign, back)]
    if _width(a_box) * _height(a_box) > theirs:
        tries[0:2] = tries[1::-1]
    fix = next((found for found in (attempt() for attempt in tries) if found), None)
    size = "height" if vertical else "width"
    # A picture keeps its proportions, so it would narrow as well; a table's
    # rows do not give.
    if fix is None and len(others) == 1 and not carried and others[0][size] > 2 * clear \
            and others[0]["type"] not in (msoTable, msoPicture, msoLinkedPicture):
        # Trimmed on the side the text is on, it keeps the rest of its place.
        b = others[0]
        trim = (f"d{field}={_num(clear)} d{size}={_num(-clear)}" if sign < 0
                else f"d{size}={_num(-clear)}")
        fix = (f"Make '{b['name']}' {_pt(clear)} {'shorter' if vertical else 'narrower'} on "
               f"the side of the text of '{a_name}', which leaves {_pt(_CLEARANCE)} between "
               f"them: nothing can move {'up or down' if vertical else 'sideways'} without "
               f"running into something else or off the slide (ppt_update_shape "
               f"shape_name='{b['name']}' {trim}){shorten}.")
    if fix is None:
        fix = (f"Make room between the text of '{a_name}' and {targets}: neither can move "
               f"{'up or down' if vertical else 'sideways'} without running into something "
               f"else or off the slide, so make one of them smaller (ppt_update_shape)"
               f"{shorten}.")
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
    return _lands([(a, a_text)], vertical, by, shapes,
                  {a["name"]} | {b["name"] for b in others}, layout)


def _lands(moving, vertical, by, shapes, skip, layout):
    """Whether shapes moved this far would run into something else, or off
    the slide. ``moving`` pairs each shape with the box it draws; ``skip``
    names what is not in their way."""
    slide_w, slide_h = layout.get("slide_width"), layout.get("slide_height")
    boxes = layout["boxes"]
    # The card a moving shape sits on is not in its way.
    skip = set(skip) | {s["name"] for shape, _ in moving for s in layout["arranged"]
                        if shape["name"] in boxes and s is not shape
                        and _belongs_to(shape, s, boxes)}
    moved = []
    for _, box in moving:
        if box is None:
            continue
        box = ((box[0], box[1] + by, box[2], box[3] + by) if vertical
               else (box[0] + by, box[1], box[2] + by, box[3]))
        if slide_w and slide_h and (
                box[0] < -_EDGE_TOLERANCE or box[1] < -_EDGE_TOLERANCE
                or box[2] > slide_w + _EDGE_TOLERANCE or box[3] > slide_h + _EDGE_TOLERANCE):
            return True
        moved.append(box)
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
        for box in moved:
            width, height = _shared((box[0] - near, box[1] - near, box[2] + near, box[3] + near), ink)
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


def _fit_row(mates, layout, low, start, end, kind, named):
    """Narrow a row (or column) of like shapes evenly to fit between ``start`` and ``end``.

    Either bound may be None, and then that end of the row stays where it is.
    The gaps stay as they are, every card takes one new size, and each card
    moves with what is on it.
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
    if gap < 0 or size < 0.5 * min(sizes) or size >= min(sizes):
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
    fix = (f"Make the {kind} fit before {named}: ppt_update_shape shape_names={members} "
           f"{field}={_num(size)}")
    if moves:
        fix += f", then move {', and '.join(moves)}"
    return fix + f". The gaps stay {_pt(gap)}."


def _template_shrink(a, union, named, keep):
    """Take back the side of a shape that crosses a template graphic.

    Only for a shape that paints, since a text frame made smaller keeps its
    text by wrapping it, and only when at least ``keep`` of it is left.
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
        if union[low] >= (box[low] + box[high]) / 2:
            target = union[low] - _CLEARANCE - box[low]
            if target >= keep * size:
                options.append((size - target, (
                    f"Make '{a['name']}' {_pt(target)} {word} (ppt_update_shape "
                    f"{field}={_num(target)}), which leaves it {_pt(_CLEARANCE)} clear of {named}.")))
        elif union[high] <= (box[low] + box[high]) / 2:
            start = union[high] + _CLEARANCE
            target = box[high] - start
            if target >= keep * size:
                options.append((size - target, (
                    f"Make '{a['name']}' {_pt(target)} {word}, keeping its far edge "
                    f"(ppt_update_shape {position}={_num(start)} {field}={_num(target)}), "
                    f"which leaves it {_pt(_CLEARANCE)} clear of {named}.")))
    return min(options)[1] if options else None


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
        lines = [(s["left"], s["left"] + s["width"], f"in line with '{s['name']}'")
                 for s in layout["arranged"] if s["name"] not in skip]
        margin = (layout.get("margins") or {}).get("horizontal")
        if margin is not None and slide_w:
            lines.append((margin, slide_w - margin, "on the slide's margin"))
        right_edge = a["left"] + a["width"]
        if by > 0:
            options = sorted((left - a["left"], f", {label}") for left, _, label in lines
                             if left - a["left"] >= by - _EDGE_TOLERANCE)
        else:
            options = sorted((right - right_edge, f", {label}") for _, right, label in lines
                             if right - right_edge <= by + _EDGE_TOLERANCE)[::-1]
        for aligned, label in options:
            if abs(aligned - by) > max(40.0, abs(by)):
                break
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
        if not backdrop or (trust and _as_designed(shape, painted_by, follows_background)):
            continue
        if isinstance(backdrop, dict):
            found = _gradient_contrast(slide_index, shape, shapes, backdrop, painted_by, painter)
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
        if ours:
            suggestion, source = ours[0], ", the colour of the other text on the slide"
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


def _gradient_contrast(slide_index, shape, shapes, paint, painted_by, painter):
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
    choices = [(c, ", the colour of the other text on the slide") for c in ours[:1]] + [
        (c, "") for c in ("#000000", "#FFFFFF")
        if min(_contrast_ratio(c, b) for b in everywhere) >= needed]
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
    shapes = _settled(slide["shapes"])
    template = slide.get("template") or []
    layout = _layout(shapes, slide_w, slide_h)
    findings = (
        _off_slide(index, shapes, slide_w, slide_h, layout)
        + _overflow(index, shapes, layout)
        + _overlap(index, shapes, layout)
        + _template_collisions(index, shapes, template, layout, slide.get("template_area"))
        + _low_contrast(index, shapes, slide.get("background_hex"), template,
                        slide.get("follows_background"), layout, trust_template,
                        slide.get("background_gradient"), slide.get("theme_colors"))
        + _cramped_text(index, layout)
        + _narrow_margins(index, layout, slide_w)
        + _misaligned(index, layout, only=only)
        + _off_centre_labels(index, layout)
        + _uneven_sizes(index, layout)
        + _sample_text(index, shapes)
    )
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
        return _sorted([
            f for f in findings
            if f["type"] in SHAPE_FINDINGS and (
                any(name in only for name in f["shapes"])
                or (f["type"] == "low_contrast" and f["shapes"][0] in drawn_on))
            and not (f["type"] == "uneven_sizes" and len(f["shapes"]) == 2)
        ])
    spacing = _uneven_spacing(index, layout, slide_w)
    # A row whose gaps are reported is put back between the margins by that
    # finding's fix, so its cards are not told about the margin as well.
    spaced = {name for f in spacing for name in f["shapes"]}
    findings = [f for f in findings
                if not (f["type"] == "narrow_margin" and f["shapes"][0] in spaced)]
    findings += spacing + _empty_placeholders(index, shapes)
    return _sorted(findings)


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


def deck_findings(slides, slide_h, slide_w=None):
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
    titled = []
    for slide in slides:
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
    return facts


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
        touched = None if detail is None else [s for s in shapes if s["name"] in detail]
        try:
            follows = slide.FollowMasterBackground == msoTrue
        except Exception:
            follows = None
        template = _template_facts(slide, touched, template_cache)
        background_hex, background_gradient = (
            _background(slide) if needs_background(shapes, detail) else (None, None))
        facts["slides"].append({
            "slide_index": slide_index,
            "design": design,
            "layout": layout,
            "background_hex": background_hex,
            "background_gradient": background_gradient,
            "theme_colors": _theme_colours(slide, template_cache),
            "follows_background": follows,
            "template": template,
            "template_area": _template_area(slide, template, template_cache),
            "shapes": shapes,
        })
    return facts


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


def _issues_for(slide_index, shapes):
    """The per-shape findings on one slide for the shapes a write touched.

    ``shapes`` are names, or 1-based positions for a tool that was given one;
    a position is turned into its name from a read of the slide's geometry.
    """
    names = {s for s in shapes if isinstance(s, str) and s}
    positions = {s for s in shapes if isinstance(s, int) and not isinstance(s, bool)}
    if positions:
        facts = ppt.execute(_slide_facts_impl, [slide_index], False, [], None)
        names |= {s["name"] for s in facts["slides"][0]["shapes"] if s["z"] in positions}
    if not names:
        return []
    facts = ppt.execute(_slide_facts_impl, [slide_index], False, sorted(names), None)
    return slide_findings(facts["slides"][0], facts["slide_width"], facts["slide_height"],
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


def report_slides(result: dict, slide_indices) -> dict:
    """report_layout for slides a write made whole, such as copies: every
    shape on each is checked as the write's own."""
    if not CHECK_EACH_WRITE or not isinstance(result, dict) or "error" in result:
        return result
    touched = {}
    try:
        for index in slide_indices:
            facts = ppt.execute(_slide_facts_impl, [index], False, [], None)
            touched[index] = [s["name"] for s in facts["slides"][0]["shapes"]]
    except Exception as e:  # the write itself already succeeded
        logger.debug("Layout check after a write failed", exc_info=True)
        result["layout_check_failed"] = str(e)
        return result
    return report_layouts(result, touched)


def report_layouts(result: dict, touched) -> dict:
    """report_layout for a write that touched shapes on several slides.

    ``touched`` maps each slide index to the shapes touched on it.
    """
    if not CHECK_EACH_WRITE:
        return result
    touched = {index: [s for s in shapes if s not in (None, "")]
               for index, shapes in (touched or {}).items()}
    touched = {index: shapes for index, shapes in touched.items() if shapes}
    if not touched or not isinstance(result, dict) or "error" in result:
        return result
    issues = []
    try:
        for slide_index in sorted(touched):
            issues += _issues_for(slide_index, touched[slide_index])
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
from backend import IS_MACOS, use_mac_impls  # noqa: E402

if IS_MACOS:  # pragma: no cover - platform specific
    from ppt_mac import layout_check as _mac_layout_check
    use_mac_impls(globals(), _mac_layout_check)
