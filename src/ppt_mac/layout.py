"""Layout tools, on Apple Events.

Mirrors ``ppt_com/layout.py``. Same function names, same signatures, same
returned shapes.

``group``, ``ungroup`` and the shape range they need cannot be reached here.
The only ``shape range`` PowerPoint for Mac hands out is the one already
selected in a window, its dictionary has an ``unselect`` command and no
``select`` command, so a script cannot put shapes into a selection to act on
them. ``ppt_merge_shapes`` has no route either, for the simpler reason that
the Boolean merge verbs are not in the dictionary in any form. Those refuse rather
than pretend.

``align`` and ``distribute`` were in that list and are not any more. Windows
hands both to ``ShapeRange.Align`` and ``ShapeRange.Distribute``, but neither
is doing anything a script cannot do itself. Every shape's ``left position``,
``top``, ``width`` and ``height`` read and write cleanly here, and moving four
shapes is four writes. So they are computed rather than delegated, and the
result is the same picture on both platforms.

Slide height is the other gap. ``page setup`` carries ``slide width`` and no
``slide height``, and the slide master's ``height`` is read only, so a height
in points cannot be set. A preset size can, because it sets both dimensions at
once, and what comes back is read out of PowerPoint rather than assumed.
"""

import logging
import os

from appscript import k
from appscript.reference import CommandError

from backend.mac_ae import ppt, slide_at as _slide
from backend.mac_enums import MsoFlipCmd, MsoGradientStyle, to_keyword
from ppt_com.constants import (
    ALIGN_CMD_MAP,
    DISTRIBUTE_CMD_MAP,
    FLIP_CMD_MAP,
    GRADIENT_STYLE_MAP,
    MERGE_CMD_MAP,
)
from ppt_mac.shapes import _get_shape
from utils.color import hex_to_rgb_list
from utils.navigation import goto_slide

logger = logging.getLogger(__name__)

# The reason every shape range tool gives, written once because it is the same
# reason each time.
_NO_SHAPE_RANGE = (
    "PowerPoint for Mac only offers a shape range through a window's current "
    "selection, and its Apple Event dictionary has no `select` command, only "
    "`unselect`. There is no way for a script to gather shapes into a range "
    "to act on them."
)

# macOS names its slide sizes and Windows numbers them. The numbers agree, but
# the project's own SLIDE_SIZE_MAP does not agree with either (it has A3 at 8
# and 16:9 at 9), so the reverse lookup the Windows code does through that map
# is not repeated here. This table carries both answers instead, the
# PpSlideSizeType constant and the preset name the tool takes as input.
_SLIDE_SIZES = {
    k.slide_size_on_screen: (1, "4:3"),
    k.slide_size_letter_paper: (2, "letter"),
    k.slide_size_A4_paper: (3, "a4"),
    k.slide_size_35_MM: (4, "35mm"),
    k.slide_size_overhead: (5, "overhead"),
    k.slide_size_banner: (6, "banner"),
    k.slide_size_custom: (7, "custom"),
    k.slide_size_ledger_paper: (8, None),
    k.slide_size_A3_paper: (9, "a3"),
    k.slide_size_B4_ISO_paper: (10, None),
    k.slide_size_B5_ISO_paper: (11, None),
    k.slide_size_B4_JIS_paper: (12, None),
    k.slide_size_B5_JIS_paper: (13, None),
    k.slide_size_hagaki_card: (14, None),
    k.slide_size_on_screen_16x9: (15, "16:9"),
    k.slide_size_on_screen_16x10: (16, "16:10"),
}

# The presets ppt_set_slide_size accepts, as slide sizes PowerPoint knows.
_PRESET_SIZES = {
    "16:9": k.slide_size_on_screen_16x9,
    "widescreen": k.slide_size_on_screen_16x9,
    "4:3": k.slide_size_on_screen,
    "16:10": k.slide_size_on_screen_16x10,
    "a4": k.slide_size_A4_paper,
    "a3": k.slide_size_A3_paper,
    "letter": k.slide_size_letter_paper,
    "35mm": k.slide_size_35_MM,
    "overhead": k.slide_size_overhead,
    "banner": k.slide_size_banner,
}

_ORIENTATIONS = {
    "landscape": k.horizontal_orientation,
    "portrait": k.vertical_orientation,
}


def _slide_dimensions(pres) -> tuple:
    """Return the slide width and height in points.

    The height does not come from ``page setup``, which has no such property.
    The slide master carries it, and reads back the same number Windows gets
    from PageSetup.SlideHeight.
    """
    return (pres.page_setup.slide_width(), pres.slide_master.height())


def _boxes(slide_index, shape_names) -> list:
    """Read each named shape's rectangle once, in one pass.

    Every alignment decision needs all four numbers for every shape before it
    can move any of them, and reading a property is an Apple Event, so they are
    read together rather than one at a time inside the loop that writes.
    ``_get_shape`` raises for a name that is not on the slide, which is the
    same error Windows gives.
    """
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    boxes = []
    for name in shape_names:
        shape = _get_shape(slide, name)
        boxes.append({
            "shape": shape,
            "left": shape.left_position(),
            "top": shape.top(),
            "width": shape.width(),
            "height": shape.height(),
        })
    return boxes


def _slide_size() -> tuple:
    """The slide's width and height in points."""
    return _slide_dimensions(ppt._get_pres_impl())


# How far a shape may sit from where it was sent and still count as arrived.
# PowerPoint rounds a position to the nearest fraction of a point and a shape
# locked by its placeholder does not move at all, so the tolerance separates
# rounding from refusal rather than being a margin of comfort.
_POSITION_TOLERANCE = 0.5


def _verify_moved(wanted, horizontal: bool) -> tuple:
    """Split shapes into the ones that moved and the ones that did not.

    ``wanted`` is a list of (box, target) pairs. One Apple Event per shape,
    which is what turns a count of writes sent into a count of shapes standing
    where they were asked to stand.
    """
    landed, stayed = [], []
    for box, target in wanted:
        try:
            now = box["shape"].left_position() if horizontal else box["shape"].top()
        except (CommandError, AttributeError):
            stayed.append(box["shape"].name())
            continue
        if now is None or abs(now - target) > _POSITION_TOLERANCE:
            stayed.append(box["shape"].name())
        else:
            landed.append(box["shape"].name())
    return landed, stayed


def _did_not_move_warning(stayed, horizontal: bool) -> str:
    """The sentence both layout tools carry for a shape that stayed put."""
    edge = "left" if horizontal else "top"
    return (
        f"{', '.join(stayed)} did not move. PowerPoint reported no error and "
        f"the {edge} edge reads back where it was, which is what a locked "
        "shape or a placeholder driven by its layout does."
    )


# ---------------------------------------------------------------------------
# Apple Event implementation functions
# ---------------------------------------------------------------------------
def _align_shapes_impl(slide_index, shape_names, align_to, relative_to_slide):
    """Align shapes by moving them, since there is no shape range to hand it to.

    PowerPoint aligns to the bounding box of the shapes themselves, or to the
    slide when asked to. Both are arithmetic on four readable properties, so
    both are done here rather than refused.
    """
    align_key = align_to.strip().lower()
    if align_key not in ALIGN_CMD_MAP:
        raise ValueError(
            f"Unknown align_to '{align_to}'. "
            f"Valid values: {list(ALIGN_CMD_MAP.keys())}"
        )

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    boxes = _boxes(slide_index, shape_names)

    if relative_to_slide:
        width, height = _slide_size()
        left, top, right, bottom = 0.0, 0.0, width, height
    else:
        left = min(b["left"] for b in boxes)
        top = min(b["top"] for b in boxes)
        right = max(b["left"] + b["width"] for b in boxes)
        bottom = max(b["top"] + b["height"] for b in boxes)

    horizontal = align_key in ("left", "center", "right")
    wanted = []
    for box in boxes:
        if align_key == "left":
            target = left
        elif align_key == "center":
            target = (left + right) / 2 - box["width"] / 2
        elif align_key == "right":
            target = right - box["width"]
        elif align_key == "top":
            target = top
        elif align_key == "middle":
            target = (top + bottom) / 2 - box["height"] / 2
        else:  # bottom
            target = bottom - box["height"]
        (box["shape"].left_position if horizontal else box["shape"].top).set(target)
        wanted.append((box, target))

    landed, stayed = _verify_moved(wanted, horizontal)

    result = {
        "success": bool(landed) or not shape_names,
        "aligned_count": len(landed),
        "align_to": align_key,
        "relative_to_slide": relative_to_slide,
    }
    if stayed:
        result["warnings"] = [_did_not_move_warning(stayed, horizontal)]
    return result


def _distribute_shapes_impl(slide_index, shape_names, direction, relative_to_slide):
    """Space shapes so the gaps between them are equal.

    PowerPoint equalises the gaps between edges, not the distance between
    centres, and it leaves the two outermost shapes where they are. Asked to
    work relative to the slide, it spreads them from edge to edge instead.
    Fewer than three shapes have nothing to distribute, which is also what
    Windows does with them.
    """
    dir_key = direction.strip().lower()
    if dir_key not in DISTRIBUTE_CMD_MAP:
        raise ValueError(
            f"Unknown direction '{direction}'. "
            f"Valid values: {list(DISTRIBUTE_CMD_MAP.keys())}"
        )

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    boxes = _boxes(slide_index, shape_names)

    horizontal = dir_key == "horizontal"
    span = "width" if horizontal else "height"
    start = "left" if horizontal else "top"
    boxes.sort(key=lambda b: b[start])

    if relative_to_slide:
        width, height = _slide_size()
        first, last = 0.0, (width if horizontal else height)
    else:
        first = boxes[0][start]
        last = boxes[-1][start] + boxes[-1][span]

    used = sum(b[span] for b in boxes)
    gap = (last - first - used) / (len(boxes) - 1) if len(boxes) > 1 else 0.0

    cursor = first
    wanted = []
    for box in boxes:
        if horizontal:
            box["shape"].left_position.set(cursor)
        else:
            box["shape"].top.set(cursor)
        wanted.append((box, cursor))
        cursor += box[span] + gap

    landed, stayed = _verify_moved(wanted, horizontal)

    result = {
        "success": bool(landed) or not shape_names,
        "distributed_count": len(landed),
        "direction": dir_key,
        "relative_to_slide": relative_to_slide,
    }
    if stayed:
        result["warnings"] = [_did_not_move_warning(stayed, horizontal)]
    return result


def _merge_shapes_impl(slide_index, shape_names, merge_type, primary_shape):
    """Refuse, because the Boolean merge verbs are not in the dictionary."""
    merge_key = merge_type.strip().lower()
    if merge_key not in MERGE_CMD_MAP:
        raise ValueError(
            f"Unknown merge_type '{merge_type}'. "
            f"Valid values: {list(MERGE_CMD_MAP.keys())}"
        )
    return {
        "error": "ppt_merge_shapes is not available on macOS",
        "reason": (
            "PowerPoint for Mac exposes no union, combine, intersect, subtract "
            "or fragment command to Apple Events, and the shapes to merge "
            "would have to be gathered into a shape range in the first place, "
            "which a script cannot do. " + _NO_SHAPE_RANGE
        ),
        "platform": "macOS",
        "alternatives": ["ppt_add_shape"],
    }


def _get_slide_size_impl():
    ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    ps = pres.page_setup

    width_pt, height_pt = _slide_dimensions(pres)
    size_word = ps.slide_size()
    orientation = ps.slide_orientation()

    slide_size, preset_name = _SLIDE_SIZES.get(size_word, (None, None))

    return {
        "success": True,
        "width_points": round(width_pt, 2),
        "height_points": round(height_pt, 2),
        "width_inches": round(width_pt / 72.0, 4),
        "height_inches": round(height_pt / 72.0, 4),
        "slide_size_type": slide_size,
        "slide_size_name": preset_name,
        "orientation": (
            "landscape" if orientation == k.horizontal_orientation else "portrait"
        ),
    }


def _set_slide_size_impl(width, height, preset, orientation):
    ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    ps = pres.page_setup

    # Refused before anything is written, because half applying this would
    # leave a deck the caller was told it did not get.
    if height is not None:
        return {
            "error": "ppt_set_slide_size cannot set the slide height on macOS",
            "reason": (
                "PowerPoint for Mac's `page setup` carries `slide width` and no "
                "`slide height`, and the slide master's `height` is read only, "
                "so a height in points has nowhere to go. A preset size sets "
                "both dimensions at once and does work."
            ),
            "platform": "macOS",
            "alternatives": ["ppt_set_slide_size with preset instead of height"],
        }

    # Set preset FIRST (it changes both dimensions)
    if preset is not None:
        preset_key = preset.strip().lower()
        wanted = _PRESET_SIZES.get(preset_key)
        if wanted is None:
            raise ValueError(
                f"Unknown preset '{preset}'. "
                f"Valid values: {list(_PRESET_SIZES.keys())}"
            )
        # A preset here is PowerPoint's own named size rather than the point
        # pair the Windows code sets, so the deck ends up at PowerPoint's
        # dimensions for that name (720 by 405 for 16:9, not 960 by 540). The
        # aspect is what the caller asked for, and every number below is read
        # back rather than assumed.
        before = _slide_dimensions(pres)
        ps.slide_size.set(wanted)
        if _slide_dimensions(pres) == before and ps.slide_size() != wanted:
            raise RuntimeError(
                f"PowerPoint accepted the '{preset_key}' slide size and neither "
                "dimension changed, which is how it reports a change it did not "
                "make."
            )

    # Then set an explicit width (overrides the preset width)
    if width is not None:
        ps.slide_width.set(width)

    # Then set orientation
    if orientation is not None:
        orient_key = orientation.strip().lower()
        wanted_orientation = _ORIENTATIONS.get(orient_key)
        if wanted_orientation is None:
            raise ValueError(
                f"Unknown orientation '{orientation}'. "
                f"Valid values: 'landscape', 'portrait'"
            )
        ps.slide_orientation.set(wanted_orientation)

    # Read back final values
    width_pt, height_pt = _slide_dimensions(pres)
    return {
        "success": True,
        "width_points": round(width_pt, 2),
        "height_points": round(height_pt, 2),
        "width_inches": round(width_pt / 72.0, 4),
        "height_inches": round(height_pt / 72.0, 4),
    }


def _set_slide_background_impl(slide_index, fill_type, color,
                                gradient_color1, gradient_color2,
                                gradient_style, image_path, transparency,
                                slide_indices=None):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()

    # Determine target slides
    targets = slide_indices if slide_indices else [slide_index]

    # Validate parameters once before the loop
    fill_key = fill_type.strip().lower()

    if fill_key == "solid":
        if color is None:
            raise ValueError("color is required for solid fill")
        color_rgb = hex_to_rgb_list(color)
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
        style_word = to_keyword(MsoGradientStyle, style_val, "gradient style")
        color1_rgb = hex_to_rgb_list(gradient_color1)
        color2_rgb = hex_to_rgb_list(gradient_color2)
    elif fill_key == "picture":
        if image_path is None:
            raise ValueError("image_path is required for picture fill")
        # A POSIX path, always. An HFS colon path is taken as a literal file
        # name and lands as a file called "Macintosh HD:Users:..." somewhere
        # nobody will look for it.
        abs_path = os.path.abspath(os.path.expanduser(image_path))
        if not os.path.isfile(abs_path):
            raise ValueError(f"Image file not found: {abs_path}")
    elif fill_key not in ("none", "master"):
        raise ValueError(
            f"Unknown fill_type '{fill_type}'. "
            f"Valid values: 'solid', 'gradient', 'picture', 'none', 'master'"
        )

    applied = []
    missed = []
    clamped = []
    for idx in targets:
        goto_slide(app, idx)
        slide = _slide(pres, idx)

        if fill_key == "master":
            slide.follow_master_background.set(True)
        else:
            # Detach from master background
            slide.follow_master_background.set(False)
            # The background is a shape here, and its fill is that shape's
            # fill format.
            fill = slide.background.fill_format

            if fill_key == "solid":
                fill.solid()
                fill.fore_color.set(color_rgb)

            elif fill_key == "gradient":
                fill.two_color_gradient(style=style_word, variant=1)
                fill.fore_color.set(color1_rgb)
                fill.back_color.set(color2_rgb)

            elif fill_key == "picture":
                fill.user_picture(picture_file=abs_path)
                # A picture fill is one of the operations that reports success
                # and does nothing, usually because PowerPoint's sandbox has no
                # grant for the folder the image sits in.
                if fill.fill_format_type() != k.fill_picture:
                    raise RuntimeError(
                        f"PowerPoint reported success but slide {idx} has no "
                        f"picture fill. It is sandboxed and may have no access "
                        f"to {os.path.dirname(abs_path)}; copying the image "
                        "into a folder PowerPoint has already opened a file "
                        "from is the usual fix."
                    )

            elif fill_key == "none":
                # No Fill.Background() on macOS; the fill's own visibility flag
                # is what "no fill" means in the interface.
                fill.visible.set(False)

            # Apply transparency if specified
            if transparency is not None and fill_key not in ("none", "master"):
                fill.transparency.set(transparency)

        # One call can name twenty slides, and this used to record all twenty
        # as repainted whether or not any of them were. Each one is measured
        # before it is counted.
        failure = _background_landed(slide, fill_key)
        if failure is None:
            applied.append(idx)
        else:
            missed.append(f"slide {idx} ({failure})")

        # Asked separately from whether the slide was repainted at all,
        # because PowerPoint keeps a transparency inside its own range without
        # saying so, and a clamped number is not a background that did not
        # take.
        wrote_transparency = (
            transparency is not None and fill_key not in ("none", "master")
        )
        if failure is None and wrote_transparency:
            held = _background_transparency(slide)
            if held is None or abs(held - transparency) > 0.01:
                clamped.append(f"slide {idx} holds {held}")

    result = {
        # A call that repainted nothing is not a success, whatever it reported.
        "success": bool(applied),
        "slide_indices": applied,
        "fill_type": fill_key,
    }
    warnings = []
    if missed:
        warnings.append(
            "PowerPoint reported no error and these slides did not take the "
            f"background: {'; '.join(missed)}."
        )
    if clamped:
        warnings.append(
            f"The background was painted and a transparency of {transparency} "
            f"is not what came back: {'; '.join(clamped)}. PowerPoint keeps "
            "the value inside its own range and says nothing about doing it."
        )
    if warnings:
        result["warnings"] = warnings
    # Backward compatibility: include slide_index when called with single
    # target. It names the slide that was asked for, so a call where nothing
    # landed still says which slide it was about.
    if slide_indices is None:
        result["slide_index"] = targets[0]
    return result


def _background_landed(slide, fill_key: str):
    """Read a slide's background back, and say what is wrong with it.

    Returns None when it took, and a short reason when it did not. The picture
    branch is not checked here because it checks itself where it is written and
    raises, which is the one case where carrying on would leave the caller with
    an image PowerPoint's sandbox never let it read.

    ``visible`` and ``fill format type`` are the pair ``ppt_get_shape_info``
    reads off a shape's fill, and a slide's background is a shape, so the same
    two answer this. Each read is guarded on its own; a background that will
    not answer is reported as unverified rather than as a failure.
    """
    if fill_key == "master":
        try:
            return None if slide.follow_master_background() else (
                "it is still painted with its own background"
            )
        except (CommandError, AttributeError):
            return "PowerPoint would not say whether it follows the master"

    fill = slide.background.fill_format
    try:
        visible = bool(fill.visible())
    except (CommandError, AttributeError):
        visible = None

    if fill_key == "none":
        if visible is None:
            return "PowerPoint would not say whether the background is hidden"
        return None if visible is False else "the background is still filled"

    if fill_key in ("solid", "gradient"):
        expected = k.fill_solid if fill_key == "solid" else k.fill_gradient
        try:
            word = fill.fill_format_type()
        except (CommandError, AttributeError):
            return "PowerPoint would not say what fill the background carries"
        if word != expected:
            return f"the background reads back as {_keyword_name(word)}"

    return None


def _background_transparency(slide):
    """Read a slide background's transparency, or None when it will not say."""
    try:
        return slide.background.fill_format.transparency()
    except (CommandError, AttributeError):
        return None


def _keyword_name(value) -> str:
    """Return an appscript keyword's own name, as it reads in the dictionary."""
    return str(value).replace("k.", "").replace("_", " ")


def _flip_shape_impl(slide_index, shape_name_or_index, direction):
    # Before goto_slide, so a misspelled direction costs neither an Apple Event
    # nor a jump to a slide the caller was not looking at.
    dir_key = direction.strip().lower()
    flip_cmd = FLIP_CMD_MAP.get(dir_key)
    if flip_cmd is None:
        raise ValueError(
            f"Unknown direction '{direction}'. "
            f"Valid values: {list(FLIP_CMD_MAP.keys())}"
        )
    flip_word = to_keyword(MsoFlipCmd, flip_cmd, "flip direction")

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    shape.flip(direction=flip_word)

    # Read back flip state
    return {
        "success": True,
        "shape_name": shape.name(),
        "horizontal_flip": bool(shape.horizontal_flip()),
        "vertical_flip": bool(shape.vertical_flip()),
    }
