"""Layout tools, on Apple Events.

Mirrors ``ppt_com/layout.py``. Same function names, same signatures, same
returned shapes.

Three of these tools cannot be done at all here, and they share one cause.
``align``, ``distribute``, ``group`` and ``ungroup`` all take a ``shape range``,
and the only ``shape range`` PowerPoint for Mac hands out is the one already
selected in a window. Its dictionary has an ``unselect`` command and no
``select`` command, so a script cannot put shapes into a selection to act on
them. ``ppt_merge_shapes`` has no route either, for a simpler reason: the
Boolean merge verbs are not in the dictionary in any form. All three refuse
rather than pretend.

Slide height is the other gap. ``page setup`` carries ``slide width`` and no
``slide height``, and the slide master's ``height`` is read only, so a height
in points cannot be set. A preset size can, because it sets both dimensions at
once, and what comes back is read out of PowerPoint rather than assumed.
"""

import logging
import os

from appscript import k

from backend.mac_ae import ppt
from backend.mac_enums import MsoFlipCmd, MsoGradientStyle, to_keyword
from ppt_com.constants import (
    ALIGN_CMD_MAP,
    DISTRIBUTE_CMD_MAP,
    FLIP_CMD_MAP,
    GRADIENT_STYLE_MAP,
    MERGE_CMD_MAP,
)
from ppt_mac.shapes import _get_shape, _slide
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
# is not repeated here. This table carries both answers instead: the
# PpSlideSizeType constant, and the preset name the tool takes as input.
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


# ---------------------------------------------------------------------------
# Apple Event implementation functions
# ---------------------------------------------------------------------------
def _align_shapes_impl(slide_index, shape_names, align_to, relative_to_slide):
    """Refuse, because a shape range cannot be built from a script here."""
    # Validated first so a caller with a typo hears about the typo too.
    align_key = align_to.strip().lower()
    if align_key not in ALIGN_CMD_MAP:
        raise ValueError(
            f"Unknown align_to '{align_to}'. "
            f"Valid values: {list(ALIGN_CMD_MAP.keys())}"
        )
    return {
        "error": "ppt_align_shapes is not available on macOS",
        "reason": _NO_SHAPE_RANGE,
        "platform": "macOS",
        "alternatives": ["ppt_update_shape", "ppt_get_shape_info"],
    }


def _distribute_shapes_impl(slide_index, shape_names, direction, relative_to_slide):
    """Refuse, for the same reason as align."""
    dir_key = direction.strip().lower()
    if dir_key not in DISTRIBUTE_CMD_MAP:
        raise ValueError(
            f"Unknown direction '{direction}'. "
            f"Valid values: {list(DISTRIBUTE_CMD_MAP.keys())}"
        )
    return {
        "error": "ppt_distribute_shapes is not available on macOS",
        "reason": _NO_SHAPE_RANGE,
        "platform": "macOS",
        "alternatives": ["ppt_update_shape", "ppt_get_shape_info"],
    }


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

        applied.append(idx)

    result = {
        "success": True,
        "slide_indices": applied,
        "fill_type": fill_key,
    }
    # Backward compatibility: include slide_index when called with single target
    if slide_indices is None:
        result["slide_index"] = applied[0]
    return result


def _flip_shape_impl(slide_index, shape_name_or_index, direction):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    dir_key = direction.strip().lower()
    flip_cmd = FLIP_CMD_MAP.get(dir_key)
    if flip_cmd is None:
        raise ValueError(
            f"Unknown direction '{direction}'. "
            f"Valid values: {list(FLIP_CMD_MAP.keys())}"
        )

    shape.flip(direction=to_keyword(MsoFlipCmd, flip_cmd, "flip direction"))

    # Read back flip state
    return {
        "success": True,
        "shape_name": shape.name(),
        "horizontal_flip": bool(shape.horizontal_flip()),
        "vertical_flip": bool(shape.vertical_flip()),
    }
