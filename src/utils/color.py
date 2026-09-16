"""Color conversion utilities for PowerPoint COM automation.

PowerPoint COM uses BGR-ordered integers for color values.
The formula is: R + (G * 256) + (B * 65536).
This is the OPPOSITE of standard HTML hex notation (#RRGGBB).
"""

THEME_COLOR_MAP = {
    "dark1": 1,
    "light1": 2,
    "dark2": 3,
    "light2": 4,
    "accent1": 5,
    "accent2": 6,
    "accent3": 7,
    "accent4": 8,
    "accent5": 9,
    "accent6": 10,
    "hyperlink": 11,
    "followed_hyperlink": 12,
}


def rgb_to_int(r: int, g: int, b: int) -> int:
    """Convert RGB values (0-255 each) to PowerPoint's BGR integer format.

    This matches VBA's RGB(r, g, b) function.

    Examples:
        rgb_to_int(255, 0, 0)   -> 255       (red)
        rgb_to_int(0, 0, 255)   -> 16711680  (blue)
    """
    return r + (g << 8) + (b << 16)


def int_to_rgb(color_int: int) -> tuple[int, int, int]:
    """Convert PowerPoint's BGR integer to an (R, G, B) tuple."""
    r = color_int & 0xFF
    g = (color_int >> 8) & 0xFF
    b = (color_int >> 16) & 0xFF
    return (r, g, b)


def hex_to_rgb(hex_str: str) -> tuple[int, int, int]:
    """Convert a hex color string (#RRGGBB or RRGGBB) to (R, G, B) tuple."""
    raw = hex_str
    hex_str = hex_str.lstrip("#")
    if len(hex_str) == 3:
        hex_str = "".join(c * 2 for c in hex_str)
    if len(hex_str) != 6:
        # A theme name is the likeliest thing to arrive here, because the
        # instructions tell callers to use the deck's accent colours and only
        # some tools take one by name. `Invalid hex color: #accent1` sends
        # nobody anywhere, so say where the number lives.
        if raw.lower().replace(" ", "_").replace("-", "_") in THEME_COLOR_MAP:
            raise ValueError(
                f"'{raw}' is a theme colour name, and this argument takes a "
                "hex value like '#156082'. Read the deck's own value from "
                "ppt_get_presentation_info, which returns accent_colors as "
                "hex. The arguments that take a name instead are "
                "ppt_format_text's font_color_theme and ppt_add_svg_icon's "
                "color."
            )
        raise ValueError(f"Invalid hex color: #{hex_str}")
    return (int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16))


def hex_to_int(hex_str: str) -> int:
    """Convert a hex color string directly to PowerPoint's BGR integer."""
    r, g, b = hex_to_rgb(hex_str)
    return rgb_to_int(r, g, b)


def int_to_hex(color_int: int) -> str:
    """Convert PowerPoint's BGR integer to a #RRGGBB hex string."""
    r, g, b = int_to_rgb(color_int)
    return f"#{r:02X}{g:02X}{b:02X}"


# ---------------------------------------------------------------------------
# macOS
# ---------------------------------------------------------------------------
# Apple Events take and return a three element list of 0 to 255 integers rather
# than one packed number. Passing a 16 bit value gets clamped to 255, so the
# scale really is 8 bit. Everything above this layer keeps speaking #RRGGBB.


def hex_to_rgb_list(hex_str: str) -> list[int]:
    """Convert a hex colour string to the [R, G, B] list Apple Events take."""
    return list(hex_to_rgb(hex_str))


def rgb_list_to_hex(rgb) -> str:
    """Convert an Apple Event [R, G, B] list to a #RRGGBB string.

    PowerPoint answers ``missing value`` rather than a colour for a shape that
    has none, which arrives as None. That is not an error, so it comes back as
    None here too and the caller decides what to say about it.
    """
    if not rgb or len(rgb) < 3:
        return None
    r, g, b = (int(c) for c in rgb[:3])
    return f"#{r:02X}{g:02X}{b:02X}"


def rgb_list_to_int(rgb) -> int:
    """Convert an Apple Event [R, G, B] list to the BGR integer Windows uses.

    Tools that report a colour numerically keep doing so identically on both
    platforms, so a deck inspected on a Mac reads the same as on Windows.
    """
    if not rgb or len(rgb) < 3:
        return None
    r, g, b = (int(c) for c in rgb[:3])
    return rgb_to_int(r, g, b)


def get_theme_color_index(name: str) -> int:
    """Convert a theme color name to its MsoThemeColorIndex constant value.

    Args:
        name: Theme color name (case-insensitive), e.g. "accent1", "dark1"

    Raises:
        ValueError: If name is not a valid theme color
    """
    key = name.lower().replace(" ", "_").replace("-", "_")
    if key not in THEME_COLOR_MAP:
        raise ValueError(
            f"Unknown theme color: '{name}'. "
            f"Valid names: {list(THEME_COLOR_MAP.keys())}"
        )
    return THEME_COLOR_MAP[key]
