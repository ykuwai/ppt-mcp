"""Shared validation helpers for MCP tool functions."""

# An instruction rather than a preference, and no longer than it has to be.
#
# "Below the recommended minimum" was read as a note and argued with: callers
# shrink text to make a layout fit, then answer the warning by saying the size
# was deliberate. "You should" says what to do instead.
#
# Length matters. This fires once per undersized write and a deck can draw it
# dozens of times, so the sizing table stays in the server instructions rather
# than being repeated here.
FONT_SIZE_WARNING = (
    "Warning: font_size {size}pt is too small. You should use 16pt or more, "
    "20pt+ for body text."
)


def font_size_warning(font_size: float | None) -> str | None:
    """Return a warning string if font_size is too small, else None."""
    if font_size is not None and font_size < 16:
        return FONT_SIZE_WARNING.format(size=font_size)
    return None
