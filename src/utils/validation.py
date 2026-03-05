"""Shared validation helpers for MCP tool functions."""

FONT_SIZE_WARNING = (
    "Warning: font_size {size}pt is below the recommended minimum of 16pt. "
    "Small text is unreadable when projected. Recommended sizes: "
    "title 40-48pt, heading 24-32pt, body 20-28pt, caption 16-20pt."
)


def font_size_warning(font_size: float | None) -> str | None:
    """Return a warning string if font_size is too small, else None."""
    if font_size is not None and font_size < 16:
        return FONT_SIZE_WARNING.format(size=font_size)
    return None
