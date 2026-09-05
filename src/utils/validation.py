"""Shared validation helpers for MCP tool functions."""

# One line, and the same line every time.
#
# A deck laid out with deliberate 15pt table text drew this warning about
# seventy times in one build, and every copy carried three lines of sizing
# advice the caller had already read in the server's instructions. A warning
# that cannot be acted on and cannot be switched off buries the ones that can,
# so the advice stays where it is read once and this says only what happened.
#
# Deduplicating it by wording was the other option and is worse: the flag would
# live as long as the server process, which serves many conversations, so the
# short form would reach callers that never saw the long one. Two wordings for
# one condition also read as two different warnings to anything parsing them.
FONT_SIZE_WARNING = (
    "Warning: font_size {size}pt is below the recommended minimum of 16pt, "
    "which is where text stops being readable when projected."
)


def font_size_warning(font_size: float | None) -> str | None:
    """Return a warning string if font_size is too small, else None."""
    if font_size is not None and font_size < 16:
        return FONT_SIZE_WARNING.format(size=font_size)
    return None
