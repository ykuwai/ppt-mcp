"""Utility modules for ppt-mcp."""

from .units import (
    inches_to_points,
    points_to_inches,
    cm_to_points,
    points_to_cm,
    emu_to_points,
    points_to_emu,
)
from .color import (
    rgb_to_int,
    int_to_rgb,
    hex_to_int,
    int_to_hex,
    hex_to_rgb,
)
from .com_wrapper import PowerPointCOMWrapper, handle_com_error
