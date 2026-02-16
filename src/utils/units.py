"""Unit conversion utilities for PowerPoint COM automation.

PowerPoint COM uses points as the native unit for all positions and sizes.
1 inch = 72 points, 1 cm ≈ 28.35 points, 1 point = 12700 EMU.
"""

POINTS_PER_INCH = 72.0
CM_PER_INCH = 2.54
POINTS_PER_CM = POINTS_PER_INCH / CM_PER_INCH
EMU_PER_POINT = 12700
EMU_PER_INCH = 914400
EMU_PER_CM = 360000

# Standard slide sizes in points
SLIDE_WIDTH_16_9 = 960.0
SLIDE_HEIGHT_16_9 = 540.0
SLIDE_WIDTH_4_3 = 720.0
SLIDE_HEIGHT_4_3 = 540.0


def inches_to_points(inches: float) -> float:
    """Convert inches to points. 1 inch = 72 points."""
    return inches * POINTS_PER_INCH


def points_to_inches(points: float) -> float:
    """Convert points to inches."""
    return points / POINTS_PER_INCH


def cm_to_points(cm: float) -> float:
    """Convert centimeters to points."""
    return cm * POINTS_PER_CM


def points_to_cm(points: float) -> float:
    """Convert points to centimeters."""
    return points / POINTS_PER_CM


def emu_to_points(emu: int) -> float:
    """Convert EMU (English Metric Units) to points."""
    return emu / EMU_PER_POINT


def points_to_emu(points: float) -> int:
    """Convert points to EMU."""
    return int(round(points * EMU_PER_POINT))


def inches_to_emu(inches: float) -> int:
    """Convert inches to EMU."""
    return int(round(inches * EMU_PER_INCH))


def emu_to_inches(emu: int) -> float:
    """Convert EMU to inches."""
    return emu / EMU_PER_INCH


def cm_to_emu(cm: float) -> int:
    """Convert centimeters to EMU."""
    return int(round(cm * EMU_PER_CM))


def emu_to_cm(emu: int) -> float:
    """Convert EMU to centimeters."""
    return emu / EMU_PER_CM
