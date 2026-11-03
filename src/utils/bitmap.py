"""Pixels out of a BMP file, read with the standard library.

Both platforms export a shape as an uncompressed BMP, which is the one
image format that needs no decoder: the layout checks sample a picture's
colours from it to judge text drawn on a photo.
"""

import struct


def sample_bmp(data: bytes, columns: int = 96):
    """A grid of the image's colours, ``{"columns", "rows", "colours"}``,
    the colours as "#RRGGBB" row by row from the top, or None for a file it
    cannot read. Each sample is one pixel, not an average, so a thin light
    streak in a dark photo is still seen."""
    if data[:2] != b"BM" or len(data) < 54:
        return None
    offset = struct.unpack("<I", data[10:14])[0]
    width, height = struct.unpack("<ii", data[18:26])
    bits, compression = struct.unpack("<H", data[28:30])[0], struct.unpack("<I", data[30:34])[0]
    if bits not in (24, 32) or compression not in (0, 3) or width <= 0 or height == 0:
        return None
    top_down = height < 0
    height = abs(height)
    step = bits // 8
    stride = (width * step + 3) & ~3
    columns = max(1, min(columns, width))
    rows = max(1, min(height, round(columns * height / width)))
    colours = []
    for r in range(rows):
        y = min(height - 1, int((r + 0.5) * height / rows))
        line = y if top_down else height - 1 - y
        base = offset + line * stride
        for c in range(columns):
            x = min(width - 1, int((c + 0.5) * width / columns))
            at = base + x * step
            if at + 3 > len(data):
                return None
            blue, green, red = data[at], data[at + 1], data[at + 2]
            colours.append(f"#{red:02X}{green:02X}{blue:02X}")
    return {"columns": columns, "rows": rows, "colours": colours}
