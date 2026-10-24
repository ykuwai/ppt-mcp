"""The names a font file answers to, read with the standard library.

PowerPoint for Mac carries the Office fonts inside its own bundle, where the
system's font list does not see them, and either system's list gives a font
by one of its names, so the others are read from the files.
"""

import os
import struct

# The name table's family, full name and typographic family records.
_NAME_IDS = (1, 4, 16)


def font_names(data: bytes) -> set:
    """Every family and full name in a TrueType or OpenType file, or a
    collection of them. An unreadable file answers an empty set."""
    return _names(lambda at, length: data[at:at + length])


def file_names(handle) -> set:
    """font_names for an open file, reading only its directory and name
    table: a folder of fonts runs to hundreds of megabytes."""
    def read(at, length):
        handle.seek(at)
        return handle.read(length)
    return _names(read)


def _names(read):
    names = set()
    try:
        head = read(0, 12)
        if head[:4] == b"ttcf":
            count = struct.unpack(">I", head[8:12])[0]
            offsets = struct.unpack(f">{count}I", read(12, 4 * count))
        else:
            offsets = (0,)
        for offset in offsets:
            names |= _names_at(read, offset)
    except (struct.error, IndexError, UnicodeDecodeError, ValueError):
        pass
    return names


def _names_at(read, offset):
    tables = struct.unpack(">H", read(offset + 4, 2))[0]
    directory = read(offset + 12, 16 * tables)
    for i in range(tables):
        entry = directory[16 * i:16 * i + 16]
        if entry[:4] != b"name":
            continue
        start, size = struct.unpack(">II", entry[8:16])
        table = read(start, size)
        count, strings = struct.unpack(">HH", table[2:6])
        names = set()
        for n in range(count):
            platform, _, _, name_id, length, at = struct.unpack(
                ">HHHHHH", table[6 + 12 * n:18 + 12 * n])
            if name_id not in _NAME_IDS:
                continue
            raw = table[strings + at:strings + at + length]
            text = raw.decode("utf-16-be") if platform in (0, 3) else raw.decode("mac_roman")
            if text.strip():
                names.add(text.strip())
        return names
    return set()


def names_in_folder(folder) -> set:
    """The names of every font file in a folder, not its subfolders."""
    names = set()
    try:
        entries = os.listdir(folder)
    except OSError:
        return names
    for entry in entries:
        if not entry.lower().endswith((".ttf", ".otf", ".ttc", ".otc")):
            continue
        try:
            with open(os.path.join(folder, entry), "rb") as handle:
                names |= file_names(handle)
        except OSError:
            continue
    return names
