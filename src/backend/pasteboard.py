"""The general NSPasteboard, read and written, and nothing else.

macOS only, AppKit only. No Apple Event is sent from here, which is the point:
``ppt_mac/gvml_paste.py`` calls this from the same worker job that sends the
paste, so the write and the paste cannot be separated by another job (design
section 2). Keeping this module free of PowerPoint means it can be faked in a
test with a dict.

What is here was measured before it was written (docs/gvml-design.md section
0 and 3). Reading fifteen types off a PowerPoint copy took 39 ms. Writing
every one of them back except PowerPoint's own in-process pointer types,
``com.microsoft.PowerPoint-12.0-Internal-*`` and ``com.microsoft.ole.source.*``,
gave a pasteboard PowerPoint could paste from again. The pointer types are
skipped because writing an 8 byte pointer back into a process that has since
freed it is a way to crash PowerPoint, and no reason to find out whether it
would.

Threads: every call here runs on the Apple Event worker thread, not the main
thread. NSPasteboard is documented thread safe, and the pastes measured for
the design were all sent from a worker without an autorelease pool and landed,
so none is set up here.
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional

from AppKit import NSData, NSPasteboard

# Types that are pointers into PowerPoint's own process and must not be written
# back from ours.
INTERNAL_PREFIXES = (
    "com.microsoft.PowerPoint-12.0-Internal-",
    "com.microsoft.ole.source.",
)

# How much of another application's clipboard this will read to save it. Past
# this the clipboard is left alone and the caller says so. Other applications
# promise data they render on demand, and reading a 200 MB image out of one so
# that a chart can be pasted is not a fair trade (design section 3).
SNAPSHOT_LIMIT = 20 * 1024 * 1024


@dataclass
class Snapshot:
    """What the pasteboard held, so it can be put back.

    ``kept`` is False when the read stopped at ``SNAPSHOT_LIMIT``, in which
    case ``types`` is partial and ``restore`` will refuse to write it.
    ``skipped`` names the types deliberately left out.
    """

    types: Dict[str, bytes] = field(default_factory=dict)
    change_count: int = 0
    skipped: List[str] = field(default_factory=list)
    kept: bool = True


def _board():
    return NSPasteboard.generalPasteboard()


def change_count() -> int:
    """The pasteboard's change count, which moves on every write by anyone."""
    return int(_board().changeCount())


def read(uti: str) -> Optional[bytes]:
    """The bytes of one type, or None when the pasteboard has no such type."""
    data = _board().dataForType_(uti)
    return None if data is None else bytes(data)


def write(uti: str, data: bytes) -> int:
    """Replace the pasteboard with one type and return the new change count.

    The count is what the caller compares against just before pasting: if it
    has moved, something else wrote in between and the paste is not sent.
    """
    board = _board()
    board.clearContents()
    ns_data = NSData.dataWithBytes_length_(data, len(data))
    if not board.setData_forType_(ns_data, uti):
        raise RuntimeError(f"NSPasteboard declined to take {len(data)} bytes of {uti}")
    return int(board.changeCount())


def snapshot(limit_bytes: int = SNAPSHOT_LIMIT) -> Snapshot:
    """Read every restorable type off the pasteboard.

    A type whose owner has quit answers nil and is skipped. The internal
    PowerPoint types are skipped by name. The total read stops at
    ``limit_bytes``, marking the snapshot as not kept.
    """
    board = _board()
    snap = Snapshot(change_count=int(board.changeCount()))
    total = 0
    for uti in list(board.types() or []):
        uti = str(uti)
        if uti.startswith(INTERNAL_PREFIXES):
            snap.skipped.append(uti)
            continue
        data = board.dataForType_(uti)
        if data is None:
            snap.skipped.append(uti)
            continue
        raw = bytes(data)
        total += len(raw)
        if total > limit_bytes:
            snap.kept = False
            snap.types = {}
            break
        snap.types[uti] = raw
    return snap


def restore(snap: Snapshot) -> bool:
    """Write a snapshot back. False when it was not kept, so nothing is written.

    Whether the pasteboard still holds what we wrote, and so whether writing
    over it is right, is the caller's decision; it holds the change count of
    its own write and this module does not.
    """
    if not snap.kept:
        return False
    board = _board()
    board.clearContents()
    for uti, raw in snap.types.items():
        board.setData_forType_(NSData.dataWithBytes_length_(raw, len(raw)), uti)
    return True
