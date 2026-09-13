"""The one way a GVML package gets onto a slide, and the one way one comes off.

Every tool that reaches PowerPoint through the clipboard goes through here:
``ppt_add_chart``, ``ppt_build_freeform`` and ``ppt_group_shapes`` paste a
package, ``ppt_get_chart_data``, ``ppt_get_shape_nodes`` and
``ppt_get_group_items`` copy one off. What they share is the procedure and its
checks, which are the whole reason the route is trustworthy (docs/gvml-design.md
sections 2, 3 and 8):

1. The user's clipboard is saved before anything is written and put back
   afterwards, unless something else wrote to it in between.
2. The package is checked by ``gvml`` before it is written, because
   PowerPoint checks nothing and says nothing.
3. The write, the navigation, the paste and the count happen inside one
   worker job, so no other job can put its own copy on the clipboard between
   our write and our paste. Just before the paste the change count is read
   again, and if it has moved since our write nothing is pasted.
4. The selection is cleared before the paste. A paste leaves what it pasted
   selected, and when that selection is a chart the next paste goes inside
   the chart rather than onto the slide.
5. The slide's shape count decides whether the paste happened. One new shape
   of the expected type is a success; none is the silent no-op MACOS_PORT
   section 5 describes; anything else is removed before the refusal, so what
   we placed by mistake does not stay on the slide.
6. The position is written afterwards, because a paste lands in the middle
   of the view and drifts on each repeat, and is read back.

The tools that edit a chart or a path in place have no way to do so, and go
through ``replace_shape``: copy the shape, rewrite its package, paste the
result beside the original, read it back, and only then delete the original
and put the new shape where the old one was in z order. What that costs is
said in ``warnings`` rather than hidden (design section 4): the shape is a
new object with the old name, and its animations, which the package does not
carry, are gone with the original.

None of the functions here end in ``_impl``. The tools' implementations live
in ``charts.py``, ``freeform.py`` and ``groups.py`` and call in.
"""

import logging
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import Callable, List, Optional

from appscript import k
from appscript.reference import CommandError

from backend import pasteboard
from backend.mac_ae import (
    count,
    elements,
    error_number,
    is_missing,
    shapes_of,
    target_window,
)
from backend.unsupported import refusal as _refusal
from gvml import GVML_UTI, Package, PackageError, validate

logger = logging.getLogger(__name__)

# How far a written position may differ from the one read back before it is
# reported. PowerPoint answers in single precision floats.
_POSITION_TOLERANCE = 0.05


class Refused(Exception):
    """A refusal, raised so the shared procedure can stop where it is.

    Carries the dict every tool returns on failure; the tool's ``_impl``
    catches this and returns ``payload``.
    """

    def __init__(self, payload: dict):
        super().__init__(payload.get("error", "refused"))
        self.payload = payload


class Clipboard:
    """The user's clipboard, saved on entry and restored on the way out.

    One per tool call, whatever the tool writes or copies in between (a group
    of five copies five times and restores once). ``claim`` records the change
    count after a write PowerPoint made on our behalf, ``copy shape``, so the
    restore afterwards knows the pasteboard is still ours.
    """

    def __init__(self, snapshot: pasteboard.Snapshot):
        self.snapshot = snapshot
        self.mine: Optional[int] = None
        self.warnings: List[str] = []

    @classmethod
    def take(cls) -> "Clipboard":
        return cls(pasteboard.snapshot())

    def claim(self) -> None:
        """The pasteboard was just written on our behalf; remember its count."""
        self.mine = pasteboard.change_count()

    def write(self, raw: bytes) -> int:
        self.mine = pasteboard.write(GVML_UTI, raw)
        return self.mine

    def restore(self) -> None:
        """Put the user's clipboard back, if it is still ours to overwrite."""
        if self.mine is None:
            # Nothing was written, so there is nothing to put back.
            return
        if not self.snapshot.kept:
            self.warnings.append(
                "The clipboard held more than "
                f"{pasteboard.SNAPSHOT_LIMIT // (1024 * 1024)} MB before this "
                "call, so it was not saved and could not be put back. It now "
                "holds what this tool pasted."
            )
            return
        if pasteboard.change_count() != self.mine:
            self.warnings.append(
                "Something else wrote to the clipboard while this tool was "
                "running, so what was there before it was not put back, to "
                "keep the newer copy."
            )
            return
        try:
            pasteboard.restore(self.snapshot)
        except Exception as exc:  # noqa: BLE001 - reported, never fatal
            logger.warning("Could not restore the clipboard: %s", exc)
            self.warnings.append(
                "What the clipboard held before this call could not be put back."
            )


def with_warnings(result: dict, clip: Clipboard, extra: Optional[List[str]] = None) -> dict:
    """Attach the clipboard's warnings, and any others, to a success."""
    warnings = list(result.get("warnings", [])) + list(extra or []) + clip.warnings
    if warnings:
        result["warnings"] = warnings
    return result


def unused_name(names: List[str], prefix: str) -> str:
    """``Chart 3`` when ``Chart 1`` and ``Chart 2`` are taken, Windows style."""
    n = 1
    while f"{prefix} {n}" in names:
        n += 1
    return f"{prefix} {n}"


# ---------------------------------------------------------------------------
# Copying a package off the slide
# ---------------------------------------------------------------------------

def copy_shape_package(shape, clip: Clipboard, tool_name: str) -> Package:
    """``copy shape``, then read the GVML package it left on the pasteboard.

    Raises ``Refused`` when there is no package to read, which is the only
    thing that can go wrong short of PowerPoint not answering.
    """
    shape.copy_shape()
    clip.claim()
    raw = pasteboard.read(GVML_UTI)
    if raw is None:
        raise Refused(_refusal(
            tool_name,
            f"`copy shape` on '{shape.name()}' put nothing of type {GVML_UTI} "
            "on the clipboard, and that package is the only route to what is "
            "inside the shape. PowerPoint writes it for every shape it can "
            "copy, so either the copy was refused without an error or "
            "something else replaced the clipboard in the same instant.",
            ["Try the call again", "ppt_get_shape_info"],
        ))
    try:
        return Package.from_bytes(raw)
    except PackageError as exc:
        raise Refused(_refusal(
            tool_name,
            f"The clipboard package PowerPoint wrote for '{shape.name()}' "
            f"could not be read: {exc}",
            ["ppt_get_shape_info"],
        )) from exc


# ---------------------------------------------------------------------------
# Pasting a package onto the slide
# ---------------------------------------------------------------------------

@dataclass
class Pasted:
    """What landed: the reference to it, its name, and anything to report."""

    shape: object
    name: str
    warnings: List[str] = field(default_factory=list)


def _keyword_name(word) -> str:
    return getattr(word, "AS_name", None) or str(word)


def _remove_landed(slide, before: int) -> List[str]:
    """Delete whatever the paste added beyond ``before``; return what remains."""
    while count(slide.shapes) > before:
        index = count(slide.shapes)
        try:
            slide.shapes[index].delete()
        except CommandError as exc:
            logger.warning("Could not remove shape %d after a bad paste: %s", index, exc)
            break
        if count(slide.shapes) >= index:
            break
    return [s.name() for s in shapes_of(slide)[before:]]


def paste_package(
    pres, slide, slide_index: int, raw: bytes, clip: Clipboard, tool_name: str,
    expected_type, left: float, top: float, alternatives: Optional[List[str]] = None,
) -> Pasted:
    """Write ``raw`` to the pasteboard, paste it onto ``slide``, and check.

    ``expected_type`` is the ``shape type`` keyword the new shape must report.
    ``left`` and ``top`` are written after the paste and read back. Raises
    ``Refused`` at every point where nothing, or the wrong thing, landed.
    """
    alternatives = alternatives or []

    # Our own check first. A package that fails it never reaches PowerPoint,
    # and the refusal says whose mistake it is.
    try:
        validate(raw)
    except PackageError as exc:
        raise Refused(_refusal(
            tool_name,
            f"The package this tool assembled did not pass its own check ({exc}). "
            "PowerPoint was not touched. This is a bug in ppt-mcp rather than a "
            "limit of PowerPoint for Mac; please report it with the arguments used.",
            alternatives,
            error=f"{tool_name} built a package it could not trust",
        )) from exc

    before = count(slide.shapes)
    mine = clip.write(raw)

    # Load bearing, not a courtesy: `paste object` takes a view and pastes
    # onto the slide that view is showing. `goto_slide` swallows its own
    # failures, so the navigation is sent here where it can refuse.
    window = target_window(pres)
    try:
        window.view.go_to_slide(number=slide_index)
    except CommandError as exc:
        raise Refused(_refusal(
            tool_name,
            f"PowerPoint answered {error_number(exc)} when asked to show slide "
            f"{slide_index}. `paste object` pastes onto the slide the window is "
            "showing, so nothing was pasted rather than risk landing it on "
            "whatever slide the window was on.",
            alternatives,
        )) from exc

    # A paste leaves its result selected, and a chart selection swallows the
    # next paste into the chart's own shapes. Clearing costs one event.
    try:
        window.selection.unselect()
    except CommandError as exc:
        logger.debug("Could not clear the selection before pasting: %s", exc)

    # The server cannot serialise the user. If anything wrote to the clipboard
    # between our write and here, pasting would put that on the slide.
    if pasteboard.change_count() != mine:
        raise Refused(_refusal(
            tool_name,
            "The clipboard was written to by something else between this "
            "tool's write and its paste, so nothing was pasted rather than "
            "paste whatever is there now. Nothing on the slide changed.",
            ["Call the tool again"],
        ))

    window.view.paste_object()

    # `paste object` declares no result, so the slide is the only evidence.
    after = shapes_of(slide)
    landed = len(after) - before
    if landed == 0:
        raise Refused(_refusal(
            tool_name,
            f"`paste object` returned no error and slide {slide_index} still "
            f"has {before} shape(s), so PowerPoint discarded the package "
            "without saying so. That is the silent no-op recorded in "
            "MACOS_PORT section 5, and it is what PowerPoint does with a "
            "package it cannot use.",
            alternatives,
        ))
    if landed != 1:
        remaining = _remove_landed(slide, before)
        note = (
            f" The extra shapes could not all be removed; still on the slide: "
            f"{remaining}." if remaining else " They were removed again."
        )
        raise Refused(_refusal(
            tool_name,
            f"The paste put {landed} shapes on slide {slide_index} where one "
            f"was expected.{note}",
            alternatives,
        ))

    shape = after[-1]
    actual_type = shape.shape_type()
    if actual_type != expected_type:
        remaining = _remove_landed(slide, before)
        note = (
            f" It could not be removed and is still on the slide as {remaining}."
            if remaining else " It was removed again."
        )
        raise Refused(_refusal(
            tool_name,
            f"The paste landed as '{_keyword_name(actual_type)}' rather than "
            f"'{_keyword_name(expected_type)}'.{note}",
            alternatives,
        ))
    shape = shapes_of(slide)[-1]
    name = shape.name()

    # The position is not honoured by the paste, which drops the shape in the
    # middle of the view and further down and right on each repeat, so it is
    # written now and read back.
    warnings: List[str] = []
    try:
        shape.left_position.set(left)
        shape.top.set(top)
        got_left, got_top = shape.left_position(), shape.top()
    except CommandError as exc:
        warnings.append(
            f"'{name}' was pasted but its position could not be written "
            f"({error_number(exc)}); it is wherever PowerPoint put it."
        )
    else:
        if abs(got_left - left) > _POSITION_TOLERANCE or abs(got_top - top) > _POSITION_TOLERANCE:
            warnings.append(
                f"'{name}' was asked to sit at ({left}, {top}) and reads back "
                f"at ({round(got_left, 2)}, {round(got_top, 2)})."
            )
    return Pasted(shape, name, warnings)


# ---------------------------------------------------------------------------
# Replacing a shape with a rewritten copy of itself
# ---------------------------------------------------------------------------

# Said in full the first time and in one line after that, the way the line
# visibility warning in shapes.py is. A caller editing twenty nodes should
# not read the paragraph twenty times, and the short form stands on its own
# because the server outlives any one conversation.
_RECREATED_LONG = (
    "'{name}' was recreated rather than edited in place. PowerPoint for Mac "
    "cannot reach inside a {what} from a script, so the shape was copied, its "
    "XML rewritten and pasted back, and the original deleted. Its name, "
    "position and z order were put back. Its animations were not, because "
    "the clipboard package does not carry them, and anything else that "
    "pointed at the old object (a comment anchor, an animation trigger on "
    "another shape) now points at nothing."
)
_RECREATED_SHORT = (
    "'{name}' was recreated by paste; name, position and z order kept, "
    "animations not."
)
_recreation_explained = False


def _recreated_warning(name: str, what: str) -> str:
    global _recreation_explained
    if _recreation_explained:
        return _RECREATED_SHORT.format(name=name)
    _recreation_explained = True
    return _RECREATED_LONG.format(name=name, what=what)


def effects_on(slide, name: str) -> Optional[int]:
    """How many effects of the slide's main sequence belong to a shape.

    None when the sequence cannot be read, which the caller reports as
    "could not be counted" rather than as zero. The count is asked of the
    sequence, never of its effects collection; see ``mac_ae.count_of``.
    """
    try:
        from ppt_mac.animation import _effect_count, _effect_shape_names, _main_sequence

        seq = _main_sequence(slide)
        total = _effect_count(seq)
        if not total:
            return 0
        return sum(1 for n in _effect_shape_names(seq, total) if n == name)
    except Exception as exc:  # noqa: BLE001 - reported as unknown, never fatal
        logger.debug("Could not count the effects on '%s': %s", name, exc)
        return None


@dataclass
class Replaced:
    """What stands where the original stood: its reference, its name, the
    package it reads back as, and everything to tell the caller."""

    shape: object
    name: str
    package: Package
    warnings: List[str] = field(default_factory=list)


def _names(slide) -> List[str]:
    """Every shape name on the slide in z order, in one Apple Event."""
    return ["" if is_missing(n) else n for n in elements(slide.shapes.name)]


def _send_backward(slide, index: int, target: int) -> None:
    """Move the shape at ``index`` down to ``target`` one step at a time.

    One Apple Event per step. The reference is re-made at each step because
    ``slide.shapes[i]`` is positional and resolves afresh on every send, so
    sending the same reference twice moves two different shapes.
    """
    for i in range(index, target, -1):
        slide.shapes[i].z_order(z_order_position=k.send_shape_backward)


def replace_shape(
    pres, slide, slide_index: int, original_index: int, raw: bytes, clip: Clipboard,
    tool_name: str, expected_type, left: float, top: float, what: str,
    verify: Callable[[Package], Optional[str]],
    alternatives: Optional[List[str]] = None,
) -> Replaced:
    """Paste ``raw``, check it, then delete the original and restore its place.

    ``original_index`` is the original's 1-based position on the slide, which
    is also its z order. ``verify`` is handed the package the new shape reads
    back as (one more ``copy shape``) and returns a description of what is
    missing from it, or None when the edit is there; a description removes
    the new shape and refuses with the original untouched.

    The order is the one that leaves nothing half done (design section 4):
    paste, verify, delete. A paste that is dropped leaves the original alone.
    A delete that fails leaves two shapes with one name, and says so.
    """
    alternatives = alternatives or []
    names_before = _names(slide)
    name = names_before[original_index - 1]
    effects = effects_on(slide, name)

    pasted = paste_package(
        pres, slide, slide_index, raw, clip, tool_name, expected_type, left, top, alternatives,
    )
    new_index = count(slide.shapes)

    # Read the new shape back before the original goes. The paste is the
    # only evidence there is (MACOS_PORT section 5), and the paste says
    # nothing, so what the shape reads back as is the check.
    readback = copy_shape_package(slide.shapes[new_index], clip, tool_name)
    problem = verify(readback)
    if problem is not None:
        remaining = _remove_landed(slide, new_index - 1)
        note = (
            f" The pasted copy could not be removed and is still on the slide as {remaining}."
            if remaining else " The pasted copy was removed again."
        )
        raise Refused(_refusal(
            tool_name,
            f"The rewritten '{name}' pasted as a {what}, but read back without "
            f"the change: {problem}.{note} The original is untouched.",
            alternatives,
            error=f"{tool_name} pasted a {what} that did not carry the edit",
        ))

    try:
        slide.shapes[original_index].delete()
    except CommandError as exc:
        logger.warning("Could not delete '%s' after replacing it: %s", name, exc)
    names_now = _names(slide)
    if len(names_now) != len(names_before):
        raise Refused(_refusal(
            tool_name,
            f"The rewritten '{name}' was pasted onto slide {slide_index} and "
            f"verified, but the original could not be deleted afterwards, so "
            f"two shapes named '{name}' are on the slide: the original at "
            f"position {original_index} and the new one at position "
            f"{new_index}. Delete one with ppt_delete_shape by index.",
            ["ppt_delete_shape"],
            error=f"{tool_name} left both the original and its replacement on the slide",
        ))

    warnings = [_recreated_warning(name, what)]
    if effects is None:
        warnings.append(
            f"Whether '{name}' had animations could not be read, so any it "
            "had are gone with the original."
        )
    elif effects:
        warnings.append(
            f"{effects} animation effect(s) on '{name}' were lost with the "
            "original; the clipboard package does not carry them."
        )
    warnings.extend(pasted.warnings)

    # The new shape is last. Walk it back to where the original was.
    last = len(names_now)
    try:
        _send_backward(slide, last, original_index)
    except CommandError as exc:
        logger.warning("Could not restore the z order of '%s': %s", name, exc)
    names_after = _names(slide)
    if names_after != names_before:
        at = names_after.index(name) + 1 if name in names_after else None
        warnings.append(
            f"'{name}' could not be put back at z order position "
            f"{original_index}; it is at position {at} now."
        )
        shape = slide.shapes[at] if at else slide.shapes[last]
    else:
        shape = slide.shapes[original_index]
    return Replaced(shape, name, readback, warnings)


def strip_ids(package: Package) -> None:
    """Drop the creation ids from a package about to be pasted beside its
    original, in place. See ``gvml.shapes.strip_creation_ids``."""
    from gvml import canvas as _canvas
    from gvml import shapes as _shapes

    root = ET.fromstring(package.drawing())
    for element in _canvas.children(_canvas.canvas_of(root)):
        _shapes.strip_creation_ids(element)
    package.parts[package.drawing_part()] = _canvas.serialize_drawing(root).encode("utf-8")
