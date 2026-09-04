"""Editing operation tools, on Apple Events.

Mirrors ``ppt_com/edit_ops.py``. Same function names, same signatures, same
returned shapes; what differs is the walk through PowerPoint's object model.

Four things about this group are worth knowing before reading on.

**Undo and redo are one command each and they take a count.** Windows has to
loop, asking ``GetEnabledMso("Undo")`` before every step, which is how it knows
how many actions it really undid. Here ``undo`` is a command on the
presentation with a ``times`` parameter, so the whole request is one Apple
Event. Nothing in the dictionary reports how many of them landed, and no
property reads back the undo stack, so ``actions_undone`` is the number that
was asked for and a warning says so. It is the one place in this module where
reading a change back has no route.

**A paste lands wherever the window is looking.** ``paste object`` takes a
``view`` and puts the clipboard on the slide that view is showing, so
navigating to the destination slide is the mechanism and not a courtesy. That
one navigation is the reason ``ppt_copy_shape_to_slide`` does not use
``goto_slide``, which swallows its own failures; a move that quietly did not
happen would drop the copy on the wrong slide and then be reported as nothing
having happened at all. The round trip goes through the system clipboard and
replaces whatever the user had copied, which is exactly what the Windows tool
does too.

**Format painting works in full.** ``pick up`` and ``apply`` each accept a
plain ``shape`` as well as a shape range, so unlike the tools in
``ppt_mac/layout.py`` this one needs no selection and refuses nothing.

**Two tools have no route at all.** There is no counterpart to
``StartNewUndoEntry``, and ``execute`` here runs a command bar control, which
is a different thing from an MSO command id. Both refuse and say why.
"""

import logging

from appscript.reference import CommandError

from backend.mac_ae import count, error_number, ppt, shapes_of
from ppt_mac.shapes import _get_shape, _slide
from utils.navigation import goto_slide

logger = logging.getLogger(__name__)


def _refusal(tool_name: str, reason: str, alternatives=None, error=None) -> dict:
    """The body a tool returns when macOS genuinely cannot do it.

    ``backend.unsupported.unsupported`` builds the same payload but returns it
    already encoded, and these functions hand a dict back to a caller that
    encodes it. Same keys, same reading, one less round of JSON.

    ``error`` overrides the headline for a tool that does work but has one
    argument it cannot honour, so a reader is not told to give up on the whole
    tool when only that argument has to go.
    """
    payload = {
        "error": error or f"{tool_name} is not available on macOS",
        "reason": reason,
        "platform": "macOS",
    }
    if alternatives:
        payload["alternatives"] = alternatives
    return payload


def _count_warning(verb: str, times: int) -> str:
    """The warning both counting tools carry.

    Written once because it is the same sentence twice. It is a warning rather
    than a refusal because the work itself does happen and only the count is
    beyond reach.
    """
    return (
        f"PowerPoint for Mac takes the number of actions to {verb} as one "
        "command and answers nothing about how many of them it carried out, "
        "and no property reads the undo stack back. So this number is what "
        f"was asked for rather than a measurement, and fewer than {times} may "
        "have happened."
    )


# ---------------------------------------------------------------------------
# Apple Event implementation functions
# ---------------------------------------------------------------------------
def _undo_impl(times):
    ppt._get_app_impl()
    pres = ppt._get_pres_impl()

    # One event for the whole request. `times` is optional in the dictionary
    # and defaults to one, and it is passed explicitly so the two platforms
    # agree on what a missing argument means.
    pres.undo(times=times)

    return {
        "success": True,
        "actions_undone": times,
        "warnings": [_count_warning("undo", times)],
    }


def _redo_impl(times):
    ppt._get_app_impl()
    pres = ppt._get_pres_impl()

    pres.redo(times=times)

    return {
        "success": True,
        "actions_redone": times,
        "warnings": [_count_warning("redo", times)],
    }


def _copy_shape_to_slide_impl(src_slide_index, shape_name_or_index, dst_slide_index):
    # No `goto_slide` here. The navigation this tool needs is the paste target
    # and it has to be able to fail, so it is sent below rather than through
    # the shared helper.
    ppt._get_app_impl()
    pres = ppt._get_pres_impl()

    src_slide = _slide(pres, src_slide_index)
    shape = _get_shape(src_slide, shape_name_or_index)
    dst_slide = _slide(pres, dst_slide_index)

    before = count(dst_slide.shapes)

    # `copy shape` puts the shape on the system clipboard, replacing whatever
    # was there. The Windows tool does the same, so this is parity rather than
    # a reason to refuse.
    shape.copy_shape()

    # Load bearing, not a courtesy. `paste object` takes a view and pastes onto
    # the slide that view is showing, so the window has to be on the
    # destination slide before the paste rather than after it. This is the same
    # Apple Event `goto_slide` sends, written out here because that helper
    # swallows its own failures, and a navigation that quietly did not happen
    # would drop the copy on whichever slide the window was still showing.
    try:
        pres.document_windows[1].view.go_to_slide(number=dst_slide_index)
    except CommandError as exc:
        return _refusal(
            "ppt_copy_shape_to_slide",
            f"PowerPoint answered {error_number(exc)} when asked to show slide "
            f"{dst_slide_index}. `paste object` takes a view and pastes onto "
            "the slide that view is showing, so nothing was pasted rather than "
            "risk landing the copy on whatever slide the window was on.",
            ["ppt_add_shape", "ppt_duplicate_slide"],
        )

    pres.document_windows[1].view.paste_object()

    # Nothing is trusted because it did not raise. `paste object` declares no
    # result at all, so the shape count is the only evidence there is.
    shapes_now = shapes_of(dst_slide)
    if len(shapes_now) != before + 1:
        return _refusal(
            "ppt_copy_shape_to_slide",
            f"PowerPoint reported success and slide {dst_slide_index} still "
            f"has {len(shapes_now)} shapes, which is the silent no-op recorded "
            "in MACOS_PORT section 5. The clipboard round trip is the only "
            "route to a copy here, and it can be lost to whatever else has "
            "written to the clipboard in between.",
            ["ppt_add_shape", "ppt_duplicate_slide"],
        )

    # The pasted shape lands at the end of the z order, and it is fetched
    # through a reference this module built rather than through anything
    # PowerPoint handed back.
    return {
        "success": True,
        "new_shape_name": shapes_now[-1].name(),
        "destination_slide": dst_slide_index,
    }


def _copy_formatting_impl(slide_index, source_shape, target_shapes):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)

    # Every shape is resolved before any formatting is applied, so a name that
    # is not on the slide stops the whole call instead of leaving half the
    # targets painted.
    src = _get_shape(slide, source_shape)
    targets = [_get_shape(slide, target_id) for target_id in target_shapes]

    goto_slide(app, slide_index)

    # `pick up` and `apply` both accept a plain shape here, so no selection and
    # no shape range is needed, which is what stops this tool sharing the fate
    # of the ones in ppt_mac/layout.py.
    src.pick_up()

    applied_to = []
    for target in targets:
        target.apply()
        applied_to.append(target.name())

    return {
        "success": True,
        "source": src.name(),
        "applied_to": applied_to,
    }


def _start_undo_entry_impl():
    """Refuse, because nothing here groups several changes into one undo step.

    Windows calls ``Application.StartNewUndoEntry``. PowerPoint for Mac's
    dictionary has ``undo`` and ``redo`` and nothing that opens an undo entry,
    so there is no way to make the next batch of edits collapse into a single
    step.
    """
    return _refusal(
        "ppt_start_undo_entry",
        "PowerPoint for Mac exposes `undo` and `redo` to Apple Events and "
        "nothing that starts an undo entry, so a batch of edits cannot be "
        "grouped into one undo step. Undoing a batch means undoing each edit, "
        "which ppt_undo does in one call because it takes a count.",
        ["ppt_undo with times set to the number of edits"],
    )


def _execute_mso_impl(command_name, check_enabled):
    """Refuse, because an MSO command id has nothing to resolve against here.

    Windows drives ``CommandBars.ExecuteMso("SelectAll")``, where the string is
    an idMso, a stable identifier that is the same in every language. The
    ``execute`` command in this dictionary runs a ``command bar control``,
    which is found by its caption, is translated with the interface, and
    carries no idMso at all. There is no ``GetEnabledMso`` either, so
    ``check_enabled`` has nothing to ask.
    """
    return _refusal(
        "ppt_execute_mso",
        "PowerPoint for Mac has no ExecuteMso. Its `execute` command runs a "
        "command bar control, which is addressed by the caption shown in the "
        "interface and translated with it, so an idMso such as 'SelectAll' or "
        "'SlideShowFromBeginning' names nothing. There is no GetEnabledMso "
        "either, so check_enabled cannot be answered.",
        ["ppt_undo", "ppt_redo", "ppt_slideshow_start"],
    )
