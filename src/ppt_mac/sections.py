"""Section tools, on Apple Events.

Mirrors ``ppt_com/sections.py``. Same function names, same signatures, same
returned shapes; what differs is the walk through PowerPoint's object model.

Three things about sections on this side are worth knowing before reading on.

**Sections are worked through commands, not through properties.** That is
unusual for this dictionary. ``presentation`` has a ``section properties``
property whose class is completely empty, no properties and no elements, so it
is never read and only ever handed to a command as its direct parameter.
``get count of sections``, ``get name of section``, ``insert section``,
``rename section``, ``move section`` and ``delete section`` are what a section
actually is here.

**Whether the count answers at all is the open question.** MACOS_PORT section 9
records ``get count of sections`` answering through appscript and failing with
-1708 through AppleScript, which cannot both be true of one PowerPoint. The
appscript form is what runs here, and a failure is turned into a refusal that
names the number rather than reaching the caller as a raw error, so the one
unknown produces one readable answer.

**Nothing is trusted because it did not raise.** Every one of these commands
answers nothing, or answers an integer, so each change is read back. Sections
carry an id through ``get id of section``, which is what a move is checked
against, because two sections are allowed to share a name.
"""

import logging

from appscript.reference import CommandError

from backend.mac_ae import (
    AE_NO_SUCH_OBJECT,
    AE_PARAMETER,
    count,
    error_number,
    is_missing,
    ppt,
)
from backend.unsupported import refusal as _refusal
from utils.navigation import goto_slide

logger = logging.getLogger(__name__)


def _count_sections(sp) -> int:
    """How many sections the deck has.

    A deck with no sections is the ordinary case, and PowerPoint answers -1728
    or ``missing value`` for an empty collection rather than zero, which is the
    same habit ``elements`` absorbs. Anything else is raised, so the caller can
    turn it into a refusal that names the number.
    """
    try:
        total = sp.get_count_of_sections()
    except CommandError as exc:
        if error_number(exc) == AE_NO_SUCH_OBJECT:
            return 0
        raise
    return 0 if is_missing(total) else int(total)


def _unreachable(tool_name: str, exc: CommandError) -> dict:
    """The refusal for a PowerPoint whose sections will not answer at all."""
    return _refusal(
        tool_name,
        f"PowerPoint answered {error_number(exc)} when asked how many sections "
        "the deck has. Sections are reached here only through commands on "
        "`section properties`, so a command that is not handled leaves no "
        "other route to them. MACOS_PORT section 9 records this answer varying "
        "between appscript and AppleScript on the same build.",
        ["ppt_list_slides", "ppt_reorder_slides"],
    )


def _name_of(sp, position: int) -> str:
    """A section's name, with ``missing value`` read as an unnamed section."""
    name = sp.get_name_of_section(at_position=position)
    return "" if is_missing(name) else str(name)


def _id_of(sp, position: int):
    """A section's id, or None when this build will not give one.

    The id is what a move is verified against. It is allowed to be missing,
    because the check falls back to the name when it is.
    """
    try:
        value = sp.get_id_of_section(at_position=position)
    except CommandError:
        return None
    return None if is_missing(value) else str(value)


def _check_position(total: int, position: int, what: str) -> None:
    """Reject a section position before it becomes an unreadable -1728.

    An out of range position does not fail where it is written, it fails later
    with a number and no mention of which argument was wrong, so the range is
    checked here while the number is still in hand.
    """
    if total == 0:
        raise ValueError("The presentation has no sections.")
    if position < 1 or position > total:
        raise ValueError(
            f"{what} {position} is out of range. "
            f"The presentation has {total} sections (1-based)."
        )


def _locate_section(sp, total: int, name: str, slide_index: int):
    """Find the section just inserted, by reading the deck back.

    ``insert section`` answers with an integer and that answer is not taken on
    trust, for the same reason no reference PowerPoint hands back is. The
    section wanted is the one carrying the name that was asked for and starting
    at the slide that was asked for; a name that repeats falls back to the
    first match, which is all there is to go on.
    """
    fallback = None
    for position in range(1, total + 1):
        if _name_of(sp, position) != name:
            continue
        if sp.get_first_slide_of_section(at_position=position) == slide_index:
            return position
        if fallback is None:
            fallback = position
    return fallback


# ---------------------------------------------------------------------------
# Apple Event implementation functions
# ---------------------------------------------------------------------------
def _add_section_impl(name, slide_index):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    sp = pres.section_properties

    total_slides = count(pres.slides)
    if slide_index < 1 or slide_index > total_slides:
        raise ValueError(
            f"Slide index {slide_index} is out of range. "
            f"The presentation has {total_slides} slides (1-based)."
        )

    try:
        before = _count_sections(sp)
    except CommandError as exc:
        return _unreachable("ppt_add_section", exc)

    # After everything that would refuse or raise, so a call that never runs
    # does not move the user's view.
    goto_slide(app, slide_index)

    # `before slide` rather than `before section`. Windows names the slide the
    # section starts at, and this is the parameter that says the same thing;
    # `before section` counts in sections and would land somewhere else.
    returned = sp.insert_section(before_slide=slide_index, titled=name)

    # Not `before + 1`. A section inserted anywhere but the first slide brings
    # a second one with it, because the slides in front of it need a section
    # too, so a deck with none goes straight to two. Measured on a four slide
    # deck, inserting at slide 3 left `既定のセクション` and the new one. The
    # question worth asking is whether the section that was asked for is there.
    after = _count_sections(sp)
    section_index = _locate_section(sp, after, name, slide_index)
    if section_index is None:
        return _refusal(
            "ppt_add_section",
            f"PowerPoint reported success and no section called '{name}' is "
            f"in the deck, which now has {after} sections where it had "
            f"{before}. This is the silent no-op recorded in MACOS_PORT "
            "section 5.",
        )

    return {
        "success": True,
        "section_index": section_index,
        "name": _name_of(sp, section_index),
        # Read back rather than echoed. PowerPoint decides where a section
        # begins once the deck already has sections around it.
        "slide_index": sp.get_first_slide_of_section(at_position=section_index),
    }


def _list_sections_impl():
    ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    sp = pres.section_properties

    try:
        total = _count_sections(sp)
    except CommandError as exc:
        return _unreachable("ppt_list_sections", exc)

    sections = []
    for position in range(1, total + 1):
        sections.append({
            "index": position,
            "name": _name_of(sp, position),
            "first_slide": sp.get_first_slide_of_section(at_position=position),
            "slides_count": sp.get_slide_count_of_section(at_position=position),
        })

    return {
        "success": True,
        "sections_count": total,
        "sections": sections,
    }


def _manage_section_impl(section_index, action, new_name, move_to_index):
    # The arguments are settled before PowerPoint is asked anything, so a call
    # that was never going to run costs no Apple Event and moves nothing.
    action_key = action.strip().lower()
    if action_key not in ("rename", "move", "delete"):
        raise ValueError(
            f"Unknown action '{action}'. Use: 'rename', 'move', or 'delete'"
        )
    if action_key == "rename" and new_name is None:
        raise ValueError("new_name is required for 'rename' action")
    if action_key == "move" and move_to_index is None:
        raise ValueError("move_to_index is required for 'move' action")

    ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    sp = pres.section_properties

    try:
        total = _count_sections(sp)
    except CommandError as exc:
        return _unreachable("ppt_manage_section", exc)

    _check_position(total, section_index, "Section index")
    if action_key == "move":
        _check_position(total, move_to_index, "move_to_index")

    if action_key == "rename":
        sp.rename_section(at_position=section_index, to=new_name)
        landed = _name_of(sp, section_index)
        if landed != new_name:
            return _refusal(
                "ppt_manage_section",
                f"PowerPoint reported success and section {section_index} is "
                f"still called {landed!r}, which is the silent no-op recorded "
                "in MACOS_PORT section 5.",
                error="ppt_manage_section could not rename the section on macOS",
            )
        return {
            "success": True,
            "action": "rename",
            "section_index": section_index,
            "new_name": new_name,
        }

    if action_key == "move":
        # The id rather than the name, because two sections are allowed to
        # share a name and then a name check would pass on the wrong one.
        moving_id = _id_of(sp, section_index)
        moving_name = _name_of(sp, section_index)

        sp.move_section(at_position=section_index, to_position=move_to_index)

        landed_id = _id_of(sp, move_to_index)
        landed_name = _name_of(sp, move_to_index)
        arrived = (
            landed_id == moving_id if moving_id is not None
            else landed_name == moving_name
        )
        if not arrived:
            return _refusal(
                "ppt_manage_section",
                f"PowerPoint reported success and position {move_to_index} "
                f"still holds {landed_name!r} rather than {moving_name!r}, "
                "which is the silent no-op recorded in MACOS_PORT section 5.",
                error="ppt_manage_section could not move the section on macOS",
            )
        return {
            "success": True,
            "action": "move",
            "section_index": section_index,
            "moved_to": move_to_index,
        }

    section_name = _name_of(sp, section_index)
    # `deleting slides` is not optional in the dictionary, and false is what
    # Windows passes, so the slides stay and only the grouping goes.
    try:
        sp.delete_section(at_position=section_index, deleting_slides=False)
    except CommandError as exc:
        if error_number(exc) != AE_PARAMETER or section_index >= total:
            raise
        # Only the first of several answers this. Removing it would leave the
        # slides in front of the next section belonging to nothing, and
        # PowerPoint will not have that. Deleting from the back works, and so
        # does deleting the only section a deck has. Measured both ways.
        return _refusal(
            "ppt_manage_section",
            f"PowerPoint will not delete '{section_name}' while section "
            f"{section_index + 1} is still there, because the slides in front "
            "of that one would then belong to no section. Delete the later "
            "sections first, or rename this one instead.",
            [
                "ppt_manage_section with action='delete' on the last section",
                "ppt_manage_section with action='rename'",
            ],
            error=(
                "ppt_manage_section cannot delete a section that has another "
                "after it on macOS"
            ),
        )

    after = _count_sections(sp)
    if after != total - 1:
        return _refusal(
            "ppt_manage_section",
            f"PowerPoint reported success and the deck still has {after} "
            "sections, which is the silent no-op recorded in MACOS_PORT "
            "section 5.",
            error="ppt_manage_section could not delete the section on macOS",
        )
    return {
        "success": True,
        "action": "delete",
        "deleted_section": section_name,
    }
