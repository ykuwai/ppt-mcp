"""Slide-level operations, on Apple Events.

Mirrors ``ppt_com/slides.py``. Same function names, same signatures, same
returned shapes; what differs is the walk through PowerPoint's object model.

Four things about PowerPoint for Mac shape everything below.

**Slides cannot be addressed by name.** The dictionary declares no ``name`` on
``slide``, on ``custom layout`` or on ``design``, so the fields Windows fills
with "Slide 3" or "Title and Content" come back as null here. An invented value
would read as real and be wrong, which is worse than an honest null. The one
exception is a built-in layout ``ppt_add_slide`` was given by name, where the
name is what the caller passed rather than anything read back.

**There is no FindBySlideID.** Following a slide across a move means reading
every slide ID and looking the wanted one up. That is one Apple Event for the
whole deck, so it is cheap, and it is what ``_slide_ids`` does.

**Slides are only ever moved backwards.** ``move ... to before slide N`` reads
N against the deck as it stands before the move, so a forward move would land
one place short. Every reorder here instead builds the deck left to right,
where the slide wanted next is always at or after the position it is going to.
See ``_realize_order``.

**Nothing is trusted because it did not raise.** Every mutation is read back.
"""

import logging
from typing import Optional

from appscript import k

from backend.mac_ae import (
    count,
    elements,
    full_names as _full_names,
    is_missing,
    osascript,
    positional,
    ppt,
    resolve_presentation as _resolve_presentation,
    target_window,
    windows_constant,
)
from backend.mac_enums import (
    PpEntryEffect,
    PpPlaceholderType,
    PpSlideLayout,
    to_keyword,
)
from ppt_com.constants import (
    ppLayoutBlank,
    ppPlaceholderCenterTitle,
    ppPlaceholderTitle,
    ppPlaceholderVerticalTitle,
)
from utils.color import hex_to_rgb_list
from utils.navigation import goto_slide as nav_goto_slide

logger = logging.getLogger(__name__)

# The placeholder that carries speaker notes, in both object models. Windows
# reaches it as NotesPage.Shapes.Placeholders(2).
_NOTES_PLACEHOLDER = 2

# Placeholder types that count as a slide title, so `has_title` answers what
# Windows' Shapes.HasTitle answers. There is no HasTitle here. Translated once
# at import, so a table that ever stopped carrying one of them fails on load
# rather than turning ppt_get_slide_info into an error at the worst moment.
_TITLE_PLACEHOLDERS = frozenset(
    to_keyword(PpPlaceholderType, value, "placeholder type")
    for value in (ppPlaceholderTitle, ppPlaceholderCenterTitle,
                  ppPlaceholderVerticalTitle)
)


# ---------------------------------------------------------------------------
# Small helpers over the object model
# ---------------------------------------------------------------------------
def _presentation_index(app, pres) -> int:
    """The 1-based position of a presentation among the open ones."""
    full_name = pres.full_name()
    names = _full_names(app)
    if full_name not in names:
        raise RuntimeError(
            "The presentation is no longer open in PowerPoint."
        )
    return names.index(full_name) + 1


def _slide_count(pres) -> int:
    """How many slides the deck has.

    Named rather than inlined because several callers here take a parameter
    called ``count``, which would otherwise shadow the helper of that name.
    """
    return count(pres.slides)


def _shape_count(container) -> int:
    """How many shapes a slide or a custom layout holds.

    Named for the same reason as ``_slide_count``, so that a caller taking a
    parameter called ``count`` can still reach it.
    """
    return count(container.shapes)


def _slide_ids(pres) -> list:
    """Every slide ID in order, in one Apple Event.

    This is the stand-in for FindBySlideID. Reading the whole list and
    indexing into it costs one event, where following a slide by walking the
    deck would cost one per slide.
    """
    return [int(value) for value in elements(pres.slides.slide_ID)]


def _windows_constant(table: dict, keyword):
    """``windows_constant`` with the missing value folded into the same answer.

    The fields this module reads come back as ``missing value`` when PowerPoint
    has nothing to report, and a caller here wants None for that as much as it
    wants None for a keyword the table does not carry. A plausible wrong
    constant is the failure this whole layer exists to avoid, so neither case
    reaches for the nearest match.
    """
    if is_missing(keyword):
        return None
    return windows_constant(table, keyword)


def _custom_layout_name(layout_ref):
    """A custom layout's name, or None when PowerPoint will not say.

    ``custom layout`` declares no ``name`` in PowerPoint for Mac's dictionary,
    so this normally answers nothing at all. It is still asked, so that a
    PowerPoint which grows the property starts working with no code change,
    and a missing value counts as absent rather than as a match.
    """
    try:
        value = layout_ref.name()
    except Exception:
        return None
    if is_missing(value):
        return None
    text = str(value).strip()
    return text or None


def _notes_text_range(slide):
    """The text range holding a slide's speaker notes.

    ``notes page`` is itself a slide here, so the placeholder walk is the same
    one Windows makes through NotesPage.Shapes.Placeholders(2).
    """
    return (
        slide.notes_page
        .place_holders[_NOTES_PLACEHOLDER]
        .text_frame
        .text_range
    )


def _realize_order(pres, final_ids: list) -> None:
    """Put the deck's slides into the given order, and prove that it took.

    Every move made here is backwards, from a higher index to a lower one.
    Working left to right guarantees it. Once positions 1 to f-1 hold their
    final slide, the slide wanted at f is at some index at or after f. Placing
    a slide further back still works, because the slides that belong ahead of
    it move past it one at a time.

    The order is simulated locally as the moves go out, so only one extra
    Apple Event is spent, on the check at the end.
    """
    ids = _slide_ids(pres)
    if len(ids) != len(final_ids):
        raise RuntimeError(
            "The deck changed while its slides were being reordered "
            f"({len(ids)} slides, expected {len(final_ids)})."
        )

    for position in range(1, len(final_ids) + 1):
        wanted = final_ids[position - 1]
        current = ids.index(wanted) + 1
        if current == position:
            continue
        pres.slides[current].move(to=pres.slides[position].before)
        ids.insert(position - 1, ids.pop(current - 1))

    if _slide_ids(pres) != list(final_ids):
        raise RuntimeError(
            "PowerPoint accepted the moves but the slides did not end up in "
            "the requested order."
        )


def _duplicate_once(app, pres, slide_index: int) -> int:
    """Duplicate one slide and return the new slide's ID.

    ``ref.duplicate()`` raises "unpack requires a buffer of 4 bytes" here,
    a struct error rather than an Apple Event one, because appscript falls back
    to a code PowerPoint does not answer. The AppleScript form does work, so
    that is what runs, and the copy is then found by diffing the slide IDs
    rather than by trusting a return value.
    """
    before = _slide_ids(pres)
    pres_index = _presentation_index(app, pres)
    osascript(
        'tell application "Microsoft PowerPoint" to duplicate '
        "slide %d of presentation %d" % (slide_index, pres_index)
    )
    after = _slide_ids(pres)
    fresh = [sid for sid in after if sid not in set(before)]
    if len(fresh) != 1:
        raise RuntimeError(
            "PowerPoint reported no error but did not duplicate slide "
            f"{slide_index} (the deck went from {len(before)} to "
            f"{len(after)} slides)."
        )
    return fresh[0]


def _friendly_layout(key: str) -> Optional[int]:
    """Look a friendly layout name up in the table ppt_com/slides.py owns.

    Imported when called rather than at module scope. ppt_com/slides.py imports
    this module from its own last line, so a module level import here would run
    while that module is still half built.
    """
    from ppt_com.slides import LAYOUT_NAME_MAP

    return LAYOUT_NAME_MAP.get(key)


# ---------------------------------------------------------------------------
# Helper to resolve a presentation
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# Implementation functions (run on the Apple Event thread via ppt.execute)
# ---------------------------------------------------------------------------
def _add_slide_impl(
    position: Optional[int],
    layout: Optional[int],
    layout_name: Optional[str],
    design_index: Optional[int] = None,
    count: int = 1,
    like_slide_index: Optional[int] = None,
) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(app)
    total = _slide_count(pres)

    if position is None:
        position = total + 1

    if position < 1 or position > total + 1:
        raise ValueError(f"Position {position} out of range (1-{total + 1})")

    # Resolve the layout once, before the loop.
    custom_layout = None
    layout_keyword = None
    resolved_layout_name = None

    if like_slide_index is not None:
        # Highest precedence, and the only way to reach a design's own layouts
        # here. The layout is carried as an object reference, so it needs no
        # name, which is exactly what PowerPoint for Mac declines to give.
        if like_slide_index < 1 or like_slide_index > total:
            raise ValueError(
                f"like_slide_index {like_slide_index} out of range (1-{total})"
            )
        custom_layout = pres.slides[like_slide_index].custom_layout.get()
        resolved_layout_name = _custom_layout_name(
            pres.slides[like_slide_index].custom_layout
        )
    elif layout_name:
        friendly_key = layout_name.lower().strip().replace(" ", "_")
        friendly = _friendly_layout(friendly_key)
        if friendly is not None:
            layout_keyword = to_keyword(PpSlideLayout, friendly, "slide layout")
            # Report the name that was asked for and applied. `custom layout`
            # declares no name here, so reading it back answers nothing, and
            # answering null for a layout this call chose itself told the
            # caller their own argument had been ignored.
            resolved_layout_name = friendly_key
        else:
            designs = elements(pres.designs)
            if design_index is not None and (
                design_index < 1 or design_index > len(designs)
            ):
                raise ValueError(
                    f"Design index {design_index} out of range (1-{len(designs)})"
                )
            # Every route into a design's custom layouts goes through their
            # names, and PowerPoint for Mac publishes none, so the search
            # cannot even be attempted honestly. Say so, and name the two
            # routes that do work.
            raise ValueError(
                f"Layout '{layout_name}' cannot be looked up on macOS. "
                "PowerPoint for Mac's dictionary declares no name on "
                "'custom layout' or on 'design', so custom layouts cannot be "
                "matched by name. Pass like_slide_index to inherit the exact "
                "layout of an existing slide, or use a built-in layout "
                "('blank', 'title', 'title_only', 'section_header', "
                "'comparison', 'content_with_caption', 'picture_with_caption') "
                "or a PpSlideLayout constant."
            )
    else:
        layout_keyword = to_keyword(
            PpSlideLayout,
            layout if layout is not None else ppLayoutBlank,
            "slide layout",
        )

    # Create every slide at the end of the deck, then put the block where it
    # was asked for. `at=pres.end` is the working form; `at=pres.slides.end`
    # raises -1708.
    existing_ids = _slide_ids(pres)
    new_ids = []
    for _ in range(count):
        expected = _slide_count(pres) + 1
        if custom_layout is not None:
            # Windows makes the slide and applies the layout in one call.
            # There is no such form here, so the slide is made and the layout
            # assigned to it, and the placeholders that are supposed to come
            # with it are checked for below.
            created = app.make(new=k.slide, at=pres.end)
            created.custom_layout.set(custom_layout)
        else:
            created = app.make(
                new=k.slide, at=pres.end, with_properties={k.layout: layout_keyword}
            )
        if _slide_count(pres) != expected:
            raise RuntimeError(
                "PowerPoint reported no error but did not add a slide."
            )
        new_ids.append(int(created.slide_ID()))

    final_ids = (
        existing_ids[: position - 1] + new_ids + existing_ids[position - 1:]
    )
    _realize_order(pres, final_ids)

    ids_now = _slide_ids(pres)
    created_slides = [
        {"slide_index": ids_now.index(slide_id) + 1, "slide_id": slide_id}
        for slide_id in new_ids
    ]

    # Navigate to the last created slide.
    nav_goto_slide(app, created_slides[-1]["slide_index"])

    first_slide = pres.slides[created_slides[0]["slide_index"]]
    final_layout_name = resolved_layout_name
    if final_layout_name is None:
        final_layout_name = _custom_layout_name(first_slide.custom_layout)

    # Assigning a layout after the fact is not the same call Windows makes, so
    # check that the layout's placeholders actually came across. A slide that
    # names the right layout and carries none of its boxes would otherwise
    # look like a success.
    warning = None
    if custom_layout is not None and _shape_count(first_slide) == 0:
        if _shape_count(custom_layout) > 0:
            warning = (
                "The layout was applied but the new slide has no shapes, so "
                "its placeholders did not come across. PowerPoint for Mac has "
                "no way to create a slide with a layout in one step. Use "
                "ppt_duplicate_slide to get a copy that keeps the boxes."
            )

    return {
        "success": True,
        "slides_created": len(created_slides),
        "slides": created_slides,
        # "layout" is always the PpSlideLayout integer constant, translated
        # back from the enumerator PowerPoint answers with.
        "layout": _windows_constant(PpSlideLayout, first_slide.layout()),
        "layout_name": final_layout_name,
        # `design` carries no name in the dictionary, and a made-up one would
        # read as real. Null is the honest answer.
        "design_name": None,
        **({"warning": warning} if warning else {}),
        # Backward compatibility: a single slide also reports at the top level.
        **(
            {
                "slide_index": created_slides[0]["slide_index"],
                "slide_id": created_slides[0]["slide_id"],
            }
            if count == 1
            else {}
        ),
    }


def _delete_slide_impl(
    slide_index: Optional[int] = None,
    slide_indices: Optional[list] = None,
    from_index: Optional[int] = None,
    to_index: Optional[int] = None,
) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(app)
    total = _slide_count(pres)

    # Normalise the three input forms into a sorted, de-duplicated index set.
    if slide_index is not None:
        targets = [slide_index]
    elif slide_indices is not None:
        targets = list(slide_indices)
    else:
        targets = list(range(from_index, to_index + 1))

    targets = sorted(set(targets))
    if not targets:
        raise ValueError("No slides to delete")

    out_of_range = [i for i in targets if i < 1 or i > total]
    if out_of_range:
        raise ValueError(f"Slide index(es) {out_of_range} out of range (1-{total})")
    if len(targets) >= total:
        raise ValueError(
            "Cannot delete every slide. A presentation must keep at least "
            f"one slide (requested {len(targets)} of {total})"
        )

    # Navigate to the lowest target before deleting so the editor follows.
    nav_goto_slide(app, targets[0])

    # Delete from the highest index first so earlier indices stay valid.
    for idx in sorted(targets, reverse=True):
        pres.slides[idx].delete()

    remaining = _slide_count(pres)
    if remaining != total - len(targets):
        raise RuntimeError(
            "PowerPoint reported no error but the deck still has "
            f"{remaining} slides, where {total - len(targets)} were expected."
        )

    result = {
        "success": True,
        "deleted_indices": targets,
        "deleted_count": len(targets),
        "remaining_count": remaining,
    }
    # Backward compatibility for single-slide callers.
    if slide_index is not None:
        result["deleted_index"] = slide_index
    return result


def _duplicate_slide_impl(
    slide_index: int,
    insert_at: Optional[int] = None,
    count: int = 1,
) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(app)
    total = _slide_count(pres)

    if slide_index < 1 or slide_index > total:
        raise ValueError(f"Slide index {slide_index} out of range (1-{total})")

    # Validate an explicit insert_at up front (raise rather than silently
    # clamp, consistent with ppt_copy_slide). Up to one past the end is fine.
    if insert_at is not None and insert_at != -1 and insert_at > total + 1:
        raise ValueError(f"insert_at {insert_at} out of range (1-{total + 1})")

    nav_goto_slide(app, slide_index)

    # Track the source by slide ID. Once copies start moving around, its index
    # shifts, so a fixed index is unsafe across iterations.
    src_id = _slide_ids(pres)[slide_index - 1]

    new_ids = []
    for i in range(count):
        ids = _slide_ids(pres)
        cur_src_idx = ids.index(src_id) + 1
        new_id = _duplicate_once(app, pres, cur_src_idx)

        # The copy lands immediately after the source. Resolve where it should
        # finish, then rebuild the whole order rather than nudging it, because
        # only a full left-to-right pass keeps every move backwards.
        ids = _slide_ids(pres)
        if insert_at is None:
            target = cur_src_idx + 1 + i
        elif insert_at == -1:
            target = len(ids)
        else:
            target = min(insert_at + i, len(ids))

        without = [sid for sid in ids if sid != new_id]
        _realize_order(pres, without[: target - 1] + [new_id] + without[target - 1:])
        new_ids.append(new_id)

    ids_now = _slide_ids(pres)
    new_indices = [ids_now.index(sid) + 1 for sid in new_ids]

    nav_goto_slide(app, new_indices[-1])

    result = {
        "success": True,
        "count": len(new_indices),
        "new_slide_indices": new_indices,
        "new_slide_ids": new_ids,
    }
    # Backward compatibility for single-copy callers.
    if count == 1:
        result["new_slide_index"] = new_indices[0]
        result["new_slide_id"] = new_ids[0]
    return result


def _move_slide_impl(
    new_position: int,
    slide_index: Optional[int] = None,
    slide_indices: Optional[list] = None,
) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(app)
    total = _slide_count(pres)

    # Normalise to a sorted, de-duplicated list of source indices.
    sources = [slide_index] if slide_index is not None else list(slide_indices)
    sources = sorted(set(sources))

    out_of_range = [i for i in sources if i < 1 or i > total]
    if out_of_range:
        raise ValueError(f"Slide index(es) {out_of_range} out of range (1-{total})")

    block = len(sources)
    max_start = total - block + 1
    if new_position < 1 or new_position > max_start:
        raise ValueError(
            f"new_position {new_position} out of range (1-{max_start}) "
            f"for a block of {block} slide(s)"
        )

    nav_goto_slide(app, sources[0])

    # Build the full desired final order of slide IDs, then realise it. A
    # direction heuristic based on sources[0] is not enough, because a
    # non-contiguous selection that straddles the target can leave a member
    # outside the block (for example [1,4,5] to position 2).
    # _compute_final_order is the same pure function the Windows side uses,
    # imported when called because ppt_com/slides.py imports this module from
    # its own last line.
    from ppt_com.slides import _compute_final_order

    all_ids = _slide_ids(pres)
    src_ids = [all_ids[i - 1] for i in sources]
    _realize_order(pres, _compute_final_order(all_ids, src_ids, new_position))

    result = {
        "success": True,
        "moved_slide_ids": src_ids,
        "moved_count": block,
        "new_start_position": new_position,
    }
    # Backward compatibility for single-slide callers.
    if slide_index is not None:
        result["moved_from"] = slide_index
        result["moved_to"] = new_position
    return result


def _copy_slide_impl(
    slide_index: Optional[int],
    slide_indices: Optional[list],
    source_presentation_index: Optional[int],
    source_presentation_name: Optional[str],
    to_presentation_index: Optional[int],
    to_presentation_name: Optional[str],
    insert_at: Optional[int],
) -> dict:
    app = ppt._get_app_impl()

    src_pres = _resolve_presentation(
        app,
        presentation_index=source_presentation_index,
        presentation_name=source_presentation_name,
    )
    dst_pres = _resolve_presentation(
        app,
        presentation_index=to_presentation_index,
        presentation_name=to_presentation_name,
    )

    if src_pres.full_name() != dst_pres.full_name():
        # Windows copies between decks through the clipboard. The Mac verbs
        # for that are `copy object` and `paste object`, and `paste object`
        # declares a direct parameter whose type does not resolve to any class
        # in the dictionary, so there is no way to say which deck it should
        # paste into. Guessing would paste into whichever deck PowerPoint
        # happens to consider current, which is a silent write to the wrong
        # file, so this refuses instead.
        return {
            "error": "ppt_copy_slide across presentations is not available on macOS",
            "reason": (
                "PowerPoint for Mac's 'paste object' command takes no "
                "addressable destination, so a slide cannot be pasted into a "
                "chosen presentation. Copying within one presentation works."
            ),
            "platform": "macOS",
            "alternatives": [
                "ppt_copy_slide within a single presentation",
                "ppt_duplicate_slide",
            ],
        }

    sources = [slide_index] if slide_index is not None else list(slide_indices)

    src_total = _slide_count(src_pres)
    out_of_range = [i for i in sources if i < 1 or i > src_total]
    if out_of_range:
        raise ValueError(
            f"Source slide index(es) {out_of_range} out of range (1-{src_total})"
        )

    append = insert_at is None or insert_at == -1
    if not append:
        # Allow inserting anywhere from the front up to one past the end.
        if insert_at < 1 or insert_at > src_total + 1:
            raise ValueError(f"insert_at {insert_at} out of range (1-{src_total + 1})")

    # Within one deck a copy is a duplicate, which is a verified route here.
    src_ids = [_slide_ids(src_pres)[i - 1] for i in sources]

    new_ids = []
    for sid in src_ids:
        ids = _slide_ids(src_pres)
        new_ids.append(_duplicate_once(app, src_pres, ids.index(sid) + 1))

    ids = _slide_ids(src_pres)
    without = [s for s in ids if s not in set(new_ids)]
    target = len(without) + 1 if append else insert_at
    _realize_order(
        src_pres, without[: target - 1] + new_ids + without[target - 1:]
    )

    ids_now = _slide_ids(src_pres)
    new_indices = [ids_now.index(sid) + 1 for sid in new_ids]

    # Navigate the window to the last copy.
    try:
        nav_goto_slide(app, new_indices[-1])
    except Exception:
        pass

    return {
        "success": True,
        "copied_count": len(new_indices),
        "new_slide_indices": new_indices,
        "target_presentation": dst_pres.name(),
        "source_presentation": src_pres.name(),
    }


def _list_slides_impl(
    presentation_index: Optional[int],
    presentation_name: Optional[str],
) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(
        app,
        presentation_index=presentation_index,
        presentation_name=presentation_name,
    )

    slides = []
    for slide in elements(pres.slides):
        has_notes = False
        try:
            notes_text = _notes_text_range(slide).content()
            has_notes = not is_missing(notes_text) and len(notes_text.strip()) > 0
        except Exception:
            pass

        slides.append({
            "index": slide.slide_index(),
            "slide_id": int(slide.slide_ID()),
            # `slide` declares no name in PowerPoint for Mac's dictionary.
            # Windows reports "Slide 3" here; inventing that would read as
            # real and would not survive being passed back in.
            "name": None,
            "layout": _windows_constant(PpSlideLayout, slide.layout()),
            "layout_name": _custom_layout_name(slide.custom_layout),
            "hidden": bool(slide.slide_show_transition.hidden()),
            "shapes_count": count(slide.shapes),
            "has_notes": has_notes,
        })

    return {"slides_count": len(slides), "slides": slides}


def _get_slide_info_impl(slide_index: int) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(app)
    total = _slide_count(pres)

    if slide_index < 1 or slide_index > total:
        raise ValueError(f"Slide index {slide_index} out of range (1-{total})")

    slide = pres.slides[slide_index]
    trans = slide.slide_show_transition

    # There is no Shapes.HasTitle here, so the title is found the way Windows
    # would have to without it, by looking for a title placeholder.
    has_title = False
    title_text = ""
    try:
        for holder in positional(slide.place_holders):
            if holder.placeholder_type() in _TITLE_PLACEHOLDERS:
                has_title = True
                content = holder.text_frame.text_range.content()
                title_text = "" if is_missing(content) else content
                break
    except Exception:
        pass

    notes_text = ""
    try:
        content = _notes_text_range(slide).content()
        notes_text = "" if is_missing(content) else content
    except Exception:
        pass

    return {
        "index": slide.slide_index(),
        "slide_id": int(slide.slide_ID()),
        "slide_number": slide.slide_number(),
        # See _list_slides_impl for why these three are null on macOS.
        "name": None,
        "layout": _windows_constant(PpSlideLayout, slide.layout()),
        "layout_name": _custom_layout_name(slide.custom_layout),
        "hidden": bool(trans.hidden()),
        "shapes_count": count(slide.shapes),
        "has_title": has_title,
        "title_text": title_text,
        "notes_text": notes_text,
        "follow_master_background": bool(slide.follow_master_background()),
        "transition_effect": _windows_constant(PpEntryEffect, trans.entry_effect()),
        "advance_on_click": bool(trans.advance_on_click()),
        "advance_on_time": bool(trans.advance_on_time()),
        "advance_time": trans.advance_time(),
        "design_name": None,
    }


def _set_slide_notes_impl(
    slide_index: int,
    notes_text: Optional[str],
    font_name: Optional[str],
    font_name_fareast: Optional[str],
    font_size: Optional[float],
    bold: Optional[bool],
    italic: Optional[bool],
    color: Optional[str],
) -> dict:
    app = ppt._get_app_impl()
    nav_goto_slide(app, slide_index)
    pres = _resolve_presentation(app)
    total = _slide_count(pres)

    if slide_index < 1 or slide_index > total:
        raise ValueError(f"Slide index {slide_index} out of range (1-{total})")

    text_range = _notes_text_range(pres.slides[slide_index])

    if notes_text is not None:
        text_range.content.set(notes_text)
        written = text_range.content()
        # Compared for presence rather than character by character, because
        # PowerPoint rewrites paragraph breaks as it stores them. Text that
        # went in and came back empty is the failure worth catching.
        if notes_text.strip() and (is_missing(written) or not written.strip()):
            raise RuntimeError(
                "PowerPoint reported no error but the speaker notes are empty."
            )

    # Apply formatting to the whole notes text range.
    font = text_range.font
    if font_name is not None:
        font.font_name.set(font_name)
        if font_name_fareast is None:
            font.east_asian_name.set(font_name)
    if font_name_fareast is not None:
        font.east_asian_name.set(font_name_fareast)
    if font_size is not None:
        font.font_size.set(font_size)
    if bold is not None:
        font.bold.set(bool(bold))
    if italic is not None:
        font.italic.set(bool(italic))
    if color is not None:
        font.font_color.set(hex_to_rgb_list(color))

    return {"success": True}


def _get_slide_notes_impl(slide_index: int) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(app)
    total = _slide_count(pres)

    if slide_index < 1 or slide_index > total:
        raise ValueError(f"Slide index {slide_index} out of range (1-{total})")

    try:
        content = _notes_text_range(pres.slides[slide_index]).content()
        notes_text = "" if is_missing(content) else content
    except Exception:
        notes_text = ""

    return {"slide_index": slide_index, "notes_text": notes_text}


def _goto_slide_impl(slide_index: int) -> dict:
    pres = ppt._get_pres_impl()
    total = _slide_count(pres)
    if slide_index < 1 or slide_index > total:
        raise ValueError(f"Slide index {slide_index} out of range (1-{total})")

    # Showing a slide needs a window, and a deck can outlive its own.
    # `target_window` explains that state and says how to get out of it, which
    # the bare -1728 from `document_windows[1]` does not.
    view = target_window(pres).view
    view.go_to_slide(number=slide_index)

    # Read the window back. A view that did not move is the silent no-op this
    # tool exists to avoid, and it is one cheap event to rule out.
    try:
        landed = view.slide.slide_index()
    except Exception:
        landed = None
    if landed is not None and landed != slide_index:
        raise RuntimeError(
            "PowerPoint accepted the navigation but the window is showing "
            f"slide {landed}, not slide {slide_index}."
        )

    return {
        "success": True,
        "active_slide_index": slide_index,
    }
