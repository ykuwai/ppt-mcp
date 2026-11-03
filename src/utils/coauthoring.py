"""Presentations that other people edit at the same time.

A deck opened from SharePoint or OneDrive (its FullName is a URL) can be
edited by several people at once, in desktop PowerPoint or in PowerPoint for
the web.  Without AutoSave, PowerPoint takes in what the others have saved
only when the deck is saved; until then it shows "Updates available", and
everything read or written through COM is this copy's own, older state.
Measured against SharePoint Server 2016 with PowerPoint for the web as the
second author:

* Presentation.Coauthoring.PendingUpdates is what the status bar shows.  It
  turns true only once PowerPoint has talked to the server, typically while
  uploading after a save; a window in the background does not poll.  Save()
  returns at once and the upload runs afterwards.
* Save() while PendingUpdates is true merges the other authors' changes
  synchronously.  SlideIDs stay, slide indices can shift.
* When both sides changed the same object, PendingUpdates stays true after
  Save(), the file is not uploaded, and PowerPoint opens a modal "Resolve
  conflict" dialog (window class NUIDialog, owned by the deck's window).
  COM keeps working and shows this copy's version.  Coauthoring.MergeMode
  stays false and FavorServerEditsDuringMerge changes nothing, so the object
  model offers no way to answer the dialog.

What this module does with that, for a deck on a server only:

* before a tool call, merge pending changes with Save() when nothing in this
  copy is unsaved, so that a read sees the server's state;
* refuse a writing call whose slides or shapes the merge moved, because its
  indices were read from the older state (MovedByMerge, nothing changed);
* report the state in the call's result under "coauthoring": pending changes
  that still wait (something here is unsaved), a merge that just happened, or
  a conflict;
* on a conflict, hand over to a resolver registered with
  set_conflict_resolver() if there is one, and report which slides then look
  different from this copy's version; otherwise leave the dialog to the user
  and stop saving that deck until it is answered.

server.py marks each presentation-bound tool call (call_coauthoring) and the
COM wrapper runs the call's operations through bind_call_coauthoring() on the
worker thread, so the merge, the operation and the report happen in one go.
"""

from __future__ import annotations

import functools
import logging
import threading
from contextvars import ContextVar
from dataclasses import dataclass, field
from typing import Any, Callable, Optional

logger = logging.getLogger(__name__)

FIELD = "coauthoring"

PENDING_NOTE = (
    "Other authors have saved changes to this file on the server that are not "
    "in this copy yet (PowerPoint shows 'Updates available'). They come in "
    "with the next ppt_save_presentation, which also saves the changes made "
    "here; re-read the slides you work on afterwards (slide indices can "
    "shift, slide IDs stay)."
)
MERGED_NOTE = (
    "Changes other authors had saved on the server were merged into this copy "
    "first; slide indices may differ from earlier reads."
)
CONFLICT_OPEN_NOTE = (
    "Merging the other authors' changes hit a conflict: someone else changed "
    "the same object. PowerPoint asks in a 'Resolve conflict' dialog which "
    "version to keep and does not upload the file until it is answered. Ask "
    "the user to resolve it in the PowerPoint window, then re-read the slides."
)
CONFLICT_RESOLVED_NOTE = (
    "Merging the other authors' changes hit a conflict and it was resolved. "
    "'overwritten' lists each slide that now differs from this copy's version, "
    "with the changed shapes and the text they had here. Re-apply what is "
    "still wanted on the current slides, or ask the user."
)
MOVED_MESSAGE = (
    "Changes other authors had saved were merged before this call, and they "
    "moved slides or shapes, so the indices of this call may point elsewhere. "
    "Nothing was changed by this call. Re-read the slides and repeat it."
)


class MovedByMerge(RuntimeError):
    """A merge before a writing call moved slides or shapes."""

    def __init__(self, state: dict):
        super().__init__(MOVED_MESSAGE)
        self.state = state


# ---------------------------------------------------------------------------
# Conflict resolver hook
# ---------------------------------------------------------------------------
ConflictResolver = Callable[[Any], bool]
_resolver: Optional[ConflictResolver] = None


def set_conflict_resolver(resolver: Optional[ConflictResolver]) -> None:
    """Register what answers PowerPoint's conflict dialog, or None for the user.

    The resolver is called on the COM worker thread with the Presentation
    after a Save() that left PendingUpdates true, while the dialog is open or
    about to open.  It returns True once the dialog is answered (for example
    through UI Automation); the deck is then saved once more, and the result
    lists in "overwritten" each slide that now differs from this copy's
    version.  False or an exception leaves the dialog to the user.

    It runs inside the call's COM budget, so it must finish in well under
    30 seconds.
    """
    global _resolver
    _resolver = resolver


def conflict_resolver() -> Optional[ConflictResolver]:
    return _resolver


# Decks (FullName, lower case) whose conflict dialog waits for the user: no
# further Save() from a tool call until nothing is pending any more.
_waiting_for_user: set[str] = set()
_waiting_lock = threading.Lock()


def _key(pres) -> str:
    return str(pres.FullName).lower()


def _mark_waiting(pres, waiting: bool) -> None:
    with _waiting_lock:
        if waiting:
            _waiting_for_user.add(_key(pres))
        else:
            _waiting_for_user.discard(_key(pres))


def _is_waiting(pres) -> bool:
    with _waiting_lock:
        return _key(pres) in _waiting_for_user


# ---------------------------------------------------------------------------
# State of a deck
# ---------------------------------------------------------------------------
def on_server(pres) -> bool:
    """True for a deck opened from a server (SharePoint, OneDrive)."""
    try:
        return str(pres.FullName).lower().startswith(("http://", "https://"))
    except Exception:
        return False


def status(pres, with_count: bool = False) -> Optional[dict]:
    """The co-authoring state of a deck on a server, None for a local file.

    with_count adds coauthor_count, PowerPoint's CoauthorCount: the authors
    with the file open including this one.  It can lag until the next sync.
    """
    if not on_server(pres):
        return None
    coauthoring = pres.Coauthoring
    state: dict = {"pending_updates": bool(coauthoring.PendingUpdates)}
    if with_count:
        state["coauthor_count"] = int(coauthoring.CoauthorCount)
    if state["pending_updates"]:
        state["note"] = (CONFLICT_OPEN_NOTE if _is_waiting(pres) else PENDING_NOTE)
        if _is_waiting(pres):
            state["conflict"] = True
    else:
        _mark_waiting(pres, False)
    return state


def _shape_state(shape) -> tuple:
    text = ""
    try:
        if shape.HasTextFrame and shape.TextFrame.HasText:
            text = shape.TextFrame.TextRange.Text
    except Exception:
        pass
    try:
        place = tuple(round(float(getattr(shape, k)), 1)
                      for k in ("Left", "Top", "Width", "Height"))
    except Exception:
        place = ()
    return (shape.Name, shape.Type, place, text)


def _slide_title(slide) -> str:
    try:
        if slide.Shapes.HasTitle:
            return slide.Shapes.Title.TextFrame.TextRange.Text
    except Exception:
        pass
    return ""


def snapshot(pres) -> dict:
    """SlideID -> index, title and the state of each shape."""
    return {
        slide.SlideID: {
            "index": slide.SlideIndex,
            "title": _slide_title(slide),
            "shapes": tuple(_shape_state(s) for s in slide.Shapes),
        }
        for slide in pres.Slides
    }


def overwritten(before: dict, after: dict) -> list:
    """Slides that differ from this copy's version after a conflict, with the
    changed shapes and the text they had here."""
    entries = []
    for slide_id, old in before.items():
        new = after.get(slide_id)
        if new is not None and new["shapes"] == old["shapes"]:
            continue
        old_shapes = {s[0]: s for s in old["shapes"]}
        new_shapes = {s[0]: s for s in new["shapes"]} if new else {}
        changed = sorted(n for n in set(old_shapes) | set(new_shapes)
                         if old_shapes.get(n) != new_shapes.get(n))
        entry = {
            "slide_id": slide_id,
            "slide_index": new["index"] if new else None,
            "title_here": old["title"],
            "changed_shapes": changed,
            "text_here": {n: old_shapes[n][3] for n in changed
                          if n in old_shapes and old_shapes[n][3]},
        }
        if new is None:
            entry["removed_by_others"] = True
        entries.append(entry)
    return entries


def addresses(pres) -> list:
    """Slide order and, per slide, the order of its shapes."""
    return [(slide.SlideID, tuple(s.Name for s in slide.Shapes)) for slide in pres.Slides]


def moved(before: list, after: list) -> bool:
    """Did a slide that was there before get another index, or other or
    differently ordered shapes?  Slides added at the end move nothing."""
    now = {slide_id: (i, shapes) for i, (slide_id, shapes) in enumerate(after)}
    return any(now.get(slide_id) != (i, shapes)
               for i, (slide_id, shapes) in enumerate(before))


# ---------------------------------------------------------------------------
# Merging
# ---------------------------------------------------------------------------
def after_save(pres, pending_before: bool) -> dict:
    """The state after a Save(): merged, in conflict, or unchanged."""
    if not pres.Coauthoring.PendingUpdates:
        _mark_waiting(pres, False)
        if pending_before:
            return {"pending_updates": False, "merged_updates": True, "note": MERGED_NOTE}
        return {"pending_updates": False}
    if not pending_before:
        # The upload after this save may only now have met the others' changes.
        return status(pres)
    return _resolve_conflict(pres)


def _resolve_conflict(pres) -> dict:
    resolver = _resolver
    if resolver is not None:
        before = snapshot(pres)
        try:
            answered = bool(resolver(pres))
        except Exception:
            logger.exception("Co-authoring conflict resolver failed")
            answered = False
        if answered:
            pres.Save()  # upload, and take in whatever arrived meanwhile
            pending = bool(pres.Coauthoring.PendingUpdates)
            _mark_waiting(pres, False)
            return {
                "pending_updates": pending,
                "merged_updates": True,
                "conflict": True,
                "overwritten": overwritten(before, snapshot(pres)),
                "note": CONFLICT_RESOLVED_NOTE + (" " + PENDING_NOTE if pending else ""),
            }
    _mark_waiting(pres, True)
    return {"pending_updates": True, "conflict": True, "note": CONFLICT_OPEN_NOTE}


def before_call(pres, writes: bool) -> Optional[dict]:
    """Merge pending changes before a tool call when nothing here is unsaved.

    Returns the state to report, or None for a local file.  Raises
    MovedByMerge when the call writes and the merge moved slides or shapes.
    """
    state = status(pres)
    if not state or not state["pending_updates"] or state.get("conflict"):
        return state
    if not pres.Saved:
        return state
    before = addresses(pres) if writes else None
    pres.Save()
    state = after_save(pres, pending_before=True)
    if writes and moved(before, addresses(pres)):
        state["moved"] = True
        raise MovedByMerge(state)
    return state


# ---------------------------------------------------------------------------
# Binding to a tool call
# ---------------------------------------------------------------------------
@dataclass
class CallCoauthoring:
    """What server.py knows about a presentation-bound tool call.

    writes: the tool changes the deck (its readOnlyHint is not true).
    presentation_name / presentation_index: the tool's own target fields
    (ppt_get_presentation_info, ppt_copy_slide's destination ...), which win
    over the `presentation` argument and the session target.
    """

    writes: bool
    presentation_name: Optional[str] = None
    presentation_index: Optional[int] = None
    # Filled on the worker: what an earlier operation of the same call merged.
    reported: dict = field(default_factory=dict)


call_coauthoring: ContextVar = ContextVar("call_coauthoring", default=None)


def _target(wrapper, app, call: CallCoauthoring):
    if call.presentation_index is not None:
        return app.Presentations(call.presentation_index)
    if call.presentation_name is not None:
        matches = [app.Presentations(i) for i in range(1, app.Presentations.Count + 1)
                   if app.Presentations(i).Name == call.presentation_name]
        return matches[0] if len(matches) == 1 else None
    if app.Presentations.Count == 0:
        return None
    return wrapper._get_pres_impl()


def _report(result: Any, state: Optional[dict]) -> Any:
    if not isinstance(result, dict) or not state:
        return result
    existing = result.get(FIELD)
    result[FIELD] = {**existing, **state} if isinstance(existing, dict) else dict(state)
    return result


def bind_call_coauthoring(wrapper, call: CallCoauthoring, func: Callable) -> Callable:
    """Wrap func so that it runs between a merge and a report (see module doc).

    Finding the deck or reading its state never fails the operation: the
    operation then runs as if the deck were local.  MovedByMerge does.
    """
    @functools.wraps(func)
    def bound(*args, **kwargs):
        pres = None
        state = None
        try:
            pres = _target(wrapper, wrapper._get_app_impl(), call)
            if pres is not None and on_server(pres):
                state = before_call(pres, call.writes)
        except MovedByMerge:
            raise
        except Exception as exc:
            logger.debug("co-authoring check before the call skipped: %s", exc)
            pres = None
        if state and (state.get("merged_updates") or state.get("conflict")):
            call.reported.update(state)

        result = func(*args, **kwargs)

        try:
            after = status(pres) if pres is not None and call.writes else state
        except Exception as exc:
            logger.debug("co-authoring check after the call skipped: %s", exc)
            after = state
        if call.reported:
            after = {**call.reported,
                     "pending_updates": bool(after and after["pending_updates"])}
        if after and (after["pending_updates"] or after.get("merged_updates")
                      or after.get("conflict")):
            result = _report(result, after)
        return result

    return bound
