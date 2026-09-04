"""Application-level tools, on Apple Events.

Mirrors ``ppt_com/app.py``. Same function names, same signatures, same returned
shapes; what differs is the walk through PowerPoint's object model.

Two things do not carry over from Windows and are handled here rather than
being pretended away. PowerPoint for Mac has no ``Application.Visible`` and no
headless mode, so ``visible`` is reported as always true. And the application
window state is not scriptable, so ``ppt_set_window_state`` refuses instead of
quietly doing nothing.
"""

import logging
from typing import Optional

from backend.mac_ae import count, elements, is_missing, ppt, shapes_of

logger = logging.getLogger(__name__)

# The Apple Event selection enumerators, keyed the way ppt_com/app.py keys the
# Windows ones so both platforms return the same words.
_SELECTION_NAMES = {
    "selection type none": "none",
    "selection type slides": "slides",
    "selection type shapes": "shapes",
    "selection type text": "text",
}


def _keyword_name(value) -> str:
    """Return an appscript keyword's own name, as it reads in the dictionary."""
    return str(value).replace("k.", "").replace("_", " ")


def _connect_impl(visible: Optional[bool]) -> dict:
    # Same contract as Windows. This is the user's explicit request to start or
    # attach PowerPoint, so launching is allowed.
    app = ppt._connect_impl(visible, allow_launch=True)
    return {
        "success": True,
        "name": app.name(),
        "version": app.version(),
        # PowerPoint for Mac has no Visible property and no headless mode. It
        # is on screen whenever it is running, so saying anything else here
        # would be a lie the caller acts on.
        "visible": True,
        "presentations_count": count(app.presentations),
    }


def _get_app_info_impl() -> dict:
    app = ppt._get_app_impl()
    info = {
        "name": app.name(),
        "version": app.version(),
        "visible": True,
        # No scriptable application window state on macOS; see
        # _set_window_state_impl.
        "window_state": "unknown",
        "presentations_count": count(app.presentations),
        "active_presentation": None,
    }
    if info["presentations_count"] > 0:
        try:
            info["active_presentation"] = ppt._get_pres_impl().name()
        except Exception:
            pass
    return info


def _get_active_window_info_impl() -> dict:
    app = ppt._get_app_impl()
    windows = elements(app.document_windows)
    if not windows:
        return {"error": "No windows are open in PowerPoint."}

    win = app.active_window
    result = {
        "caption": None,
        "view_type": None,
        "active_slide_index": None,
        "selection_type": "none",
        "selected_shapes": [],
        "selected_text": None,
    }

    try:
        result["caption"] = win.caption()
    except Exception:
        result["caption"] = windows[0].caption()
        win = windows[0]

    try:
        result["view_type"] = _keyword_name(win.view_type())
    except Exception:
        pass

    try:
        result["active_slide_index"] = win.view.slide.slide_index()
    except Exception:
        pass

    try:
        selection = win.selection
        name = _keyword_name(selection.selection_type())
        result["selection_type"] = _SELECTION_NAMES.get(name, "unknown")

        if result["selection_type"] == "shapes":
            shapes = []
            try:
                # `shapes_of`, not `elements(...shapes)`. PowerPoint addresses
                # the shapes of a range by subclass, so a selection holding a
                # text box and an autoshape answers `text_boxes[1]` and
                # `shapes[2]`, and the second does not resolve. Counting and
                # then indexing is the route the rest of the port takes.
                for shape in shapes_of(selection.shape_range):
                    shapes.append({
                        "name": shape.name(),
                        "type": _keyword_name(shape.shape_type()),
                        # macOS exposes no shape id, only stacking order, and a
                        # made-up id would be worse than an honest null.
                        "id": None,
                        "z_order_position": shape.z_order_position(),
                    })
            except Exception as exc:
                # An empty list here used to read as an empty selection, which
                # is a different answer from PowerPoint refusing to describe
                # one, and only the second is worth a caller's attention.
                logger.warning("Could not read the selected shapes: %s", exc)
                result["warnings"] = [
                    "Shapes are selected and PowerPoint would not say which "
                    f"ones ({exc}), so selected_shapes is empty because the "
                    "question could not be answered rather than because "
                    "nothing is selected."
                ]
            result["selected_shapes"] = shapes

        elif result["selection_type"] == "text":
            content = selection.text_range.content()
            result["selected_text"] = None if is_missing(content) else content
    except Exception:
        pass

    return result


def _list_presentations_impl() -> dict:
    app = ppt._get_app_impl()
    presentations = []
    for index, pres in enumerate(elements(app.presentations), start=1):
        path = pres.path()
        presentations.append({
            "index": index,
            "name": pres.name(),
            "full_name": pres.full_name(),
            # An unsaved deck answers `missing value` for its path rather than
            # an empty string, which would otherwise reach the caller as the
            # literal text "k.missing_value".
            "path": None if is_missing(path) else path,
            "slides_count": count(pres.slides),
            "read_only": bool(pres.read_only()),
            "saved": bool(pres.saved()),
        })
    return {"presentations": presentations, "count": len(presentations)}


def _set_window_state_impl(window_state: str) -> dict:
    """Refuse, because macOS has no scriptable application window state.

    The Windows side sets ``Application.WindowState``. PowerPoint for Mac has
    no such property; its ``document window`` carries position and size, but
    nothing that maximises or minimises the application, and there is no
    single application window to act on in the first place.
    """
    return {
        "error": "ppt_set_window_state is not available on macOS",
        "reason": (
            "PowerPoint for Mac exposes no application window state to Apple "
            "Events. Individual presentation windows can be moved and resized, "
            "but there is nothing to maximise or minimise."
        ),
        "platform": "macOS",
    }
