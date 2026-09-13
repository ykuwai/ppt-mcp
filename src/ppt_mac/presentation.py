"""Presentation-level operations, on Apple Events.

Mirrors ``ppt_com/presentation.py``. Same function names, same signatures, same
returned shapes; what differs is the walk through PowerPoint's object model.

Three things about PowerPoint for Mac shape everything below.

**A deck with no file path cannot be saved.** ``save`` on an unsaved deck
writes nothing and can hang past forty seconds without reporting anything. So
the path is checked before the save, and everything that would produce a
pathless deck produces a real file instead.

**PowerPoint is sandboxed.** It can write where it has a grant and nowhere
else, and a write to a directory it has never touched blocks for tens of
seconds and then kills the application. There is no way to ask it what it is
allowed to write, so a save is made and then the file is looked for on disk;
when it is not there the error names the sandbox and the one directory that
always works.

**Page setup has a slide width but no slide height.** Height is read from the
slide master and cannot be set at all.
"""

import glob as glob_mod
import logging
import os
import shutil
from typing import Optional

from appscript import k, mactypes

from backend.mac_ae import (
    EXPORT_STAGING_DIR,
    count,
    elements,
    full_names as _full_names,
    is_missing,
    ppt,
    resolve_presentation as _resolve_presentation,
)
from backend.mac_enums import PpSaveAsFileType, to_keyword
from ppt_com.constants import (
    ppSaveAsDefault,
    ppSaveAsJPG,
    ppSaveAsOpenXMLPresentation,
    ppSaveAsPDF,
    ppSaveAsPNG,
)
from utils.color import rgb_list_to_hex
from utils.onedrive import resolve_local_path
from utils.units import (
    SLIDE_HEIGHT_4_3,
    SLIDE_HEIGHT_16_9,
    SLIDE_WIDTH_4_3,
    SLIDE_WIDTH_16_9,
)

logger = logging.getLogger(__name__)

# The same tables ppt_com/presentation.py builds, rebuilt from the same
# sources. They are not imported from there because that module imports this
# one from its own last line, so it is only half built when this one loads.
SAVE_FORMAT_MAP = {
    "pptx": ppSaveAsOpenXMLPresentation,
    "pdf": ppSaveAsPDF,
    "png": ppSaveAsPNG,
    "jpg": ppSaveAsJPG,
    "default": ppSaveAsDefault,
}

SLIDE_SIZE_PRESETS = {
    "16:9": (SLIDE_WIDTH_16_9, SLIDE_HEIGHT_16_9),
    "4:3": (SLIDE_WIDTH_4_3, SLIDE_HEIGHT_4_3),
}

# Whole-deck image export reports success in a fifth of a second and writes no
# folder and no files, measured three separate times. It is refused rather
# than attempted, because a caller that believes the images exist keeps
# building on top of that belief.
_IMAGE_FORMATS = {"png", "jpg", "jpeg"}

# Where Office keeps personal templates on macOS. There is no registry to ask,
# so these are checked in order and the first that exists wins. The localised
# names are what a Japanese or French system actually has on disk.
_TEMPLATE_DIR_CANDIDATES = (
    "~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/"
    "Templates.localized",
    "~/Library/Group Containers/UBF8T346G9.Office/User Content/Templates",
    "~/Library/Application Support/Microsoft/Office/User Templates/My Templates",
)


# ---------------------------------------------------------------------------
# Small helpers over the object model and the filesystem
# ---------------------------------------------------------------------------
def _index_of(app, full_name: str) -> Optional[int]:
    """The 1-based position of a presentation among the open ones."""
    names = _full_names(app)
    return names.index(full_name) + 1 if full_name in names else None


def _local_path(full_name) -> Optional[str]:
    """A presentation's path when it is one a Python process can stat.

    An unsaved deck answers its bare name rather than a path, and that is not
    something to go looking for on disk.
    """
    if is_missing(full_name):
        return None
    text = str(full_name)
    return text if text.startswith("/") else None


def _open_and_find(app, path: str):
    """Open a file and return the presentation that appeared.

    The deck is found by diffing the open presentations before and after
    rather than by matching the path that was asked for. Symlinks and the
    ``/System/Volumes/Data`` prefix make two spellings of the same file look
    different, and that mismatch would be silent.
    """
    before = set(_full_names(app))
    ppt.open_presentation(path)

    presentations = elements(app.presentations)
    names = _full_names(app)
    fresh = [i for i, name in enumerate(names) if name not in before]
    if fresh:
        return presentations[fresh[-1]]

    # Nothing new appeared. Either the file was already open, which is the
    # ordinary idempotent case, or PowerPoint said nothing and did nothing.
    basename = os.path.basename(path)
    for pres in presentations:
        if pres.name() == basename:
            return pres
    raise RuntimeError(
        f"PowerPoint reported no error but did not open {path}."
    )


def _stage_template_copy(template_path: str) -> str:
    """Copy a template into PowerPoint's container and return the copy's path.

    Windows makes an untitled presentation from the template. There is no
    untitled route worth taking here, because a deck with no file path cannot
    be saved through Apple Events at all. Copying the template first and
    opening the copy gives the same deck and a file that can actually be
    saved.

    The copy goes into PowerPoint's own container because that is the one
    directory it is always allowed to read and write, and an unsandboxed
    Python process can reach it freely.
    """
    os.makedirs(EXPORT_STAGING_DIR, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(template_path))
    dest = os.path.join(EXPORT_STAGING_DIR, stem + ext)
    serial = 2
    while os.path.exists(dest):
        dest = os.path.join(EXPORT_STAGING_DIR, f"{stem} {serial}{ext}")
        serial += 1
    shutil.copyfile(template_path, dest)
    if not os.path.exists(dest) or os.path.getsize(dest) == 0:
        raise RuntimeError(f"Could not stage a copy of the template at {dest}")
    return dest


def _presentation_at(app, path: str):
    """The open presentation whose file is ``path``, or None.

    Saving a deck renames it, and every reference held before the save is
    addressed by the old name, so it stops resolving. Scanning is the only way
    back to it, and the file just written is the one certain fact to scan for.
    """
    wanted = os.path.abspath(path)
    for index in range(1, count(app.presentations) + 1):
        candidate = app.presentations[index]
        try:
            full_name = candidate.full_name()
        except Exception:
            continue
        if not is_missing(full_name) and os.path.abspath(str(full_name)) == wanted:
            return candidate
    return None


def _sandbox_error(target: str) -> RuntimeError:
    """The error for a save PowerPoint accepted and did not perform."""
    return RuntimeError(
        f"PowerPoint reported no error but nothing was written to {target}. "
        "PowerPoint for Mac is sandboxed and can only write where it already "
        "has a grant, and it reports success either way. Save into "
        f"{EXPORT_STAGING_DIR}, which it can always write, and move the file "
        "afterwards."
    )


# ---------------------------------------------------------------------------
# Helper to resolve a presentation by index or active
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# Implementation functions (run on the Apple Event thread via ppt.execute)
# ---------------------------------------------------------------------------
def _create_presentation_impl(
    template_path: Optional[str],
    slide_width: Optional[float],
    slide_height: Optional[float],
    preset: Optional[str],
    activate: bool,
) -> dict:
    # Creating a deck legitimately needs PowerPoint, so launch it if it is not
    # already running. It is not brought forward, though. A new window arrives
    # on screen by itself, and taking the front from whatever the user is
    # working in is more than was asked for. `activate` is honoured below only
    # when the caller asked for it.
    app = ppt._get_app_impl(allow_launch=True)

    warnings = []
    target_height = None

    if template_path:
        abs_path = os.path.abspath(os.path.expanduser(template_path))
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"Template not found: {abs_path}")
        staged = _stage_template_copy(abs_path)
        pres = _open_and_find(app, staged)
        warnings.append(
            "This deck is a copy of the template at "
            f"{staged}, not an untitled presentation, because PowerPoint for "
            "Mac cannot save a deck that has no file path. Use "
            "ppt_save_presentation_as to put it where you want it."
        )
    else:
        # The deck is found by diffing the open presentations rather than by
        # keeping what `make` handed back, so a PowerPoint that answers
        # nothing is caught here rather than three lines later.
        before_names = set(_full_names(app))
        app.make(new=k.presentation)
        names_now = _full_names(app)
        fresh = [i for i, name in enumerate(names_now) if name not in before_names]
        if not fresh:
            raise RuntimeError(
                "PowerPoint reported no error but did not create a presentation."
            )
        pres = elements(app.presentations)[fresh[-1]]

        target_width = None
        if preset:
            preset_key = preset.strip()
            if preset_key not in SLIDE_SIZE_PRESETS:
                raise ValueError(
                    f"Unknown preset '{preset}'. "
                    f"Supported: {list(SLIDE_SIZE_PRESETS.keys())}"
                )
            target_width, target_height = SLIDE_SIZE_PRESETS[preset_key]
        elif slide_width is not None and slide_height is not None:
            target_width, target_height = slide_width, slide_height

        if target_width is not None:
            # Only the width is settable. Both supported presets are 540 points
            # high, and so is every deck PowerPoint starts from, so setting the
            # width alone lands them exactly; anything else is reported below
            # rather than silently left wrong.
            pres.page_setup.slide_width.set(target_width)

    actual_width = pres.page_setup.slide_width()
    # `page setup` carries no slide height at all, so the height comes from the
    # slide master, which is read only.
    actual_height = pres.slide_master.height()

    if target_height is not None and abs(actual_height - target_height) > 0.5:
        warnings.append(
            f"The slide height is {actual_height} points, not {target_height}. "
            "PowerPoint for Mac's page setup has no slide height property and "
            "the slide master's height is read only, so only the width could "
            "be set."
        )

    template_name = ""
    try:
        value = pres.template_name()
        template_name = "" if is_missing(value) else value
    except Exception:
        pass

    full_name = pres.full_name()
    pres_index = _index_of(app, full_name)
    if pres_index is None:
        raise RuntimeError(
            "PowerPoint reported no error but the new presentation is not open."
        )

    if activate:
        try:
            pres.document_windows[1].activate()
        except Exception as exc:  # noqa: BLE001 - the deck still exists
            logger.warning("Could not activate new presentation window: %s", exc)
        ppt._target_pres_full_name = full_name

    result = {
        "success": True,
        "presentation_index": pres_index,
        "name": pres.name(),
        "slides_count": count(pres.slides),
        "slide_width": actual_width,
        "slide_height": actual_height,
        "template_name": template_name,
        "activated": activate,
    }
    if warnings:
        result["warning"] = " ".join(warnings)
    return result


def _open_presentation_impl(
    file_path: str,
    read_only: bool,
    with_window: bool,
    activate: bool,
) -> dict:
    path = os.path.abspath(os.path.expanduser(file_path))
    if not os.path.exists(path):
        raise FileNotFoundError(f"File not found: {file_path}")

    # Opening a file legitimately needs PowerPoint, so launch it if not running.
    # `with_window` says the deck wants a window, not that PowerPoint should
    # take the front from whatever the user is in the middle of. The window
    # appears either way, and `activate` below is the argument that means it.
    app = ppt._get_app_impl(allow_launch=True)

    pres = _open_and_find(app, path)

    warnings = []
    actual_read_only = bool(pres.read_only())
    if read_only and not actual_read_only:
        # The dictionary declares no `open`, so opening goes through
        # AppleScript's own verb, which takes no read-only option.
        warnings.append(
            "read_only was requested but the deck is open for editing. "
            "PowerPoint for Mac's open verb takes no read-only option."
        )
    if not with_window:
        warnings.append(
            "with_window=false has no effect. PowerPoint for Mac has no "
            "headless mode and always shows a window."
        )

    full_name = pres.full_name()
    pres_index = _index_of(app, full_name)

    if activate:
        if with_window:
            try:
                pres.document_windows[1].activate()
            except Exception as exc:  # noqa: BLE001 - the deck is still open
                logger.warning(
                    "Could not activate opened presentation window: %s", exc
                )
        ppt._target_pres_full_name = full_name

    result = {
        "success": True,
        "presentation_index": pres_index,
        "name": pres.name(),
        "full_name": full_name,
        "slides_count": count(pres.slides),
        "read_only": actual_read_only,
        "activated": activate,
    }
    if warnings:
        result["warning"] = " ".join(warnings)
    return result


def _save_presentation_impl(
    presentation_index: Optional[int],
    presentation_name: Optional[str],
) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(
        app,
        presentation_index=presentation_index,
        presentation_name=presentation_name,
    )

    path = pres.path()
    if is_missing(path) or not str(path).strip():
        raise RuntimeError(
            "This presentation has never been saved to a file, and PowerPoint "
            "for Mac does not answer a save on a deck with no file path; it "
            "writes nothing and can hang for a minute rather than failing. "
            "Use ppt_save_presentation_as with a full path instead."
        )

    full_name = pres.full_name()
    local = _local_path(full_name)
    before_mtime = (
        os.path.getmtime(local) if local and os.path.exists(local) else None
    )

    pres.save()

    saved = bool(pres.saved())
    if local:
        if not os.path.exists(local) or os.path.getsize(local) == 0:
            raise _sandbox_error(local)
        # A deck with nothing to write leaves the file untouched and comes
        # back already marked saved, so either signal is enough.
        if os.path.getmtime(local) <= (before_mtime or 0) and not saved:
            raise _sandbox_error(local)
    elif not saved:
        raise RuntimeError(
            f"PowerPoint reported no error but {full_name} is still unsaved."
        )

    result = {
        "success": True,
        "name": pres.name(),
        "saved": saved,
    }
    refreshed = _refresh_external_copy(local)
    if refreshed:
        result["also_copied_to"] = refreshed
    return result


# Where each deck inside PowerPoint's container has a copy of its own outside
# it. PowerPoint for Mac cannot hold a document anywhere but its container, so
# ppt_save_presentation_as leaves the deck there and carries a copy out. Saving
# after that used to refresh only the container's file, so a caller following
# the advice to save at every break watched their own file fall further behind
# without being told. Keyed by the container path, which is what the open deck
# answers for itself.
_EXTERNAL_COPIES: dict = {}


def _free_staged_path(app, target: str) -> str:
    """A container path for `target` that no other open deck is already using.

    The basename alone is not unique. Two decks called `report.pptx` living in
    different folders staged to the same container path, so saving the second
    wrote over the first while it was still open, and both then shared one key
    in `_EXTERNAL_COPIES`, which left a save on either one refreshing whichever
    copy registered last.

    The plain name is kept whenever it is free, because PowerPoint puts it in
    the title bar and a suffix on every deck would be noise. Saving the same
    deck to the same place twice reuses its own staged file rather than
    growing a new one each time.
    """
    base = os.path.basename(target)
    stem, ext = os.path.splitext(base)
    wanted = os.path.abspath(target)
    taken = set(_full_names(app))

    for attempt in range(1, 100):
        name = base if attempt == 1 else f"{stem}-{attempt}{ext}"
        candidate = os.path.join(EXPORT_STAGING_DIR, name)
        key = os.path.abspath(candidate)
        registered = _EXTERNAL_COPIES.get(key)
        if registered is not None and registered != wanted:
            # Held by a different deck that still wants its own copy refreshed.
            continue
        if registered is None and candidate in taken:
            # Open under this name without a copy outside; do not write over it.
            continue
        return candidate

    raise RuntimeError(
        f"A hundred decks called {base} are already staged in PowerPoint's "
        "container. Close some, or save under a different name."
    )


def _refresh_external_copy(local: Optional[str]) -> Optional[str]:
    """Bring the caller's own copy back up to date after a save.

    Returns where it was copied to, or None when this deck has no copy outside
    the container. A copy that cannot be written is not an error; the save
    itself landed, and saying where it did not reach beats failing the call.
    """
    if not local:
        return None
    target = _EXTERNAL_COPIES.get(os.path.abspath(local))
    if not target:
        return None
    try:
        shutil.copy2(local, target)
    except OSError:
        logger.debug("Could not refresh the copy at %s", target)
        return None
    return target


# Said in full once and then in one line. Saving at every natural break is the
# advice, so the long form arrives on every save and a reader stops reading it,
# while dropping it altogether would hide that the copy outside the container
# goes stale.
#
# The short form has to stand on its own. It first said the reason had been
# given "the first time", and the flag behind that lives as long as the server
# process, which serves one conversation after another. A caller's very first
# save was answered with a pointer to an explanation given to somebody else.
_container_copy_explained = False


def _container_copy_warning(staged: str, target: str) -> str:
    """Say the open deck is the container's copy, at length the first time."""
    global _container_copy_explained
    if _container_copy_explained:
        return (
            f"PowerPoint holds {staged} and {target} is a copy of it, because "
            "it cannot hold a document outside its container. "
            "ppt_save_presentation keeps the copy up to date from here."
        )
    _container_copy_explained = True
    return (
        f"The open deck is {staged}, inside PowerPoint's container, and "
        f"{target} is a copy of it. PowerPoint for Mac cannot hold a document "
        "outside its container. There is nothing more to do about it though. "
        "ppt_save_presentation now refreshes this copy every time it saves, "
        "and says so in `also_copied_to`."
    )


def _save_presentation_as_impl(
    file_path: str,
    format: Optional[str],
    presentation_index: Optional[int],
    presentation_name: Optional[str],
) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(
        app,
        presentation_index=presentation_index,
        presentation_name=presentation_name,
    )

    target = os.path.abspath(os.path.expanduser(file_path))
    parent = os.path.dirname(target)
    if parent and not os.path.isdir(parent):
        raise FileNotFoundError(f"Directory not found: {parent}")

    fmt_key = None
    if format:
        fmt_key = format.lower().strip()
        if fmt_key not in SAVE_FORMAT_MAP:
            raise ValueError(
                f"Unknown format '{format}'. "
                f"Supported: {list(SAVE_FORMAT_MAP.keys())}"
            )
    else:
        # No format given, so PowerPoint infers from the extension. Work out
        # the same answer here, only to catch the image formats below.
        fmt_key = os.path.splitext(target)[1].lstrip(".").lower() or None

    if fmt_key in _IMAGE_FORMATS:
        return {
            "error": "ppt_save_presentation_as to an image format is not "
                     "available on macOS",
            "reason": (
                "PowerPoint for Mac's dictionary has no export command, and a "
                "whole-deck save as PNG or JPG returns success in a fifth of "
                "a second while writing no folder and no files. Save the deck "
                "as PDF instead, which is verified working."
            ),
            "platform": "macOS",
            "alternatives": ["ppt_save_presentation_as with format='pdf'"],
        }

    kwargs = {}
    if format:
        kwargs["as_"] = to_keyword(
            PpSaveAsFileType, SAVE_FORMAT_MAP[fmt_key], "save format"
        )

    before_mtime = os.path.getmtime(target) if os.path.exists(target) else None

    # Two things are needed together and neither is optional. The destination
    # has to be a file reference rather than a path string, and a format has to
    # be named. `save in: "<path>"` writes nothing at all, anywhere, including
    # inside PowerPoint's own container, and reports success. With
    # `mactypes.File` and an explicit format it lands.
    #
    # And it only lands inside the container. Documents, Desktop and Downloads
    # were each tried and each hung PowerPoint until the call timed out, which
    # is the sandbox behaviour MACOS_PORT section 5.3 measured. So the deck is
    # saved into the container and Python carries the file the rest of the way.
    #
    # A POSIX path, always. An HFS colon path is taken as a literal filename
    # and produces a file called "Macintosh HD:Users:..." in the container root.
    os.makedirs(EXPORT_STAGING_DIR, exist_ok=True)
    staged = _free_staged_path(app, target)
    if "as_" not in kwargs:
        kwargs["as_"] = to_keyword(
            PpSaveAsFileType,
            SAVE_FORMAT_MAP.get(fmt_key, ppSaveAsOpenXMLPresentation),
            "save format",
        )
    pres.save(in_=mactypes.File(staged), **kwargs)

    if not os.path.exists(staged) or os.path.getsize(staged) == 0:
        raise _sandbox_error(staged)

    # The caller is allowed to name the container path itself, and a caller who
    # read the warning below is likely to. Then source and destination are one
    # file and `copy2` raises SameFileError, whose message prints the same path
    # twice and reads like nonsense for a save that in fact landed.
    saved_in_place = os.path.abspath(staged) == os.path.abspath(target)
    if not saved_in_place:
        shutil.copy2(staged, target)
    if not os.path.exists(target) or os.path.getsize(target) == 0:
        raise RuntimeError(
            f"The deck was saved to {staged} but could not be copied to "
            f"{target}. The staged file is still there."
        )

    # The old reference is addressed by name and saving renamed the deck, so
    # asking it anything now answers -1728. It is found again by the file
    # PowerPoint just wrote, which is the one thing known for certain.
    pres = _presentation_at(app, staged) or pres
    full_name = pres.full_name()

    # Saving renames the deck, and the session target is held by name. Without
    # this the target no longer matches anything open, so the next call falls
    # back to the active presentation, which with two decks open is the wrong
    # one and says nothing about the switch.
    if ppt._target_pres_full_name:
        ppt._target_pres_full_name = str(full_name)

    warnings = []
    if os.path.abspath(str(full_name)) == os.path.abspath(staged):
        # PowerPoint is holding the staged file now, not the caller's. Saying
        # so matters, because its own File then Save writes to the container
        # from here on and the caller's copy would quietly stop keeping up.
        if not saved_in_place:
            # Remember it, so an ordinary save from here on refreshes this copy
            # too rather than leaving it behind in the container.
            _EXTERNAL_COPIES[os.path.abspath(staged)] = os.path.abspath(target)
        if saved_in_place:
            warnings.append(
                f"The deck was saved to {target}, which is inside PowerPoint's "
                "container, so there is no copy anywhere else. Save it again "
                "to a path of your own when you want one outside the container."
            )
        else:
            warnings.append(_container_copy_warning(staged, target))
    else:
        try:
            os.remove(staged)
        except OSError:
            logger.debug("Could not remove the staged copy at %s", staged)

    result = {
        "success": True,
        "name": pres.name(),
        "full_name": full_name,
    }
    if before_mtime is not None and os.path.getmtime(target) <= before_mtime:
        # Overwriting a file that was already there, and its timestamp did not
        # move. That is ambiguous rather than a failure, because PowerPoint may
        # have had nothing to rewrite, so it is reported instead of raised.
        warnings.append(
            f"{target} already existed and its timestamp did not change, so it "
            "may not have been rewritten. Check the file before relying on it."
        )
    if warnings:
        result["warnings"] = warnings
    return result


def _close_presentation_impl(
    save_changes: bool,
    presentation_index: Optional[int],
    presentation_name: Optional[str],
) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(
        app,
        presentation_index=presentation_index,
        presentation_name=presentation_name,
    )
    name = pres.name()
    full_name = pres.full_name()
    local = _local_path(full_name)
    closed_copy = None

    if save_changes:
        path = pres.path()
        if is_missing(path) or not str(path).strip():
            raise RuntimeError(
                "This presentation has never been saved to a file, so saving "
                "it on the way out would write nothing and could hang. Save "
                "it with ppt_save_presentation_as first, or close it with "
                "save_changes=false."
            )
        pres.save()
        if local and (not os.path.exists(local) or os.path.getsize(local) == 0):
            raise _sandbox_error(local)
        # The last save of all, and the one most worth carrying out of the
        # container. Without this the deck closes with the caller's own file
        # holding everything except the edits they just asked to keep.
        closed_copy = _refresh_external_copy(local)
    else:
        # Suppress the "save changes?" sheet, which would otherwise leave
        # PowerPoint waiting on a click nobody is there to make.
        pres.saved.set(True)

    pres.close()

    if full_name in _full_names(app):
        raise RuntimeError(
            f"PowerPoint reported no error but {name} is still open. It is "
            "probably waiting on a dialog."
        )

    if local:
        _EXTERNAL_COPIES.pop(os.path.abspath(local), None)

    result = {"success": True, "closed": name}
    if closed_copy:
        result["also_copied_to"] = closed_copy
    return result


def _get_presentation_info_impl(
    presentation_index: Optional[int],
    presentation_name: Optional[str],
) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(
        app,
        presentation_index=presentation_index,
        presentation_name=presentation_name,
    )
    page = pres.page_setup

    template_name = ""
    try:
        value = pres.template_name()
        template_name = "" if is_missing(value) else value
    except Exception:
        pass

    # Fonts: title/body, Latin/East Asian from the theme font scheme. Windows
    # reads MinorFont(1) and MinorFont(3); the collections here are in the same
    # order, so the positions carry over.
    fonts = {
        "title_latin": None,
        "title_east_asian": None,
        "body_latin": None,
        "body_east_asian": None,
    }
    try:
        scheme = pres.slide_master.theme.theme_font_scheme

        def _clean_font(name):
            if is_missing(name):
                return None
            text = str(name)
            if text.startswith("+"):
                return None
            return text or None

        minor = elements(scheme.minor_theme_fonts)
        major = elements(scheme.major_theme_fonts)
        if len(minor) >= 1:
            fonts["body_latin"] = _clean_font(minor[0].name())
        if len(minor) >= 3:
            fonts["body_east_asian"] = _clean_font(minor[2].name())
        if len(major) >= 1:
            fonts["title_latin"] = _clean_font(major[0].name())
        if len(major) >= 3:
            fonts["title_east_asian"] = _clean_font(major[2].name())
    except Exception:
        pass

    # Accent colours. macOS orders its scheme colours exactly as Windows
    # numbers them, dark1, light1, dark2, light2, then accent1 to accent6, so
    # the accents are positions 5 to 10.
    #
    # Asked for one at a time rather than read out of the materialised
    # collection. `theme colors` answers twelve elements happily enough, and
    # then `RGB` on one of them raises, which is the rule this whole port runs
    # on: a reference PowerPoint hands back is not trusted. Every accent came
    # back null for it, on decks whose palette reads perfectly well through
    # `theme_colors[5]`.
    accent_colors = {}
    scheme = pres.slide_master.theme.theme_color_scheme
    for position, key in enumerate(
        ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"],
        start=5,
    ):
        try:
            accent_colors[key] = rgb_list_to_hex(scheme.theme_colors[position].RGB())
        except Exception:
            # A deck with no palette to read is the ordinary case for a
            # template that carries none, so null rather than an error.
            accent_colors[key] = None

    full_name = pres.full_name()
    local_path = resolve_local_path(full_name)
    local_dir = os.path.dirname(local_path) if local_path else None

    path = pres.path()
    slide_width = page.slide_width()
    slide_height = pres.slide_master.height()

    return {
        "name": pres.name(),
        "full_name": full_name,
        "local_path": local_path,
        "local_dir": local_dir,
        # An unsaved deck answers `missing value` rather than an empty string,
        # which would otherwise reach the caller as the text "k.missing_value".
        "path": None if is_missing(path) else path,
        "slides_count": count(pres.slides),
        "read_only": bool(pres.read_only()),
        "saved": bool(pres.saved()),
        "slide_width": slide_width,
        # There is no slide height in `page setup`, so this comes from the
        # slide master.
        "slide_height": slide_height,
        "slide_width_inches": round(slide_width / 72.0, 3),
        "slide_height_inches": round(slide_height / 72.0, 3),
        "first_slide_number": page.first_slide_number(),
        "template_name": template_name,
        "fonts": fonts,
        "accent_colors": accent_colors,
    }


def _list_templates_impl(templates_dir: Optional[str]) -> dict:
    """List PowerPoint template files in a directory.

    Touches the filesystem only, never PowerPoint. The Windows version reads
    the personal templates folder out of the registry, which has no macOS
    counterpart, so the known Office locations are checked instead.
    """
    if templates_dir is None:
        for candidate in _TEMPLATE_DIR_CANDIDATES:
            expanded = os.path.expanduser(candidate)
            if os.path.isdir(expanded):
                templates_dir = expanded
                break

    if templates_dir is None:
        return {
            "templates_dir": None,
            "count": 0,
            "templates": [],
            "error": (
                "Could not find a templates directory. Specify templates_dir "
                "explicitly."
            ),
        }

    if not os.path.isdir(templates_dir):
        return {
            "templates_dir": templates_dir,
            "count": 0,
            "templates": [],
            "error": f"Directory not found: {templates_dir}",
        }

    templates = []
    for ext in ("*.potx", "*.potm"):
        pattern = os.path.join(templates_dir, ext)
        for filepath in glob_mod.glob(pattern):
            templates.append({
                "name": os.path.splitext(os.path.basename(filepath))[0],
                "file_name": os.path.basename(filepath),
                "file_path": os.path.abspath(filepath),
            })

    templates.sort(key=lambda t: t["name"])
    return {
        "templates_dir": templates_dir,
        "count": len(templates),
        "templates": templates,
    }
