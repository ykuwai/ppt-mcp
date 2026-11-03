"""ppt-mcp: The world's best PowerPoint MCP server.

Real-time PowerPoint control via COM automation.
"""

import json
import logging
import os
import sys
import tempfile
from pathlib import Path

from pydantic import BaseModel, ConfigDict, Field

# When installed via PyPI (entry point: src.server:main), ensure the src/
# directory is in sys.path so that internal imports like
# `from backend import ppt` resolve correctly.
_src_dir = str(Path(__file__).parent)
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)

from contextlib import asynccontextmanager

# mcp 2.0 renamed mcp.server.fastmcp to mcp.server.mcpserver, and FastMCP to
# MCPServer, without leaving a compatibility shim behind. Import from whichever
# module the installed major provides so the server runs on both 1.x and 2.x.
#
# Catch ModuleNotFoundError rather than ImportError: the fallback should only
# trigger when the 2.x module is absent. A broken 2.x install that raises
# ImportError from one of its own imports must surface that traceback instead
# of being masked by the 1.x fallback failing afterwards.
try:
    from mcp.server.mcpserver import Image, MCPServer  # mcp >= 2.0
    from mcp.server.mcpserver.exceptions import ToolError
except ModuleNotFoundError:
    from mcp.server.fastmcp import Image  # mcp 1.x
    from mcp.server.fastmcp import FastMCP as MCPServer
    from mcp.server.fastmcp.exceptions import ToolError
from pydantic import ValidationError

from utils.arguments import describe_unknown_top_level, describe_validation_error
from utils.com_wrapper import UndoEntry, call_presentation, call_undo_entry
from utils.coauthoring import CallCoauthoring, call_coauthoring

# Configure logging to stderr (stdout is used for MCP protocol)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger("ppt-mcp")


@asynccontextmanager
async def app_lifespan(server: MCPServer):
    """Manage COM lifecycle for the MCP server."""
    from backend import ppt

    from backend import AUTO_DISMISS_DIALOG
    logger.info("AUTO_DISMISS_DIALOG=%s (set PPT_AUTO_DISMISS_DIALOG=true to enable)", AUTO_DISMISS_DIALOG)
    logger.info("Starting PowerPoint COM worker thread...")
    ppt.start()
    # Do NOT connect to PowerPoint here. Connecting at startup would launch
    # PowerPoint.exe the moment the MCP client boots, even when the user never
    # invokes a ppt_* tool. The COM connection is established lazily on the
    # first tool call instead (see PowerPointCOMWrapper._get_app_impl).
    try:
        yield {}
    finally:
        logger.info("Shutting down PowerPoint COM worker thread...")
        ppt.stop()


# Appended to the server instructions only when the backend is the Apple Event
# one. On Windows it is empty, so nothing about it reaches a Windows reader.
_MACOS_NOTE = """
## On macOS

This server is driving PowerPoint for Mac over Apple Events rather than COM.
Almost everything works the same, and where it does not, a tool answers with
`error`, `reason`, `platform` and often `alternatives` instead of doing
something unexpected. Read the `reason`; it says what PowerPoint for Mac does
not have and what to do instead, and one retry along that route usually works.

**Call these tools one after another, not several in the same turn.** Only one
request at a time reaches PowerPoint here, so the rest queue behind it, and the
order they come out in is whatever order they arrived. A call that adds a slide
and a call that reads it, sent together, can run the wrong way round. A call
left waiting too long is taken back and says so, having changed nothing.

A refusal whose `error` names an argument rather than the tool means the tool
itself is fine and only that argument has to go. `ppt_add_table_row` cannot
insert at a `position`, for instance, but appending still works.

Some results carry `warnings` alongside `success`. Those are parts of the
request that did not land, listed rather than silently dropped.

A note about Japanese fonts. PowerPoint for Mac will not set the East Asian
font to a font that has no East Asian glyphs. It accepts the write, keeps the
old value and reports nothing, so `font_name` alone leaves Japanese text in
whatever font it was in. That now comes back as a warning; pass
`font_name_fareast` to choose the Japanese font yourself.

**Save early and at each natural break.** A deck that has never been saved has
been seen to disappear mid-session on macOS, taking every slide with it, and
the cause is not yet known (issue #191). `ppt_save_presentation_as` right after
creating a deck, and `ppt_save_presentation` at each break, costs almost
nothing and is the difference between a hiccup and starting over.

Charts, freeform paths and groups go through the clipboard here, because
PowerPoint for Mac's scripting dictionary has no words for them. Creating and
reading them works as on Windows. Editing an existing chart or path
(`ppt_set_chart_data`, `ppt_change_chart_type`, `ppt_format_chart`,
`ppt_format_chart_axis`, `ppt_set_chart_series`, and the node tools) works by
copying the shape, rewriting it and pasting it back, then deleting the
original: the name, position and z order are kept, animations on that shape
are lost, and `warnings` says so each time. Set the data and formatting of a
chart before animating it. A node's editing type cannot be set; SmartArt has
no route at all. The user's clipboard is saved and put back around each of
these calls.

A few chart arguments have no place in that XML and are refused by name rather
than ignored: `chart_style`, `legend_font_size`, explicit legend and title
coordinates and the corner legend positions on `ppt_format_chart`, and
`tick_label_font_size` on `ppt_format_chart_axis`. Everything else those two
take works. Drop the named argument and call again.
""" if sys.platform == "darwin" else ""


# Naming a font PowerPoint cannot find does not fail. The system quietly
# substitutes one, so a deck built from the Windows recommendation on a Mac
# comes out in something nobody chose. The advice therefore has to be right
# the first time it is read, rather than corrected further down (#194).
from ppt_com.layout_check import CHECK_EACH_WRITE

# Step 7 of the instructions, which PPT_LAYOUT_CHECKS changes.
_WRITE_CHECKS = (
    "The tools that add, copy, move, resize or retext a shape, or fill a placeholder, "
    "check what they touched and report any problem in their result under "
    "`layout_issues`, each with a `fix` naming the call and its values: act on them "
    "before moving on."
    if CHECK_EACH_WRITE else
    "This server's write tools do not check what they touched (PPT_LAYOUT_CHECKS=off), "
    "so run `ppt_check_layout` on a slide after each few changes to it, and act on "
    "each issue's `fix` before moving on."
)

_PREFERRED_FONTS = (
    "Preferred fonts: Hiragino Sans (Japanese, ヒラギノ角ゴシック in the font "
    "menu) with Hiragino Sans W6 for headings, and Helvetica Neue or Arial "
    "(Latin). Segoe UI and BIZ UDPゴシック are Windows fonts and are not on "
    "macOS; BIZ UDGothic ships here, without the P."
    if sys.platform == "darwin"
    else "Preferred fonts: BIZ UDPゴシック (Japanese) + Segoe UI (Latin)."
)


# The argument every presentation-bound tool accepts next to `params`, and the
# tools that do not take it because they act on PowerPoint itself, on the
# running slide show, or set the session target in the first place.
PRESENTATION_ARG = "presentation"
_TOOLS_WITHOUT_PRESENTATION_ARG = frozenset({
    "ppt_connect", "ppt_get_app_info", "ppt_get_active_window",
    "ppt_set_window_state", "ppt_list_presentations", "ppt_create_presentation",
    "ppt_open_presentation", "ppt_activate_presentation", "ppt_list_templates",
    "ppt_search_icons", "ppt_slideshow_stop", "ppt_slideshow_next",
    "ppt_slideshow_previous", "ppt_slideshow_goto", "ppt_slideshow_get_status",
})
# Type only. The same entry is copied into every bound tool's schema, and a
# description there would be sent 141 times per connection (about 80,000
# characters); the server instructions explain the argument once instead.
_PRESENTATION_ARG_SCHEMA = {"type": "string"}


class _PowerPointServer(MCPServer):
    """MCPServer that accepts a per-call `presentation` argument.

    The session target set by ppt_activate_presentation is one value for the
    whole process, and one process can serve several conversations at once:
    Claude Desktop starts a server once and routes every chat through it. When
    a second conversation activates its deck, the first one's next call lands
    there. A call that names its deck is immune to that.

    The argument is taken out here, before the SDK validates `params`, and
    held in utils.com_wrapper.call_presentation while the tool runs. Both SDK
    majors dispatch tools/call through this method, so no SDK internals are
    involved; the schema side is added by _add_presentation_arg().

    It is also where an argument the tool does not take is refused. The input
    models forbid unknown keys, but the SDK's own argument model around
    `params` ignores them, so a key passed next to `params` is checked here
    against the tool's declared arguments. A validation failure inside
    `params` is rewritten from pydantic's report into one sentence naming the
    tool, the argument, the likely intended name and the valid ones. Both
    still answer as a tool error (isError true).

    A tool that answers a failure, a JSON object with a top-level "error", is
    marked isError true here too (#249), with its text left exactly as the
    tool wrote it; see _flag_failure.
    """

    async def call_tool(self, name, arguments, *args, **kwargs):
        wanted = None
        if isinstance(arguments, dict) and PRESENTATION_ARG in arguments:
            arguments = dict(arguments)
            wanted = arguments.pop(PRESENTATION_ARG)
            if wanted is not None and not isinstance(wanted, str):
                raise ValueError(
                    f"'{PRESENTATION_ARG}' must be a string (a file name or "
                    f"full path), got {type(wanted).__name__}"
                )
            # Harmless where it means nothing, and clients told to pass it on
            # every call will, so it is dropped there rather than refused.
            if name in _TOOLS_WITHOUT_PRESENTATION_ARG:
                wanted = None
            wanted = (wanted or "").strip() or None
        tool = self._registered_tool(name)
        if tool is not None and isinstance(arguments, dict):
            _refuse_unknown_arguments(name, tool, arguments)
        token = call_presentation.set(wanted)
        coauthoring_token = call_coauthoring.set(_coauthoring_call(name, tool, arguments))
        undo_token = call_undo_entry.set(_undo_entry_for(name, tool, arguments))
        try:
            return _flag_failure(await super().call_tool(name, arguments, *args, **kwargs))
        except ToolError as exc:
            readable = _readable_validation_error(name, tool, exc)
            if readable is None:
                raise
            raise ToolError(readable) from exc.__cause__
        finally:
            call_undo_entry.reset(undo_token)
            call_coauthoring.reset(coauthoring_token)
            call_presentation.reset(token)

    async def list_tools(self, *args, **kwargs):
        tools = await super().list_tools(*args, **kwargs)
        # Catalog mode (PPT_TOOL_MODE=catalog) lists the three tools that reach
        # the rest; every tool stays registered, so ppt_run dispatches through
        # call_tool above. Full mode, the default, lists every tool but those
        # three.
        if CATALOG_MODE:
            return [t for t in tools if t.name in _CATALOG_TOOLS]
        return [t for t in tools if t.name not in _CATALOG_TOOLS]

    def _registered_tool(self, name):
        """The SDK's Tool for `name`, or None.

        _tool_manager is an SDK internal (see _inline_schema_defs); if it moves,
        calls go through unchecked rather than failing.
        """
        try:
            return self._tool_manager.get_tool(name)
        except Exception:
            return None


def _failure_payload(text):
    """The JSON object `text` holds if it is a failure (a top-level "error"), else None."""
    if not text or not text.lstrip().startswith("{"):
        return None
    try:
        payload = json.loads(text)
    except ValueError:
        return None
    return payload if isinstance(payload, dict) and "error" in payload else None


def _sdk_passes_call_tool_result():
    """Whether the SDK sends a CallToolResult from call_tool as it is.

    mcp 1.19 was the first to. Before it, the low-level handler checks the
    result against the tool's output schema and answers "Output validation
    error" instead of the tool's text, so failures there stay as they were.
    """
    from importlib.metadata import PackageNotFoundError, version

    try:
        major, minor = (int(part) for part in version("mcp").split(".")[:2])
    except (PackageNotFoundError, ValueError):
        return True
    return (major, minor) >= (1, 19)


_SDK_PASSES_CALL_TOOL_RESULT = _sdk_passes_call_tool_result()


def _flag_failure(result):
    """`result` as a tool error (isError true) if the tool answered a failure.

    Every tool answers a failure as json.dumps({"error": ...}), which the SDK
    sends as a success. Raising ToolError instead would mark it, but the SDK
    puts "Error executing tool <name>: " in front of the text, and the JSON no
    longer parses. A CallToolResult with isError set keeps the text as it is.
    It carries no structured content, the shape #249 measured: a client does
    not check an error result against the tool's output schema.
    """
    if hasattr(result, "content") and not isinstance(result, (list, tuple, dict)):
        # mcp 2.x answers a CallToolResult already.
        if getattr(result, "is_error", False) or _failure_payload(_text_of(result.content)) is None:
            return result
        return result.model_copy(update={"is_error": True, "structured_content": None})
    # mcp 1.x answers (blocks, structured) for a tool with an output schema,
    # which is every tool returning str, and the blocks alone for the others.
    # A dict is structured content only, and no tool here answers one.
    blocks = result[0] if isinstance(result, tuple) else result
    if (not _SDK_PASSES_CALL_TOOL_RESULT or isinstance(result, dict)
            or _failure_payload(_text_of(blocks or [])) is None):
        return result
    from mcp.types import CallToolResult

    return CallToolResult(content=list(blocks), isError=True)


# Presentation-bound tools that handle co-authoring themselves (save) or have
# no deck left to look at afterwards (save as under a new name, close).
_TOOLS_WITHOUT_COAUTHORING = frozenset({
    "ppt_save_presentation", "ppt_save_presentation_as", "ppt_close_presentation",
})


# The tools that work the undo stack themselves, and so must not start an entry
# of their own first: an entry started just before Undo is the one it undoes.
_TOOLS_THAT_DRIVE_UNDO = frozenset({"ppt_undo", "ppt_redo", "ppt_start_undo_entry"})
# The ribbon commands that do the same through ppt_execute_mso.
_MSO_COMMANDS_THAT_DRIVE_UNDO = frozenset({"undo", "redo"})


def _is_read_only(tool) -> bool:
    annotations = getattr(tool, "annotations", None)
    read_only = getattr(annotations, "readOnlyHint", None)
    if read_only is None:
        read_only = getattr(annotations, "read_only_hint", None)
    return bool(read_only)


def _call_params(arguments) -> dict:
    """The call's ``params``, as a dict, whether sent as an object or as JSON."""
    params = arguments.get("params") if isinstance(arguments, dict) else None
    if isinstance(params, str):
        try:
            params = json.loads(params)
        except ValueError:
            params = None
    return params if isinstance(params, dict) else {}


def _undo_entry_for(name, tool, arguments=None):
    """A fresh undo step for a tool call that changes the deck, or None.

    PowerPoint gathers every change made through COM since the last
    StartNewUndoEntry into one undo entry, so without one per call a single
    Undo took back everything the tools had done, slides included (#298).
    utils.com_wrapper starts it before the call's first COM operation. macOS
    has no StartNewUndoEntry and its backend does not read it.
    """
    if name in _TOOLS_THAT_DRIVE_UNDO or _is_read_only(tool):
        return None
    if name == "ppt_execute_mso":
        command = str(_call_params(arguments).get("command_name") or "").strip().lower()
        if command in _MSO_COMMANDS_THAT_DRIVE_UNDO:
            return None
    return UndoEntry()


def _coauthoring_call(name, tool, arguments):
    """What utils/coauthoring.py needs to know about this call, or None.

    Only presentation-bound tools take part. A tool writes unless its
    readOnlyHint says otherwise; its own target fields (the destination of
    ppt_copy_slide first) name the deck it works on.
    """
    if name in _TOOLS_WITHOUT_PRESENTATION_ARG or name in _TOOLS_WITHOUT_COAUTHORING:
        return None
    annotations = getattr(tool, "annotations", None)
    read_only = getattr(annotations, "readOnlyHint", None)
    if read_only is None:
        read_only = getattr(annotations, "read_only_hint", None)
    params = arguments.get("params") if isinstance(arguments, dict) else None
    if isinstance(params, str):
        try:
            params = json.loads(params)
        except ValueError:
            params = None
    if not isinstance(params, dict):
        params = {}
    for prefix in ("to_", ""):
        index = params.get(prefix + "presentation_index")
        name_ = params.get(prefix + "presentation_name")
        if index is not None or name_ is not None:
            return CallCoauthoring(writes=not read_only, presentation_name=name_,
                                   presentation_index=index)
    return CallCoauthoring(writes=not read_only)


def _arg_model(tool):
    """The SDK's per-tool argument model (the one whose field is `params`)."""
    return getattr(getattr(tool, "fn_metadata", None), "arg_model", None)


def _refuse_unknown_arguments(name, tool, arguments):
    """Refuse keys next to `params` that the tool does not declare.

    The SDK's argument model ignores them, so `{"params": {...},
    "presntation": "x"}` used to run against the session target and answer
    success. `presentation` itself has already been taken out by call_tool.
    """
    declared = list((tool.parameters or {}).get("properties") or {})
    # An empty `params` on a tool that takes none carries nothing that could
    # be lost, and a caller used to sending it everywhere is not wrong.
    unknown = [
        key for key, value in arguments.items()
        if key not in declared and not (key == "params" and value in (None, {}))
    ]
    if not unknown:
        return
    if name not in _TOOLS_WITHOUT_PRESENTATION_ARG and PRESENTATION_ARG not in declared:
        # A server built without _add_presentation_arg (the tests' probes).
        declared.append(PRESENTATION_ARG)
    params_model = None
    arg_model = _arg_model(tool)
    if arg_model is not None and "params" in arg_model.model_fields:
        annotation = arg_model.model_fields["params"].annotation
        if isinstance(annotation, type) and issubclass(annotation, BaseModel):
            params_model = annotation
    raise ToolError(describe_unknown_top_level(name, unknown, declared, params_model))


def _readable_validation_error(name, tool, exc):
    """A caller-facing message for a ToolError caused by argument validation.

    Returns None for any other ToolError, including a ValidationError raised
    from inside a tool body, which is recognised by its title: pydantic titles
    the error after the model that was validated, and for the arguments that
    is the SDK's per-tool argument model.
    """
    cause = exc.__cause__
    if not isinstance(cause, ValidationError):
        return None
    arg_model = _arg_model(tool)
    if arg_model is None or cause.title != arg_model.__name__:
        return None
    try:
        return describe_validation_error(name, arg_model, cause)
    except Exception:
        logger.warning("Could not rewrite a validation error for %s", name, exc_info=True)
        return None


mcp = _PowerPointServer(
    "powerpoint_mcp",
    lifespan=app_lifespan,
    instructions="""
## Getting started

1. Call `ppt_activate_presentation` first — locks all tools to a specific file and prevents accidental edits to the wrong presentation. The activated presentation is shared by every conversation this server serves, so also pass its `full_name` as `presentation` on every further call: that keeps the call on your file even when another conversation activates a different one. Every tool that works on a presentation takes this optional `presentation` argument next to `params`. It accepts a `full_name` (as returned by `ppt_activate_presentation` or `ppt_open_presentation`), a file name, or a file name without extension; a new unsaved deck goes by the `name` that `ppt_create_presentation` returns. It applies to that call only and never changes the activated presentation. A presentation that is not open is an error, never a fallback to another one.
2. Call `ppt_get_presentation_info` to understand the presentation — slide count, dimensions, template, current default fonts, and accent colors. Use this to inform all subsequent decisions. When saving files (e.g., exported markdown, images), use the presentation's `local_dir` from `ppt_get_presentation_info` as the default save directory. To read the existing slide content, call `ppt_get_all_text` — it returns all text as pseudo-Markdown with layout analysis, heading detection, and formatting markers.
3. When adding slides, use `ppt_add_slide` with `count` to create multiple slides at once instead of calling it repeatedly.
4. After placing text, set fonts explicitly with `ppt_batch_apply_formatting` or `ppt_set_default_fonts`. On Japanese-locale systems, the slide master default is often 游ゴシック, which renders thin and illegible when projected. """ + _PREFERRED_FONTS + """
5. For visual symbols, `ppt_search_icons` + `ppt_add_svg_icon` produce crisper, scalable results than emoji characters and are generally preferred in presentations.
6. Use `ppt_get_slide_preview` to visually inspect slides as you work. The image is usually larger than the slide is in points; its result gives `pixels_per_point` to convert.
7. """ + _WRITE_CHECKS + """ When a slide is finished, run `ppt_check_layout` on it, and before calling the deck done, run it once more with no `slide_index`, which also compares the titles across slides. A slide is done when it reports no issues, or only ones you mean to keep: if the same finding comes back after two attempts to fix it, treat it as intended (a photo behind white text, an overlap you designed, something the template draws) and move on. The write tools check only the shapes they touched, and a slide copied into another open deck is not checked, so run `ppt_check_layout` there after copying. The check measures rather than looks, so it works with or without image input.
8. Build a slide at a time, not the whole deck at once. Before the first shape, fix the grid (margins, column positions, a type scale and a palette) and keep every shape on it: most alignment issues are coordinates a few points apart. Decide on photos up front, since adding them later changes the layout. Finish each slide, with its issues acted on, `ppt_check_layout` clean, or its remaining findings intended, and a preview checked if you can see images, before starting the next. Fix problems in place with the call each `fix` names. Deleting and rebuilding a slide or the deck to move something a few points costs many calls and finds the same problems again.

## Shared files

On Windows, a presentation opened from a server (SharePoint, OneDrive) can be edited by other people at the same time. PowerPoint takes in their saved changes only when the file is saved: tools do that before they run whenever nothing here is unsaved, and `ppt_save_presentation` always does. Results for such a file carry `coauthoring`: `pending_updates` = changes of others are waiting, `merged_updates` = they were just merged (re-read before relying on slide indices; a writing call whose slides or shapes the merge moved is refused without changes). A `conflict` means both sides changed the same object; unless it was resolved (`overwritten` then lists this copy's version), PowerPoint asks the user in a dialog. Save after each finished step so that conflicts stay small.

## Key tools

- `ppt_batch_apply_formatting`: Apply fill, line, shadow, text formatting to multiple shapes at once. Use whenever styling multiple shapes consistently.
- `ppt_set_default_shape_style`: Set default fill/border/font before inserting shapes. Can capture all properties from a template shape including gradients.
- `ppt_add_shape`: Supports inline text styling (`font_name`, `font_size`, `bold`, `font_color`, `align`) — no need for a separate `ppt_format_text` call.
- `ppt_add_slide`: Accepts `count` to create multiple slides at once.
- `ppt_set_slide_background` and `ppt_set_slide_transition`: Accept `slide_indices` to change several slides in one call.

## Design thinking

Before placing any shapes, take a moment to plan:
- **Slide structure**: decide the narrative arc and number of slides upfront
- **Color system**: use the deck's accent colors instead of hardcoding RGB values: pass the name, such as `"accent1"`, where a fill, line, text or icon color is taken, and the shape follows the theme if it changes. Pick one accent and use it consistently for icons, headings, and highlights.
- **Icons**: whenever content has distinct points, steps, or features, add an icon for each — use `ppt_search_icons` + `ppt_add_svg_icon`
- **One message per slide**: decide what this slide communicates, then place only what supports that message

Standard 16:9 slide = 960 × 540 pt. Default to light backgrounds unless the user specifies otherwise. Always ensure sufficient contrast between text and background.

**Consistency across slides:** use the same heading size, card style, and spacing throughout. `ppt_set_default_shape_style` and `ppt_batch_apply_formatting` are your tools for this. For rounded rectangles, use `corner_radius_pt` (e.g. `10`) to keep the same corner radius across shapes.

**Font sizes** — you should use these. They are what makes a slide readable when projected:
- Slide title: 40–48 pt
- Section heading / subheading: 24–32 pt
- Body text: 20–28 pt
- Caption / annotation: 16–20 pt
- **Never go below 16 pt.** Smaller text is unreadable when projected.
"""
    + _MACOS_NOTE,
)


# =============================================================================
# App tools
# =============================================================================
from utils.offload import run_offloaded
from ppt_com.app import (
    ConnectInput,
    SetWindowStateInput,
    connect_to_powerpoint,
    get_app_info,
    get_active_window_info,
    list_presentations,
    set_window_state,
)


@mcp.tool(
    name="ppt_connect",
    annotations={
        "title": "Connect to PowerPoint",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def tool_ppt_connect(params: ConnectInput) -> str:
    """Connect to a running PowerPoint instance or launch a new one.

    Attempts to connect to an already-running PowerPoint via COM.
    If no instance is found, launches a new one.
    Set visible=false for headless mode (background operation).
    """
    return await run_offloaded(connect_to_powerpoint, params)


@mcp.tool(
    name="ppt_get_app_info",
    annotations={
        "title": "Get PowerPoint App Info",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def tool_ppt_get_app_info() -> str:
    """Get information about the connected PowerPoint application.

    Returns version, visibility, window state, presentation count,
    and active presentation name.
    """
    return await run_offloaded(get_app_info)


@mcp.tool(
    name="ppt_get_active_window",
    annotations={
        "title": "Get Active Window Info",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def tool_ppt_get_active_window() -> str:
    """Get info about the active PowerPoint window and current selection.

    Returns window caption, view type, current slide index,
    and what is selected (shapes, text, or nothing).
    """
    return await run_offloaded(get_active_window_info)


@mcp.tool(
    name="ppt_list_presentations",
    annotations={
        "title": "List Open Presentations",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def tool_ppt_list_presentations() -> str:
    """List all currently open presentations in PowerPoint.

    Returns name, path, slide count, and status for each.
    """
    return await run_offloaded(list_presentations)


@mcp.tool(
    name="ppt_set_window_state",
    annotations={
        "title": "Set PowerPoint Window State",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    },
)
async def tool_ppt_set_window_state(params: SetWindowStateInput) -> str:
    """Set the PowerPoint application window state.

    Controls whether the PowerPoint window is maximized, minimized, or
    restored to normal size.
    """
    return await run_offloaded(set_window_state, params)


# =============================================================================
# Import and register additional tool modules as they are implemented.
# Each module registers its tools below.
# =============================================================================

# Presentation tools
try:
    from ppt_com.presentation import register_tools as register_presentation_tools
    register_presentation_tools(mcp)
except ImportError:
    logger.debug("presentation module not yet available")

# Slide tools
try:
    from ppt_com.slides import register_tools as register_slide_tools
    register_slide_tools(mcp)
except ImportError:
    logger.debug("slides module not yet available")

# Shape tools
try:
    from ppt_com.shapes import register_tools as register_shape_tools
    register_shape_tools(mcp)
except ImportError:
    logger.debug("shapes module not yet available")

# Text tools
try:
    from ppt_com.text import register_tools as register_text_tools
    register_text_tools(mcp)
except ImportError:
    logger.debug("text module not yet available")

# Placeholder tools
try:
    from ppt_com.placeholders import register_tools as register_placeholder_tools
    register_placeholder_tools(mcp)
except ImportError:
    logger.debug("placeholders module not yet available")

# Formatting tools
try:
    from ppt_com.formatting import register_tools as register_formatting_tools
    register_formatting_tools(mcp)
except ImportError:
    logger.debug("formatting module not yet available")

# Table tools
try:
    from ppt_com.tables import register_tools as register_table_tools
    register_table_tools(mcp)
except ImportError:
    logger.debug("tables module not yet available")

# Export tools
try:
    from ppt_com.export import register_tools as register_export_tools
    register_export_tools(mcp)
except ImportError:
    logger.debug("export module not yet available")

# SlideShow tools
try:
    from ppt_com.slideshow import register_tools as register_slideshow_tools
    register_slideshow_tools(mcp)
except ImportError:
    logger.debug("slideshow module not yet available")

# Groups tools
try:
    from ppt_com.groups import register_tools as register_groups_tools
    register_groups_tools(mcp)
except ImportError:
    logger.debug("groups module not yet available")

# Connectors tools
try:
    from ppt_com.connectors import register_tools as register_connectors_tools
    register_connectors_tools(mcp)
except ImportError:
    logger.debug("connectors module not yet available")

# Hyperlinks tools
try:
    from ppt_com.hyperlinks import register_tools as register_hyperlinks_tools
    register_hyperlinks_tools(mcp)
except ImportError:
    logger.debug("hyperlinks module not yet available")

# Sections tools
try:
    from ppt_com.sections import register_tools as register_sections_tools
    register_sections_tools(mcp)
except ImportError:
    logger.debug("sections module not yet available")

# Properties tools
try:
    from ppt_com.properties import register_tools as register_properties_tools
    register_properties_tools(mcp)
except ImportError:
    logger.debug("properties module not yet available")

# Charts tools
try:
    from ppt_com.charts import register_tools as register_charts_tools
    register_charts_tools(mcp)
except ImportError:
    logger.debug("charts module not yet available")

# Animation tools
try:
    from ppt_com.animation import register_tools as register_animation_tools
    register_animation_tools(mcp)
except ImportError:
    logger.debug("animation module not yet available")

# Themes tools
try:
    from ppt_com.themes import register_tools as register_themes_tools
    register_themes_tools(mcp)
except ImportError:
    logger.debug("themes module not yet available")

# Media tools
try:
    from ppt_com.media import register_tools as register_media_tools
    register_media_tools(mcp)
except ImportError:
    logger.debug("media module not yet available")

# SmartArt tools
try:
    from ppt_com.smartart import register_tools as register_smartart_tools
    register_smartart_tools(mcp)
except ImportError:
    logger.debug("smartart module not yet available")

# Edit operations tools (undo, redo, clipboard, format copy)
try:
    from ppt_com.edit_ops import register_tools as register_edit_ops_tools
    register_edit_ops_tools(mcp)
except ImportError:
    logger.debug("edit_ops module not yet available")

# Layout tools (align, distribute, slide size, background, flip)
try:
    from ppt_com.layout import register_tools as register_layout_tools
    register_layout_tools(mcp)
except ImportError:
    logger.debug("layout module not yet available")

# Layout check (measured problems on a slide or across the deck)
try:
    from ppt_com.layout_check import register_tools as register_layout_check_tools
    register_layout_check_tools(mcp)
except ImportError:
    logger.debug("layout_check module not yet available")

# Visual effects tools (glow, reflection, soft edge)
try:
    from ppt_com.effects import register_tools as register_effects_tools
    register_effects_tools(mcp)
except ImportError:
    logger.debug("effects module not yet available")

# Comments tools
try:
    from ppt_com.comments import register_tools as register_comments_tools
    register_comments_tools(mcp)
except ImportError:
    logger.debug("comments module not yet available")

# Advanced operations tools (tags, fonts, crop, merge, export, hidden, select, view)
try:
    from ppt_com.advanced_ops import register_tools as register_advanced_ops_tools
    register_advanced_ops_tools(mcp)
except ImportError:
    logger.debug("advanced_ops module not yet available")

# Batch apply formatting tools
try:
    from ppt_com.batch_apply import register_tools as register_batch_apply_tools
    register_batch_apply_tools(mcp)
except ImportError:
    logger.debug("batch_apply module not yet available")

# Freeform shape tools
try:
    from ppt_com.freeform import register_tools as register_freeform_tools
    register_freeform_tools(mcp)
except ImportError:
    logger.debug("freeform module not yet available")


# =============================================================================
# Tools: Slide Preview (Visual Inspection)
# =============================================================================


class GetSlidePreviewInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    slide_index: int = Field(1, ge=1, description="1-based slide index")


@mcp.tool(
    name="ppt_get_slide_preview",
    annotations={
        "title": "Get Slide Preview Image",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
# No return annotation: it answers an image and a text block, and the SDK
# builds an output schema from `list[Image | str]` and fails on Image.
async def tool_ppt_get_slide_preview(params: GetSlidePreviewInput):
    """Get a visual preview of a PowerPoint slide as an image, with its scale in pixels per point.

    This is the RECOMMENDED way to visually inspect slides for appearance, design,
    layout, colors, text readability, and overall quality. Much more efficient
    than exporting all slides to files.

    Also navigates the PowerPoint editor window to the target slide so the user
    can see which slide is being inspected.

    Returns:
        Image: PNG image of the slide for visual inspection, followed by JSON
        giving its scale: image_width_px and image_height_px, slide_width_pt
        and slide_height_pt, and pixels_per_point. Multiply a position or size
        in points by pixels_per_point to find it in the image; the image is
        usually larger than the slide is in points.
    """
    from backend import IS_MACOS, ppt
    from utils.navigation import goto_slide

    def _export_slide_impl(slide_idx: int):
        image_data = _export_slide_png(slide_idx)
        return image_data, _slide_size_in_points()

    def _slide_size_in_points():
        # Looked up through the module, so macOS gets its own. A slide size
        # that cannot be read costs the scale, not the preview.
        from ppt_com import layout

        try:
            size = layout._get_slide_size_impl()
            return size["width_points"], size["height_points"]
        except Exception:
            logger.debug("Cannot read the slide size for the preview", exc_info=True)
            return None

    def _export_slide_png(slide_idx: int):
        if IS_MACOS:
            # PowerPoint for Mac has no Slide.Export. The preview is rendered
            # from the deck's PDF export instead; see ppt_mac/export.py.
            from ppt_mac.export import render_slide_png

            goto_slide(ppt._get_app_impl(), slide_idx)
            return render_slide_png(slide_idx)

        app = ppt._get_app_impl()
        pres = ppt._get_pres_impl()
        goto_slide(app, slide_idx)

        # Validate slide
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            raise ValueError(
                f"Slide index {slide_idx} out of range (1-{pres.Slides.Count})"
            )

        slide = pres.Slides(slide_idx)

        # Generate temp file path
        temp_dir = tempfile.gettempdir()
        temp_file = os.path.join(temp_dir, f"ppt_slide_preview_{slide_idx}.png")

        try:
            # Export slide as PNG
            slide.Export(temp_file, "PNG")

            # Read binary data
            with open(temp_file, "rb") as f:
                image_data = f.read()

            return image_data
        finally:
            # Cleanup
            if os.path.exists(temp_file):
                os.remove(temp_file)

    image_data, slide_size = await run_offloaded(
        ppt.execute, _export_slide_impl, params.slide_index)
    return [Image(data=image_data, format="png"),
            json.dumps(preview_scale(params.slide_index, image_data, slide_size))]


def preview_scale(slide_index, image_data, slide_size):
    """What ppt_get_slide_preview says beside its image, to convert between
    the image's pixels and the slide's points.

    The image of a 960 pt slide is 1920 px across on Windows, and a caller
    measuring a chart in it took the pixels for points twice and concluded
    the chart was ignoring its width (#194).
    """
    said = {"slide_index": slide_index}
    size = png_size(image_data)
    if size:
        said["image_width_px"], said["image_height_px"] = size
    if slide_size:
        said["slide_width_pt"], said["slide_height_pt"] = slide_size
    if size and slide_size and slide_size[0]:
        said["pixels_per_point"] = round(size[0] / slide_size[0], 4)
    return said


def png_size(data):
    """(width, height) in pixels from a PNG's header, or None."""
    import struct

    if not data or data[:8] != b"\x89PNG\r\n\x1a\n" or data[12:16] != b"IHDR":
        return None
    return struct.unpack(">II", data[16:24])


# =============================================================================
# Tool schema post-processing
# =============================================================================


def _inline_schema_defs(server) -> int:
    """Inline ``$defs`` references in every registered tool's input schema.

    Every tool takes a single Pydantic model (``params: XInput``), and the SDK
    renders that as ``{"params": {"$ref": "#/$defs/XInput"}}`` with the model's
    fields tucked away under ``$defs``. MCP clients such as Claude Desktop do
    not dereference ``$ref``, so ``params`` showed up as an empty object and
    callers had to guess field names from validation errors. Inlining keeps
    the wire schema self-contained; validation still runs against the Pydantic
    model and is unaffected.

    Returns the number of tools whose schema was rewritten.
    """

    def _resolve(node, defs, stack):
        if isinstance(node, list):
            return [_resolve(item, defs, stack) for item in node]
        if not isinstance(node, dict):
            return node
        ref = node.get("$ref")
        if isinstance(ref, str) and ref.startswith("#/$defs/"):
            name = ref[len("#/$defs/"):]
            target = defs.get(name)
            if target is not None and name not in stack:
                merged = dict(target)
                # Keep sibling keys (e.g. a field-level description) next to
                # the inlined definition.
                merged.update({k: v for k, v in node.items() if k != "$ref"})
                return _resolve(merged, defs, stack + [name])
            # Unknown target or a recursive schema: leave the reference alone.
        return {key: _resolve(value, defs, stack) for key, value in node.items()}

    rewritten = 0
    for tool in server._tool_manager.list_tools():
        schema = tool.parameters
        if not isinstance(schema, dict) or "$defs" not in schema:
            continue
        defs = schema["$defs"]
        inlined = _resolve({k: v for k, v in schema.items() if k != "$defs"}, defs, [])
        if "$ref" in json.dumps(inlined):
            # Something could not be resolved (recursive model); keep $defs so
            # the remaining references stay valid.
            inlined["$defs"] = defs
        tool.parameters = inlined
        rewritten += 1
    return rewritten


try:
    _inlined_tool_count = _inline_schema_defs(mcp)
    logger.debug("Inlined $defs in %d tool schemas", _inlined_tool_count)
except Exception:
    # _tool_manager and Tool.parameters are SDK internals, and the SDK has
    # renamed internals across majors before (see the MCPServer import above).
    # Losing the inlining costs clients an extra round trip; losing the import
    # would cost them every tool, so degrade instead of failing to start.
    logger.warning(
        "Could not inline $defs in tool schemas; they keep their $ref form",
        exc_info=True,
    )
    _inlined_tool_count = 0


def _slim(node):
    """A JSON schema without what Pydantic adds and no model reads.

    Every field gets a ``title`` that repeats its name, and every optional one
    is ``anyOf: [X, {"type": "null"}]`` with ``default: null``. Clients send
    these schemas with every request, and the two made up about a sixth of
    the 112k tokens the tools cost. A field left out still takes its default,
    so saying only X loses nothing; one whose null means something other than
    leaving it out, a default that is not null, keeps its anyOf.
    """
    if isinstance(node, list):
        return [_slim(item) for item in node]
    if not isinstance(node, dict):
        return node
    slim = {}
    for key, value in node.items():
        if key == "title" and isinstance(value, str):
            continue
        if key in ("properties", "$defs") and isinstance(value, dict):
            # Names here are the caller's: a parameter may be called "title".
            slim[key] = {name: _slim(field) for name, field in value.items()}
        else:
            slim[key] = _slim(value)
    branches = slim.get("anyOf")
    if (isinstance(branches, list) and "default" in slim and slim["default"] is None
            and {"type": "null"} in branches):
        rest = [b for b in branches if b != {"type": "null"}]
        if len(rest) == 1 and isinstance(rest[0], dict) and not set(rest[0]) & set(slim):
            slim.pop("anyOf")
            slim = {**rest[0], **slim}
    return slim


def _slim_tool_schemas(server) -> int:
    """Run every tool's input schema through _slim. Returns how many changed."""
    changed = 0
    for tool in server._tool_manager.list_tools():
        slim = _slim(tool.parameters)
        if slim != tool.parameters:
            tool.parameters = slim
            changed += 1
    return changed


try:
    _slimmed_tool_count = _slim_tool_schemas(mcp)
except Exception:
    # The same SDK internals as _inline_schema_defs. The full schemas are
    # still correct, only larger, so degrade instead of failing to start.
    logger.warning("Could not slim the tool schemas; they keep their full form", exc_info=True)
    _slimmed_tool_count = 0


def _add_presentation_arg(server) -> int:
    """Declare the per-call `presentation` argument in the input schemas.

    _PowerPointServer.call_tool is what honours it; this only makes it visible
    to clients. Tools listed in _TOOLS_WITHOUT_PRESENTATION_ARG are skipped.
    The argument sits next to `params` and is never required.

    Returns the number of tools that gained the argument.
    """
    added = 0
    for tool in server._tool_manager.list_tools():
        if tool.name in _TOOLS_WITHOUT_PRESENTATION_ARG:
            continue
        schema = dict(tool.parameters)
        properties = dict(schema.get("properties") or {})
        if PRESENTATION_ARG in properties:
            raise RuntimeError(f"{tool.name} already has a '{PRESENTATION_ARG}' argument")
        properties[PRESENTATION_ARG] = dict(_PRESENTATION_ARG_SCHEMA)
        schema["properties"] = properties
        tool.parameters = schema
        added += 1
    return added


try:
    _presentation_arg_count = _add_presentation_arg(mcp)
except Exception:
    # Same SDK internals as _inline_schema_defs, and the same reasoning: the
    # argument still works when a client sends it, it is only not advertised.
    logger.warning(
        "Could not add the 'presentation' argument to tool schemas",
        exc_info=True,
    )
    _presentation_arg_count = 0


# ---------------------------------------------------------------------------
# Catalog mode, opted into with PPT_TOOL_MODE=catalog
# ---------------------------------------------------------------------------
# A client sends every listed tool's schema with every request, and the 157
# come to about 112k tokens. In catalog mode the server lists three tools: one
# that answers the list of tools, a line each, and full schemas on request,
# one that runs any tool, and one that runs a list of calls. The tools
# themselves are unchanged and run through the same call_tool, so their
# validation, their checks and their layout findings are all as before. The
# three are registered either way; full mode leaves them off the list.
#
# Full mode is the default: catalog mode lists other names than the tools'
# own, so a client's rules that approve or block a tool by name would match
# nothing once a release turned it on unasked.

def _tool_mode(value):
    """'catalog' or 'full', from PPT_TOOL_MODE."""
    value = value.strip().lower()
    if value == "catalog":
        return "catalog"
    if value in ("full", "all", ""):
        return "full"
    # Read at import, so a typo must not stop the server from starting.
    logger.warning(
        "PPT_TOOL_MODE=%r is not a setting; use 'full' (the default) to list every "
        "tool, or 'catalog' to list three. Using 'full'.", value)
    return "full"


CATALOG_MODE = _tool_mode(os.getenv("PPT_TOOL_MODE", "full")) == "catalog"
_CATALOG_TOOLS = frozenset({"ppt_describe", "ppt_run", "ppt_run_steps"})


def _summary(description):
    """The first sentence of a tool's description, for its catalog line."""
    text = " ".join((description or "").split())
    for end in (". ", "! ", "? "):
        if end in text:
            text = text.split(end, 1)[0] + end.strip()
            break
    return text[:120]


def _catalog_lines():
    return "\n".join(
        f"{tool.name}: {_summary(tool.description)}"
        for tool in mcp._tool_manager.list_tools()
        if tool.name not in _CATALOG_TOOLS
    )


def _content_of(result):
    """What call_tool answered, as a list of content blocks under either SDK major."""
    if hasattr(result, "content") and not isinstance(result, (list, tuple, dict)):
        # mcp 2.x answers a CallToolResult, and an error in it is not raised.
        # A tool's own failure (see _flag_failure) is passed on as its blocks,
        # so ppt_run answers the same JSON and is marked failed in turn; a
        # ToolError raised here would reach the caller with the SDK's prefix.
        if getattr(result, "isError", False) or getattr(result, "is_error", False):
            text = _text_of(result.content)
            if _failure_payload(text) is None:
                raise ToolError(text or "The tool failed.")
        result = result.content
    if isinstance(result, tuple):  # mcp 1.x: (unstructured, structured)
        result = result[0]
    if isinstance(result, dict):
        return [json.dumps(result, ensure_ascii=False)]
    return list(result or [])


def _text_of(blocks):
    return "".join(getattr(b, "text", "") or "" for b in blocks if getattr(b, "type", "") == "text")


async def _run(tool, params, presentation):
    if not isinstance(tool, str) or not tool:
        raise ToolError("A step needs a 'tool': the name of the tool to run, "
                        "such as \"ppt_add_textbox\".")
    target = mcp._registered_tool(tool)
    if target is None or tool in _CATALOG_TOOLS:
        import difflib
        names = [t.name for t in mcp._tool_manager.list_tools() if t.name not in _CATALOG_TOOLS]
        close = difflib.get_close_matches(tool, names, n=3)
        hint = f" Did you mean {', '.join(close)}?" if close else ""
        raise ToolError(f"No tool called '{tool}'.{hint} The list is in ppt_describe's description.")
    arguments = {}
    declared = (target.parameters or {}).get("properties") or {}
    if "params" in declared:
        arguments["params"] = params or {}
    elif params:
        raise ToolError(f"{tool} takes no params.")
    if presentation is not None:
        arguments[PRESENTATION_ARG] = presentation
    return _content_of(await mcp.call_tool(tool, arguments))


def _resolve_refs(value, results):
    """Replace "$prev.key" and "$N.key" (N counts steps from 1) with what that
    step answered, read from its JSON result."""
    if isinstance(value, dict):
        return {k: _resolve_refs(v, results) for k, v in value.items()}
    if isinstance(value, list):
        return [_resolve_refs(v, results) for v in value]
    if not (isinstance(value, str) and value.startswith("$") and "." in value):
        return value
    which, key = value[1:].split(".", 1)
    if which == "prev":
        index = len(results) - 1
    elif which.isdigit():
        index = int(which) - 1
    else:
        return value
    if not 0 <= index < len(results) or results[index] is None:
        raise ToolError(f"'{value}' refers to a step that has no result.")
    node = results[index]
    for part in key.split("."):
        if isinstance(node, dict) and part in node:
            node = node[part]
        elif isinstance(node, list) and part.isdigit() and int(part) < len(node):
            node = node[int(part)]
        else:
            raise ToolError(f"'{value}': the step's result has no '{key}'.")
    return node


@mcp.tool(name="ppt_describe")
async def ppt_describe(names: list[str] | None = None) -> str:
    """The PowerPoint tools: call with no names first for the list of all
    of them, one line each; then with the names you are about to use, once
    each, for their full description and argument schema.

    Every PowerPoint tool is run through ppt_run (one call) or
    ppt_run_steps (several), not called directly, with exactly the
    arguments its schema gives. Start a task with ppt_describe().
    """
    # The list is an answer, not part of this description: clients cut
    # long descriptions short (Claude Code kept 54 of 157 lines), while a
    # tool's answer reaches the model whole.
    if not names:
        return _catalog_lines()
    found, missing = {}, []
    for name in names:
        tool = mcp._registered_tool(name)
        if tool is None or name in _CATALOG_TOOLS:
            missing.append(name)
            continue
        schema = dict(tool.parameters or {})
        schema.get("properties", {}).pop(PRESENTATION_ARG, None)
        params = (schema.get("properties") or {}).get("params")
        found[name] = {"description": tool.description, "params": params}
    result = {"tools": found}
    if missing:
        result["unknown"] = missing
    return json.dumps(result, ensure_ascii=False)

@mcp.tool(name="ppt_run")
async def ppt_run(tool: str, params: dict | None = None, presentation: str | None = None):
    """Run one PowerPoint tool: `tool` is its name, `params` its arguments
    as ppt_describe gives them, `presentation` the deck it works on
    (optional, as on every tool). Answers exactly what the tool answers."""
    return await _run(tool, params, presentation)

@mcp.tool(name="ppt_run_steps")
async def ppt_run_steps(steps: list[dict], presentation: str | None = None,
                        stop_on_error: bool = True):
    """Run several PowerPoint tools in order, in one call.

    Each step is {"tool": name, "params": {...}}. A value "$prev.key" is
    the previous step's result field, and "$2.shape_name" the second
    step's, so a step can use the shape the one before it made. The answer
    lists each step's result; images (slide previews) follow it. A step
    fails when its tool raises or answers {"error": ...}; with
    stop_on_error, the first failing step ends the run and says why.
    """
    report, results, images, failed = [], [], [], []
    for number, step in enumerate(steps, start=1):
        step = step if isinstance(step, dict) else {}
        tool = step.get("tool")
        try:
            params = _resolve_refs(step.get("params") or {}, results)
            blocks = await _run(tool, params, step.get("presentation", presentation))
        except Exception as exc:  # noqa: BLE001 - reported per step
            report.append({"step": number, "tool": tool, "error": str(exc)})
            results.append(None)
            failed.append((number, tool, str(exc)))
            if stop_on_error:
                break
            continue
        text = _text_of(blocks)
        try:
            data = json.loads(text) if text else {}
        except ValueError:
            data = {"text": text}
        if isinstance(data, dict) and data.get("error"):
            # Most tools answer a failure with {"error": ...} rather than
            # raising it (#249), so that answer fails the step too: it stops
            # the run, and a reference to it says the step has no result.
            entry = {"step": number, "tool": tool, "error": data["error"]}
            details = {k: v for k, v in data.items() if k != "error"}
            if details:
                entry["details"] = details
            report.append(entry)
            results.append(None)
            failed.append((number, tool, str(data["error"])))
            if stop_on_error:
                break
            continue
        report.append({"step": number, "tool": tool, "result": data})
        images += [b for b in blocks if getattr(b, "type", "") == "image"]
        results.append(data)
    summary = {"steps": report, "ran": len(report), "of": len(steps)}
    if failed:
        # A top-level "error", so the run is marked failed (isError true) by
        # the same rule as any tool's own answer; see _flag_failure.
        number, tool, message = failed[0]
        more = f" ({len(failed)} steps failed)" if len(failed) > 1 else ""
        summary["error"] = f"Step {number} ({tool}) failed: {message}{more}"
    summary = json.dumps(summary, ensure_ascii=False)
    return [summary, *images] if images else summary

if CATALOG_MODE:
    try:
        # MCPServer.instructions is read-only under both majors; the low-level
        # server it reads from is _mcp_server under 1.x, _lowlevel_server under 2.x.
        lowlevel = getattr(mcp, "_lowlevel_server", None) or mcp._mcp_server
        lowlevel.instructions = (
            "Catalog mode: call ppt_describe() for the list of PowerPoint tools, "
            "ppt_describe(names) for their arguments, and run them with ppt_run or "
            "ppt_run_steps. Wherever the instructions say to call a tool, do it that way.\n\n"
            + (mcp.instructions or "")
        )
    except Exception:
        logger.warning("Could not add the catalog note to the instructions", exc_info=True)


def main():
    """Entry point for the MCP server."""
    mcp.run()


if __name__ == "__main__":
    main()
