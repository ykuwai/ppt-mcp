"""Text content, formatting and manipulation tools, on Apple Events.

Mirrors ``ppt_com/text.py``. Same function names, same signatures, same
returned shapes; what differs is the walk through PowerPoint's object model.
The pseudo Markdown exporter and the typography checker are mostly arithmetic
over text, so their algorithms are untouched and only the accessors moved.

Four things about text on this side shape the whole module.

**Font formatting only exists under ``font of text range``.** Reaching a font
any other way answers -10006, so every font write here goes through
``text_range.font``.

**There is no ``run``.** ``text range`` has characters, words, sentences, lines
and paragraphs, and nothing between a character and a paragraph. Runs are
rebuilt by reading each character's formatting in one bulk event and grouping
the neighbours that match. When PowerPoint refuses the bulk read the range is
reported as uniformly formatted rather than having runs invented for it.

**There is no Find or Replace.** Both are done in Python over the text that
comes back, and a replacement is written into just its own character range so
the formatting on either side of it survives.

**``paragraph format`` lost the indents.** ``first line indent``, ``left
indent``, ``right indent`` and tab stops are all gone. ``indent level``
survives, on ``text range`` rather than on ``paragraph format``, which is also
where the COM code reads it, so bullet nesting still works. Anything asking for
a real indent is refused rather than quietly dropped.
"""

import logging
import re

from appscript import k
from appscript.reference import CommandError

from backend.mac_ae import (
    count,
    elements,
    is_missing,
    ppt,
    shape_by_name_or_index as _get_shape,
    shapes_of,
    windows_constant as _windows_constant,
)
from backend.mac_enums import (
    MsoShapeType,
    MsoTextOrientation,
    MsoThemeColorIndex,
    MsoVerticalAnchor,
    PpAutoSize,
    PpBulletType,
    PpNumberedBulletStyle,
    PpParagraphAlignment,
    PpPlaceholderType,
    to_keyword,
)
from backend.unsupported import refusal as _refusal
from utils.color import (
    get_theme_color_index,
    hex_to_rgb_list,
    rgb_list_to_hex,
)
from utils.navigation import goto_slide
from ppt_com.constants import (
    msoGroup, msoPlaceholder,
    ppAutoSizeNone, ppAutoSizeTextToFitShape,
    ppBulletNone, ppBulletNumbered,
)

logger = logging.getLogger(__name__)

# scripts/gen_mac_enums.py did not pair ppAutoSizeTextToFitShape, so the
# generated table stops at "shape to fit text". macOS does have the value. The
# sdef carries it twice under two spellings, `ppAutoSizeTextToFitShape` and
# `text to fit shape`, and both share the code 0x00e50002, so this is a gap in
# the generator rather than a gap in PowerPoint. Patched here so shrink to fit
# keeps working, and it should disappear the next time the table is
# regenerated.
_AUTO_SIZE = dict(PpAutoSize)
_AUTO_SIZE.setdefault(ppAutoSizeTextToFitShape, k.ppAutoSizeTextToFitShape)

# At module scope rather than inside the branch that uses it, so the name a
# caller passes can be checked before anything is written.
_VERTICAL_ANCHOR_MAP = {
    "top": 1,       # msoAnchorTop
    "middle": 3,    # msoAnchorMiddle
    "bottom": 4,    # msoAnchorBottom
}


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------
def _clean(value):
    """Turn `missing value` into None. Everything else passes through."""
    return None if is_missing(value) else value


def _text_of(text_range) -> str:
    """Read a range's content, treating `missing value` as empty."""
    return _clean(text_range.content()) or ""


def _require_text_frame(shape):
    """Raise the same message COM raises when a shape holds no text."""
    if not shape.has_text_frame():
        raise ValueError(f"Shape '{shape.name()}' does not have a text frame")


def _target_range(text_range, paragraph_index):
    """Return one paragraph, or the whole range when no index was given.

    Out of range says so. Indexing the reference directly would leave the
    caller with -1728, which names nothing.
    """
    if paragraph_index is None:
        return text_range
    paragraphs = elements(text_range.paragraphs)
    if paragraph_index < 1 or paragraph_index > len(paragraphs):
        raise ValueError(
            f"Paragraph index {paragraph_index} out of range "
            f"(1-{len(paragraphs)})"
        )
    return paragraphs[paragraph_index - 1]


# ---------------------------------------------------------------------------
# Runs, rebuilt from characters
# ---------------------------------------------------------------------------
def _character_runs(text_range):
    """Group a range's characters into runs of identical formatting.

    Returns a list of run dicts, or None when PowerPoint will not answer the
    bulk read. None means "no run detail available", not "one plain run", so
    the caller decides how to degrade rather than having a guess handed to it.

    Eight events cover the whole range no matter how long it is. The
    alternative, one event per character per property, is what makes this worth
    doing, because a 200 character shape would otherwise cost 1,600 round
    trips.
    """
    try:
        chars = text_range.characters
        texts = chars.content.get()
        if not isinstance(texts, list) or not texts:
            return None
        font = chars.font
        columns = [
            font.bold.get(),
            font.italic.get(),
            font.underline.get(),
            font.font_name.get(),
            font.east_asian_name.get(),
            font.font_size.get(),
            font.font_color.get(),
        ]
    except Exception:
        logger.debug("Bulk character read refused; runs unavailable",
                     exc_info=True)
        return None

    # A scalar instead of a list, or a short list, means PowerPoint answered
    # something other than one value per character. Zipping that would quietly
    # truncate the shape's text, so it is treated as no answer at all.
    for column in columns:
        if not isinstance(column, list) or len(column) != len(texts):
            return None

    bolds, italics, underlines, names, fareast, sizes, colors = columns

    runs = []
    position = 1
    for i, char in enumerate(texts):
        colour = colors[i]
        key = (
            _clean(bolds[i]) is True,
            _clean(italics[i]) is True,
            _clean(underlines[i]) is True,
            _clean(names[i]),
            _clean(fareast[i]),
            _clean(sizes[i]),
            tuple(colour) if isinstance(colour, list) else None,
        )
        if runs and runs[-1]["_key"] == key:
            runs[-1]["text"] += char
            runs[-1]["length"] += 1
        else:
            runs.append({
                "_key": key,
                "text": char,
                "start": position,
                "length": 1,
                "font_name": key[3],
                "font_name_fareast": key[4],
                "font_size": key[5],
                "bold": key[0],
                "italic": key[1],
                "underline": key[2],
                "color_hex": rgb_list_to_hex(colour),
            })
        position += 1

    for run in runs:
        del run["_key"]
    return runs


def _uniform_run(text_range):
    """Describe a range as one run, reading its formatting as a whole.

    The honest fallback when per character detail is unavailable. A property
    that differs across the range reads back as `missing value`, which arrives
    here as None rather than as a value that was never true of the whole range.
    """
    text = _text_of(text_range)
    font = text_range.font

    def read(accessor):
        """One property, or None when PowerPoint will not answer for a range."""
        try:
            return _clean(accessor())
        except CommandError:
            return None

    return [{
        "text": text,
        "start": 1,
        "length": len(text),
        "font_name": read(font.font_name),
        "font_name_fareast": read(font.east_asian_name),
        "font_size": read(font.font_size),
        "bold": read(font.bold) is True,
        "italic": read(font.italic) is True,
        "underline": read(font.underline) is True,
        "color_hex": rgb_list_to_hex(read(font.font_color)),
    }]


# ---------------------------------------------------------------------------
# Font and character range writing
# ---------------------------------------------------------------------------
def _apply_font_props(font, font_name, font_name_fareast, font_size, bold,
                      italic, underline, color, font_color_theme):
    """Apply font properties to a `font of text range` reference.

    Returns a list of sentences for anything that did not land, empty when it
    all did.
    """
    warnings = []
    if font_name is not None:
        font.font_name.set(font_name)
        # Windows' documented behaviour, that a Latin name also becomes the
        # East Asian one unless overridden. It does not always take here.
        # PowerPoint for Mac accepts the write and keeps the old value when the
        # font has no East Asian glyphs, saying nothing about it. Setting
        # Arial left Hiragino Sans in place; setting Meiryo replaced it. So it
        # is asked for and then read back, and a refusal is reported rather
        # than assumed to have worked.
        font.east_asian_name.set(font_name)
        if font_name_fareast is None and not _east_asian_took(font, font_name):
            warnings.append(
                f"The East Asian font was left as it was. '{font_name}' has no "
                "East Asian glyphs, and PowerPoint for Mac keeps the old East "
                "Asian font in that case without reporting anything. Japanese "
                "and Chinese text is still in the previous font. Pass "
                "font_name_fareast to choose one."
            )
    if font_name_fareast is not None:
        font.east_asian_name.set(font_name_fareast)  # override East Asian independently
        if not _east_asian_took(font, font_name_fareast):
            warnings.append(
                f"font_name_fareast '{font_name_fareast}' was not applied. "
                "PowerPoint for Mac keeps the East Asian font it had when the "
                "one it is given has no East Asian glyphs, and says nothing. "
                "Check the spelling, and that the font is installed."
            )
    if font_size is not None:
        font.font_size.set(font_size)
    if bold is not None:
        font.bold.set(bool(bold))
    if italic is not None:
        font.italic.set(bool(italic))
    if underline is not None:
        font.underline.set(bool(underline))
    if color is not None:
        font.font_color.set(hex_to_rgb_list(color))
    if font_color_theme is not None:
        font.font_color_theme_index.set(
            to_keyword(
                MsoThemeColorIndex,
                get_theme_color_index(font_color_theme),
                "theme color",
            )
        )
    return warnings


def _east_asian_took(font, wanted: str) -> bool:
    """Whether the East Asian font is the one that was just asked for."""
    try:
        current = font.east_asian_name()
    except CommandError:
        # No way to tell, so nothing is claimed either way. A warning that
        # might be wrong is worse than none.
        return True
    return is_missing(current) or str(current) == wanted


def _character_range(text_range, start, length):
    """Reference the characters COM would call ``Characters(start, length)``.

    AppleScript's ``thru`` is inclusive at both ends, so a run of ``length``
    characters beginning at ``start`` ends at ``start + length - 1``.
    """
    return text_range.characters[start:start + length - 1]


def _apply_font_to_range(text_range, start, length, props):
    """Format one character range, one character at a time if it has to be.

    A range specifier is a single reference, so the whole span is normally one
    event. PowerPoint is not documented to accept a property write against a
    plural reference though, so a refusal falls back to writing each character
    on its own rather than reporting a success that never happened.
    """
    if length < 1:
        return []
    try:
        return _apply_font_props(
            _character_range(text_range, start, length).font, **props
        )
    except CommandError:
        logger.debug("Range font write refused; falling back per character",
                     exc_info=True)
    warnings = []
    for i in range(start, start + length):
        for note in _apply_font_props(text_range.characters[i].font, **props):
            # The same refusal once per character is noise, not information.
            if note not in warnings:
                warnings.append(note)
    return warnings


def _apply_highlight(text_range, highlight_color, start=None, length=None):
    """Apply a text highlight, or say why clearing one is not on offer.

    Returns None when the highlight was applied, or a sentence explaining the
    refusal. macOS has a real ``highlight color`` on ``font``, so setting one
    is a single write and needs none of the Windows ClearFormatting dance. What
    it has no word for is removing one. There is no "no highlight" value in the
    dictionary, and guessing at ``missing value`` would be exactly the write
    that reports success and does nothing.
    """
    if highlight_color.lower() == "clear":
        return (
            "highlight_color='clear' was not applied. PowerPoint for Mac's "
            "font carries a highlight color but no word for removing one, and "
            "the Windows workaround goes through a ribbon command that has no "
            "Apple Event equivalent. The existing highlight is unchanged."
        )

    rgb = hex_to_rgb_list(highlight_color)
    if start is not None and length is not None:
        if length < 1:
            return None
        try:
            _character_range(text_range, start, length).font.highlight_color.set(rgb)
        except CommandError:
            for i in range(start, start + length):
                text_range.characters[i].font.highlight_color.set(rgb)
    else:
        text_range.font.highlight_color.set(rgb)
    return None


# ---------------------------------------------------------------------------
# Helpers for ppt_get_all_text
# ---------------------------------------------------------------------------
def _is_all_bold(shape) -> bool:
    """Check if ALL text in a shape is bold.

    One read rather than a walk over runs. A range whose characters disagree
    answers `missing value`, which is precisely the "not all bold" case, so the
    cheap question and the right question are the same one here.
    """
    try:
        tr = shape.text_frame.text_range
        text = _text_of(tr).strip()
        if not text:
            return False
        return _clean(tr.font.bold()) is True
    except Exception:
        return False


def _runs_to_markdown(paragraph) -> str:
    """Convert a paragraph's runs to Markdown with bold/italic markers.

    Merges consecutive runs with identical formatting to avoid
    fragmented markers like **word1****word2**.
    """
    runs = _character_runs(paragraph)
    if runs is None:
        # No per character detail. Describe the paragraph as one run using its
        # formatting as a whole, which loses mixed emphasis rather than
        # inventing it.
        try:
            runs = _uniform_run(paragraph)
        except Exception:
            return _text_of(paragraph).replace("\r", "").replace("\v", "\n")

    raw = []
    for run in runs:
        text = run["text"].replace("\r", "").replace("\v", "\n")
        if not text:
            continue
        raw.append({"text": text, "bold": run["bold"], "italic": run["italic"]})

    if not raw:
        return _text_of(paragraph).replace("\r", "").replace("\v", "\n")

    # Merge consecutive runs with identical formatting
    merged = []
    for r in raw:
        if merged and merged[-1]["bold"] == r["bold"] and merged[-1]["italic"] == r["italic"]:
            merged[-1]["text"] += r["text"]
        else:
            merged.append(dict(r))

    # Format
    parts = []
    for m in merged:
        t = m["text"]
        if m["bold"] and m["italic"]:
            parts.append(f"***{t}***")
        elif m["bold"]:
            parts.append(f"**{t}**")
        elif m["italic"]:
            parts.append(f"*{t}*")
        else:
            parts.append(t)

    return "".join(parts)


def _plain_text(text_range) -> str:
    """Extract plain text from a text range, stripping formatting markers."""
    try:
        return _text_of(text_range).replace("\r", " ").replace("\v", " ").strip()
    except Exception:
        return ""


def _shape_paragraphs_to_markdown(shape, as_heading: str = "") -> str:
    """Convert a shape's paragraphs to Markdown text.

    Args:
        shape: shape reference with a text frame
        as_heading: If set (e.g. "#" or "##"), render as heading
    """
    try:
        tr = shape.text_frame.text_range
    except Exception:
        return ""

    if as_heading:
        # For headings, use plain text (bold markers are redundant for # and ##)
        text = _plain_text(tr)
        if not text:
            return ""
        return f"{as_heading} {text}"

    numbered = to_keyword(PpBulletType, ppBulletNumbered, "bullet type")

    lines = []
    for para in elements(tr.paragraphs):
        text = _runs_to_markdown(para).strip()
        if not text:
            # Preserve empty paragraph as blank line
            lines.append("")
            continue

        # Detect bullet. `indent level` lives on the text range here rather
        # than on the paragraph format, which is also where COM reads it.
        indent_level = _clean(para.indent_level()) or 1
        bullet_prefix = ""
        try:
            bullet = para.paragraph_format.bullet_format
            if bullet.visible() is True:
                indent = "  " * max(0, indent_level - 1)
                if bullet.bullet_type() == numbered:
                    bullet_prefix = f"{indent}1. "
                else:
                    bullet_prefix = f"{indent}- "
        except Exception:
            pass

        lines.append(f"{bullet_prefix}{text}")

    return "\n".join(lines)


def _table_to_markdown(shape) -> str:
    """Convert a table shape to a Markdown table.

    Note: Cell text is extracted as plain text; inline bold/italic
    formatting within table cells is not preserved.
    """
    try:
        table = shape.table_object
        rows = shape.number_of_rows()
        cols = shape.number_of_columns()

        md_rows = []
        for r in range(1, rows + 1):
            cells = []
            for c in range(1, cols + 1):
                try:
                    # By index, never `get cell from`. The command answers
                    # with a reference that will not resolve, so every cell
                    # came back empty here. See ppt_mac/tables.py.
                    cell = table.rows[r].cells[c]
                    text = _text_of(cell.shape.text_frame.text_range)
                    text = text.replace("\r", " ").replace("\v", " ").replace("|", "\\|").strip()
                except Exception:
                    text = ""
                cells.append(text)
            md_rows.append("| " + " | ".join(cells) + " |")

            # Add header separator after first row
            if r == 1:
                md_rows.append("| " + " | ".join(["---"] * cols) + " |")

        return "\n".join(md_rows)
    except Exception:
        return ""


def _collect_text_shapes(slide) -> list:
    """Collect all text-bearing shapes from a slide with position info.

    Returns a list of dicts with keys:
        shape, top, left, width, height, is_title, is_subtitle,
        has_table, is_group
    Skips SlideNumber, Header, Footer, Date placeholders.
    """
    from ppt_com.text import (
        _SKIP_PLACEHOLDER_TYPES,
        _SUBTITLE_PLACEHOLDER_TYPES,
        _TITLE_PLACEHOLDER_TYPES,
    )
    placeholder_kind = to_keyword(MsoShapeType, msoPlaceholder, "shape type")
    group_kind = to_keyword(MsoShapeType, msoGroup, "shape type")

    shapes = []

    def _process_shape(shape, offset_top=0.0, offset_left=0.0):
        """Process a single shape (may be called recursively for groups)."""
        # Check placeholder skip / classify
        is_title = False
        is_subtitle = False
        shape_kind = shape.shape_type()
        if shape_kind == placeholder_kind:
            try:
                # `placeholder type` is declared on `place holder`, and a shape
                # reached through `shapes` still answers it when the object
                # really is one. The try/except is what covers the case where
                # it does not, exactly as on the COM side.
                ph_type = _windows_constant(
                    PpPlaceholderType, shape.placeholder_type()
                )
                if ph_type in _SKIP_PLACEHOLDER_TYPES:
                    return
                is_title = ph_type in _TITLE_PLACEHOLDER_TYPES
                is_subtitle = ph_type in _SUBTITLE_PLACEHOLDER_TYPES
            except Exception:
                pass

        # Recurse into groups early (no need to build info dict).
        # Pass group's position as offset since child coordinates are
        # relative to the group, not the slide.
        if shape_kind == group_kind:
            try:
                g_top = shape.top()
                g_left = shape.left_position()
                for child in shapes_of(shape):
                    _process_shape(child, offset_top + g_top, offset_left + g_left)
            except Exception:
                pass
            return

        info = {
            "shape": shape,
            "top": shape.top() + offset_top,
            "left": shape.left_position() + offset_left,
            "width": shape.width(),
            "height": shape.height(),
            "is_title": is_title,
            "is_subtitle": is_subtitle,
            "has_table": False,
        }

        # Check for table
        try:
            if shape.has_table():
                info["has_table"] = True
                shapes.append(info)
                return
        except Exception:
            pass

        # Check for text
        try:
            if shape.has_text_frame():
                if shape.text_frame.has_text():
                    shapes.append(info)
        except Exception:
            pass

    for shape in shapes_of(slide):
        _process_shape(shape)

    return shapes


def _shape_info_to_markdown(info: dict, subheading_level: str = "##") -> str:
    """Convert a single shape_info dict to Markdown text."""
    shape = info["shape"]

    # Table
    if info["has_table"]:
        return _table_to_markdown(shape)

    # Title
    if info["is_title"]:
        return _shape_paragraphs_to_markdown(shape, as_heading="#")

    # A subtitle is plain text, with no heading marker
    if info["is_subtitle"]:
        return _shape_paragraphs_to_markdown(shape)

    # All-bold → subheading (level depends on context)
    if _is_all_bold(shape):
        return _shape_paragraphs_to_markdown(shape, as_heading=subheading_level)

    # Regular text
    return _shape_paragraphs_to_markdown(shape)


def _slide_to_markdown(slide, slide_index: int) -> str:
    """Convert a single slide to pseudo-Markdown.

    The layout algorithm is the COM module's, unchanged; only the reads moved.
    """
    # The row and column grouping is pure arithmetic over the dicts above, so
    # it is shared rather than copied. Imported here rather than at module
    # scope because ppt_com/text.py imports this module at the bottom of its
    # own file, so an "import ppt_mac.text first" ordering would otherwise run
    # that swap block against a module that has defined nothing yet, and the
    # swap would silently not happen. By call time both are fully loaded.
    from ppt_com.text import _group_into_columns, _group_into_rows

    # Check if slide is hidden
    hidden = ""
    try:
        if slide.slide_show_transition.hidden():
            hidden = " (hidden)"
    except Exception:
        pass
    parts = [f"== Slide {slide_index}{hidden} =="]

    shape_infos = _collect_text_shapes(slide)
    if not shape_infos:
        parts.append("(no text)")
        return "\n".join(parts)

    rows = _group_into_rows(shape_infos)
    has_multi_shape_rows = any(len(row) > 1 for row in rows)

    if not has_multi_shape_rows:
        # Simple layout: all single-shape rows
        for row in rows:
            md = _shape_info_to_markdown(row[0])
            if md.strip():
                parts.append(md)
    else:
        # Mixed layout: interleave full-width and column groups
        # in original Y-order.  Consecutive multi-shape rows are
        # collected and flushed as a column group together.
        pending_column_shapes = []

        def _flush_columns():
            """Group pending column shapes by X and append to parts."""
            if not pending_column_shapes:
                return
            columns = _group_into_columns(pending_column_shapes)
            for col_idx, col in enumerate(columns):
                if col_idx > 0:
                    parts.append("")  # blank line between columns
                for info in col:
                    md = _shape_info_to_markdown(info, subheading_level="###")
                    if md.strip():
                        parts.append(md)
            pending_column_shapes.clear()

        for row in rows:
            if len(row) == 1:
                # Flush any pending column shapes before this full-width row
                _flush_columns()
                md = _shape_info_to_markdown(row[0])
                if md.strip():
                    parts.append(md)
            else:
                # Collect column shapes from consecutive multi-shape rows
                pending_column_shapes.extend(row)

        # Flush remaining column shapes at the end
        _flush_columns()

    return "\n".join(parts)


def _get_all_text_impl(slide_indices) -> str:
    """Extract all text from the presentation as pseudo-Markdown.

    Runs on the Apple Event thread.
    """
    pres = ppt._get_pres_impl()
    total_slides = count(pres.slides)

    if slide_indices is None:
        indices = list(range(1, total_slides + 1))
    else:
        indices = slide_indices

    slide_parts = []
    for idx in indices:
        if idx < 1 or idx > total_slides:
            slide_parts.append(
                f"== Slide {idx} ==\n(invalid slide index, "
                f"presentation has {total_slides} slides)"
            )
            continue
        slide = pres.slides[idx]
        slide_parts.append(_slide_to_markdown(slide, idx))

    return "\n\n".join(slide_parts)


# ---------------------------------------------------------------------------
# Apple Event implementation functions
# ---------------------------------------------------------------------------
def _set_text_impl(slide_index: int, shape_name_or_index, text: str) -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]
    shape = _get_shape(slide, shape_name_or_index)

    _require_text_frame(shape)

    tr = shape.text_frame.text_range
    text = text.replace('\n', '\r')  # \n -> paragraph break, CR here too
    # \v (vertical tab) -> line break (Shift+Enter), passed through as it is
    tr.content.set(text)

    # Nothing is trusted because it did not raise.
    written = _clean(tr.content())
    if text and not written:
        raise RuntimeError(
            f"PowerPoint reported no error but shape '{shape.name()}' is still "
            "empty. The text was not written."
        )

    return {
        "status": "success",
        "slide_index": slide_index,
        "shape_name": shape.name(),
        "text_length": tr.text_length(),
        "paragraph_count": count(tr.paragraphs),
    }


def _get_text_impl(slide_index: int, shape_name_or_index) -> dict:
    ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]
    shape = _get_shape(slide, shape_name_or_index)

    _require_text_frame(shape)

    tr = shape.text_frame.text_range

    result = {
        "status": "success",
        "shape_name": shape.name(),
        "text": _text_of(tr),
        "text_length": tr.text_length(),
        "paragraph_count": count(tr.paragraphs),
    }

    paragraphs = []
    for i, para in enumerate(elements(tr.paragraphs), start=1):
        paragraphs.append({
            "index": i,
            "text": _text_of(para),
            "indent_level": _clean(para.indent_level()),
            "alignment": _windows_constant(
                PpParagraphAlignment, para.paragraph_format.alignment()
            ),
        })
    result["paragraphs"] = paragraphs

    runs = _character_runs(tr)
    if runs is None:
        runs = _uniform_run(tr)
        result["runs_note"] = (
            "PowerPoint has no run element on macOS, and it refused the bulk "
            "character read this shape's runs are rebuilt from, so the text is "
            "reported as one run using the formatting of the range as a whole."
        )
    result["runs"] = [dict(run, index=i) for i, run in enumerate(runs, start=1)]

    return result


def _format_text_impl(slide_index, shape_name_or_index,
                      font_name, font_name_fareast, font_size, bold, italic, underline,
                      color, font_color_theme, highlight_color) -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]
    shape = _get_shape(slide, shape_name_or_index)

    _require_text_frame(shape)

    tr = shape.text_frame.text_range
    warnings = _apply_font_props(
        tr.font, font_name, font_name_fareast, font_size, bold, italic,
        underline, color, font_color_theme,
    )

    warning = None
    if highlight_color is not None:
        warning = _apply_highlight(tr, highlight_color)

    result = {
        "status": "success",
        "shape_name": shape.name(),
        "formatted_text": _text_of(tr),
        # COM reports the range's own start and length; on macOS those are
        # `offset` and `text length` of the same range.
        "start": tr.offset(),
        "length": tr.text_length(),
    }
    if warning:
        result["partial"] = True
        result["unsupported"] = ["highlight_color=clear"]
        warnings.append(warning)
    if warnings:
        result["warnings"] = warnings
    return result


def _format_text_range_impl(slide_index, shape_name_or_index, start, length,
                            search_text, occurrence,
                            font_name, font_name_fareast, font_size, bold, italic, underline,
                            color, font_color_theme, highlight_color) -> dict:
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]
    shape = _get_shape(slide, shape_name_or_index)

    _require_text_frame(shape)

    tr = shape.text_frame.text_range
    full_text = _text_of(tr)

    # Resolve search_text to start/length if provided
    if search_text is not None:
        pos = -1
        search_from = 0
        for i in range(occurrence):
            pos = full_text.find(search_text, search_from)
            if pos == -1:
                if i == 0:
                    raise ValueError(
                        f"search_text '{search_text}' not found in shape '{shape.name()}'"
                    )
                else:
                    raise ValueError(
                        f"search_text '{search_text}' has only {i} occurrence(s) "
                        f"in shape '{shape.name()}', but occurrence={occurrence} was requested"
                    )
            search_from = pos + len(search_text)
        # Character positions are 1-based here too
        start = pos + 1
        length = len(search_text)

    warnings = _apply_font_to_range(tr, start, length, {
        "font_name": font_name,
        "font_name_fareast": font_name_fareast,
        "font_size": font_size,
        "bold": bold,
        "italic": italic,
        "underline": underline,
        "color": color,
        "font_color_theme": font_color_theme,
    })

    warning = None
    if highlight_color is not None:
        warning = _apply_highlight(tr, highlight_color, start, length)

    result = {
        "status": "success",
        "shape_name": shape.name(),
        # Sliced from the text already read rather than asked for again. A
        # `thru` range answers one value per element, so reading its content
        # back would give a list of single characters instead of a string.
        "formatted_text": full_text[start - 1:start - 1 + length],
        "start": start,
        "length": length,
    }
    if warning:
        result["partial"] = True
        result["unsupported"] = ["highlight_color=clear"]
        warnings.append(warning)
    if warnings:
        result["warnings"] = warnings
    return result


def _set_paragraph_format_impl(slide_index, shape_name_or_index, paragraph_index,
                               alignment, line_spacing, space_before, space_after,
                               indent_level, first_line_indent) -> dict:
    from ppt_com.text import ALIGNMENT_MAP

    # Translated before goto_slide, so a misspelled alignment costs neither an
    # Apple Event nor a jump to a slide the caller was not looking at.
    align_word = None
    if alignment is not None:
        align_val = ALIGNMENT_MAP.get(alignment)
        if align_val is None:
            raise ValueError(
                f"Invalid alignment '{alignment}'. "
                f"Valid values: {list(ALIGNMENT_MAP.keys())}"
            )
        align_word = to_keyword(
            PpParagraphAlignment, align_val, "paragraph alignment"
        )

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]
    shape = _get_shape(slide, shape_name_or_index)

    _require_text_frame(shape)

    tr = shape.text_frame.text_range

    target = _target_range(tr, paragraph_index)

    pf = target.paragraph_format
    applied = 0

    if align_word is not None:
        pf.alignment.set(align_word)
        applied += 1

    if line_spacing is not None:
        pf.line_rule_within.set(True)
        pf.space_within.set(line_spacing)
        applied += 1

    if space_before is not None:
        pf.line_rule_before.set(False)
        pf.space_before.set(space_before)
        applied += 1

    if space_after is not None:
        pf.line_rule_after.set(False)
        pf.space_after.set(space_after)
        applied += 1

    if indent_level is not None:
        # `indent level` is on the text range here, not on the paragraph
        # format, which is also where the COM code sets it.
        target.indent_level.set(indent_level)
        applied += 1

    if first_line_indent is None:
        return {
            "status": "success",
            "shape_name": shape.name(),
            "paragraph_index": paragraph_index or "all",
        }

    reason = (
        "PowerPoint for Mac's paragraph format has no first line indent, and "
        "none of left indent, right indent or tab stops either. The text "
        "frame's ruler carries a first margin per indent level, but that "
        "applies to every paragraph at that level rather than to this one, so "
        "it is not used as a substitute here."
    )
    if applied == 0:
        return _refusal("ppt_set_paragraph_format", reason)

    return {
        "status": "success",
        "shape_name": shape.name(),
        "paragraph_index": paragraph_index or "all",
        "partial": True,
        "unsupported": ["first_line_indent"],
        "warnings": [reason],
    }


def _set_bullet_impl(slide_index, shape_name_or_index, paragraph_index,
                     bullet_type, bullet_char, bullet_start_value,
                     indent_level, color, size, font_name,
                     numbered_style, use_text_color, use_text_font) -> dict:
    from ppt_com.text import BULLET_TYPE_MAP, NUMBERED_STYLE_MAP

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]
    shape = _get_shape(slide, shape_name_or_index)

    _require_text_frame(shape)

    tr = shape.text_frame.text_range

    target = _target_range(tr, paragraph_index)

    # bullet_type has already been coerced to "numbered" at the Pydantic
    # layer when numbered_style is set, so we can trust it here.
    bullet_type_val = BULLET_TYPE_MAP[bullet_type]

    bullet = target.paragraph_format.bullet_format

    if bullet_type_val == ppBulletNone:
        bullet.visible.set(False)
    else:
        bullet.visible.set(True)
        bullet.bullet_type.set(
            to_keyword(PpBulletType, bullet_type_val, "bullet type")
        )

    if bullet_char is not None:
        # COM takes the character's code point; macOS takes the character.
        bullet.bullet_character.set(bullet_char[0])

    if bullet_start_value is not None:
        bullet.bullet_start_value.set(bullet_start_value)

    if indent_level is not None:
        target.indent_level.set(indent_level)

    if numbered_style is not None:
        bullet.bullet_style.set(
            to_keyword(
                PpNumberedBulletStyle,
                NUMBERED_STYLE_MAP[numbered_style],
                "numbered bullet style",
            )
        )

    if size is not None:
        bullet.relative_size.set(size)

    # Apply UseText* toggles first; an explicit color/font_name below
    # supersedes them.
    if use_text_color is not None:
        bullet.use_text_color.set(bool(use_text_color))
    if use_text_font is not None:
        bullet.use_text_font.set(bool(use_text_font))

    if color is not None:
        bullet.bullet_font.font_color.set(hex_to_rgb_list(color))

    if font_name is not None:
        bullet.bullet_font.font_name.set(font_name)

    # An explicit color or font_name turns the corresponding UseText* flag off.
    # Reflect that effective state so callers see what was applied rather than
    # what they asked for.
    effective_use_text_color = False if color is not None else use_text_color
    effective_use_text_font = False if font_name is not None else use_text_font

    # Read back the two that decide whether the paragraph is bulleted at all,
    # plus the indent when one was asked for. Those are the properties this
    # module already reads elsewhere, so the route is known to work. The rest
    # of the bullet format is written blind and the warning below says so
    # rather than letting the echo pass for evidence.
    measured_type = _measured_bullet_type(bullet)
    measured_indent = indent_level
    if indent_level is not None:
        try:
            measured_indent = _clean(target.indent_level()) or indent_level
        except CommandError:
            pass

    result = {
        "status": "success",
        "shape_name": shape.name(),
        "paragraph_index": paragraph_index or "all",
        "bullet_type": measured_type or bullet_type,
        "numbered_style": numbered_style,
        "indent_level": measured_indent,
        "color_hex": color,
        "size": size,
        "font_name": font_name,
        "use_text_color": effective_use_text_color,
        "use_text_font": effective_use_text_font,
    }

    warnings = []
    if measured_type is not None and measured_type != bullet_type:
        warnings.append(
            f"A {bullet_type} bullet was asked for and the paragraph reads "
            f"back as {measured_type}, so PowerPoint did not take it."
        )
    unmeasured = [
        name for name, value in (
            ("numbered_style", numbered_style), ("color_hex", color),
            ("size", size), ("font_name", font_name),
            ("use_text_color", use_text_color), ("use_text_font", use_text_font),
        )
        if value is not None
    ]
    if unmeasured:
        warnings.append(
            f"{', '.join(unmeasured)} is what was asked for rather than a "
            "measurement. PowerPoint for Mac answers nothing for those parts "
            "of a bullet format, so only the bullet type and the indent level "
            "above were read back."
        )
    if warnings:
        result["warnings"] = warnings
    return result


def _measured_bullet_type(bullet):
    """Read a bullet back, in the words the tool takes it in.

    ``visible`` and ``bullet type`` are the pair ``ppt_get_slide_markdown``
    already reads, which is why these two are read and the rest of the bullet
    format is not. None means PowerPoint would not say.
    """
    from ppt_com.text import BULLET_TYPE_MAP

    try:
        if not bullet.visible():
            return "none"
        word = bullet.bullet_type()
    except (CommandError, AttributeError):
        return None
    if is_missing(word):
        return None
    for value, keyword in PpBulletType.items():
        if keyword == word:
            for name, mapped in BULLET_TYPE_MAP.items():
                if mapped == value:
                    return name
            return None
    return None


def _replace_characters(text_range, start, length, new_text):
    """Overwrite one character range, keeping the formatting around it.

    Returns True when the range write was accepted, False when PowerPoint
    refused it and the caller has to rewrite the whole frame instead, which
    flattens the shape's formatting to whatever the first run carries.
    """
    try:
        _character_range(text_range, start, length).content.set(new_text)
        return True
    except CommandError:
        logger.debug("Character range write refused; rewriting the frame",
                     exc_info=True)
        return False


def _find_replace_text_impl(
    find_text,
    replace_text,
    dry_run,
    match_case,
    whole_words,
    slide_indices,
    shape_name,
    context_chars,
) -> dict:
    from ppt_com.text import _build_context

    pres = ppt._get_pres_impl()
    find_only = replace_text is None or dry_run

    # PowerPoint's dictionary has no Find and no Replace on `text range`, so
    # both are done here over the text the shape hands back. match_case and
    # whole_words become regex flags, which is as close as the two get.
    pattern = re.escape(find_text)
    if whole_words:
        pattern = r"\b" + pattern + r"\b"
    matcher = re.compile(pattern, 0 if match_case else re.IGNORECASE)

    total = count(pres.slides)
    if slide_indices is not None:
        for i in slide_indices:
            if i > total:
                raise ValueError(
                    f"slide_indices entry {i} out of range (1-{total})"
                )
        indices = list(slide_indices)
    else:
        indices = list(range(1, total + 1))

    hits = []
    warnings = []
    for slide_index in indices:
        slide = pres.slides[slide_index]
        for shape in shapes_of(slide):
            if shape_name is not None and shape.name() != shape_name:
                continue
            if not shape.has_text_frame():
                continue
            tr = shape.text_frame.text_range
            content = _text_of(tr)
            if not content:
                continue

            if find_only:
                for match in matcher.finditer(content):
                    hit = {
                        "slide_index": slide_index,
                        "shape_name": shape.name(),
                        "start": match.start() + 1,
                        "length": len(match.group(0)),
                    }
                    if context_chars > 0:
                        hit["context"] = _build_context(
                            content, match.start() + 1, len(match.group(0)),
                            context_chars,
                        )
                    hits.append(hit)
            else:
                # Walk forward the way COM does, so a replacement containing
                # the search text ("foo" -> "foobar") cannot loop forever, and
                # a deletion still makes progress.
                cursor = 0
                while True:
                    match = matcher.search(content, cursor)
                    if match is None:
                        break
                    start = match.start()
                    matched_length = match.end() - match.start()

                    if not _replace_characters(
                        tr, start + 1, matched_length, replace_text
                    ):
                        rewritten = (
                            content[:start] + replace_text
                            + content[start + matched_length:]
                        )
                        tr.content.set(rewritten)
                        note = (
                            f"Shape '{shape.name()}' on slide {slide_index} "
                            "refused a character range write, so its whole "
                            "text frame was rewritten and any mixed formatting "
                            "inside it is now uniform."
                        )
                        if note not in warnings:
                            warnings.append(note)

                    # Nothing is trusted because it did not raise.
                    content = _text_of(tr)
                    if content[start:start + len(replace_text)] != replace_text:
                        raise RuntimeError(
                            "PowerPoint reported no error but the replacement "
                            f"did not land in shape '{shape.name()}' on slide "
                            f"{slide_index}."
                        )

                    hit = {
                        "slide_index": slide_index,
                        "shape_name": shape.name(),
                        "start": start + 1,
                        "length": len(replace_text),
                    }
                    if context_chars > 0:
                        hit["context"] = _build_context(
                            content, start + 1, len(replace_text), context_chars
                        )
                    hits.append(hit)
                    cursor = start + max(len(replace_text), 1)

    result = {
        "status": "success",
        "mode": "find" if find_only else "replace",
        "find_text": find_text,
        "replace_text": replace_text,
        "match_count": len(hits),
        "matches": hits,
    }
    if warnings:
        result["warnings"] = warnings
    return result


def _set_textframe_impl(slide_index, shape_name_or_index,
                        auto_size, word_wrap,
                        margin_left, margin_right, margin_top, margin_bottom,
                        orientation, vertical_anchor) -> dict:
    from ppt_com.text import AUTO_SIZE_MAP, ORIENTATION_MAP

    # All three names are checked before the first write. Checked where they
    # were used, a misspelled one handed the caller an error and a text frame
    # that had already taken four new margins and a new wrap setting.
    orient_val = None
    if orientation is not None:
        orient_val = ORIENTATION_MAP.get(orientation)
        if orient_val is None:
            raise ValueError(
                f"Invalid orientation '{orientation}'. "
                f"Valid values: {list(ORIENTATION_MAP.keys())}"
            )
    auto_size_val = None
    if auto_size is not None:
        auto_size_val = AUTO_SIZE_MAP.get(auto_size)
        if auto_size_val is None:
            raise ValueError(
                f"Invalid auto_size '{auto_size}'. "
                f"Valid values: {list(AUTO_SIZE_MAP.keys())}"
            )
    anchor_val = None
    if vertical_anchor is not None:
        anchor_val = _VERTICAL_ANCHOR_MAP.get(vertical_anchor.lower())
        if anchor_val is None:
            raise ValueError(
                f"Invalid vertical_anchor '{vertical_anchor}'. "
                f"Must be one of: {sorted(_VERTICAL_ANCHOR_MAP)}"
            )

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]
    shape = _get_shape(slide, shape_name_or_index)

    _require_text_frame(shape)

    tf = shape.text_frame

    if margin_left is not None:
        tf.margin_left.set(margin_left)
    if margin_right is not None:
        tf.margin_right.set(margin_right)
    if margin_top is not None:
        tf.margin_top.set(margin_top)
    if margin_bottom is not None:
        tf.margin_bottom.set(margin_bottom)
    if word_wrap is not None:
        tf.word_wrap.set(bool(word_wrap))
    if orient_val is not None:
        # `orientation` is read only here; `text orientation` is the settable
        # one and carries the same enumerators.
        tf.text_orientation.set(
            to_keyword(MsoTextOrientation, orient_val, "text orientation")
        )

    if auto_size_val is not None:
        # No TextFrame2 detour is needed, because `auto size` is on the text
        # frame itself and covers shrink to fit as well.
        tf.auto_size.set(to_keyword(_AUTO_SIZE, auto_size_val, "auto size"))

    if anchor_val is not None:
        tf.vertical_anchor.set(
            to_keyword(MsoVerticalAnchor, anchor_val, "vertical anchor")
        )

    return {
        "status": "success",
        "shape_name": shape.name(),
    }


# ---------------------------------------------------------------------------
# Typography check (widow line detection)
# ---------------------------------------------------------------------------
def _line_texts(text_range) -> list:
    """Return the text of every line in a range, in order.

    One bulk read where PowerPoint allows it. The widow fix re-measures a shape
    after every one point of extra width, so this runs up to twenty times per
    shape and the difference between one event and one per line is the
    difference between a check that finishes and one that hits the timeout.
    """
    try:
        values = text_range.lines.content.get()
    except CommandError:
        values = None
    if isinstance(values, list):
        return ["" if is_missing(v) else v for v in values]
    if values is not None and not is_missing(values):
        return [values]
    return [_text_of(line) for line in elements(text_range.lines)]


def _get_widows(shape, max_chars, max_words):
    """Return list of widow issues for a single shape."""
    from ppt_com.text import _is_latin

    lines = _line_texts(shape.text_frame.text_range)
    if len(lines) < 2:
        return []

    widows = []
    for li in range(2, len(lines) + 1):
        prev_text = lines[li - 2]
        # An explicit break (\r = paragraph, \n = soft return) is not a widow.
        if prev_text.endswith("\r") or prev_text.endswith("\n"):
            continue

        cur_text = lines[li - 1].rstrip("\r\n")
        if not cur_text:
            continue

        is_widow = False
        if _is_latin(cur_text):
            if len(cur_text.split()) <= max_words:
                is_widow = True
        else:
            if len(cur_text) <= max_chars:
                is_widow = True

        if is_widow:
            widows.append({
                "line_index": li,
                "line_text": cur_text,
                "char_count": len(cur_text),
                "prev_line_text": prev_text.rstrip("\r\n"),
            })
    return widows


def _get_short_vbreaks(shape, max_chars, max_words):
    """Return list of lines after an explicit \\v that are too short."""
    from ppt_com.text import _is_latin

    lines = _line_texts(shape.text_frame.text_range)
    if len(lines) < 2:
        return []

    short_breaks = []
    for li in range(2, len(lines) + 1):
        prev_text = lines[li - 2]
        # Only flag lines after an explicit \v (which reads back as \n).
        # \r is a paragraph break, which is deliberate structure, so it is skipped.
        if not prev_text.endswith("\n"):
            continue

        cur_text = lines[li - 1].rstrip("\r\n")
        if not cur_text:
            continue

        is_short = False
        if _is_latin(cur_text):
            if len(cur_text.split()) <= max_words:
                is_short = True
        else:
            if len(cur_text) <= max_chars:
                is_short = True

        if is_short:
            short_breaks.append({
                "line_index": li,
                "line_text": cur_text,
                "char_count": len(cur_text),
                "prev_line_text": prev_text.rstrip("\r\n"),
                "issue_type": "short_after_vbreak",
            })
    return short_breaks


def _right_neighbor_gap(shape, slide):
    """Find the gap (pt) to the nearest shape on the right that vertically overlaps."""
    s_left = shape.left_position()
    s_right = s_left + shape.width()
    s_top = shape.top()
    s_bottom = s_top + shape.height()
    name = shape.name()
    min_gap = float("inf")

    for other in shapes_of(slide):
        if other.name() == name:
            continue
        o_top = other.top()
        o_left = other.left_position()
        # Must vertically overlap
        if o_top + other.height() <= s_top or o_top >= s_bottom:
            continue
        # Must be to the right
        if o_left > s_right - 1:
            gap = o_left - s_right
            if gap < min_gap:
                min_gap = gap

    return min_gap


# A shape is allowed to sit right on the slide edge, so a fraction of a point
# past it is rounding rather than a mistake.
_OFF_SLIDE_TOLERANCE = 0.5


def _off_slide_edges(shape, slide_w, slide_h) -> list:
    """Which slide edges a shape hangs over, by name, or an empty list."""
    if not slide_w or not slide_h:
        return []
    try:
        left, top = _shape_left_top(shape)
        width, height = _shape_width_height(shape)
    except Exception:  # noqa: BLE001 - a shape that will not answer is skipped
        return []
    edges = []
    if left < -_OFF_SLIDE_TOLERANCE:
        edges.append("left")
    if top < -_OFF_SLIDE_TOLERANCE:
        edges.append("top")
    if left + width > slide_w + _OFF_SLIDE_TOLERANCE:
        edges.append("right")
    if top + height > slide_h + _OFF_SLIDE_TOLERANCE:
        edges.append("bottom")
    return edges


def _shape_left_top(shape):
    return shape.left_position(), shape.top()


def _shape_width_height(shape):
    return shape.width(), shape.height()


def _check_typography_impl(slide_indices, max_chars, max_words,
                           fix, max_expand_pt):
    """Scan shapes for widow lines and for text that does not fit its box."""
    from ppt_com.text import _find_best_vbreak

    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    total_slides = count(pres.slides)
    slide_w = pres.page_setup.slide_width()
    slide_h = pres.slide_master.height()
    issues = []
    groups_passed = []
    fixed = []

    shrink_to_fit = to_keyword(_AUTO_SIZE, ppAutoSizeTextToFitShape, "auto size")
    no_auto_size = to_keyword(_AUTO_SIZE, ppAutoSizeNone, "auto size")

    for si in slide_indices:
        if si < 1 or si > total_slides:
            continue
        goto_slide(app, si)
        slide = pres.slides[si]

        for shape in shapes_of(slide):

            # A group is walked past, not into. On macOS a group answers no
            # members over Apple Events, so there is no way in from here, and
            # the text in one is simply not looked at. That is survivable; what
            # is not is answering "no issues" for a slide whose only problem is
            # in a group, which is what this did. Counted, and said out loud
            # below, the way ppt_replace_font says it.
            try:
                if shape.shape_type() == MsoShapeType[msoGroup]:
                    groups_passed.append(shape.name())
                    continue
            except CommandError:
                pass

            # Before the text frame check, because a picture hanging off the
            # slide is as wrong as a paragraph doing it. A box set to grow with
            # its text is the usual way in: nothing overflows, because the box
            # keeps growing, and it walks off the bottom of the slide instead.
            edges = _off_slide_edges(shape, slide_w, slide_h)
            if edges:
                issues.append({
                    "slide_index": si,
                    "shape_name": shape.name(),
                    "shape_width": round(shape.width(), 2),
                    "type": "off_slide",
                    "edges": edges,
                    "fixable": False,
                })

            if not shape.has_text_frame():
                continue
            tf = shape.text_frame
            tr = tf.text_range
            if not _text_of(tr).strip():
                continue

            # Detect auto-shrink, only when text is actually being
            # compressed (natural height exceeds available space).
            # The same measurement answers two questions. Text that does not
            # fit is either being shrunk to make it fit, which is worth saying
            # because the reader gets smaller than the deck was designed for,
            # or it is spilling out of the box, which is worse and used to go
            # unreported. Only shrink-to-fit has to be turned off first, so
            # that the natural height is what gets measured.
            try:
                shrinking = tf.auto_size() == shrink_to_fit
                if shrinking:
                    tf.auto_size.set(no_auto_size)
                try:
                    natural_h = tr.bounds_height()
                    margin_h = tf.margin_top() + tf.margin_bottom()
                    avail_h = shape.height() - margin_h
                finally:
                    if shrinking:
                        tf.auto_size.set(shrink_to_fit)
                if natural_h > avail_h:
                    issues.append({
                        "slide_index": si,
                        "shape_name": shape.name(),
                        "shape_width": round(shape.width(), 2),
                        "type": "auto_shrink" if shrinking else "overflow",
                        "natural_height": round(natural_h, 2),
                        "available_height": round(avail_h, 2),
                        "fixable": False,
                    })
            except Exception:
                logger.debug("Cannot measure text height for shape '%s'",
                             shape.name(), exc_info=True)

            widows = _get_widows(shape, max_chars, max_words)

            # Detect short lines after explicit \v breaks
            # (when fix=True, we re-check after fix and report then)
            if not fix:
                vbreak_shorts = _get_short_vbreaks(shape, max_chars, max_words)
                for vb in vbreak_shorts:
                    issues.append({
                        "slide_index": si,
                        "shape_name": shape.name(),
                        "shape_width": round(shape.width(), 2),
                        **vb,
                    })

            if not widows:
                continue

            if fix:
                # Calculate safe expansion room
                gap = _right_neighbor_gap(shape, slide)
                room = min(gap - 2, max_expand_pt)  # 2pt margin
                if room < 1:
                    room = 0  # skip widen step, go straight to \v

                original_width = shape.width()
                resolved = False
                for step in range(1, int(room) + 1):
                    shape.width.set(original_width + step)
                    remaining = _get_widows(shape, max_chars, max_words)
                    if not remaining:
                        fixed.append({
                            "slide_index": si,
                            "shape_name": shape.name(),
                            "old_width": round(original_width, 2),
                            "new_width": round(shape.width(), 2),
                            "expanded_by": step,
                        })
                        resolved = True
                        break

                if not resolved:
                    # Revert the width and try inserting a soft return instead
                    shape.width.set(original_width)
                    remaining = widows
                    # Strategy 2: insert \v at word boundary
                    # Process widows in reverse order (later positions first)
                    # so that earlier character positions remain valid.
                    remaining.sort(
                        key=lambda w: w["line_index"], reverse=True,
                    )
                    vbreak_applied = False
                    for w in remaining:
                        brk = _find_best_vbreak(
                            w["prev_line_text"], w["line_text"],
                        )
                        if brk is None:
                            issues.append({
                                "slide_index": si,
                                "shape_name": shape.name(),
                                "shape_width": round(original_width, 2),
                                "fix_status": "no_break_point",
                                **w,
                            })
                            continue
                        # Find the fragment in the full text and get the
                        # character position for insertion.
                        # NOTE: find() returns the first occurrence of old_frag.
                        # If identical text appears multiple times in the shape,
                        # the wrong position may be used. This is rare in practice.
                        brk_pos, before, after = brk
                        old_frag = w["prev_line_text"] + w["line_text"]
                        full_text = _text_of(tr)
                        idx = full_text.find(old_frag)
                        if idx == -1:
                            issues.append({
                                "slide_index": si,
                                "shape_name": shape.name(),
                                "shape_width": round(original_width, 2),
                                "fix_status": "text_not_found",
                                **w,
                            })
                            continue
                        # 1-based character position for the break point. COM
                        # inserts into a zero length range; a `thru` range
                        # cannot be empty here, so this inserts before the
                        # character at that position, which is the same edit.
                        char_pos = idx + brk_pos + 1
                        tr.characters[char_pos].insert_text_text_range(
                            insert_where=k.insert_before, new_text="\v",
                        )
                        vbreak_applied = True
                        fixed.append({
                            "slide_index": si,
                            "shape_name": shape.name(),
                            "fix_method": "soft_return",
                            "before": before,
                            "after": after,
                        })
                    # After \v insertions, re-check for remaining widows
                    # and new short_after_vbreak issues
                    if vbreak_applied:
                        still_remaining = _get_widows(
                            shape, max_chars, max_words,
                        )
                        for w in still_remaining:
                            issues.append({
                                "slide_index": si,
                                "shape_name": shape.name(),
                                "shape_width": round(shape.width(), 2),
                                "fix_status": "remaining",
                                **w,
                            })

                # Always report short_after_vbreak in fix mode
                # (covers both width-expanded and \v-inserted shapes)
                post_fix_vbreaks = _get_short_vbreaks(
                    shape, max_chars, max_words,
                )
                for vb in post_fix_vbreaks:
                    issues.append({
                        "slide_index": si,
                        "shape_name": shape.name(),
                        "shape_width": round(shape.width(), 2),
                        **vb,
                    })
            else:
                for w in widows:
                    issues.append({
                        "slide_index": si,
                        "shape_name": shape.name(),
                        "shape_width": round(shape.width(), 2),
                        **w,
                    })

    result = {"issues": issues, "total": len(issues)}
    if groups_passed:
        named = ", ".join(sorted(set(groups_passed))[:5])
        more = "" if len(set(groups_passed)) <= 5 else ", and others"
        result["warnings"] = [
            f"{len(set(groups_passed))} grouped shape(s) were not looked "
            f"inside ({named}{more}). PowerPoint for Mac reports no members "
            "for a group over Apple Events, so their text was not checked and "
            "a problem in one would not appear above. Ungroup with "
            "ppt_ungroup_shapes to include it."
        ]
    if fix:
        result["fixed"] = fixed
        result["fixed_count"] = len(fixed)
        result["remaining"] = len(issues)
    return result


def _slide_count_impl() -> int:
    """How many slides the target presentation has.

    The counterpart of the COM helper of the same name. It exists on both sides
    because the public functions need a slide count before they know which
    slide to look at, and a lambda there could not be swapped.
    """
    return count(ppt._get_pres_impl().slides)
