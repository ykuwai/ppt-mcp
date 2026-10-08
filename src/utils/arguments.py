"""Refusals for tool arguments, written for the caller rather than a developer.

Every input model forbids keys it does not declare, so a misspelled or borrowed
argument is refused instead of being dropped while the call answers success.
Pydantic's own report of that ("1 validation error for
tool_ppt_add_textboxArguments ... Extra inputs are not permitted") names the
key but not the tool, not what the key should have been, and not what the
model does take. The helpers here turn a validation failure into a sentence
that does, so the caller can correct the call in one step.
"""

from __future__ import annotations

import difflib
import types
from collections.abc import Iterable
from typing import Annotated, Any, Union, get_args, get_origin

from pydantic import BaseModel, ValidationError

# Where one idea goes by two names across these tools. The value is the name
# the models on the receiving end use; the key is what a caller arrives with,
# having just used a sibling tool. difflib is no guide for these: it answers
# `font_color_theme` for `font_color`, which is a different thing entirely.
SIBLING_NAMES = {
    "font_color": "color",
    "line_visible": "visible",
    "line_color": "color",
    "line_weight": "weight",
    "fill_color": "color",
    "fill_transparency": "transparency",
}


def suggest(key: str, known: Iterable[str]) -> str | None:
    """The argument `key` was most likely meant to be, or None."""
    known = list(known)
    sibling = SIBLING_NAMES.get(key)
    if sibling in known:
        return sibling
    close = difflib.get_close_matches(key, known, n=1, cutoff=0.6)
    return close[0] if close else None


def _flatten(annotation: Any) -> list[Any]:
    """The types an annotation may hold, with Optional, Union and Annotated
    taken apart."""
    origin = get_origin(annotation)
    if origin is Annotated:
        return _flatten(get_args(annotation)[0])
    if origin is Union or origin is types.UnionType:
        out = []
        for arg in get_args(annotation):
            out.extend(_flatten(arg))
        return out
    return [annotation]


def _is_model(tp: Any) -> bool:
    return isinstance(tp, type) and issubclass(tp, BaseModel)


def _walk(root: type[BaseModel], loc: tuple) -> tuple[str, type[BaseModel] | None]:
    """Follow a pydantic error location from `root`.

    Returns the location as a caller would write it (`params.ranges[0]`),
    leaving out the union member tags pydantic puts in, and the model found
    there, or None when the location does not end on a model.
    """
    current: list[Any] = [root]
    path = ""
    for part in loc:
        if isinstance(part, int):
            items = []
            for tp in current:
                if get_origin(tp) in (list, tuple, set, frozenset):
                    for arg in get_args(tp):
                        if arg is not Ellipsis:
                            items.extend(_flatten(arg))
            if items:
                current = items
            path += f"[{part}]"
            continue
        owners = [tp for tp in current if _is_model(tp) and part in tp.model_fields]
        if owners:
            current = _flatten(owners[0].model_fields[part].annotation)
            path += f".{part}" if path else part
            continue
        # Not a field: a union member tag such as 'TextRangeSpec' or 'str'.
        # Narrow to that member when it is a model, and keep it out of the path.
        tagged = [tp for tp in current if _is_model(tp) and tp.__name__ == part]
        if tagged:
            current = tagged
    models = [tp for tp in current if _is_model(tp)]
    return path, (models[0] if len(models) == 1 else None)


def _got(value: Any) -> str:
    """' (got 7)' for a short scalar input; nothing for anything bigger."""
    if isinstance(value, (str, int, float, bool)) or value is None:
        text = repr(value)
        if len(text) <= 60:
            return f" (got {text})"
    return ""


def _unknown(keys: list[str], where: str, model: type[BaseModel] | None) -> str:
    """The refusal for keys a model does not take, with a guess for each."""
    valid = list(model.model_fields) if model is not None else []
    named = []
    for key in keys:
        guess = suggest(key, valid) if valid else None
        named.append(f"{key!r}" + (f" (did you mean {guess!r}?)" if guess else ""))
    noun = "argument" if len(keys) == 1 else "arguments"
    text = f"unknown {noun} {', '.join(named)} in {where}."
    if valid:
        text += " Valid arguments: " + ", ".join(valid) + "."
    return text


def _other(root: type[BaseModel], err: dict) -> str:
    loc = tuple(err.get("loc", ()))
    kind = err.get("type", "")
    if kind == "missing" and loc:
        parent, _ = _walk(root, loc[:-1])
        return f"missing required argument {str(loc[-1])!r} in {parent or 'the call'}."
    path, _ = _walk(root, loc)
    msg = str(err.get("msg", "invalid value"))
    for prefix in ("Value error, ", "Assertion failed, "):
        if msg.startswith(prefix):
            msg = msg[len(prefix):]
    # A validator's own message already says what was wrong with the value.
    got = "" if kind in ("value_error", "assertion_error") else _got(err.get("input"))
    return f"{path or 'arguments'}: {msg}{got}"


def describe_validation_error(tool_name: str, root: type[BaseModel], error: ValidationError) -> str:
    """One readable message for a tool call whose arguments failed validation.

    `root` is the model the arguments were validated against, the SDK's
    per-tool argument model whose one field is normally `params`. Unknown
    keys are gathered per place they were passed, so the valid arguments
    there are listed once.
    """
    lines: list[str] = []
    unknown: dict[str, tuple[list[str], type[BaseModel] | None, int]] = {}
    for err in error.errors():
        loc = tuple(err.get("loc", ()))
        if err.get("type") == "extra_forbidden" and loc:
            where, model = _walk(root, loc[:-1])
            where = where or "the call"
            if where not in unknown:
                unknown[where] = ([], model, len(lines))
                lines.append("")  # filled in below, keeping the order
            unknown[where][0].append(str(loc[-1]))
            continue
        line = _other(root, err)
        if line not in lines:
            lines.append(line)
    for where, (keys, model, at) in unknown.items():
        lines[at] = _unknown(keys, where, model)
    if len(lines) == 1:
        return f"{tool_name}: {lines[0]} Nothing was applied."
    return (
        f"{tool_name}: {len(lines)} problems with the arguments, nothing was applied.\n"
        + "\n".join(f"- {line}" for line in lines)
    )


def describe_unknown_top_level(
    tool_name: str,
    unknown: Iterable[str],
    declared: Iterable[str],
    params_model: type[BaseModel] | None = None,
) -> str:
    """The refusal for keys passed next to `params` that the tool does not take.

    A key that is one of the `params` fields was put at the wrong level, which
    is what a caller who flattened the call did, so that is what it is told.
    """
    unknown = list(unknown)
    declared = list(declared)
    inner = list(params_model.model_fields) if params_model is not None else []
    misplaced = [key for key in unknown if key in inner]
    named = []
    for key in unknown:
        if key in misplaced:
            continue
        guess = suggest(key, declared)
        if guess:
            named.append(f"{key!r} (did you mean {guess!r}?)")
            continue
        guess = suggest(key, inner)
        named.append(f"{key!r}" + (f" (did you mean {guess!r} inside params?)" if guess else ""))
    parts = []
    if named:
        noun = "argument" if len(named) == 1 else "arguments"
        parts.append(f"unknown {noun} {', '.join(named)}.")
    if misplaced:
        verb = "belongs" if len(misplaced) == 1 else "belong"
        parts.append(f"{', '.join(repr(k) for k in misplaced)} {verb} inside params, not next to it.")
    if declared:
        parts.append("Valid arguments: " + ", ".join(declared) + ". Nothing was applied.")
    else:
        parts.append("This tool takes no arguments. Nothing was applied.")
    return f"{tool_name}: " + " ".join(parts)
