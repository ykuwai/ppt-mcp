"""Hyperlink tools, on Apple Events.

Mirrors ``ppt_com/hyperlinks.py``. Same function names, same signatures, same
returned shapes; what differs is the walk through PowerPoint's object model.

Three things about hyperlinks on this side are worth knowing before reading on.

**A shape's click action is reached by a command, and its result has to be
trusted.** ``shape`` declares no ``action setting`` property and no
``action setting`` element, so there is nothing to count and nothing to index.
The only route in the dictionary is ``get action setting for``, which takes the
shape and an ``EPPMouseActivation`` word and answers with an ``action setting``
reference. Everywhere else in this port a reference PowerPoint hands back is
thrown away and rebuilt by hand, and here there is nothing to rebuild it from.
So the command result is used, and a reference that does not resolve is treated
as an expected outcome rather than a surprise. Both the command and the first
property access are guarded, and a failure comes back as a refusal that names
the route.

**There is no screen tip.** The whole ``hyperlink`` class is
``hyperlink address``, ``hyperlink sub address`` and a read only
``hyperlink type``. A caller that asks for a screen tip is told that one
argument has to go rather than that the tool is unavailable.

**The hyperlink type table is built here.** ``constants.py`` has no
``MsoHyperlinkType`` group, so ``scripts/gen_mac_enums.py`` had no Windows
names to pair the macOS enumeration against and generated no table for it. The
three enumerator codes in the dictionary end in 0, 1 and 2, which is the
Windows numbering as well, so the pairing below is corroborated by the
dictionary rather than assumed.

Text ranges have a route of their own, ``get text action setting``, so a link
on part of a paragraph is reachable in principle. The Windows module only ever
links whole shapes, and this module matches it rather than growing a second
surface macOS alone would have.
"""

import logging

from appscript import k
from appscript.reference import CommandError

from backend.mac_ae import (
    error_number,
    is_missing,
    positional,
    ppt,
    shape_by_name_or_index as _get_shape,
    windows_constant as _windows_constant,
)
from backend.mac_enums import PpActionType, PpMouseActivation, to_keyword
from backend.unsupported import refusal as _refusal
from utils.navigation import goto_slide
from ppt_com.constants import ppActionHyperlink, ppActionNone

logger = logging.getLogger(__name__)

# Windows MsoHyperlinkType, which constants.py never declared. See the module
# docstring for why the generated table has no entry for it. Kept in the shape
# of a generated table so `_windows_constant` reads the same as it does in
# tables.py and animation.py.
MsoHyperlinkType: dict = {
    0: k.hyperlink_type_text_range,   # msoHyperlinkRange
    1: k.hyperlink_type_shape,        # msoHyperlinkShape
    2: k.hyperlink_type_inline_shape,  # msoHyperlinkInlineShape
}

# The second route to the same three values. `hyperlink type` is declared in
# the dictionary as `mHyT` while the enumeration it means is `mHlT`, a type
# reference that points at nothing. It does not change the answer, because
# appscript names an enumerator from one table covering the whole dictionary
# and these three codes appear nowhere else in it, but if a future build breaks
# that the raw code still carries the number in its low byte.
_HYPERLINK_TYPE_CODES = {
    b"\x00\x96\x00\x00": 0,  # msoHyperlinkRange
    b"\x00\x96\x00\x01": 1,  # msoHyperlinkShape
    b"\x00\x96\x00\x02": 2,  # msoHyperlinkInlineShape
}

_WHAT_ACTION = "action type"
_WHAT_EVENT = "mouse activation"

# Said the same way by every refusal that gets this far, because the cause is
# always the one command and a reader should recognise it on sight.
_ACTION_SETTING_ROUTE = (
    "A shape's click action is reached here as `action settings` position 1 for "
    "a click and position 2 for a mouse over, built rather than asked for, and "
    "this shape's did not resolve."
)


def _event_keyword(action_on):
    """Validate ``action_on`` and return the macOS word for it.

    Raises ``ValueError`` on an unknown value, matching the Windows module,
    which the tool layer turns into the same error text on both platforms.
    """
    # Imported lazily. ppt_com/hyperlinks.py imports this module at the bottom
    # of its own file, so importing it back at module scope would let an
    # "import ppt_mac.hyperlinks first" ordering run that swap block against a
    # module that has defined nothing yet. By call time both are fully loaded.
    from ppt_com.hyperlinks import ACTION_ON_MAP

    action_key = action_on.strip().lower()
    if action_key not in ACTION_ON_MAP:
        raise ValueError(
            f"Unknown action_on '{action_on}'. Use: {', '.join(ACTION_ON_MAP.keys())}"
        )
    return action_key, to_keyword(
        PpMouseActivation, ACTION_ON_MAP[action_key], _WHAT_EVENT
    )


# Where each mouse event's action setting sits. PowerPoint keeps exactly two
# per shape, click first and mouse over second, which was checked by reading
# both and by writing to the first and seeing the slide gain a hyperlink.
_SETTING_POSITION = {"click": 1, "mouseover": 2}


def _action_setting(shape, action_key):
    """The shape's action setting for one mouse event, built here.

    Not `get action setting for`. The command is in the dictionary, takes the
    event the dictionary says it takes, and answers with a reference that does
    not resolve; every property on it comes back -1728. Building the same
    reference by index works for both events, reads the action, writes it, and
    the slide gains a hyperlink afterwards. It is the `get cell from` defect
    from tables.py again, and the rule is the same, a reference PowerPoint
    hands back is not used.
    """
    return shape.action_settings[_SETTING_POSITION[action_key]]


def _hyperlink_type(word):
    """The Windows number for the kind of thing a hyperlink is attached to.

    By name first, which is how every other enumerator in this port is read,
    and by raw code second. None rather than a near miss when neither answers,
    because a wrong type reads as a fact rather than as a gap.
    """
    value = _windows_constant(MsoHyperlinkType, word)
    if value is not None:
        return value
    return _HYPERLINK_TYPE_CODES.get(getattr(word, "code", None))


def _text_or_none(reference):
    """Read a text property, treating `missing value` and a dead reference alike."""
    try:
        value = reference()
    except CommandError:
        return None
    return None if is_missing(value) or value == "" else value


# ---------------------------------------------------------------------------
# Apple Event implementation functions
# ---------------------------------------------------------------------------
def _add_hyperlink_impl(
    slide_index, shape_name_or_index, address, sub_address, screen_tip, action_on
):
    action_key, event = _event_keyword(action_on)

    # Before goto_slide, so a call that is going to be refused does not move
    # the user's view first.
    if screen_tip is not None:
        return _refusal(
            "ppt_add_hyperlink",
            "PowerPoint for Mac's `hyperlink` class carries only its address, "
            "its sub address and a read only type. There is no screen tip to "
            "set, so the link would go in and the tip would be dropped without "
            "anyone being told.",
            ["ppt_add_hyperlink without screen_tip"],
            error="ppt_add_hyperlink cannot set screen_tip on macOS",
        )

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]
    shape = _get_shape(slide, shape_name_or_index)

    try:
        setting = _action_setting(shape, action_key)
        # The order matters here as it does on Windows. The action has to say
        # hyperlink before the address will hold.
        setting.action.set(to_keyword(PpActionType, ppActionHyperlink, _WHAT_ACTION))
        setting.hyperlink.hyperlink_address.set(address)
        if sub_address is not None:
            setting.hyperlink.hyperlink_sub_address.set(sub_address)
        written = setting.hyperlink.hyperlink_address()
        # Only read back what was written. A link that never had a sub address
        # has none to report, and Windows reports none for it too.
        written_sub = (
            setting.hyperlink.hyperlink_sub_address()
            if sub_address is not None else None
        )
    except CommandError as exc:
        logger.warning(
            "The action setting for this shape did not resolve", exc_info=True
        )
        return _refusal(
            "ppt_add_hyperlink",
            f"{_ACTION_SETTING_ROUTE} PowerPoint answered with Apple Event "
            f"error {error_number(exc)}.",
        )

    # Nothing is trusted because it did not raise. The read back is checked for
    # emptiness rather than for equality, because PowerPoint completes an
    # address it considers partial and a strict comparison would refuse a write
    # that did land.
    if is_missing(written) or written == "":
        return _refusal(
            "ppt_add_hyperlink",
            "PowerPoint reported success but the shape's action setting still "
            "holds no address, which is the silent no-op recorded in "
            "MACOS_PORT section 5.",
        )

    return {
        "success": True,
        "shape_name": shape.name(),
        # Read back rather than echoed, so that a caller is told what is now in
        # the deck. PowerPoint rewrites an address it considers incomplete.
        "address": written,
        "sub_address": (
            None if is_missing(written_sub) or written_sub == "" else written_sub
        ),
        "action_on": action_key,
    }


def _get_hyperlinks_impl(slide_index):
    ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]

    # `hyperlink` is a declared element of `slide`, so this one is a plain
    # collection walk. Positional all the same, which is the house rule for
    # every collection here.
    links = positional(slide.hyperlinks)

    hyperlinks = []
    for index, link in enumerate(links, 1):
        try:
            type_word = link.hyperlink_type()
        except CommandError:
            type_word = None
        hyperlinks.append({
            "index": index,
            "address": _text_or_none(link.hyperlink_address),
            "sub_address": _text_or_none(link.hyperlink_sub_address),
            "type": _hyperlink_type(type_word),
        })

    return {
        "success": True,
        "slide_index": slide_index,
        "hyperlinks_count": len(hyperlinks),
        "hyperlinks": hyperlinks,
    }


def _remove_hyperlink_impl(slide_index, shape_name_or_index, action_on):
    action_key, event = _event_keyword(action_on)

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.slides[slide_index]
    shape = _get_shape(slide, shape_name_or_index)

    try:
        setting = _action_setting(shape, action_key)
        setting.action.set(to_keyword(PpActionType, ppActionNone, _WHAT_ACTION))
        after = setting.action()
    except CommandError as exc:
        logger.warning(
            "The action setting for this shape did not resolve", exc_info=True
        )
        return _refusal(
            "ppt_remove_hyperlink",
            f"{_ACTION_SETTING_ROUTE} PowerPoint answered with Apple Event "
            f"error {error_number(exc)}.",
        )

    # `action type unset` is what a shape that never carried an action reads
    # back as, and it is as cleared as `action type none` is. Requiring only
    # the second would refuse the ordinary case.
    if after not in (k.action_type_none, k.action_type_unset):
        return _refusal(
            "ppt_remove_hyperlink",
            "PowerPoint reported success but the shape's action setting still "
            "reads back as something other than cleared, which is the silent "
            "no-op recorded in MACOS_PORT section 5.",
        )

    return {
        "success": True,
        "shape_name": shape.name(),
        "action_on": action_key,
    }
