"""Connector tools, on Apple Events.

Mirrors ``ppt_com/connectors.py``. Same function names, same signatures, same
returned shapes; what differs is the walk through PowerPoint's object model.

Three things about connectors on this side are worth knowing before reading on.

**A connection site can only be given as a number.** ``shape`` carries
``connection site count`` and the dictionary has no ``connection site`` class
and no element for one, so the coordinates of a site cannot be read. Windows
resolves 'top' or 'left' by comparing every site against the shape's centre,
and that comparison has nothing to run on here. Guessing which index is the top
of a shape would attach the connector to the wrong edge and report success, so a
direction name is refused by name while the integer form goes through.

**Connecting is a command on the connector format, not a property.** ``begin
connect`` and ``end connect`` take the shape and the site, exactly as
``ConnectorFormat.BeginConnect`` does, and ``begin connected`` reads back
whether they landed. That read back is the check, because a connector that came
out unattached looks identical to one that worked until someone opens the deck.

**Everything else is the line format.** Colour, weight, dash style and the six
arrowhead properties are all on ``line format`` and all writable. Two of their
names do not match each other; the dictionary spells the begin one ``begin arrow
head length`` and the end one ``end arrowhead length``, so anything that assumes
symmetry raises ``AttributeError`` on one of the two.
"""

import logging

from appscript import k
from appscript.reference import CommandError

from backend.mac_ae import ppt, raw, shapes_of, slide_at as _slide
from backend.mac_enums import (
    MsoArrowheadStyle,
    MsoConnectorType,
    MsoLineDashStyle,
    to_keyword,
)
from backend.unsupported import refusal as _refusal
from ppt_mac.shapes import _DASH_STYLE, _get_shape
from utils.color import hex_to_rgb_list
from utils.navigation import goto_slide

logger = logging.getLogger(__name__)

_ARROWHEAD_STYLES = MsoArrowheadStyle

# Arrowhead length and width never reached the generated table at all, because
# the generator works from the banner sections of ppt_com/constants.py and
# constants.py names neither enumeration. Both are in PowerPoint's dictionary in
# full, and the numbers are the ones ppt_com/connectors.py already uses.
_ARROWHEAD_LENGTHS = {
    1: k.short_arrowhead,
    2: k.medium_arrowhead,
    3: k.long_arrowhead,
}

_ARROWHEAD_WIDTHS = {
    1: k.narrow_width_arrowhead,
    2: k.medium_width_arrowhead,
    3: k.wide_arrowhead,
}

# The rectangle a new connector is made at, before the two ends pull it into
# place. The Windows call passes the same throwaway numbers to AddConnector.
_PLACEHOLDER_BOX = {"left": 0.0, "top": 0.0, "width": 100.0, "height": 100.0}

# Said the same way by both tools, because it is the same limit each time.
_NO_SITE_NAMES = (
    "PowerPoint for Mac tells a script how many connection sites a shape has "
    "and nothing about where they are. There is no `connection site` class in "
    "the dictionary and no element for one, so 'top' or 'left' cannot be "
    "matched to a site number, and picking one by guesswork would attach the "
    "connector to the wrong edge without saying so."
)


def _named_sites(**sites) -> list:
    """Return the site arguments that were given as a direction name.

    Checked on the type rather than on the value, because both site arguments
    default to the integer 1 and that path has to keep working.
    """
    return [name for name, value in sites.items() if isinstance(value, str)]


def _site_refusal(tool_name: str, named: list) -> dict:
    """The answer a direction name gets, naming the argument and not the tool."""
    return _refusal(
        tool_name,
        _NO_SITE_NAMES + " Give the site as a 1-based number instead; a plain "
        "rectangle has four.",
        [f"{tool_name} with {named[0]} as a number"],
        error=f"{tool_name} cannot take {', '.join(named)} as a name on macOS",
    )


def _resolve_site(shape, site: int) -> int:
    """Check a 1-based site number against the shape it is meant for.

    ``connection site count`` is the one thing the dictionary does say about
    sites, so it is worth spending an Apple Event on. A shape that will not
    answer it leaves the number unchecked rather than blocking the call.
    """
    if site < 1:
        raise ValueError(f"Connection site {site} must be 1 or greater")
    try:
        total = shape.connection_site_count()
    except CommandError:
        logger.warning(
            "Could not read the connection site count of a shape", exc_info=True
        )
        return site
    if total and site > total:
        raise ValueError(
            f"Connection site {site} is out of range for shape "
            f"'{shape.name()}', which has {total} sites (1-based)."
        )
    return site


# ---------------------------------------------------------------------------
# Apple Event implementation functions
# ---------------------------------------------------------------------------
def _add_connector_impl(slide_index, connector_type, begin_shape, begin_site,
                         end_shape, end_site):
    # Imported lazily. ppt_com/connectors.py imports this module at the bottom
    # of its own file, so importing it back at module scope would let an
    # "import ppt_mac.connectors first" ordering run that swap block against a
    # module that has defined nothing yet, and the swap would silently not
    # happen. By call time both modules are fully loaded.
    from ppt_com.connectors import CONNECTOR_TYPE_MAP

    type_key = connector_type.strip().lower()
    type_int = CONNECTOR_TYPE_MAP.get(type_key)
    if type_int is None:
        raise ValueError(
            f"Unknown connector_type '{connector_type}'. "
            f"Valid values: {list(CONNECTOR_TYPE_MAP.keys())}"
        )
    type_word = to_keyword(MsoConnectorType, type_int, "connector type")

    named = _named_sites(begin_site=begin_site, end_site=end_site)
    if named:
        return _site_refusal("ppt_add_connector", named)

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)

    # Both ends are found before anything is created, so a name that is not on
    # the slide leaves no stray connector behind.
    begin = _get_shape(slide, begin_shape)
    end = _get_shape(slide, end_shape)
    resolved_begin = _resolve_site(begin, begin_site)
    resolved_end = _resolve_site(end, end_site)

    before = len(shapes_of(slide))
    # The insertion location is the slide itself, never its shapes. `at` given
    # as `slide.shapes.end` raises -1708, and the error does not say why.
    try:
        app.make(
            new=k.connector,
            at=slide.end,
            with_properties={
                k.left_position: _PLACEHOLDER_BOX["left"],
                k.top: _PLACEHOLDER_BOX["top"],
                k.width: _PLACEHOLDER_BOX["width"],
                k.height: _PLACEHOLDER_BOX["height"],
            },
        )
    except CommandError as exc:
        return _refusal(
            "ppt_add_connector",
            "PowerPoint refused to make a connector and answered "
            f"{exc}. The dictionary declares no command for adding one, so "
            "this goes through the Standard Suite `make` that creates every "
            "other shape here, and that is the whole of the route.",
            ["ppt_add_shape with a line shape"],
        )

    # Nothing is trusted because it did not raise, and `make`'s own return
    # value is not used either. A new shape lands at the end of the z order, so
    # it is fetched again by a route that is known to work.
    shapes_now = shapes_of(slide)
    if len(shapes_now) != before + 1:
        return _refusal(
            "ppt_add_connector",
            "PowerPoint reported success but the slide gained no shape, which "
            "is the silent no-op recorded in MACOS_PORT section 5.",
            ["ppt_add_shape with a line shape"],
        )

    connector = shapes_now[-1]
    if not connector.is_connector():
        return _refusal(
            "ppt_add_connector",
            "PowerPoint made a shape that is not a connector, so it has no "
            "ends to attach and nothing was connected. A `make` that falls "
            "through leaves a plain autoshape behind.",
            ["ppt_add_shape with a line shape"],
        )

    fmt = connector.connector_format
    # `connector type` is read only on the connector itself and writable on its
    # format, which is the only reason an elbow or a curve is reachable at all.
    fmt.connector_type.set(type_word)
    fmt.begin_connect(connected_shape=begin, connection_site=resolved_begin)
    fmt.end_connect(connected_shape=end, connection_site=resolved_end)
    connector.reroute_connections()

    name = connector.name()
    if not (fmt.begin_connected() and fmt.end_connected()):
        return _refusal(
            "ppt_add_connector",
            f"PowerPoint made connector '{name}' and reported both connections "
            "as done, but it reads back as unattached at one end or both. The "
            "connector is on the slide at its placeholder size and can be "
            "attached by hand.",
            ["Attach the connector by hand in PowerPoint"],
            error="ppt_add_connector could not attach the connector on macOS",
        )

    return {
        "success": True,
        "shape_name": name,
        "connector_type": type_key,
    }


def _format_connector_impl(slide_index, shape_name_or_index,
                             color, weight, dash_style,
                             begin_arrow, begin_arrow_length, begin_arrow_width,
                             end_arrow, end_arrow_length, end_arrow_width,
                             begin_shape, begin_site,
                             end_shape, end_site):
    # Lazy import for the same reason as in _add_connector_impl.
    from ppt_com.connectors import (
        ARROW_LENGTH_MAP,
        ARROW_STYLE_MAP,
        ARROW_WIDTH_MAP,
        DASH_STYLE_MAP,
    )

    def _lookup(value, mapping, argument):
        """Turn one of the tool's own words into the Windows constant for it."""
        if value is None:
            return None
        found = mapping.get(value.strip().lower())
        if found is None:
            raise ValueError(
                f"Unknown {argument} '{value}'. "
                f"Valid values: {list(mapping.keys())}"
            )
        return found

    # Every word is checked before the view moves and before anything is
    # written, so a typo in the last argument does not leave the first four
    # applied.
    dash_val = _lookup(dash_style, DASH_STYLE_MAP, "dash_style")
    begin_arrow_val = _lookup(begin_arrow, ARROW_STYLE_MAP, "begin_arrow")
    begin_length_val = _lookup(
        begin_arrow_length, ARROW_LENGTH_MAP, "begin_arrow_length"
    )
    begin_width_val = _lookup(
        begin_arrow_width, ARROW_WIDTH_MAP, "begin_arrow_width"
    )
    end_arrow_val = _lookup(end_arrow, ARROW_STYLE_MAP, "end_arrow")
    end_length_val = _lookup(end_arrow_length, ARROW_LENGTH_MAP, "end_arrow_length")
    end_width_val = _lookup(end_arrow_width, ARROW_WIDTH_MAP, "end_arrow_width")

    named = _named_sites(begin_site=begin_site, end_site=end_site)
    if named:
        return _site_refusal("ppt_format_connector", named)

    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    line = shape.line_format

    if color is not None:
        line.fore_color.set(hex_to_rgb_list(color))

    if weight is not None:
        line.line_weight.set(weight)

    if dash_val is not None:
        # `dash style` is a name AppleScript's own vocabulary already owns, so
        # appscript cannot reach PowerPoint's property by it. The code can.
        raw(line, _DASH_STYLE).set(
            to_keyword(MsoLineDashStyle, dash_val, "dash style")
        )

    if begin_arrow_val is not None:
        line.begin_arrowhead_style.set(
            to_keyword(_ARROWHEAD_STYLES, begin_arrow_val, "arrowhead style")
        )

    if begin_length_val is not None:
        # The dictionary spells this one `begin arrow head length` and its
        # opposite number `end arrowhead length`. The asymmetry is real.
        line.begin_arrow_head_length.set(
            to_keyword(_ARROWHEAD_LENGTHS, begin_length_val, "arrowhead length")
        )

    if begin_width_val is not None:
        line.begin_arrowhead_width.set(
            to_keyword(_ARROWHEAD_WIDTHS, begin_width_val, "arrowhead width")
        )

    if end_arrow_val is not None:
        line.end_arrowhead_style.set(
            to_keyword(_ARROWHEAD_STYLES, end_arrow_val, "arrowhead style")
        )

    if end_length_val is not None:
        line.end_arrowhead_length.set(
            to_keyword(_ARROWHEAD_LENGTHS, end_length_val, "arrowhead length")
        )

    if end_width_val is not None:
        line.end_arrowhead_width.set(
            to_keyword(_ARROWHEAD_WIDTHS, end_width_val, "arrowhead width")
        )

    name = shape.name()
    reroute = False
    if begin_shape is not None or end_shape is not None:
        if not shape.is_connector():
            raise ValueError(
                f"Shape '{name}' is not a connector, so it has no ends to "
                "reconnect. Only a connector can take begin_shape or end_shape."
            )
        fmt = shape.connector_format
        if begin_shape is not None:
            target = _get_shape(slide, begin_shape)
            site = _resolve_site(target, begin_site if begin_site is not None else 1)
            fmt.begin_connect(connected_shape=target, connection_site=site)
            reroute = True
        if end_shape is not None:
            target = _get_shape(slide, end_shape)
            site = _resolve_site(target, end_site if end_site is not None else 1)
            fmt.end_connect(connected_shape=target, connection_site=site)
            reroute = True

    if reroute:
        shape.reroute_connections()
        # Read back, because an end that did not take looks like one that did
        # until someone opens the deck.
        attached = (
            (begin_shape is None or fmt.begin_connected())
            and (end_shape is None or fmt.end_connected())
        )
        if not attached:
            return _refusal(
                "ppt_format_connector",
                f"The line formatting landed on '{name}', but PowerPoint "
                "reported the reconnection as done and the connector reads "
                "back as unattached at one end or both.",
                ["Attach the connector by hand in PowerPoint"],
                error=(
                    "ppt_format_connector could not reconnect the connector "
                    "on macOS"
                ),
            )

    return {
        "success": True,
        "shape_name": name,
    }
