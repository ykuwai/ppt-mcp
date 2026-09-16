"""Tests for the macOS effect, group and connector tools.

Pure unit tests against a fake object graph. None of this launches PowerPoint,
and none of it may; the live behaviour is covered by MACOS_PORT.md and by
running the server. appscript only installs on macOS, so everything that
imports it is skipped elsewhere.

What is worth testing here is what the three modules had to decide. Which
Windows constant becomes which macOS enumerator, what a tool says when macOS
cannot honour one argument, and that a refusal is answered before PowerPoint is
touched at all, so a refused call never moves the view or edits the deck.
"""

import sys

import pytest

sys.path.insert(0, "src")

macos_only = pytest.mark.skipif(
    sys.platform != "darwin", reason="the Apple Event backend needs appscript"
)


# ---------------------------------------------------------------------------
# Enumerator tables
# ---------------------------------------------------------------------------
@macos_only
class TestEffectEnums:
    """Reflection and soft edge are spelled out by hand, so they are checked."""

    def test_reflection_presets_are_the_dictionary_names(self):
        from appscript import k

        from ppt_mac.effects import _REFLECTION_TYPES

        assert _REFLECTION_TYPES[0] == k.reflection_type_none
        assert _REFLECTION_TYPES[1] == k.reflection_type1
        assert _REFLECTION_TYPES[9] == k.reflection_type9
        assert len(_REFLECTION_TYPES) == 10

    def test_soft_edge_presets_run_off_to_fifty_points(self):
        from appscript import k

        from ppt_mac.effects import _SOFT_EDGE_PRESETS, _nearest_soft_edge

        assert [points for points, _, _ in _SOFT_EDGE_PRESETS] == [
            0.0, 1.0, 2.5, 5.0, 10.0, 25.0, 50.0
        ]
        assert _nearest_soft_edge(0)[1] == k.no_soft_edge
        assert _nearest_soft_edge(5)[1] == k.soft_edge_type3
        assert _nearest_soft_edge(1000)[1] == k.soft_edge_type6

    def test_a_radius_between_presets_takes_the_nearer_one(self):
        from ppt_mac.effects import _nearest_soft_edge

        assert _nearest_soft_edge(4)[0] == 5.0
        assert _nearest_soft_edge(3)[0] == 2.5
        assert _nearest_soft_edge(0.4)[0] == 0.0


@macos_only
class TestConnectorEnums:
    """Three arrowhead tables, one generated and two written out by hand."""

    def test_the_arrowhead_the_generator_used_to_miss_is_there(self):
        """Windows says msoArrowheadNone, macOS says `no arrowhead`.

        The name matcher cannot see through that, so the generator carries an
        override for it now and the table comes straight from the generated
        one rather than from a local patch.
        """
        from appscript import k

        from backend.mac_enums import MsoArrowheadStyle
        from ppt_mac.connectors import _ARROWHEAD_STYLES

        assert _ARROWHEAD_STYLES is MsoArrowheadStyle
        assert _ARROWHEAD_STYLES[1] == k.no_arrowhead
        assert _ARROWHEAD_STYLES[2] == k.triangle_arrowhead
        assert _ARROWHEAD_STYLES[6] == k.oval_arrowhead

    def test_lengths_and_widths_are_the_dictionary_names(self):
        from appscript import k

        from ppt_mac.connectors import _ARROWHEAD_LENGTHS, _ARROWHEAD_WIDTHS

        assert _ARROWHEAD_LENGTHS == {
            1: k.short_arrowhead, 2: k.medium_arrowhead, 3: k.long_arrowhead
        }
        assert _ARROWHEAD_WIDTHS == {
            1: k.narrow_width_arrowhead,
            2: k.medium_width_arrowhead,
            3: k.wide_arrowhead,
        }

    def test_connector_types_pair_with_the_windows_map(self):
        from appscript import k

        from backend.mac_enums import MsoConnectorType, to_keyword
        from ppt_com.connectors import CONNECTOR_TYPE_MAP

        assert to_keyword(
            MsoConnectorType, CONNECTOR_TYPE_MAP["elbow"], "connector type"
        ) == k.elbow
        assert MsoConnectorType[CONNECTOR_TYPE_MAP["curve"]] == k.curve


# ---------------------------------------------------------------------------
# Effects
# ---------------------------------------------------------------------------
@macos_only
class TestGlow:
    """Radius and colour land; transparency has nowhere to go and says so."""

    def test_radius_and_colour_are_written(self):
        from ppt_mac.effects import _set_glow_impl

        with _fake_deck(["Title"]) as deck:
            result = _set_glow_impl(1, "Title", 12.0, "#1F6FEB", None)

        glow = deck.shape("Title").glow_format
        assert glow.radius() == 12.0
        assert glow.color() == [31, 111, 235]
        assert result["status"] == "success"
        assert result["shape_name"] == "Title"
        assert "warnings" not in result

    def test_the_radius_reported_is_the_one_powerpoint_kept(self):
        """Read back rather than echoed, so a clamped radius is visible."""
        from ppt_mac.effects import _set_glow_impl

        with _fake_deck(["Title"]) as deck:
            deck.shape("Title").glow_format.radius.clamp_to = 8.0
            result = _set_glow_impl(1, "Title", 400.0, None, None)

        assert result["glow_radius"] == 8.0

    def test_transparency_is_named_in_a_warning_and_the_glow_still_lands(self):
        from ppt_mac.effects import _set_glow_impl

        with _fake_deck(["Title"]) as deck:
            result = _set_glow_impl(1, "Title", 6.0, None, 0.4)

        assert deck.shape("Title").glow_format.radius() == 6.0
        assert result["status"] == "success"
        assert "transparency=0.4" in result["warnings"][0]
        assert "no transparency property" in result["warnings"][0]


@macos_only
class TestReflection:
    """One preset is the whole of it, and a call that would do nothing refuses."""

    def test_a_preset_is_translated_to_its_enumerator(self):
        from appscript import k

        from ppt_mac.effects import _set_reflection_impl

        with _fake_deck(["Title"]) as deck:
            result = _set_reflection_impl(1, "Title", 3, None, None, None, None)

        assert deck.shape("Title").reflection_format.reflection_type() == (
            k.reflection_type3
        )
        assert result == {"status": "success", "shape_name": "Title"}

    def test_zero_turns_the_reflection_off(self):
        from appscript import k

        from ppt_mac.effects import _set_reflection_impl

        with _fake_deck(["Title"]) as deck:
            _set_reflection_impl(1, "Title", 0, None, None, None, None)

        assert deck.shape("Title").reflection_format.reflection_type() == (
            k.reflection_type_none
        )

    def test_a_preset_macos_does_not_have_says_the_range(self):
        from ppt_mac.effects import _set_reflection_impl

        with _fake_deck(["Title"]):
            with pytest.raises(ValueError, match="Unknown reflection_type 11"):
                _set_reflection_impl(1, "Title", 11, None, None, None, None)

    def test_the_extras_alongside_a_preset_come_back_as_warnings(self):
        from ppt_mac.effects import _set_reflection_impl

        with _fake_deck(["Title"]) as deck:
            result = _set_reflection_impl(1, "Title", 2, 5.0, 3.0, 50.0, 0.3)

        assert deck.shape("Title").reflection_format.reflection_type is not None
        assert result["status"] == "success"
        warning = result["warnings"][0]
        for argument in ("blur", "offset", "size", "transparency"):
            assert argument in warning

    def test_the_extras_on_their_own_are_refused_by_name(self):
        """Nothing would land, so a success would be a lie."""
        from ppt_mac.effects import _set_reflection_impl

        with _fake_deck(["Title"]):
            result = _set_reflection_impl(1, "Title", None, 5.0, None, None, 0.3)

        assert result["error"] == (
            "ppt_set_reflection cannot set blur, transparency on macOS"
        )
        assert "is not available" not in result["error"]
        assert result["platform"] == "macOS"
        assert "reflection type" in result["reason"]
        assert result["alternatives"] == ["ppt_set_reflection with reflection_type"]

    def test_that_refusal_never_reaches_powerpoint(self):
        from ppt_mac.effects import _set_reflection_impl

        with _no_powerpoint():
            result = _set_reflection_impl(1, "Title", None, 5.0, None, None, None)

        assert "error" in result


@macos_only
class TestSoftEdge:
    """A radius in points has to become one of six presets."""

    def test_a_radius_on_a_preset_is_applied_without_comment(self):
        from appscript import k

        from ppt_mac.effects import _set_soft_edge_impl

        with _fake_deck(["Title"]) as deck:
            result = _set_soft_edge_impl(1, "Title", 5.0)

        assert deck.shape("Title").soft_edge_format.soft_edge_type() == (
            k.soft_edge_type3
        )
        assert result["soft_edge_radius"] == 5.0
        assert "warnings" not in result

    def test_zero_removes_the_soft_edge(self):
        from appscript import k

        from ppt_mac.effects import _set_soft_edge_impl

        with _fake_deck(["Title"]) as deck:
            result = _set_soft_edge_impl(1, "Title", 0)

        assert deck.shape("Title").soft_edge_format.soft_edge_type() == (
            k.no_soft_edge
        )
        assert "warnings" not in result

    def test_a_radius_between_presets_says_where_it_landed(self):
        from ppt_mac.effects import _set_soft_edge_impl

        with _fake_deck(["Title"]):
            result = _set_soft_edge_impl(1, "Title", 4.0)

        assert result["soft_edge_radius"] == 5.0
        assert "radius=4.0" in result["warnings"][0]
        assert "5 points" in result["warnings"][0]


# ---------------------------------------------------------------------------
# Groups
# ---------------------------------------------------------------------------
@macos_only
class TestGroupItems:
    """The shape check comes first. Reading the members is in test_mac_gvml."""

    def test_a_shape_that_is_not_a_group_says_so_with_its_type(self):
        from ppt_mac.groups import _get_group_items_impl

        with _fake_deck(["Title"]):
            with pytest.raises(ValueError, match=r"is not a group \(type=1\)"):
                _get_group_items_impl(1, "Title")


@macos_only
class TestUngrouping:
    """Tried rather than refused, and the slide decides whether it worked."""

    def test_a_group_that_comes_apart_reports_its_members(self):
        from ppt_mac.groups import _ungroup_shapes_impl

        with _fake_deck(["Title", "Diagram"]) as deck:
            deck.make_group("Diagram", ["Box", "Circle"], comes_apart=True)
            result = _ungroup_shapes_impl(1, "Diagram")

        assert result["success"] is True
        assert result["ungrouped_count"] == 2
        assert result["shape_names"] == ["Box", "Circle"]
        # PowerPoint renames the members on the way out, and every tool here
        # tells callers to use names because indices shift. A caller who noted
        # the names before this call finds none of them afterwards.
        assert "renamed the members" in result["warnings"][0]

    def test_a_group_that_stays_put_is_a_refusal_not_a_success(self):
        from ppt_mac.groups import _ungroup_shapes_impl

        with _fake_deck(["Title", "Diagram"]) as deck:
            deck.make_group("Diagram", ["Box", "Circle"], comes_apart=False)
            result = _ungroup_shapes_impl(1, "Diagram")

        assert result["error"] == "ppt_ungroup_shapes is not available on macOS"
        assert "'Diagram' is still on slide 1 as one shape" in result["reason"]
        assert "silent no-op" in result["reason"]
        assert result["platform"] == "macOS"

    def test_an_unreadable_reply_is_not_taken_for_a_failure(self):
        """PowerPoint answers with a shape range, which may not come back."""
        from ppt_mac.groups import _ungroup_shapes_impl

        with _fake_deck(["Diagram"]) as deck:
            deck.make_group(
                "Diagram", ["Box", "Circle"], comes_apart=True, reply_fails=True
            )
            result = _ungroup_shapes_impl(1, "Diagram")

        assert result["success"] is True

    def test_a_failed_call_that_changed_nothing_quotes_powerpoint(self):
        from ppt_mac.groups import _ungroup_shapes_impl

        with _fake_deck(["Diagram"]) as deck:
            deck.make_group(
                "Diagram", ["Box"], comes_apart=False, reply_fails=True
            )
            result = _ungroup_shapes_impl(1, "Diagram")

        assert "stub error -1708" in result["reason"]

    def test_a_shape_that_is_not_a_group_is_never_ungrouped(self):
        from ppt_mac.groups import _ungroup_shapes_impl

        with _fake_deck(["Title"]) as deck:
            with pytest.raises(ValueError, match=r"is not a group \(type=1\)"):
                _ungroup_shapes_impl(1, "Title")

        assert deck.ungrouped == []


# ---------------------------------------------------------------------------
# Connectors
# ---------------------------------------------------------------------------
@macos_only
class TestConnectorSites:
    """A site can only be a number, because macOS will not say where sites are."""

    def test_a_direction_name_is_refused_by_name(self):
        from ppt_mac.connectors import _add_connector_impl

        result = _add_connector_impl(1, "elbow", "Box", "top", "Circle", 1)

        assert result["error"] == (
            "ppt_add_connector cannot take begin_site as a name on macOS"
        )
        assert "is not available" not in result["error"]
        assert "no `connection site` class" in result["reason"]
        assert result["alternatives"] == [
            "ppt_add_connector with begin_site as a number"
        ]

    def test_both_names_are_listed_when_both_are_names(self):
        from ppt_mac.connectors import _add_connector_impl

        result = _add_connector_impl(1, "elbow", "Box", "top", "Circle", "left")

        assert "begin_site, end_site" in result["error"]

    def test_the_default_numeric_site_still_goes_through(self):
        """Both site arguments default to 1, so the common path must survive."""
        from ppt_mac.connectors import _add_connector_impl

        with _fake_deck(["Box", "Circle"]):
            result = _add_connector_impl(1, "elbow", "Box", 1, "Circle", 1)

        assert result["success"] is True

    def test_a_site_past_the_end_of_the_shape_says_the_range(self):
        from ppt_mac.connectors import _add_connector_impl

        with _fake_deck(["Box", "Circle"]):
            with pytest.raises(ValueError, match="has 4 sites"):
                _add_connector_impl(1, "elbow", "Box", 9, "Circle", 1)

    def test_the_site_refusal_never_reaches_powerpoint(self):
        from ppt_mac.connectors import _add_connector_impl, _format_connector_impl

        with _no_powerpoint():
            assert "error" in _add_connector_impl(
                1, "elbow", "Box", "top", "Circle", 1
            )
            assert "error" in _format_connector_impl(
                1, "Connector", None, None, None, None, None, None,
                None, None, None, "Box", "right", None, None,
            )


@macos_only
class TestAddConnector:
    """Made with `make`, attached with two commands, then read back."""

    def test_a_connector_is_created_attached_and_named(self):
        from appscript import k

        from ppt_mac.connectors import _add_connector_impl

        with _fake_deck(["Box", "Circle"]) as deck:
            result = _add_connector_impl(1, "elbow", "Box", 2, "Circle", 3)

        connector = deck.shape("Shape_3")
        assert connector.connector_format.connector_type() == k.elbow
        assert connector.connector_format.connections == [
            ("begin", "Box", 2), ("end", "Circle", 3),
        ]
        assert connector.rerouted == 1
        assert result == {
            "success": True,
            "shape_name": "Shape_3",
            "connector_type": "elbow",
        }

    def test_an_unknown_type_is_rejected_before_anything_is_made(self):
        from ppt_mac.connectors import _add_connector_impl

        with _no_powerpoint():
            with pytest.raises(ValueError, match="Unknown connector_type"):
                _add_connector_impl(1, "squiggle", "Box", 1, "Circle", 1)

    def test_a_shape_name_that_is_not_there_leaves_no_stray_connector(self):
        from ppt_mac.connectors import _add_connector_impl

        with _fake_deck(["Box"]) as deck:
            with pytest.raises(ValueError, match="'Circle' not found"):
                _add_connector_impl(1, "elbow", "Box", 1, "Circle", 1)

        assert len(deck.slide_shapes) == 1

    def test_a_make_that_adds_nothing_is_the_silent_no_op(self):
        from ppt_mac.connectors import _add_connector_impl

        with _fake_deck(["Box", "Circle"]) as deck:
            deck.make_does_nothing = True
            result = _add_connector_impl(1, "elbow", "Box", 1, "Circle", 1)

        assert result["error"] == "ppt_add_connector is not available on macOS"
        assert "gained no shape" in result["reason"]

    def test_a_make_that_lands_a_plain_shape_is_not_reported_as_a_connector(self):
        from ppt_mac.connectors import _add_connector_impl

        with _fake_deck(["Box", "Circle"]) as deck:
            deck.make_is_a_connector = False
            result = _add_connector_impl(1, "elbow", "Box", 1, "Circle", 1)

        assert "not a connector" in result["reason"]

    def test_a_make_powerpoint_refuses_outright_says_what_it_answered(self):
        from ppt_mac.connectors import _add_connector_impl

        with _fake_deck(["Box", "Circle"]) as deck:
            deck.make_raises = True
            result = _add_connector_impl(1, "elbow", "Box", 1, "Circle", 1)

        assert "stub error -1708" in result["reason"]
        assert result["alternatives"] == ["ppt_add_shape with a line shape"]

    def test_an_end_that_did_not_attach_names_the_argument_not_the_tool(self):
        from ppt_mac.connectors import _add_connector_impl

        with _fake_deck(["Box", "Circle"]) as deck:
            deck.connections_take = False
            result = _add_connector_impl(1, "elbow", "Box", 1, "Circle", 1)

        assert result["error"] == (
            "ppt_add_connector could not attach the connector on macOS"
        )
        assert "Shape_3" in result["reason"]


@macos_only
class TestFormatConnector:
    """Colour, weight, dash and six arrowhead properties, all on line format."""

    def test_line_properties_are_translated_to_enumerators(self, monkeypatch):
        from appscript import k

        from ppt_mac import connectors as mac_connectors

        # `raw` reaches a property by its four character code and needs a real
        # appscript reference to do it, so the fake is handed the same job.
        monkeypatch.setattr(
            mac_connectors, "raw", lambda ref, code: ref.dash_style_raw
        )

        with _fake_deck(["Connector"]) as deck:
            result = mac_connectors._format_connector_impl(
                1, "Connector", "#FF0000", 2.5, "dash",
                "none", "long", "wide",
                "triangle", "short", "narrow",
                None, None, None, None,
            )

        line = deck.shape("Connector").line_format
        assert line.fore_color() == [255, 0, 0]
        assert line.line_weight() == 2.5
        assert line.dash_style_raw() == k.line_dash_style_dash
        assert line.begin_arrowhead_style() == k.no_arrowhead
        assert line.begin_arrow_head_length() == k.long_arrowhead
        assert line.begin_arrowhead_width() == k.wide_arrowhead
        assert line.end_arrowhead_style() == k.triangle_arrowhead
        assert line.end_arrowhead_length() == k.short_arrowhead
        assert line.end_arrowhead_width() == k.narrow_width_arrowhead
        assert result == {"success": True, "shape_name": "Connector"}

    @pytest.mark.parametrize(
        "argument,position",
        [("dash_style", 4), ("begin_arrow", 5), ("end_arrow_width", 10)],
    )
    def test_a_word_that_is_not_in_the_map_is_rejected_first(self, argument, position):
        """A typo in the last argument must not leave the first four applied."""
        from ppt_mac.connectors import _format_connector_impl

        args = [1, "Connector"] + [None] * 13
        args[position] = "nonsense"

        with _fake_deck(["Connector"]) as deck:
            with pytest.raises(ValueError, match=f"Unknown {argument}"):
                _format_connector_impl(*args)

        assert deck.shape("Connector").line_format.line_weight() is None

    def test_reconnecting_goes_through_the_connector_format(self):
        from ppt_mac.connectors import _format_connector_impl

        with _fake_deck(["Box", "Circle", "Connector"]) as deck:
            deck.shape("Connector").connector = True
            result = _format_connector_impl(
                1, "Connector", None, None, None, None, None, None,
                None, None, None, "Box", 2, "Circle", None,
            )

        connector = deck.shape("Connector")
        assert connector.connector_format.connections == [
            ("begin", "Box", 2), ("end", "Circle", 1),
        ]
        assert connector.rerouted == 1
        assert result["success"] is True

    def test_a_shape_that_is_not_a_connector_cannot_be_reconnected(self):
        from ppt_mac.connectors import _format_connector_impl

        with _fake_deck(["Box", "Circle"]):
            with pytest.raises(ValueError, match="is not a connector"):
                _format_connector_impl(
                    1, "Box", None, None, None, None, None, None,
                    None, None, None, "Circle", None, None, None,
                )

    def test_a_reconnection_that_did_not_take_says_the_formatting_landed(self):
        from ppt_mac.connectors import _format_connector_impl

        with _fake_deck(["Box", "Connector"]) as deck:
            deck.shape("Connector").connector = True
            deck.connections_take = False
            result = _format_connector_impl(
                1, "Connector", None, 3.0, None, None, None, None,
                None, None, None, "Box", None, None, None,
            )

        assert result["error"] == (
            "ppt_format_connector could not reconnect the connector on macOS"
        )
        assert "line formatting landed" in result["reason"]
        assert deck.shape("Connector").line_format.line_weight() == 3.0


# ---------------------------------------------------------------------------
# A slide made of stand-ins, so none of this needs PowerPoint. Only the parts
# the three modules touch are modelled, and the collections behave the way
# PowerPoint's do, including answering -1728 for an empty one.
# ---------------------------------------------------------------------------
class _Prop:
    """A property that remembers what was written to it."""

    def __init__(self, value=None):
        self.value = value
        self.writes: list = []
        self.clamp_to = None

    def __call__(self):
        return self.value

    def get(self):
        return self.value

    def set(self, value):
        self.writes.append(value)
        self.value = self.clamp_to if self.clamp_to is not None else value


def _bag(**properties):
    """A stand-in for one of PowerPoint's little property-only classes."""
    holder = type("Bag", (), {})()
    for name, value in properties.items():
        setattr(holder, name, _Prop(value))
    return holder


class _FakeList:
    def __init__(self, items):
        self._items = list(items)

    def get(self):
        if not self._items:
            # PowerPoint raises rather than answering an empty list.
            raise _command_error(-1728)
        return list(self._items)

    def __getitem__(self, index):
        if index < 1 or index > len(self._items):
            raise _command_error(-1728)
        return self._items[index - 1]


class _FakeShapes(_FakeList):
    """A shapes collection, which can also be asked for every name at once."""

    @property
    def name(self):
        return _FakeList([shape.name() for shape in self._items])


class _FakeConnectorFormat:
    def __init__(self, deck):
        self._deck = deck
        self.connector_type = _Prop()
        self.connections: list = []

    def begin_connect(self, connected_shape=None, connection_site=None):
        self.connections.append(("begin", connected_shape.name(), connection_site))

    def end_connect(self, connected_shape=None, connection_site=None):
        self.connections.append(("end", connected_shape.name(), connection_site))

    def begin_connected(self):
        return self._deck.connections_take

    def end_connected(self):
        return self._deck.connections_take


class _FakeShape:
    def __init__(self, deck, name, left=10.0, top=20.0, width=30.0, height=40.0):
        from appscript import k

        self._deck = deck
        self._name = name
        self._type = k.shape_type_auto
        self.children: list = []
        self.connector = False
        self.rerouted = 0
        self.glow_format = _bag(radius=None, color=None)
        self.reflection_format = _bag(reflection_type=None)
        self.soft_edge_format = _bag(soft_edge_type=None)
        self.line_format = _bag(
            fore_color=None, line_weight=None, dash_style_raw=None,
            begin_arrowhead_style=None, begin_arrow_head_length=None,
            begin_arrowhead_width=None, end_arrowhead_style=None,
            end_arrowhead_length=None, end_arrowhead_width=None,
        )
        self.connector_format = _FakeConnectorFormat(deck)
        self._box = (left, top, width, height)

    def name(self):
        return self._name

    def shape_type(self):
        return self._type

    def left_position(self):
        return self._box[0]

    def top(self):
        return self._box[1]

    def width(self):
        return self._box[2]

    def height(self):
        return self._box[3]

    def is_connector(self):
        return self.connector

    def connection_site_count(self):
        return 4

    def reroute_connections(self):
        self.rerouted += 1

    @property
    def shapes(self):
        return _FakeShapes(self.children)

    def ungroup(self):
        self._deck.ungroup(self)


class _FakeDeck:
    """One slide holding one shape per name given, plus what `make` adds."""

    def __init__(self, shape_names):
        self.slide_shapes = [_FakeShape(self, name) for name in shape_names]
        self.ungrouped: list = []
        self.connections_take = True
        self.make_does_nothing = False
        self.make_is_a_connector = True
        self.make_raises = False
        self._comes_apart: dict = {}
        self._reply_fails: set = set()

    def shape(self, name):
        for shape in self.slide_shapes:
            if shape.name() == name:
                return shape
        for parent in self.slide_shapes:
            for child in parent.children:
                if child.name() == name:
                    return child
        raise KeyError(name)

    def make_group(self, name, member_names, comes_apart=False, reply_fails=False):
        """Turn one of the slide's shapes into a group holding these members."""
        from appscript import k

        group = self.shape(name)
        group._type = k.shape_type_group
        group.children = [_FakeShape(self, member) for member in member_names]
        self._comes_apart[name] = comes_apart
        if reply_fails:
            self._reply_fails.add(name)

    def ungroup(self, group):
        self.ungrouped.append(group.name())
        if self._comes_apart.get(group.name()):
            index = self.slide_shapes.index(group)
            self.slide_shapes[index:index + 1] = group.children
        if group.name() in self._reply_fails:
            # The command did its work and then answered with something
            # appscript could not unpack, which is the case worth separating.
            raise _command_error(-1708)

    def make(self, new=None, at=None, with_properties=None):
        if self.make_raises:
            raise _command_error(-1708)
        if self.make_does_nothing:
            return None
        shape = _FakeShape(self, f"Shape_{len(self.slide_shapes) + 1}")
        shape.connector = self.make_is_a_connector
        self.slide_shapes.append(shape)
        return object()

    @property
    def slide(self):
        deck = self

        class _Slide:
            end = object()

            @property
            def shapes(self):
                return _FakeShapes(deck.slide_shapes)

        return _Slide()

    @property
    def presentation(self):
        deck = self
        return type("Pres", (), {"slides": _FakeList([deck.slide])})()


class _fake_deck:  # noqa: N801 - reads as a context manager, not a class
    """Point the wrapper at a fake slide for the length of a `with` block."""

    def __init__(self, shape_names):
        self._deck = _FakeDeck(shape_names)

    def __enter__(self):
        from backend.mac_ae import ppt
        from ppt_mac import effects as mac_effects

        self._ppt = ppt
        self._app = ppt._get_app_impl
        self._pres = ppt._get_pres_impl
        ppt._get_app_impl = lambda *a, **kw: self._deck
        ppt._get_pres_impl = lambda *a, **kw: self._deck.presentation

        # `soft edge format` has to be addressed by its four character code,
        # because `font` has a property of the same name whose code wins the
        # name lookup. The fake answers the code the same way PowerPoint would.
        self._effects = mac_effects
        self._raw = mac_effects.raw
        mac_effects.raw = lambda shape, code: (
            shape.soft_edge_format if code == b"DSeF" else self._raw(shape, code)
        )
        return self._deck

    def __exit__(self, *exc):
        self._ppt._get_app_impl = self._app
        self._ppt._get_pres_impl = self._pres
        self._effects.raw = self._raw
        return False


class _no_powerpoint:  # noqa: N801 - reads as a context manager, not a class
    """Make any approach to PowerPoint fail, so a refusal has to be answered first."""

    def __enter__(self):
        from backend.mac_ae import ppt

        def _explode(*args, **kwargs):
            raise AssertionError("a refusal must not touch PowerPoint")

        self._ppt = ppt
        self._app = ppt._get_app_impl
        self._pres = ppt._get_pres_impl
        ppt._get_app_impl = _explode
        ppt._get_pres_impl = _explode
        return self

    def __exit__(self, *exc):
        self._ppt._get_app_impl = self._app
        self._ppt._get_pres_impl = self._pres
        return False


def _command_error(number):
    """Build an appscript CommandError carrying a given OSError number.

    ``errornumber`` is a read-only property on the real class, so the stub
    subclasses it rather than assigning through.
    """
    from appscript.reference import CommandError

    class _Stub(CommandError):
        def __init__(self):
            Exception.__init__(self, f"stub error {number}")

        @property
        def errornumber(self):
            return number

        @property
        def errormessage(self):
            return f"stub error {number}"

        def __str__(self):
            return f"stub error {number}"

    return _Stub()
