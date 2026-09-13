"""Tests for the three macOS modules that mostly refuse.

Charts, SmartArt and freeform paths are the areas PowerPoint for Mac's
dictionary has the fewest words for, so ``ppt_mac/charts.py``,
``ppt_mac/smartart.py`` and ``ppt_mac/freeform.py`` are almost entirely
refusals. What is worth testing is therefore not what they do but what they
say, and in particular the four things a bad refusal gets wrong.

It must not look like a success, so no refusal carries a ``success`` key. It
must not swallow a caller's mistake, so a wrong shape name or a misspelled
argument still raises the error Windows raises. It must not cost the user
anything, so a refusal never moves the view and never edits the deck. And it
must arrive in the shape the tool wrapper expects, which is a dict for charts
and SmartArt and a JSON string for freeform, because those wrappers differ.

Pure unit tests against a fake object graph. None of this launches PowerPoint,
and none of it may. appscript only installs on macOS, so everything that
imports it is skipped elsewhere.
"""

import json
import sys

import pytest

sys.path.insert(0, "src")

macos_only = pytest.mark.skipif(
    sys.platform != "darwin", reason="the Apple Event backend needs appscript"
)

SDEF = "/Applications/Microsoft PowerPoint.app/Contents/Resources/PowerPoint.sdef"


# ---------------------------------------------------------------------------
# The dictionary these refusals cite
# ---------------------------------------------------------------------------
@macos_only
@pytest.mark.skipif(
    not __import__("os").path.exists(SDEF),
    reason="PowerPoint for Mac is not installed",
)
class TestTheDictionaryStillSaysThis:
    """Every reason in the three modules is a claim about PowerPoint.sdef.

    Reading the file is safe and reading it is not the same as talking to
    PowerPoint, so the claims are checked rather than trusted. appscript parses
    the sdef into the same tables it would build from a live connection, which
    is the closest thing to primary evidence available without launching
    anything.
    """

    @staticmethod
    def _tables():
        """Return the two tables keyed by name, dropping the two keyed by code.

        ``tablesforsdef`` answers with type by code, type by name, reference by
        code and reference by name, in that order.
        """
        from appscript.terminology import tablesforsdef

        with open(SDEF, "rb") as handle:
            _, typebyname, _, refbyname = tablesforsdef(handle.read())
        return typebyname, refbyname

    def test_no_word_for_a_chart_a_smart_art_or_a_node(self):
        """The reference table holds every property, element and command."""
        _, refbyname = self._tables()

        assert "chart" not in refbyname
        assert "smart_art" not in refbyname
        assert "nodes" not in refbyname
        assert "vertices" not in refbyname
        assert "build_freeform" not in refbyname

    def test_the_only_near_misses_are_the_ones_the_modules_name(self):
        """`chart unit effect` and `smart quotes` are the false friends."""
        _, refbyname = self._tables()

        chartish = [n for n in refbyname if "chart" in n]
        smartish = [n for n in refbyname if "smart" in n]

        assert chartish == ["chart_unit_effect"]
        assert sorted(smartish) == ["smart_cut_paste", "smart_quotes"]

    def test_a_shape_has_no_has_chart_beside_its_has_table(self):
        _, refbyname = self._tables()

        assert "has_table" in refbyname
        assert "has_chart" not in refbyname

    def test_the_smart_art_node_words_are_declared_and_taken_by_nothing(self):
        """The clearest evidence in the port that this was removed.

        The enumerators for placing a node into a SmartArt graphic are all
        still in the type table, and the reference table takes none of them.
        """
        typebyname, refbyname = self._tables()

        for word in (
            "after_node", "before_node", "above_node", "below_node",
            "default_node", "assistant_node",
        ):
            assert word in typebyname, word
            assert word not in refbyname, word

    def test_the_three_shape_types_are_real_and_carry_the_windows_number(self):
        """An existing chart, SmartArt or freeform is still a first class shape.

        The low byte of the enumerator code is the Windows constant, which is
        what lets ``_win_constant`` recognise all three.
        """
        typebyname, _ = self._tables()

        assert typebyname["shape_type_chart"].code[-1] == 3
        assert typebyname["shape_type_free_form"].code[-1] == 5
        assert typebyname["shape_type_smartart_graphic"].code[-1] == 24

    def test_the_near_misses_freeform_names_are_really_there(self):
        """`line shape`, `motion effect` path and `path format` all exist.

        The freeform module says each of these was checked and rejected. If one
        of them stopped existing the wording would be wrong, and if one gained a
        node the refusal would be wrong.
        """
        _, refbyname = self._tables()

        assert "begin_line_X" in refbyname
        assert "end_line_Y" in refbyname
        assert "motion_effect" in refbyname
        assert "path" in refbyname
        assert "path_format" in refbyname
        assert "adjustments" in refbyname
        assert "adjustment_value" in refbyname


# ---------------------------------------------------------------------------
# The payload every refusal shares
# ---------------------------------------------------------------------------
@macos_only
class TestTheRefusalHelperIsOneHelper:
    """Eighteen modules name ``_refusal`` and all eighteen mean one function."""

    def test_every_module_names_the_same_function(self):
        import importlib

        from backend.unsupported import refusal

        modules = (
            "advanced_ops",
            "animation",
            "charts",
            "comments",
            "connectors",
            "edit_ops",
            "effects",
            "freeform",
            "groups",
            "hyperlinks",
            "media",
            "properties",
            "sections",
            "slideshow",
            "smartart",
            "tables",
            "text",
            "themes",
        )
        for name in modules:
            module = importlib.import_module(f"ppt_mac.{name}")
            assert module._refusal is refusal, name

    def test_the_encoded_form_cannot_drift_from_the_dict_one(self):
        """``unsupported`` is the JSON of the same payload, built by the same call."""
        import json

        from backend.unsupported import refusal, unsupported

        encoded = unsupported("ppt_x", "because", ["ppt_y"])
        assert json.loads(encoded) == refusal("ppt_x", "because", ["ppt_y"])

    def test_it_builds_the_documented_shape(self):
        from ppt_mac.charts import _refusal

        assert _refusal("ppt_x", "because") == {
            "error": "ppt_x is not available on macOS",
            "reason": "because",
            "platform": "macOS",
        }

    def test_alternatives_are_omitted_rather_than_padded(self):
        from ppt_mac.charts import _refusal

        assert "alternatives" not in _refusal("ppt_x", "because", [])

    def test_the_error_override_names_an_argument_instead_of_a_tool(self):
        """No tool in these three modules needs it, so it is checked here.

        All seventeen refuse wholesale. Other modules do use the override, and
        this keeps it covered from the side that reads the shared helper.
        """
        from ppt_mac.charts import _refusal

        payload = _refusal("ppt_x", "because", error="width is not honoured")
        assert payload["error"] == "width is not honoured"


@macos_only
class TestNothingHereEditsASlide:
    """A refused call must not move the user's view.

    ``goto_slide`` is what moves it, and the cheapest way to be sure none of
    the seventeen calls it is that none of the three modules imports it or
    calls it. The docstrings say so in prose, which is why the check looks for
    the call and the import rather than for the name.
    """

    def test_no_module_reaches_for_goto_slide(self):
        import inspect

        from ppt_mac import charts, freeform, smartart

        for module in (charts, smartart, freeform):
            source = inspect.getsource(module)
            assert "goto_slide(" not in source, module.__name__
            assert "import goto_slide" not in source, module.__name__

    def test_no_module_reaches_for_the_application_either(self):
        """``_get_app_impl`` is what ``goto_slide`` needs on the Windows path."""
        import inspect

        from ppt_mac import charts, freeform, smartart

        for module in (charts, smartart, freeform):
            assert "_get_app_impl" not in inspect.getsource(module), module.__name__


@macos_only
class TestTheSwapActuallyHappened:
    """A refusal nobody reaches is the same as a hidden tool."""

    @pytest.mark.parametrize(
        "module_name,impl_name",
        [
            ("ppt_com.charts", "_add_chart_impl"),
            ("ppt_com.charts", "_format_chart_axis_impl"),
            ("ppt_com.smartart", "_modify_smartart_impl"),
            ("ppt_com.freeform", "_build_freeform_impl"),
            ("ppt_com.freeform", "_set_segment_type_impl"),
        ],
    )
    def test_the_com_module_answers_with_the_apple_event_one(self, module_name, impl_name):
        import importlib

        module = importlib.import_module(module_name)
        impl = getattr(module, impl_name)
        assert impl.__module__ == module_name.replace("ppt_com", "ppt_mac")


# ---------------------------------------------------------------------------
# Charts
# ---------------------------------------------------------------------------
@macos_only
class TestAddChart:
    """The chart is pasted from XML now; the typo check still comes first.

    What happens after the type resolves is in test_mac_gvml.py.
    """

    def test_a_misspelled_chart_type_is_still_heard(self):
        """The typo is the caller's real problem, so it comes first."""
        from ppt_mac.charts import _add_chart_impl

        with _no_powerpoint():
            with pytest.raises(ValueError, match="Unknown chart type 'colunm'"):
                _add_chart_impl(1, "colunm", 50, 50, 500, 350)


@macos_only
class TestChartArgumentsRefusedByName:
    """The five editing tools work now (test_mac_gvml.py). What is left to
    refuse is the odd argument, and each one is named before the chart is
    touched, so dropping it and calling again costs nothing."""

    def test_formatting_refuses_the_unmapped_arguments_by_name(self):
        from ppt_mac.charts import _format_chart_impl

        with _no_powerpoint():
            result = _format_chart_impl(
                1, "Q3 Revenue", "Revenue", True, "bottom", 3, 12.0,
                None, None, None, None, None,
            )

        assert result["error"] == "ppt_format_chart cannot set chart_style, legend_font_size on macOS"
        assert "drop the argument" in result["reason"]
        assert "success" not in result

    def test_an_eight_direction_legend_position_is_refused_by_argument(self):
        from ppt_mac.charts import _format_chart_impl

        with _no_powerpoint():
            result = _format_chart_impl(
                1, "Q3 Revenue", None, None, "top-left", None, None,
                None, None, None, None, None,
            )

        assert result["error"] == "ppt_format_chart cannot set an 8-direction legend_position on macOS"
        assert "'bottom', 'left', 'right', 'top' or 'corner'" in result["alternatives"][0]

    def test_a_misspelled_legend_position_is_heard_the_windows_way(self):
        from ppt_mac.charts import _format_chart_impl

        with _no_powerpoint():
            with pytest.raises(ValueError, match="Unknown legend position 'botom'"):
                _format_chart_impl(
                    1, "Q3 Revenue", None, None, "botom", None, None,
                    None, None, None, None, None,
                )

    def test_the_axis_font_size_is_refused_by_name(self):
        from ppt_mac.charts import _format_chart_axis_impl

        with _no_powerpoint():
            result = _format_chart_axis_impl(
                1, "Q3 Revenue", "value", None,
                None, None, None, None, None, None, None, None,
                None, None, None, 9.0, None,
            )

        assert result["error"] == "ppt_format_chart_axis cannot set tick_label_font_size on macOS"

    def test_the_windows_axis_guards_come_before_powerpoint(self):
        from ppt_mac.charts import _format_chart_axis_impl

        with _no_powerpoint():
            with pytest.raises(ValueError, match="min_scale/max_scale are only valid for value axes"):
                _format_chart_axis_impl(
                    1, "Q3 Revenue", "category", None,
                    0.0, 100.0, None, None, None, None, None, None,
                    None, None, None, None, None,
                )
            with pytest.raises(ValueError, match="log_base requires log_scale=true"):
                _format_chart_axis_impl(
                    1, "Q3 Revenue", "value", None,
                    None, None, None, None, None, None, None, None,
                    None, None, 2.0, None, None,
                )

    def test_a_misspelled_axis_is_still_heard(self):
        from ppt_mac.charts import _format_chart_axis_impl

        with _no_powerpoint():
            with pytest.raises(ValueError, match="Unknown axis 'catagory'"):
                _format_chart_axis_impl(
                    1, "Q3 Revenue", "catagory", None,
                    None, None, None, None, None, None, None, None,
                    None, None, None, None, None,
                )

    def test_changing_the_type_hears_a_typo_and_refuses_an_unknown_kind_by_argument(self):
        from ppt_mac.charts import _change_chart_type_impl

        with _no_powerpoint():
            with pytest.raises(ValueError, match="Unknown chart type 'pei'"):
                _change_chart_type_impl(1, "Q3 Revenue", "pei")
            result = _change_chart_type_impl(1, "Q3 Revenue", -4100)

        assert result["error"] == "ppt_change_chart_type cannot draw chart_type -4100 on macOS"
        assert "success" not in result


# ---------------------------------------------------------------------------
# SmartArt
# ---------------------------------------------------------------------------
@macos_only
class TestTheSmartArtShapeTypeGap:
    """24 used to be missing from the generated table. The generator carries it."""

    def test_the_generated_table_has_smart_art(self):
        """Windows says msoSmartArt, macOS says `shape type smartart graphic`."""
        from appscript import k

        from backend.mac_enums import MsoShapeType
        from ppt_com.constants import msoSmartArt

        assert MsoShapeType[msoSmartArt] == k.shape_type_smartart_graphic

    def test_the_module_pairs_it_with_the_dictionary_word(self):
        from appscript import k

        from ppt_com.constants import msoSmartArt
        from ppt_mac.smartart import _SHAPE_TYPES, _WIN_SHAPE_TYPE

        assert _SHAPE_TYPES[msoSmartArt] == k.shape_type_smartart_graphic
        assert _WIN_SHAPE_TYPE[k.shape_type_smartart_graphic] == msoSmartArt

    def test_a_real_smart_art_shape_reads_back_as_its_windows_number(self):
        """The bug the generator override exists to prevent, stated as a test."""
        from appscript import k

        from backend.mac_enums import MsoShapeType
        from ppt_com.constants import msoSmartArt
        from ppt_mac.shapes import _win_constant

        generated = {word: number for number, word in MsoShapeType.items()}
        assert _win_constant(generated, k.shape_type_smartart_graphic) == msoSmartArt

    def test_the_alternative_promises_a_type_it_can_deliver(self):
        """``ppt_get_shape_info`` reads through the same generated table.

        It carries 24 now, so both halves of the answer arrive and the
        alternative can say so without a caveat.
        """
        from ppt_mac.smartart import _SHAPE_TOOLS

        assert "name, box and type" in _SHAPE_TOOLS[0]
        assert "null" not in _SHAPE_TOOLS[0]


@macos_only
class TestSmartArt:
    """No class, and the words for its nodes outlived the thing itself."""

    def test_adding_refuses_and_names_the_orphaned_vocabulary(self):
        from ppt_mac.smartart import _add_smartart_impl

        with _no_powerpoint():
            result = _add_smartart_impl(
                1, "Basic Process", None, 50, 50, 400, 300,
                ["a", "b"], None, None, None, None, None, None,
            )

        assert result["error"] == "ppt_add_smartart is not available on macOS"
        assert "no `smart art` class" in result["reason"]
        assert "`after node`" in result["reason"]
        assert "success" not in result

    def test_modifying_refuses_and_names_the_shape_and_the_action(self):
        from ppt_mac.smartart import _modify_smartart_impl

        with _fake_deck(["Process"]) as deck:
            deck.set_type("Process", "smartart")
            result = _modify_smartart_impl(
                1, "Process", "set_text", 2, "Step two",
                None, None, None, None, None, None, None, None,
                None, None, None,
            )

        assert result["error"] == "ppt_modify_smartart is not available on macOS"
        assert "'Process' on slide 1 is a SmartArt graphic" in result["reason"]
        assert "the action 'set_text'" in result["reason"]
        assert "ppt_ungroup_shapes" in result["alternatives"][1]
        assert "success" not in result

    def test_a_misspelled_action_is_still_heard(self):
        from ppt_mac.smartart import _modify_smartart_impl

        with _fake_deck(["Process"]) as deck:
            deck.set_type("Process", "smartart")
            with pytest.raises(ValueError, match="Unknown action 'set_txet'"):
                _modify_smartart_impl(
                    1, "Process", "set_txet", 2, "Step two",
                    None, None, None, None, None, None, None, None,
                    None, None, None,
                )

    def test_a_shape_that_is_not_smart_art_says_what_it_is(self):
        from ppt_mac.smartart import _modify_smartart_impl

        with _fake_deck(["Title"]):
            with pytest.raises(ValueError, match=r"is not a SmartArt graphic \(type=AutoShape\)"):
                _modify_smartart_impl(
                    1, "Title", "set_text", 1, "x",
                    None, None, None, None, None, None, None, None,
                    None, None, None,
                )

    def test_the_action_check_runs_after_the_shape_check(self):
        """Two mistakes at once, and the shape is the one worth naming."""
        from ppt_mac.smartart import _modify_smartart_impl

        with _fake_deck(["Title"]):
            with pytest.raises(ValueError, match="is not a SmartArt graphic"):
                _modify_smartart_impl(
                    1, "Title", "set_txet", 1, "x",
                    None, None, None, None, None, None, None, None,
                    None, None, None,
                )

    def test_the_eight_actions_match_the_message_that_lists_them(self):
        from ppt_mac.smartart import _ACTIONS, _modify_smartart_impl

        with _fake_deck(["Process"]) as deck:
            deck.set_type("Process", "smartart")
            for action in _ACTIONS:
                result = _modify_smartart_impl(
                    1, "Process", action, 1, "x",
                    None, None, None, None, None, None, None, None,
                    None, None, None,
                )
                assert f"the action '{action}'" in result["reason"]


@macos_only
class TestListingSmartArtLayouts:
    """Servable from a static table and deliberately not served."""

    def test_it_refuses_rather_than_inventing_indices(self):
        from ppt_mac.smartart import _list_smartart_options_impl

        with _no_powerpoint():
            result = _list_smartart_options_impl("layouts", None, None, False)

        assert result["error"] == (
            "ppt_list_smartart_layouts is not available on macOS"
        )
        assert "numbers that mean nothing" in result["reason"]

    def test_it_carries_no_success_and_no_catalogue(self):
        """A list of layouts nothing can use would read as a working tool."""
        from ppt_mac.smartart import _list_smartart_options_impl

        with _no_powerpoint():
            result = _list_smartart_options_impl("layouts", None, None, False)

        assert "success" not in result
        assert "layouts" not in result
        assert "total_count" not in result

    @pytest.mark.parametrize("list_type", ["layouts", "colors", "styles", "categories"])
    def test_all_four_list_types_refuse(self, list_type):
        from ppt_mac.smartart import _list_smartart_options_impl

        with _no_powerpoint():
            assert "error" in _list_smartart_options_impl(list_type, None, None, False)

    def test_a_misspelled_list_type_is_still_heard(self):
        from ppt_mac.smartart import _list_smartart_options_impl

        with _no_powerpoint():
            with pytest.raises(ValueError, match="Unknown list_type 'layout'"):
                _list_smartart_options_impl("layout", None, None, False)


# ---------------------------------------------------------------------------
# Freeform
# ---------------------------------------------------------------------------
@macos_only
class TestFreeformWireForm:
    """The one freeform refusal left returns a JSON string, like the rest.

    The freeform tool wrappers hand back whatever ``ppt.execute`` returns
    without encoding it, so a dict here would reach the model as a Python repr
    rather than as JSON. Worth its own test because nothing else would catch it.
    """

    def test_the_editing_type_refusal_is_a_string(self):
        result = _editing_type_refusal()
        assert isinstance(result, str)
        assert json.loads(result)["platform"] == "macOS"

    def test_it_does_not_look_like_a_success(self):
        payload = json.loads(_editing_type_refusal())
        assert "success" not in payload
        assert payload["error"].endswith("is not available on macOS")


@macos_only
class TestFreeformNodeTools:
    """Four of the five node editors work now (test_mac_gvml.py). The
    editing type is the one with nowhere to be written, and it says so."""

    def test_the_editing_type_names_the_shape_and_the_node(self):
        payload = json.loads(_editing_type_refusal())
        assert "'Arrow' on slide 1 is a freeform" in payload["reason"]
        assert "node 5" in payload["reason"]
        assert "no such attribute either" in payload["reason"]
        assert "ppt_set_node_position" in payload["reason"]

    def test_a_shape_that_is_not_a_freeform_says_so_the_way_windows_does(self):
        from ppt_mac.freeform import _get_shape_nodes_impl

        with _fake_deck(["Title"]):
            with pytest.raises(
                ValueError,
                match=r"'Title' is not a freeform \(type=1\)\. "
                      r"Only freeform shapes \(type=5\)",
            ):
                _get_shape_nodes_impl(1, "Title", None)

    def test_a_shape_index_out_of_range_says_so(self):
        from ppt_mac.freeform import _get_shape_nodes_impl

        with _fake_deck(["Title"]):
            with pytest.raises(ValueError, match="Shape index 9 is out of range"):
                _get_shape_nodes_impl(1, None, 9)

    def test_it_points_at_the_shape_tools(self):
        payload = json.loads(_editing_type_refusal())
        assert "ppt_list_shapes" in payload["alternatives"]


# ---------------------------------------------------------------------------
# The fake object graph
# ---------------------------------------------------------------------------
def _editing_type_refusal():
    """Run the one refusing freeform impl against a freeform."""
    from ppt_mac.freeform import _set_node_editing_type_impl

    with _fake_deck(["Arrow"]) as deck:
        deck.set_type("Arrow", "freeform")
        return _set_node_editing_type_impl(1, "Arrow", None, 5, 2)


def _command_error(number):
    """An appscript CommandError carrying an Apple Event error number."""
    from appscript.reference import CommandError

    error = CommandError.__new__(CommandError)
    error.errornumber = number
    error.args = (number,)
    return error


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


class _FakeShape:
    """Only what these three modules ask a shape for, which is very little."""

    def __init__(self, name):
        from appscript import k

        self._name = name
        self._type = k.shape_type_auto

    def name(self):
        return self._name

    def shape_type(self):
        return self._type


class _FakeDeck:
    """One slide holding one shape per name given."""

    _TYPES = {
        "chart": "shape_type_chart",
        "smartart": "shape_type_smartart_graphic",
        "freeform": "shape_type_free_form",
        "auto": "shape_type_auto",
    }

    def __init__(self, shape_names):
        self.slide_shapes = [_FakeShape(name) for name in shape_names]

    def set_type(self, name, kind):
        """Make one of the slide's shapes report a chart, SmartArt or freeform."""
        from appscript import k

        for shape in self.slide_shapes:
            if shape.name() == name:
                shape._type = getattr(k, self._TYPES[kind])
                return
        raise KeyError(name)

    @property
    def slide(self):
        deck = self

        class _Slide:
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

        self._ppt = ppt
        self._app = ppt._get_app_impl
        self._pres = ppt._get_pres_impl
        # Nothing in these three modules should reach for the application. Only
        # goto_slide needs it, and none of them navigates.
        ppt._get_app_impl = _explode
        ppt._get_pres_impl = lambda *a, **kw: self._deck.presentation
        return self._deck

    def __exit__(self, *exc):
        self._ppt._get_app_impl = self._app
        self._ppt._get_pres_impl = self._pres
        return False


def _explode(*args, **kwargs):
    raise AssertionError("a refusal must not touch PowerPoint")


class _no_powerpoint:  # noqa: N801 - reads as a context manager, not a class
    """Make any approach to PowerPoint fail, so a refusal has to come first."""

    def __enter__(self):
        from backend.mac_ae import ppt

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
