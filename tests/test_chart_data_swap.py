"""ppt_set_chart_data swaps a chart's data without Excel where it can (#299).

The chart is copied, its XML gets the new data, and it is pasted back as a
new shape holding it. What decides that, what
the copy takes over, and what happens when it cannot, are checked here against
stand-ins for PowerPoint. The live checks are in the PR.
"""

import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

pytestmark = pytest.mark.skipif(
    sys.platform != "win32", reason="drives the COM impls, which macOS replaces with its own")

CATS = ["A", "B"]
SERIES = [{"name": "S", "values": [1, 2]}]


class Tags:
    def __init__(self, pairs=()):
        self.pairs = list(pairs)

    @property
    def Count(self):
        return len(self.pairs)

    def Name(self, i):
        return self.pairs[i - 1][0]

    def Value(self, i):
        return self.pairs[i - 1][1]

    def Add(self, name, value):
        self.pairs = [p for p in self.pairs if p[0] != name] + [(name, value)]


class Action:
    def __init__(self, action=0, address="", sub=""):
        self.Action = action
        self.Hyperlink = SimpleNamespace(Address=address, SubAddress=sub)


def chart_shape(name="Chart 1", shape_id=3, z=2, chart_type=51, **extra):
    shape = SimpleNamespace(
        Name=name, Id=shape_id, ZOrderPosition=z, Type=3, Child=0, HasChart=True, Connector=0,
        Left=10.0, Top=20.0, Width=300.0, Height=200.0, Rotation=0.0, Visible=-1,
        LockAspectRatio=0, AlternativeText="", Title="", Tags=Tags(),
        Chart=SimpleNamespace(ChartType=chart_type, ChartData=SimpleNamespace(IsLinked=False)),
    )
    actions = {1: Action(), 2: Action()}
    shape.ActionSettings = lambda i: actions[i]
    for key, value in extra.items():
        setattr(shape, key, value)
    return shape


class Sequence:
    def __init__(self, shapes=()):
        self.effects = [SimpleNamespace(Shape=s, Timing=SimpleNamespace(TriggerShape=None))
                        for s in shapes]

    @property
    def Count(self):
        return len(self.effects)

    def __call__(self, i):
        return self.effects[i - 1]


def slide_with(shapes, animated=(), interactive=(), comments=0):
    sequences = [Sequence(interactive)] if interactive else []
    slide = SimpleNamespace(
        TimeLine=SimpleNamespace(
            MainSequence=Sequence(animated),
            InteractiveSequences=type("I", (), {
                "Count": len(sequences), "__call__": lambda self, i: sequences[i - 1]})()),
        Comments=SimpleNamespace(Count=comments),
    )
    slide.Shapes = type("Shapes", (), {
        "Count": property(lambda self: len(shapes)),
        "__call__": lambda self, i: shapes[i - 1]})()
    return slide


@pytest.fixture
def charts():
    from ppt_com import charts as module

    return module


class TestWhenExcelIsStillNeeded:
    def test_a_plain_chart_needs_none(self, charts):
        shape = chart_shape()
        assert charts._excel_needed(slide_with([shape]), shape) is None

    @pytest.mark.parametrize("chart_type, needed", [
        (51, False), (5, False), (-4120, False), (4, False), (1, False), (-4151, False),
        (-4169, False), (74, False), (15, False), (87, False), (54, False), (-4102, False),
        (-4100, False), (60, False), (-4098, False), (70, False), (68, False),
        (99, False), (92, False), (112, False), (84, False), (86, False),
        (83, True), (85, True),
        (88, True), (89, True), (90, True), (91, True), (118, True),
    ])
    def test_only_stock_filled_surface_and_unknown_types_do(self, charts, chart_type, needed):
        """Swapping rewrites the chart's own series, so a type gvml has no
        template for, such as 3-D area, cone or wireframe, needs no Excel. A
        filled surface chart comes back in other colours, so it stays on Excel."""
        shape = chart_shape(chart_type=chart_type)
        assert (charts._excel_needed(slide_with([shape]), shape) is not None) is needed

    @pytest.mark.parametrize("chart_type", [88, 89, 90, 91])
    def test_a_stock_chart_is_named_as_one(self, charts, chart_type):
        shape = chart_shape(chart_type=chart_type)
        assert "stock chart" in charts._excel_needed(slide_with([shape]), shape)

    def test_linked_data(self, charts):
        shape = chart_shape()
        shape.Chart.ChartData.IsLinked = True
        assert "linked" in charts._excel_needed(slide_with([shape]), shape)

    def test_a_placeholder_and_a_group(self, charts):
        placeholder = chart_shape(Type=14)
        assert "placeholder" in charts._excel_needed(slide_with([placeholder]), placeholder)
        grouped = chart_shape(Child=-1)
        assert "group" in charts._excel_needed(slide_with([grouped]), grouped)

    def test_animations_and_connectors_no_longer_need_it(self, charts):
        """They are handed to the copy (#343)."""
        shape = chart_shape()
        assert charts._excel_needed(slide_with([shape], animated=[shape]), shape) is None
        assert charts._excel_needed(slide_with([shape], interactive=[shape]), shape) is None
        ends = SimpleNamespace(BeginConnected=-1, BeginConnectedShape=shape, EndConnected=0)
        connector = SimpleNamespace(Type=1, Connector=-1, ConnectorFormat=ends)
        assert charts._excel_needed(slide_with([shape, connector]), shape) is None

    def test_a_check_that_cannot_be_made_means_excel(self, charts):
        class Broken:
            Name = "Chart 1"

            @property
            def Chart(self):
                raise RuntimeError("busy")

        shape = Broken()
        reason = charts._excel_needed(slide_with([]), shape)
        assert reason.startswith("it could not be checked") and "busy" in reason

    def test_comments_on_the_slide(self, charts):
        shape = chart_shape()
        assert "comments" in charts._excel_needed(slide_with([shape], comments=1), shape)


class TestCarryingOver:
    def test_everything_is_carried(self, charts):
        old = chart_shape(AlternativeText="alt", Title="t", Rotation=5.0, Tags=Tags([("K", "V")]))
        old.ActionSettings(1).Action = 7
        old.ActionSettings(1).Hyperlink.Address = "https://example.com/"
        new = chart_shape(name="Chart 9", shape_id=8, Left=0.0, Top=0.0)
        assert charts._carry_over(charts._shape_state(old), new) == []
        assert charts._shape_state(new) == charts._shape_state(old)

    def test_what_does_not_take_is_named(self, charts):
        old = chart_shape(AlternativeText="alt")

        class Stubborn(SimpleNamespace):
            def __setattr__(self, name, value):
                if name in ("AlternativeText", "Name") and hasattr(self, name):
                    return
                super().__setattr__(name, value)

        base = chart_shape(name="Chart 9")
        new = Stubborn(**vars(base))
        failed = charts._carry_over(charts._shape_state(old), new)
        assert failed == ["AlternativeText", "Name"]


WORKBOOK = "clipboard/embeddings/Microsoft_Excel_Worksheet.xlsx"


def copied(chart_type=51, workbook=True, plots=1):
    """A package as PowerPoint copies a chart: with its workbook, by default."""
    from gvml import charts as gvml_charts
    from gvml import package

    xml = gvml_charts.chart_xml(chart_type)
    if plots > 1:
        # A second plot group, as a combo chart has.
        xml = xml.replace("</c:barChart>", "</c:barChart><c:lineChart><c:grouping val=\"standard\"/>"
                          "<c:varyColors val=\"0\"/><c:axId val=\"10\"/><c:axId val=\"20\"/></c:lineChart>")
    if workbook:
        xml = xml.replace("</c:chartSpace>", '<c:externalData r:id="rId9"><c:autoUpdate val="0"/>'
                          "</c:externalData></c:chartSpace>")
    raw = package.read(gvml_charts.chart_package("Chart 1", 0, 0, 100, 100, xml))
    if workbook:
        raw.parts[WORKBOOK] = b"xlsx"
        part = raw.chart_part()
        raw.parts[part.rsplit("/", 1)[0] + "/_rels/" + part.rsplit("/", 1)[1] + ".rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet.xlsx"/>'
            "</Relationships>").encode()
    return raw


def replace(charts, monkeypatch, *, check=None, carry=None, fail_at=None, copy_z=4,
            colours=None, stuck=False, state=None, chart_type=51, package=None,
            categories=CATS, series=SERIES):
    import contextlib

    from gvml import package as gvml_package

    old = chart_shape(z=2, chart_type=chart_type)
    copy = chart_shape(name="Chart 1", shape_id=7, z=copy_z)
    copy.deleted = False
    copy.Delete = lambda: setattr(copy, "deleted", True)
    old.deleted = False

    def delete_old():
        old.deleted = True
        copy.ZOrderPosition -= 1  # everything above it moves down one

    old.Delete = delete_old

    steps = []

    def send_backward(how):
        steps.append(how)
        if len(steps) > 50:
            raise AssertionError("the stacking loop does not end")
        if not stuck:
            copy.ZOrderPosition -= 1

    copy.ZOrder = send_backward
    slide = slide_with([old, copy])
    seen = {"put": [], "pasted": False}

    def copied_package(shape):
        assert shape is old
        if fail_at == "copy":
            raise RuntimeError("Copying the chart put no Art::GVML ClipFormat on the clipboard")
        return package if package is not None else copied(chart_type)

    def paste_from(put, target):
        put()
        if fail_at == "paste back":
            raise RuntimeError("Could not paste the chart")
        seen["pasted"] = True
        return copy

    @contextlib.contextmanager
    def kept():
        yield []

    monkeypatch.setattr(charts, "_copied_package", copied_package)
    monkeypatch.setattr(charts, "_paste_from", paste_from)
    monkeypatch.setattr(charts, "kept_clipboard", kept)
    monkeypatch.setattr(charts.clipboard, "put", lambda fmt, raw: seen["put"].append(
        (fmt, gvml_package.read(raw))))
    monkeypatch.setattr(charts, "_check_chart_data", check or (lambda chart, c, s: (None, None)))
    if carry is not None:
        monkeypatch.setattr(charts, "_carry_over", carry)
    if state is not None:
        monkeypatch.setattr(charts, "_shape_state", state)
    shown = iter(colours or [[(1, 2)], [(1, 2)]])
    monkeypatch.setattr(charts, "_series_colours", lambda chart: next(shown))
    pres = SimpleNamespace(PageSetup=SimpleNamespace(SlideWidth=960, SlideHeight=540))
    run = lambda: charts._replace_chart_data("app", pres, slide, old, categories, series)  # noqa: E731
    return old, copy, seen, run


class TestReplacing:
    def test_the_copy_takes_the_old_chart_s_place(self, charts, monkeypatch):
        old, copy, seen, run = replace(charts, monkeypatch)
        new, warnings = run()
        assert new is copy and not copy.deleted
        assert old.deleted
        assert copy.ZOrderPosition == 2, "in the old chart's place in the stacking order"
        assert warnings == []

    def test_what_is_pasted_holds_the_new_data_and_no_workbook(self, charts, monkeypatch):
        from gvml import charts as gvml_charts

        old, copy, seen, run = replace(charts, monkeypatch)
        run()
        [(fmt, pasted)] = seen["put"]
        assert fmt == "Art::GVML ClipFormat"
        data = gvml_charts.read_chart(pasted.chart())
        assert data == {"categories": CATS, "series": [{"name": "S", "values": [1.0, 2.0]}]}
        assert gvml_charts.external_data_id(pasted.chart()) is None
        assert WORKBOOK not in pasted.parts, "the old workbook would say something else"

    @pytest.mark.parametrize("how, match", [
        ("check", "does not show the data"),
        ("carry", "could not carry over Name"),
        ("paste back", "Could not paste"),
        ("copy", "put no Art::GVML ClipFormat"),
    ])
    def test_a_copy_that_falls_short_is_removed_and_the_old_chart_kept(
            self, charts, monkeypatch, how, match):
        kwargs = {}
        if how == "check":
            kwargs["check"] = lambda chart, c, s: ("it has 0 categories, not 2", None)
        elif how == "carry":
            kwargs["carry"] = lambda state, shape: ["Name"]
        else:
            kwargs["fail_at"] = how
        old, copy, seen, run = replace(charts, monkeypatch, **kwargs)
        with pytest.raises(charts._NotCarried, match=match):
            run()
        assert not old.deleted
        assert copy.deleted is (how in ("check", "carry"))

    def test_animations_are_handed_over_before_the_old_chart_goes(self, charts, monkeypatch):
        calls = []
        carry = charts.animation_carry
        monkeypatch.setattr(carry, "animated", lambda slide, shape_id: True)
        monkeypatch.setattr(carry, "timeline", lambda slide, chart_id, leave_out=None: ["same"])
        monkeypatch.setattr(carry, "copy_animations", lambda slide, old, new: calls.append("copy"))
        monkeypatch.setattr(carry, "move_links", lambda slide, old, new, c, undo: calls.append("links"))
        old, copy, seen, run = replace(charts, monkeypatch)
        new, warnings = run()
        assert calls == ["copy", "links"] and old.deleted and not copy.deleted

    def test_animations_that_did_not_come_over_leave_the_old_chart_as_it_was(
            self, charts, monkeypatch):
        """The copy goes, which takes its effects along, and what was pointed
        at it is pointed back first, while it is still there to point from."""
        events = []
        carry = charts.animation_carry
        timelines = iter([["before"], ["after"]])
        monkeypatch.setattr(carry, "animated", lambda slide, shape_id: True)
        monkeypatch.setattr(carry, "timeline", lambda slide, chart_id, leave_out=None: next(timelines))
        monkeypatch.setattr(carry, "copy_animations", lambda slide, old, new: None)
        monkeypatch.setattr(carry, "move_links",
                            lambda slide, old, new, c, undo: undo.append(lambda: events.append("pointed back")))
        old, copy, seen, run = replace(charts, monkeypatch)
        copy.Delete = lambda: (events.append("copy deleted"), setattr(copy, "deleted", True))
        with pytest.raises(charts._NotCarried, match="animations did not come over"):
            run()
        assert events == ["pointed back", "copy deleted"]
        assert not old.deleted

    def test_links_that_cannot_be_pointed_back_leave_both_charts_and_say_so(
            self, charts, monkeypatch):
        carry = charts.animation_carry
        timelines = iter([["before"], ["after"]])
        monkeypatch.setattr(carry, "animated", lambda slide, shape_id: True)
        monkeypatch.setattr(carry, "timeline", lambda slide, chart_id, leave_out=None: next(timelines))
        monkeypatch.setattr(carry, "copy_animations", lambda slide, old, new: None)

        def stuck():
            raise RuntimeError("busy")

        monkeypatch.setattr(carry, "move_links", lambda slide, old, new, c, undo: undo.append(stuck))
        old, copy, seen, run = replace(charts, monkeypatch)
        with pytest.raises(RuntimeError, match=r"could not all be pointed back \(busy\)"):
            run()
        assert not old.deleted and not copy.deleted

    def test_links_that_fail_halfway_are_pointed_back(self, charts, monkeypatch):
        events = []
        carry = charts.animation_carry
        monkeypatch.setattr(carry, "animated", lambda slide, shape_id: False)

        def halfway(slide, old, new, c, undo):
            undo.append(lambda: events.append("pointed back"))
            raise RuntimeError("the second connector would not move")

        monkeypatch.setattr(carry, "move_links", halfway)
        old, copy, seen, run = replace(charts, monkeypatch)
        with pytest.raises(charts._NotCarried, match="second connector"):
            run()
        assert events == ["pointed back"] and copy.deleted and not old.deleted

    def test_an_old_chart_that_will_not_go_gets_its_links_back(self, charts, monkeypatch):
        events = []
        carry = charts.animation_carry
        monkeypatch.setattr(carry, "animated", lambda slide, shape_id: False)
        monkeypatch.setattr(carry, "move_links",
                            lambda slide, old, new, c, undo: undo.append(lambda: events.append("pointed back")))
        old, copy, seen, run = replace(charts, monkeypatch)

        def refuse():
            raise RuntimeError("locked")

        old.Delete = refuse
        copy.Delete = lambda: (events.append("copy deleted"), setattr(copy, "deleted", True))
        with pytest.raises(charts._NotCarried, match="could not be removed"):
            run()
        assert events == ["pointed back", "copy deleted"]

    def test_data_that_does_not_fit_the_chart_leaves_it_as_it_was(self, charts, monkeypatch):
        """A scatter chart given labels for x, say: Excel then, as before."""
        old, copy, seen, run = replace(charts, monkeypatch, chart_type=-4169)
        with pytest.raises(charts._NotCarried, match="'A' is not a number"):
            run()
        assert not old.deleted and not copy.deleted and not seen["pasted"], "nothing was pasted back"

    def test_a_bubble_chart_needs_its_sizes(self, charts, monkeypatch):
        old, copy, seen, run = replace(charts, monkeypatch, chart_type=15, categories=["1", "2"])
        with pytest.raises(charts._NotCarried, match="no sizes for its bubbles"):
            run()
        assert not old.deleted and not seen["pasted"]

    def test_a_bubble_chart_with_sizes_is_swapped(self, charts, monkeypatch):
        from gvml import charts as gvml_charts

        old, copy, seen, run = replace(
            charts, monkeypatch, chart_type=15, categories=["1", "2.5"],
            series=[{"name": "S", "values": [1, 2], "sizes": [3, "4"]}])
        run()
        xml = seen["put"][0][1].chart().decode("utf-8")
        assert "<c:v>2.5</c:v>" in xml and "<c:v>4</c:v>" in xml
        assert old.deleted
        assert gvml_charts.read_chart(seen["put"][0][1].chart())["categories"] == ["1.0", "2.5"]

    def test_a_3d_area_chart_is_swapped_in_its_own_kind(self, charts, monkeypatch):
        """A kind gvml cannot add is still rewritten in place (#343)."""
        from gvml import charts as gvml_charts
        from gvml import package

        xml = gvml_charts.chart_xml(1).replace("c:areaChart>", "c:area3DChart>")
        assert xml.count("c:area3DChart>") == 2
        raw = package.read(gvml_charts.chart_package("Chart 1", 0, 0, 100, 100, xml))
        old, copy, seen, run = replace(charts, monkeypatch, chart_type=-4098, package=raw)
        run()
        new = seen["put"][0][1].chart()
        assert gvml_charts.plot_kinds(new) == ["area3DChart"]
        assert gvml_charts.read_chart(new)["categories"] == CATS
        assert old.deleted

    def test_a_stock_chart_goes_through_excel(self, charts, monkeypatch):
        from gvml import charts as gvml_charts
        from gvml import package

        xml = gvml_charts.chart_xml(4).replace("c:lineChart>", "c:stockChart>")
        raw = package.read(gvml_charts.chart_package("Chart 1", 0, 0, 100, 100, xml))
        old, copy, seen, run = replace(charts, monkeypatch, package=raw)
        with pytest.raises(charts._NotCarried, match="stock chart"):
            run()
        assert not old.deleted and not seen["pasted"]

    def test_a_combo_chart_goes_through_excel(self, charts, monkeypatch):
        old, copy, seen, run = replace(charts, monkeypatch, package=copied(plots=2))
        with pytest.raises(charts._NotCarried, match="2 plot groups"):
            run()
        assert not old.deleted and not seen["pasted"]

    def test_series_that_come_back_in_other_colours_are_not_kept(self, charts, monkeypatch):
        old, copy, seen, run = replace(charts, monkeypatch, colours=[[(1, 2), (3, 4)], [(1, 2), (9, 9)]])
        with pytest.raises(charts._NotCarried, match="other colours"):
            run()
        assert copy.deleted and not old.deleted

    def test_a_series_more_or_fewer_is_no_colour_change(self, charts, monkeypatch):
        old, copy, seen, run = replace(charts, monkeypatch, colours=[[(1, 2)], [(1, 2), (5, 6)]])
        run()
        assert old.deleted

    def test_a_stacking_step_that_does_not_move_ends_the_loop(self, charts, monkeypatch):
        old, copy, seen, run = replace(charts, monkeypatch, stuck=True)
        with pytest.raises(charts._NotCarried, match="stacking order"):
            run()
        assert copy.deleted and not old.deleted

    def test_what_cannot_be_read_off_the_old_chart_is_not_carried(self, charts, monkeypatch):
        def no_title(shape):
            raise AttributeError("Title")

        old, copy, seen, run = replace(charts, monkeypatch, state=no_title)
        with pytest.raises(charts._NotCarried, match="could not read what it has to take over"):
            run()
        assert not old.deleted and seen["put"] == []

    def test_an_old_chart_and_copy_that_both_stay_name_the_copy(self, charts, monkeypatch):
        old, copy, seen, run = replace(charts, monkeypatch)

        def refuse():
            raise RuntimeError("cannot delete")

        old.Delete = refuse
        copy.Delete = refuse
        with pytest.raises(RuntimeError, match=r"delete the shape 'Chart 1' \(id 7\), which is the copy"):
            run()

    def test_a_copy_that_cannot_be_removed_is_named(self, charts, monkeypatch):
        old, copy, seen, run = replace(
            charts, monkeypatch, check=lambda chart, c, s: ("empty", None))

        def refuse():
            raise RuntimeError("cannot delete")

        copy.Delete = refuse
        with pytest.raises(RuntimeError, match=r"delete the shape 'Chart 1' \(id 7\)"):
            run()
        assert not old.deleted


class TestTheResultSaysWhatWasDone:
    def deck(self, charts, monkeypatch, shape):
        monkeypatch.setattr(charts.ppt, "_get_app_impl", lambda *a, **k: "app", raising=False)
        monkeypatch.setattr(charts.ppt, "_get_pres_impl", lambda *a, **k: SimpleNamespace(
            Slides=lambda i: "slide"), raising=False)
        monkeypatch.setattr(charts, "goto_slide", lambda *a: None)
        monkeypatch.setattr(charts, "_get_chart_shape", lambda slide, ref: shape)

    def test_replaced_without_excel(self, charts, monkeypatch):
        shape = chart_shape(shape_id=3)
        self.deck(charts, monkeypatch, shape)
        monkeypatch.setattr(charts, "_excel_needed", lambda slide, s: None)
        monkeypatch.setattr(charts, "_replace_chart_data", lambda *a: (
            SimpleNamespace(Name="Chart 1", Id=11), []))
        result = charts._set_chart_data_impl(1, "Chart 1", CATS, SERIES)
        assert result["method"] == "replaced without Excel"
        assert (result["replaced"], result["shape_id"], result["previous_shape_id"]) == (True, 11, 3)

    def test_a_copy_that_fell_short_goes_through_excel_and_says_why(self, charts, monkeypatch):
        shape = chart_shape()
        self.deck(charts, monkeypatch, shape)
        monkeypatch.setattr(charts, "_excel_needed", lambda slide, s: None)

        def fall_short(*a):
            raise charts._NotCarried("could not carry over Tags")

        monkeypatch.setattr(charts, "_replace_chart_data", fall_short)
        monkeypatch.setattr(charts, "_set_chart_data_through_excel",
                            lambda slide, s, c, ser: {"success": True, "shape_name": "Chart 1"})
        result = charts._set_chart_data_impl(1, "Chart 1", CATS, SERIES)
        assert result["method"] == "Excel"
        assert result["reason"] == "swapping it without Excel failed: could not carry over Tags"
        assert "shape_id" not in result

    def test_a_failure_through_excel_says_why_excel_was_used(self, charts, monkeypatch):
        shape = chart_shape()
        self.deck(charts, monkeypatch, shape)
        monkeypatch.setattr(charts, "_excel_needed", lambda slide, s: "it is animated")

        def fail(*a):
            raise RuntimeError("The chart did not take the data")

        monkeypatch.setattr(charts, "_set_chart_data_through_excel", fail)
        with pytest.raises(RuntimeError, match=r"since it is animated"):
            charts._set_chart_data_impl(1, "Chart 1", CATS, SERIES)

    def test_bubble_sizes_are_refused_rather_than_dropped_through_excel(self, charts, monkeypatch):
        shape = chart_shape(chart_type=15)
        self.deck(charts, monkeypatch, shape)
        monkeypatch.setattr(charts, "_excel_needed", lambda slide, s: "it is animated")
        written = []
        monkeypatch.setattr(charts, "_set_chart_data_through_excel",
                            lambda *a: written.append(a) or {"success": True})
        with pytest.raises(RuntimeError, match=r"Bubble sizes .* since it is animated"):
            charts._set_chart_data_impl(
                1, "Chart 1", ["1", "2"], [{"name": "S", "values": [1, 2], "sizes": [3, 4]}])
        assert written == [], "nothing was written"

    def test_a_bubble_chart_with_no_sizes_is_refused_too(self, charts, monkeypatch):
        shape = chart_shape(chart_type=15)
        self.deck(charts, monkeypatch, shape)
        monkeypatch.setattr(charts, "_excel_needed", lambda slide, s: "it is animated")
        written = []
        monkeypatch.setattr(charts, "_set_chart_data_through_excel",
                            lambda *a: written.append(a) or {"success": True})
        with pytest.raises(ValueError, match="Nothing was changed"):
            charts._set_chart_data_impl(1, "Chart 1", ["1", "2"], [{"name": "S", "values": [1, 2]}])
        assert written == [], "the Excel way would drop the sizes"


class Series:
    def __init__(self, name, values, labels):
        self.Name, self.Values, self.XValues = name, tuple(values), tuple(labels)


def chart_of(chart_type, *series):
    collection = type("Collection", (), {
        "Count": len(series), "__call__": lambda self, i: series[i - 1]})()
    return SimpleNamespace(
        ChartType=chart_type, SeriesCollection=lambda *i: collection(*i) if i else collection)


class TestReadingTheChartBack:
    @pytest.mark.parametrize("chart_type", [68, 71])
    def test_the_other_slice_of_a_pie_of_pie_is_left_out(self, charts, chart_type):
        """COM adds the "Other" slice as a last point with no label (seen on
        PowerPoint 16), which is not part of the data written."""
        chart = chart_of(chart_type, Series("S", [3, 5, 2, 4], ["X", "Y", "Z", ""]))
        categories, series = charts._read_chart_data(chart)
        assert categories == ["X", "Y", "Z"]
        assert series == [{"name": "S", "values": [3, 5, 2]}]
        assert charts._compare_chart_data(
            ["X", "Y", "Z"], [{"name": "S", "values": [3, 5, 2]}], categories, series) == (None, None)

    def test_a_plain_pie_keeps_every_point(self, charts):
        chart = chart_of(5, Series("S", [3, 5, 2], ["X", "Y", "Z"]))
        assert charts._read_chart_data(chart) == (
            ["X", "Y", "Z"], [{"name": "S", "values": [3, 5, 2]}])
