"""ppt_set_chart_data swaps a chart's data without Excel where it can (#299).

The chart is replaced by a copy holding the new data. What decides that, what
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
        self.effects = [SimpleNamespace(Shape=s) for s in shapes]

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
def charts(monkeypatch):
    from ppt_com import charts as module

    monkeypatch.setattr(module.chart_file, "available", lambda: True)
    return module


class TestWhenExcelIsStillNeeded:
    def test_a_plain_chart_needs_none(self, charts):
        shape = chart_shape()
        assert charts._excel_needed(slide_with([shape]), shape) is None

    @pytest.mark.parametrize("chart_type, needed", [
        (51, False), (5, False), (-4120, False), (4, False), (1, False), (-4151, False),
        (-4169, True), (15, True), (54, True), (-4102, True),
    ])
    def test_scatter_bubble_and_3d_charts_do(self, charts, chart_type, needed):
        shape = chart_shape(chart_type=chart_type)
        assert (charts._excel_needed(slide_with([shape]), shape) is not None) is needed

    def test_without_python_pptx(self, charts, monkeypatch):
        monkeypatch.setattr(charts.chart_file, "available", lambda: False)
        shape = chart_shape()
        assert "python-pptx" in charts._excel_needed(slide_with([shape]), shape)

    def test_linked_data(self, charts):
        shape = chart_shape()
        shape.Chart.ChartData.IsLinked = True
        assert "linked" in charts._excel_needed(slide_with([shape]), shape)

    def test_a_placeholder_and_a_group(self, charts):
        placeholder = chart_shape(Type=14)
        assert "placeholder" in charts._excel_needed(slide_with([placeholder]), placeholder)
        grouped = chart_shape(Child=-1)
        assert "group" in charts._excel_needed(slide_with([grouped]), grouped)

    def test_an_animation_in_either_sequence(self, charts):
        shape = chart_shape()
        assert "animated" in charts._excel_needed(slide_with([shape], animated=[shape]), shape)
        assert "animated" in charts._excel_needed(slide_with([shape], interactive=[shape]), shape)
        other = chart_shape(shape_id=9)
        assert charts._excel_needed(slide_with([shape, other], animated=[other]), shape) is None

    @pytest.mark.parametrize("end", ["Begin", "End"])
    def test_an_attached_connector(self, charts, end):
        shape = chart_shape()
        ends = SimpleNamespace(BeginConnected=0, EndConnected=0)
        setattr(ends, f"{end}Connected", -1)
        setattr(ends, f"{end}ConnectedShape", shape)
        connector = SimpleNamespace(Connector=-1, ConnectorFormat=ends)
        assert "connector" in charts._excel_needed(slide_with([shape, connector]), shape)

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


class FakePowerPoint:
    """The scratch deck, the file and the paste back, as _replace_chart_data
    meets them."""

    def __init__(self, slide, copy, fail_at=None):
        self.slide, self.copy, self.fail_at = slide, copy, fail_at
        self.closed = []
        fake = self

        class Scratch:
            PageSetup = SimpleNamespace(SlideWidth=0, SlideHeight=0)
            Slides = SimpleNamespace(Add=lambda i, layout: "scratch slide")

            def SaveAs(self, path):
                if fake.fail_at == "save":
                    raise RuntimeError("cannot save")
                Path(path).write_bytes(b"pptx")

            def Close(self):
                fake.closed.append("scratch")

        source = SimpleNamespace(
            Slides=lambda i: SimpleNamespace(Shapes=lambda j: "copy in file"),
            Close=lambda: self.closed.append("source"))
        self.app = SimpleNamespace(Presentations=SimpleNamespace(
            Add=lambda with_window: Scratch(),
            Open=lambda path, *a: source))


def replace(charts, monkeypatch, *, check=None, carry=None, fail_at=None, copy_z=4):
    import contextlib

    old = chart_shape(z=2)
    copy = chart_shape(name="Chart 1", shape_id=7, z=copy_z)
    copy.deleted = False
    copy.Delete = lambda: setattr(copy, "deleted", True)
    old.deleted = False

    def delete_old():
        old.deleted = True
        copy.ZOrderPosition -= 1  # everything above it moves down one

    old.Delete = delete_old

    def send_backward(how):
        copy.ZOrderPosition -= 1

    copy.ZOrder = send_backward
    slide = slide_with([old, copy])
    pp = FakePowerPoint(slide, copy, fail_at)

    def paste_from(source, target):
        if source == "copy in file":
            if fail_at == "paste back":
                raise RuntimeError("Could not paste the new chart")
            return copy
        return "copy in scratch"

    @contextlib.contextmanager
    def kept():
        yield []

    monkeypatch.setattr(charts, "_paste_from", paste_from)
    monkeypatch.setattr(charts, "kept_clipboard", kept)
    monkeypatch.setattr(charts.chart_file, "replace_data", lambda path, c, s: None)
    monkeypatch.setattr(charts, "_check_chart_data", check or (lambda chart, c, s: (None, None)))
    if carry is not None:
        monkeypatch.setattr(charts, "_carry_over", carry)
    pres = SimpleNamespace(PageSetup=SimpleNamespace(SlideWidth=960, SlideHeight=540))
    return old, copy, pp, (lambda: charts._replace_chart_data(pp.app, pres, slide, old, CATS, SERIES))


class TestReplacing:
    def test_the_copy_takes_the_old_chart_s_place(self, charts, monkeypatch):
        old, copy, pp, run = replace(charts, monkeypatch)
        new, warnings = run()
        assert new is copy and not copy.deleted
        assert old.deleted
        assert copy.ZOrderPosition == 2, "in the old chart's place in the stacking order"
        assert pp.closed == ["scratch", "source"]
        assert warnings == []

    @pytest.mark.parametrize("how, match", [
        ("check", "does not show the data"),
        ("carry", "could not carry over Name"),
        ("paste back", "Could not paste"),
        ("save", "cannot save"),
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
        old, copy, pp, run = replace(charts, monkeypatch, **kwargs)
        with pytest.raises(charts._NotCarried, match=match):
            run()
        assert not old.deleted
        assert copy.deleted is (how in ("check", "carry"))
        assert "scratch" in pp.closed

    def test_a_copy_that_cannot_be_removed_is_named(self, charts, monkeypatch):
        old, copy, pp, run = replace(
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
