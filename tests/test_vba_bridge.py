"""The VBA route on macOS: the bridge, the proxy, the router and the add-in.

PowerPoint for Mac's VBA reaches what its Apple Event dictionary does not, and
``backend.vba`` runs a tool's Windows implementation through it. Nothing here
reaches PowerPoint: the add-in's side of the protocol is played by fakes that
write the answer file the way ``PptMcpBridge.bas`` does. The live checks are
described in the pull request that added the route.
"""

import importlib
import inspect
import json
import re
import sys
import threading
import zipfile
from pathlib import Path

import pytest

sys.path.insert(0, "src")

from backend import vba  # noqa: E402
from backend.vba import routes  # noqa: E402

RS, US = "\x1e", "\x1f"
VBA_DIR = Path(vba.__file__).parent

macos_only = pytest.mark.skipif(sys.platform != "darwin", reason="the Apple Event backend needs appscript")


# ---------------------------------------------------------------------------
# The add-in's side, played by a fake
# ---------------------------------------------------------------------------
class FakeEvent:
    def __init__(self, target, code, params):
        self.target, self.code, self.params = target, code, params

    def send(self, timeout, flags):
        self.target.sent.append({"code": self.code, "params": self.params, "timeout": timeout, "flags": flags})
        rid = self.params[b"tpAE"][0]
        answer = self.target.answer(self.params[b"Str#"], self.params[b"tpAE"][1:])
        if answer is not None:
            (self.target.folder / f"ppt-mcp-bridge-{rid}.txt").write_bytes(answer.encode("utf-16-le"))


class FakeTarget:
    """Stands in for the aem.Application; answers as the add-in would."""

    def __init__(self, folder, answer):
        self.folder, self.answer, self.sent = folder, answer, []

    def event(self, code, params):
        return FakeEvent(self, code, params)


def decode(body_hex):
    return bytes.fromhex(body_hex).decode("utf-16-be")


def bridge_answering(tmp_path, *records):
    """A bridge whose add-in answers each Dispatch with the next records."""
    queue = list(records)
    seen = []

    def answer(macro, params):
        seen.append((macro, decode(params[0]) if params else None))
        return queue.pop(0) if queue else None

    target = FakeTarget(tmp_path, answer)
    return vba.Bridge(target, answers=tmp_path, timeout=2), target, seen


class TestTheProtocol:
    def test_a_batch_is_sent_as_utf16_hex_and_the_answers_decoded(self, tmp_path):
        bridge, target, seen = bridge_answering(
            tmp_path,
            RS.join(["ok" + US + "num" + US + "Long" + US + "3",
                     "ok" + US + "str" + US + "String" + US + "Grüße ✓",
                     "ok" + US + "bool" + US + "1",
                     "ok" + US + "obj" + US + "Shape" + US + "h7",
                     "ok" + US + "num" + US + "Single" + US + "12.5",
                     "ok" + US + "nothing"]) + RS)
        out = bridge.run([("app", "g", "Version"), ("app", "g", "Name"), ("app", "g", "Visible"),
                          ("app", "c", "Item", 1), ("app", "g", "Left"), ("app", "g", "Parent")])
        assert out[:3] == [3, "Grüße ✓", True]
        assert isinstance(out[3], vba.Proxy) and (out[3]._h, out[3]._t) == ("h7", "Shape")
        assert out[4] == 12.5 and out[5] is None
        macro, body = seen[0]
        assert macro == vba.DISPATCH
        assert body.split(RS)[3] == US.join(["app", "c", "Item", "i:1"])

    def test_arguments_are_typed(self):
        p = vba.Proxy(None, "h3", "Shape")
        assert vba.Bridge._arg(p) == "o:h3"
        assert vba.Bridge._arg(None) == "m:"
        assert vba.Bridge._arg(True) == "b:1"
        assert vba.Bridge._arg(7) == "i:7"
        assert vba.Bridge._arg(1.5) == "d:1.5"
        assert vba.Bridge._arg("x") == "s:x"
        assert vba.Bridge._arg(["A", "B"]) == "a:A\x1dB"
        assert vba.Bridge._arg([1, 2.5]) == "N:1\x1d2.5"

    def test_arrays_come_back_as_pywin32_hands_them_over(self, tmp_path):
        bridge, _, _ = bridge_answering(
            tmp_path,
            "ok" + US + "array1" + US + "3" + US + "n1" + US + "n2.5" + US + "sx" + RS
            + "ok" + US + "array2" + US + "2" + US + "2" + US + "n10" + US + "n20" + US + "n30" + US + "n40" + RS)
        one, two = bridge.run([("h1", "g", "Values"), ("h1", "g", "Vertices")])
        assert one == (1, 2.5, "x")
        assert two == ((10, 20), (30, 40))

    def test_an_unknown_member_is_an_attribute_error(self, tmp_path):
        bridge, _, _ = bridge_answering(tmp_path, "err" + US + "1" + US + "438" + US + "VBAProject" + US + "nope" + RS)
        with pytest.raises(AttributeError, match="<unknown>.MergeShapes"):
            bridge.run([("h1", "c", "MergeShapes", 1)])

    def test_a_powerpoint_failure_is_shaped_as_pywin32_raises_it(self, tmp_path):
        bridge, _, _ = bridge_answering(
            tmp_path, "ok" + US + "empty" + RS + "err" + US + "2" + US + "-2147188160"
            + US + "Microsoft PowerPoint" + US + "Shapes.Item : Integer out of range." + RS)
        with pytest.raises(vba.COMError) as info:
            bridge.run([("app", "g", "Version"), ("h1", "c", "Item", 99)])
        e = info.value
        assert e.args[0] == -2147352567 and e.args[1] == "Exception occurred."
        assert e.excepinfo[1:3] == ("Microsoft PowerPoint", "Shapes.Item : Integer out of range.")
        assert e.excepinfo[5] == -2147188160

    def test_the_event_goes_without_waiting_and_with_whole_seconds(self, tmp_path):
        bridge, target, _ = bridge_answering(tmp_path, "ok" + US + "empty" + RS)
        bridge.run([("app", "g", "Version")])
        sent = target.sent[0]
        assert sent["code"] == vba.RUN_VB_MACRO
        # aem refuses a float timeout, and a reply a hidden PowerPoint never
        # sends must not be waited for: the answer file says it is done.
        assert isinstance(sent["timeout"], int)
        assert sent["flags"] & 3 == 1   # kAENoReply, not kAEWaitReply
        assert not list(tmp_path.glob("ppt-mcp-bridge-*")), "the answer file is removed"

    def test_no_answer_is_a_timeout_that_says_what_to_do(self, tmp_path):
        bridge, _, _ = bridge_answering(tmp_path)  # answers nothing
        bridge.timeout = 0.2
        with pytest.raises(vba.BridgeTimeout, match="Close any dialog"):
            bridge.run([("app", "g", "Version")])

    def test_ping_reports_the_add_in_version(self, tmp_path):
        target = FakeTarget(tmp_path, lambda macro, params: "ok" + US + "1" + US + "16.114.0")
        bridge = vba.Bridge(target, answers=tmp_path, timeout=2)
        assert bridge.ping() == ("1", "16.114.0")
        assert target.sent[0]["params"][b"Str#"] == vba.PING


# ---------------------------------------------------------------------------
# The proxy
# ---------------------------------------------------------------------------
class RecordingBridge:
    def __init__(self, answers=None):
        self.requests, self.answers = [], answers or {}

    def one(self, handle, kind, member, *args):
        self.requests.append((handle, kind, member, args))
        result = self.answers.get((kind, member))
        if isinstance(result, Exception):
            raise result
        return result


class TestTheProxy:
    def test_a_property_is_read_when_named(self):
        b = RecordingBridge({("g", "Name"): "Box"})
        assert vba.Proxy(b, "h1", "Shape").Name == "Box"
        assert b.requests == [("h1", "g", "Name", ())]

    def test_a_method_is_handed_back_and_only_called_when_called(self):
        b = RecordingBridge()
        add = vba.Proxy(b, "h1", "Shapes").AddShape
        assert b.requests == []
        add(1, 10, 20, 30, 40)
        assert b.requests == [("h1", "c", "AddShape", (1, 10, 20, 30, 40))]

    def test_keyword_arguments_go_to_their_places(self):
        b = RecordingBridge()
        vba.Proxy(b, "h1", "Shapes").AddTextbox(1, 0, 0, Width=50, Height=20)
        assert b.requests[0][3] == (1, 0, 0, 50, 20)

    def test_a_value_is_let_and_an_object_property_is_set(self):
        b = RecordingBridge()
        shape = vba.Proxy(b, "h1", "Shape")
        shape.Left = 12.0
        assert b.requests[-1][:3] == ("h1", "l", "Left")

    def test_calling_an_object_calls_its_default_member(self):
        b = RecordingBridge()
        vba.Proxy(b, "h2", "Shapes")(3)
        assert b.requests == [("h2", "c", "Item", (3,))]

    def test_a_property_with_its_return_value_named_is_still_a_property(self):
        # The dual-interface view names a property's return value RHS; the
        # shipped table must not count it as an argument, or Count would be
        # taken for a method that needs one.
        members, _ = vba._interface("SeriesCollection")
        assert "RHS" not in members["Count"][1]
        b = RecordingBridge({("g", "Count"): 3})
        assert vba.Proxy(b, "h4", "SeriesCollection").Count == 3

    def test_an_interface_the_table_lacks_is_read_then_called(self):
        b = RecordingBridge({("g", "Refresh"): AttributeError("<unknown>.Refresh"), ("c", "Refresh"): None})
        vba.Proxy(b, "h9", "SomethingNew").Refresh
        assert [r[1] for r in b.requests] == ["g", "c"]


# ---------------------------------------------------------------------------
# The router
# ---------------------------------------------------------------------------
class FakeApp:
    def __init__(self, decks):
        self._decks = decks

    @property
    def Presentations(self):  # noqa: N802 - the COM name
        decks = self._decks

        class Coll:
            Count = len(decks)

            def __call__(self, i):
                return decks[i - 1]
        return Coll()


class FakeDeck:
    def __init__(self, name, full):
        self.Name, self.FullName = name, full


class FakeBridgeForRun:
    def __init__(self, decks):
        self.app = FakeApp(decks)
        self.released = 0

    def application(self):
        return self.app

    def release(self):
        self.released += 1


class FakeMacDeck:
    def __init__(self, name, full):
        self._n, self._f = name, full

    def name(self):
        return self._n

    def full_name(self):
        return self._f


@pytest.fixture
def route(monkeypatch):
    """A router over a COM and a Mac implementation that record their calls."""
    import backend

    calls = []
    decks = [FakeDeck("Other.pptx", "/x/Other.pptx"), FakeDeck("Deck.pptx", "/x/Deck.pptx")]
    bridge = FakeBridgeForRun(decks)

    class FakePpt:
        def _get_app_impl(self, *a, **k):
            return "app_ref"

        def _get_pres_impl(self):
            return FakeMacDeck("Deck.pptx", "/x/Deck.pptx")

    monkeypatch.setattr(backend, "ppt", FakePpt())
    monkeypatch.setattr(vba, "ENABLED", True)
    state = {"available": True}

    def get_bridge(app_ref):
        if not state["available"]:
            raise vba.VbaUnavailable("no add-in")
        return bridge

    monkeypatch.setattr(vba._SESSION, "bridge", get_bridge)

    def com_impl(slide_index, shape_name_or_index, screen_tip=None):
        call = vba.current()
        calls.append(("com", slide_index, call.pres.Name if call else None))
        return {"via": "vba"}

    def mac_impl(slide_index, shape_name_or_index, screen_tip=None):
        calls.append(("mac", slide_index, vba.current()))
        return {"via": "apple events"}

    impl = vba.routed(com_impl, mac_impl, routes.given("screen_tip"))
    return impl, calls, bridge, state, mac_impl


class TestTheRouter:
    def test_a_call_that_needs_it_runs_the_windows_code_on_the_same_deck(self, route):
        impl, calls, bridge, _, _ = route
        assert impl(2, "Box", screen_tip="tip") == {"via": "vba"}
        assert calls == [("com", 2, "Deck.pptx")]
        assert vba.current() is None, "the call's context does not outlive it"
        assert bridge.released == 1

    def test_a_call_that_does_not_need_it_stays_on_apple_events(self, route):
        impl, calls, bridge, _, _ = route
        assert impl(2, "Box") == {"via": "apple events"}
        assert calls == [("mac", 2, None)] and bridge.released == 0

    def test_without_the_add_in_the_apple_event_code_runs(self, route):
        impl, calls, _, state, _ = route
        state["available"] = False
        assert impl(2, "Box", screen_tip="tip") == {"via": "apple events"}
        assert calls == [("mac", 2, None)]

    def test_turned_off_the_apple_event_code_runs(self, route, monkeypatch):
        impl, calls, _, _, _ = route
        monkeypatch.setattr(vba, "ENABLED", False)
        assert impl(2, "Box", screen_tip="tip") == {"via": "apple events"}

    def test_a_failure_in_the_windows_code_is_raised_and_not_retried(self, route, monkeypatch):
        impl, calls, bridge, _, mac_impl = route

        def failing(slide_index, shape_name_or_index, screen_tip=None):
            raise vba.COMError(-2147188160, "Microsoft PowerPoint", "Bad value")

        broken = vba.routed(failing, mac_impl, routes.ALWAYS)
        with pytest.raises(vba.COMError):
            broken(1, "Box")
        assert calls == [], "the Apple Event code is not run after a VBA write may have landed"
        assert vba.current() is None and bridge.released == 1

    def test_the_router_unwraps_to_the_apple_event_implementation(self, route):
        impl, _, _, _, mac_impl = route
        assert inspect.unwrap(impl) is mac_impl
        assert impl.vba_route[1] is not None

    def test_same_named_decks_are_told_apart_by_full_name(self):
        app = FakeApp([FakeDeck("Deck.pptx", "/a/Deck.pptx"), FakeDeck("Deck.pptx", "/b/Deck.pptx")])
        assert vba._same_deck(app, "Deck.pptx", "/b/Deck.pptx").FullName == "/b/Deck.pptx"
        with pytest.raises(RuntimeError, match="does not see"):
            vba._same_deck(app, "Missing.pptx", "/m.pptx")


class TestTheRoutes:
    @pytest.mark.parametrize("key", sorted(routes.ROUTES))
    def test_each_route_names_both_implementations(self, key):
        module_name, impl_name = key
        source = (Path("src") / Path(*module_name.split("."))).with_suffix(".py").read_text(encoding="utf-8")
        assert re.search(rf"^def {impl_name}\(", source, re.M), f"{module_name} has no {impl_name}"
        mac = (Path("src/ppt_mac") / (module_name.split(".")[1] + ".py")).read_text(encoding="utf-8")
        assert re.search(rf"^def {impl_name}\(", mac, re.M), f"ppt_mac has no {impl_name} to fall back to"

    def test_the_conditions_ask_for_arguments_the_implementations_take(self):
        checks = {
            ("ppt_com.hyperlinks", "_add_hyperlink_impl"): {"screen_tip"},
            ("ppt_com.tables", "_add_table_row_impl"): {"position"},
            ("ppt_com.animation", "_add_animation_impl"): {"trigger_shape"},
            ("ppt_com.effects", "_set_reflection_impl"): {"blur", "offset", "size", "transparency", "target"},
            ("ppt_com.media", "_set_media_settings_impl"): {"volume", "muted", "fade_in", "fade_out", "start_point", "end_point"},
            ("ppt_com.text", "_set_textframe_impl"): {"auto_size"},
        }
        for (module_name, impl_name), names in checks.items():
            source = (Path("src") / Path(*module_name.split("."))).with_suffix(".py").read_text(encoding="utf-8")
            sig = source[source.index(f"def {impl_name}("):]
            sig = sig[:sig.index(")")]
            for n in names:
                assert re.search(rf"\b{n}\b", sig), f"{impl_name} takes no {n}"

    def test_argument_conditions(self):
        assert routes.given("screen_tip")({"screen_tip": "x"})
        assert not routes.given("screen_tip")({"screen_tip": None})
        assert routes.reflection_details({"blur": 3, "target": "shape"})
        assert routes.reflection_details({"target": "text"})
        assert not routes.reflection_details({"reflection_type": 1, "target": "shape"})
        assert routes.media_levels({"volume": 0.5})
        assert not routes.media_levels({"volume": 0.5, "start_point": 100}), "trimming stays where it is"
        assert not routes.media_levels({"loop": True})
        assert routes.shrink_to_fit({"auto_size": "shrink_to_fit"})
        assert not routes.shrink_to_fit({"auto_size": "none"})

    def test_what_is_not_routed_is_not(self):
        routed = {impl for _, impl in routes.ROUTES}
        for name in ("_merge_shapes_impl", "_execute_mso_impl", "_add_chart_impl", "_set_chart_data_impl",
                     "_add_smartart_impl", "_modify_smartart_impl", "_set_node_editing_type_impl",
                     "_set_window_state_impl", "_start_undo_entry_impl"):
            assert name not in routed


@macos_only
class TestWiredIn:
    def test_a_routed_tool_is_the_router_and_the_rest_the_apple_event_one(self):
        com = importlib.import_module("ppt_com.advanced_ops")
        mac = importlib.import_module("ppt_mac.advanced_ops")
        assert hasattr(com._set_tag_impl, "vba_route")
        assert inspect.unwrap(com._set_tag_impl) is mac._set_tag_impl
        assert com._lock_aspect_ratio_impl is mac._lock_aspect_ratio_impl

    def test_the_mac_wrapper_hands_over_the_bridge_objects_during_a_call(self):
        from backend import mac_ae

        wrapper = mac_ae.PowerPointAppleEventWrapper()
        call = vba.Call(app="bridge app", pres="bridge deck")
        token = vba._CALL.set(call)
        try:
            assert wrapper._get_app_impl() == "bridge app"
            assert wrapper._get_pres_impl() == "bridge deck"
        finally:
            vba._CALL.reset(token)
        with pytest.raises(RuntimeError, match="only through the VBA bridge"):
            wrapper._get_target_window_impl()


def test_navigation_drives_the_window_the_windows_way_during_a_call(monkeypatch):
    import utils.navigation as nav

    moved = []

    class FakeView:
        def GotoSlide(self, n):  # noqa: N802
            moved.append(n)

    class Window:
        View = FakeView()

    class FakePpt:
        def _get_target_window_impl(self):
            return Window()

    monkeypatch.setattr(nav, "IS_MACOS", True)
    monkeypatch.setattr(nav, "ppt", FakePpt())
    monkeypatch.setattr(nav, "FOLLOW_EDITS", True)
    token = vba._CALL.set(vba.Call(app=None, pres=None))
    try:
        nav.goto_slide(None, 4)
    finally:
        vba._CALL.reset(token)
    assert moved == [4]


# ---------------------------------------------------------------------------
# Loading the add-in
# ---------------------------------------------------------------------------
class FakeProperty:
    def __init__(self, value):
        self.value, self.sets = value, []

    def get(self):
        if isinstance(self.value, Exception):
            raise self.value
        return self.value

    def set(self, value):
        self.sets.append(value)
        self.value = value


class FakeAppRef:
    def __init__(self, loaded):
        self.loaded = FakeProperty(loaded)
        self.registered = []

        class AddIns(dict):
            pass
        self.add_ins = AddIns()
        self.add_ins[vba.ADDIN_NAME] = type("AddIn", (), {"loaded": self.loaded})()

    def register_add_in(self, path):
        self.registered.append(path)


@pytest.fixture
def session(monkeypatch, tmp_path):
    monkeypatch.setattr(vba, "INSTALL_DIR", tmp_path / "ppt-mcp")
    versions = []

    class Pinging:
        def __init__(self, target, *a, **k):
            pass

        def ping(self):
            return versions.pop(0) if versions else (vba.BRIDGE_VERSION, "16.114.0")

    monkeypatch.setattr(vba, "Bridge", Pinging)
    return vba._Session(), versions, tmp_path


def _app_ref(app_ref):
    app_ref.AS_appdata = type("D", (), {"target": staticmethod(lambda: "aem target")})()
    return app_ref


class TestLoading:
    def test_a_first_call_installs_registers_and_loads_the_add_in(self, session):
        s, _, tmp_path = session
        app_ref = _app_ref(FakeAppRef(loaded=RuntimeError("no such add-in")))
        s.bridge(app_ref)
        installed = tmp_path / "ppt-mcp" / vba.ADDIN_FILE.name
        assert installed.read_bytes() == vba.ADDIN_FILE.read_bytes()
        assert app_ref.registered == [str(installed)]
        assert app_ref.loaded.sets[-1] is True

    def test_a_loaded_add_in_is_used_as_it_is(self, session):
        s, _, _ = session
        app_ref = _app_ref(FakeAppRef(loaded=False))
        first = s.bridge(app_ref)
        app_ref.registered.clear()
        assert s.bridge(app_ref) is first and app_ref.registered == []

    def test_an_older_add_in_still_loaded_is_reloaded_or_refused(self, session):
        s, versions, _ = session
        versions.extend([("0", "16.114.0"), ("0", "16.114.0")])
        with pytest.raises(vba.VbaUnavailable, match="quit PowerPoint"):
            s.bridge(_app_ref(FakeAppRef(loaded=False)))

    def test_a_failure_is_not_retried_on_every_call(self, session, monkeypatch):
        s, _, _ = session
        tries = []

        def broken(app_ref):
            tries.append(1)
            raise RuntimeError("sandbox said no")
        monkeypatch.setattr(s, "_load", broken)
        for _ in range(3):
            with pytest.raises(vba.VbaUnavailable, match="sandbox said no"):
                s.bridge(object())
        assert len(tries) == 1


# ---------------------------------------------------------------------------
# What ships
# ---------------------------------------------------------------------------
class TestShipped:
    def test_the_add_in_is_a_compiled_vba_project(self):
        with zipfile.ZipFile(vba.ADDIN_FILE) as z:
            assert "ppt/vbaProject.bin" in z.namelist()

    def test_the_sources_and_the_python_side_agree(self):
        bridge = (VBA_DIR / "PptMcpBridge.bas").read_text(encoding="utf-8")
        handles = (VBA_DIR / "PptMcpHandles.bas").read_text(encoding="utf-8")
        assert f'BRIDGE_VERSION As String = "{vba.BRIDGE_VERSION}"' in bridge
        module, dispatch = vba.DISPATCH.split(".")
        assert f'Attribute VB_Name = "{module}"' in bridge
        assert re.search(rf"^Sub {dispatch}\(", bridge, re.M)
        assert re.search(rf"^Sub {vba.PING.split('.')[1]}\(", bridge, re.M)
        assert 'Attribute VB_Name = "PptMcpHandles"' in handles
        # Every entry point traps its errors from its first line: an untrapped
        # one stops PowerPoint on a dialog.
        assert "Sub Dispatch(id As String, batch As String)\n    Dim" in bridge
        assert "On Error GoTo fail" in bridge.split("Sub Dispatch")[1].split("For k")[0]

    def test_the_member_table_knows_properties_from_methods(self):
        members, default = vba._interface("Shapes")
        assert 1 in members["AddShape"][0]
        assert default == "Item"
        shape, _ = vba._interface("Shape")
        assert {2, 4} <= set(shape["Left"][0])


# ---------------------------------------------------------------------------
# PowerPoint's own error text
# ---------------------------------------------------------------------------
def _generator():
    import importlib.util

    spec = importlib.util.spec_from_file_location("gen_vba_calls", Path("scripts/gen_vba_calls.py"))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class TestDirectCalls:
    """CallByName turns every PowerPoint failure into 440, "Automation error".
    The members the routed tools use are called written out instead, which
    keeps PowerPoint's own message; the generated modules must keep up."""

    def test_the_generated_modules_are_up_to_date(self):
        expected = _generator().generate()
        on_disk = {p.stem: p.read_bytes().decode("ascii").replace("\r\n", "\n")
                   for p in VBA_DIR.glob("PptMcpCalls*.bas")}
        assert on_disk == expected, "run: uv run python scripts/gen_vba_calls.py, then rebuild the add-in on Windows"

    def test_they_are_imported_with_crlf_line_endings(self):
        for p in VBA_DIR.glob("PptMcp*.bas"):
            data = p.read_bytes()
            assert b"\r\n" in data and b"\n" not in data.replace(b"\r\n", b""), p.name

    def test_the_members_the_routed_code_uses_are_written_out(self):
        generated = "".join(p.read_text(encoding="ascii") for p in VBA_DIR.glob("PptMcpCalls?.bas"))
        for name in ("Tags", "AddTable", "Rows", "Columns", "Nodes", "SetPosition", "Insert",
                     "SeriesCollection", "Axes", "ChartType", "HasTitle", "TimeLine", "MainSequence",
                     "Select", "Group", "Range", "Item", "Count"):
            assert f'Case "{name}"' in generated, name

    def test_the_dispatcher_tries_the_written_out_call_first(self):
        body = (VBA_DIR / "PptMcpBridge.bas").read_text(encoding="utf-8").split("Sub Dispatch")[1]
        assert body.index("If Not Direct(") < body.index('Case "l": CallByName')


# ---------------------------------------------------------------------------
# Text in groups, slides by layout name, duplicating and copying slides
# ---------------------------------------------------------------------------
class TestGroupMemberRoute:
    """On the Mac a group's member is edited through the clipboard, which
    recreates the group; a top-level shape is edited in place already."""

    @pytest.fixture
    def slide_with(self, monkeypatch):
        import backend

        def make(*top_level):
            class Shapes:
                class name:  # noqa: N801 - appscript's spelling
                    @staticmethod
                    def get():
                        return list(top_level)

            class Slide:
                shapes = Shapes

            class Deck:
                slides = {2: Slide}

            class FakePpt:
                def _get_pres_impl(self):
                    return Deck()

            monkeypatch.setattr(backend, "ppt", FakePpt())
        return make

    def test_a_path_is_a_member(self, slide_with):
        slide_with("Group 4")
        assert routes.group_member({"slide_index": 2, "shape_name_or_index": "Group 4/TextBox 3"})

    def test_a_name_no_top_level_shape_has_is_a_member(self, slide_with):
        slide_with("Group 4", "Title 1")
        assert routes.group_member({"slide_index": 2, "shape_name_or_index": "TextBox 3"})

    def test_a_top_level_shape_stays_on_apple_events(self, slide_with):
        slide_with("Group 4", "Title 1")
        assert not routes.group_member({"slide_index": 2, "shape_name_or_index": "Title 1"})
        assert not routes.group_member({"slide_index": 2, "shape_name_or_index": 1})

    def test_the_new_routes_are_listed(self):
        assert routes.ROUTES[("ppt_com.text", "_set_text_impl")] is routes.group_member
        assert routes.ROUTES[("ppt_com.text", "_format_text_impl")] is routes.group_member
        assert routes.ROUTES[("ppt_com.slides", "_duplicate_slide_impl")] is routes.ALWAYS
        assert routes.ROUTES[("ppt_com.slides", "_copy_slide_impl")] is routes.ALWAYS
        add_slide = routes.ROUTES[("ppt_com.slides", "_add_slide_impl")]
        assert add_slide({"layout_name": "Title Only"}) and not add_slide({"layout": 12})


def test_temporary_files_go_where_powerpoint_may_read_them(monkeypatch, tmp_path):
    # copy_slide's Windows code saves a snapshot of the deck for PowerPoint to
    # insert slides from; outside its container macOS would ask the user.
    import tempfile

    monkeypatch.setattr(vba, "INSTALL_DIR", tmp_path / "ppt-mcp")
    before = tempfile.tempdir
    with vba._temp_files_powerpoint_can_read():
        assert tempfile.mkdtemp(prefix="ppt-mcp-copy-").startswith(str(tmp_path / "ppt-mcp" / "tmp"))
    assert tempfile.tempdir == before
