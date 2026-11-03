"""COM names bound once per interface: utils.com_wrapper.share_type_information.

Most of it runs anywhere, against a stand-in for win32com.client.dynamic, since
what it changes is which olerepr pywin32 hands each new object (#264). The last
part runs pywin32 itself on Windows, against Scripting.Dictionary, a COM server
every Windows has, to show that one object's odd answer no longer changes how
every other object of its interface is read.
"""

import sys

sys.path.insert(0, "src")

import pytest  # noqa: E402

from utils.com_wrapper import is_guess, share_type_information  # noqa: E402

FUNCDESC, INVOKE_FUNC, INVOKE_PROPERTYGET = 1, 1, 2
NULL_IID = "{00000000-0000-0000-0000-000000000000}"


class ComError(Exception):
    pass


class TypeInfo:
    def __init__(self, iid, broken=False):
        self.iid, self.broken = iid, broken

    def GetTypeAttr(self):
        if self.broken:
            raise ComError("no type attributes")
        return (self.iid, 0, 0)


class Olerepr:
    def __init__(self):
        self.mapFuncs, self.propMap = {}, {}


class Entry:
    def __init__(self, desc=None, was_property=0):
        self.desc, self.wasProperty = desc, was_property


class Desc:
    def __init__(self, desckind, invkind):
        self.desckind, self.invkind = desckind, invkind


class CDispatch:
    """Stands in for pywin32's: a read it cannot answer as a property it
    guesses to be a method and records in the olerepr, as dynamic.py does on
    ERRORS_BAD_CONTEXT; a write of a name it does not know records the dispid."""

    def __init__(self, olerepr, odd=()):
        self.__dict__["_olerepr_"] = olerepr
        self.__dict__["odd"] = set(odd)

    def __getattr__(self, attr):
        olerepr = self.__dict__["_olerepr_"]
        if attr in olerepr.mapFuncs:
            return "method"
        if attr in self.__dict__["odd"]:
            olerepr.mapFuncs[attr] = Entry()
            return "method"
        return "value"

    def __setattr__(self, attr, value):
        self.__dict__["_olerepr_"].propMap[attr] = Entry()


class Dynamic:
    def __init__(self):
        self.built = []
        self.CDispatch = CDispatch

        def MakeOleRepr(dispatch, typeinfo, typecomp):
            self.built.append(dispatch)
            return Olerepr()

        self.MakeOleRepr = MakeOleRepr


def shared_dynamic():
    dynamic = Dynamic()
    share_type_information(dynamic, ComError, NULL_IID, FUNCDESC, INVOKE_FUNC)
    return dynamic


def test_objects_of_one_interface_share_one_olerepr():
    dynamic = shared_dynamic()
    shape1 = dynamic.MakeOleRepr("shape 1", TypeInfo("{Shape}"), "comp")
    shape2 = dynamic.MakeOleRepr("shape 2", TypeInfo("{Shape}"), "comp")
    frame = dynamic.MakeOleRepr("frame", TypeInfo("{TextFrame}"), "comp")
    assert shape1 is shape2
    assert frame is not shape1
    assert dynamic.built == ["shape 1", "frame"]


def test_an_object_without_type_information_keeps_its_own():
    dynamic = shared_dynamic()
    a = dynamic.MakeOleRepr("a", None, None)
    b = dynamic.MakeOleRepr("b", None, None)
    c = dynamic.MakeOleRepr("c", TypeInfo("{Shape}"), None)
    d = dynamic.MakeOleRepr("d", TypeInfo("{Shape}", broken=True), "comp")
    e = dynamic.MakeOleRepr("e", TypeInfo("{Shape}", broken=True), "comp")
    assert len({id(x) for x in (a, b, c, d, e)}) == 5
    assert dynamic.built == ["a", "b", "c", "d", "e"]


def test_an_interface_that_says_iid_null_is_not_shared():
    dynamic = shared_dynamic()
    a = dynamic.MakeOleRepr("a", TypeInfo(NULL_IID), "comp")
    b = dynamic.MakeOleRepr("b", TypeInfo(NULL_IID), "comp")
    assert a is not b


def test_one_objects_guess_does_not_reach_the_others():
    dynamic = shared_dynamic()
    olerepr = dynamic.MakeOleRepr("shape 1", TypeInfo("{Shape}"), "comp")
    odd = dynamic.CDispatch(olerepr, odd={"TextFrame"})
    other = dynamic.CDispatch(dynamic.MakeOleRepr("shape 2", TypeInfo("{Shape}"), "comp"))
    assert odd.TextFrame == "method"
    assert "TextFrame" not in olerepr.mapFuncs
    assert other.TextFrame == "value"


def test_a_guessed_write_is_not_shared_either():
    dynamic = shared_dynamic()
    olerepr = dynamic.MakeOleRepr("shape 1", TypeInfo("{Shape}"), "comp")
    dynamic.CDispatch(olerepr).Unknown = 1
    assert "Unknown" not in olerepr.propMap


def test_an_olerepr_of_its_own_keeps_pywin32s_guess():
    dynamic = shared_dynamic()
    olerepr = dynamic.MakeOleRepr("no type info", None, None)
    odd = dynamic.CDispatch(olerepr, odd={"TextFrame"})
    assert odd.TextFrame == "method"
    assert "TextFrame" in olerepr.mapFuncs


def test_what_the_type_library_says_is_not_a_guess():
    assert not is_guess(Entry(Desc(FUNCDESC, INVOKE_FUNC)), FUNCDESC, INVOKE_FUNC)
    # A property that takes arguments, moved to the methods to be callable.
    assert not is_guess(Entry(Desc(FUNCDESC, INVOKE_PROPERTYGET), was_property=1),
                        FUNCDESC, INVOKE_FUNC)
    assert is_guess(Entry(None), FUNCDESC, INVOKE_FUNC)
    assert is_guess(Entry(Desc(FUNCDESC, INVOKE_PROPERTYGET)), FUNCDESC, INVOKE_FUNC)


@pytest.mark.skipif(sys.platform != "win32", reason="runs pywin32 against a real COM server")
class TestPywin32ItSelf:
    """utils.com_wrapper has already shared type information in this process."""

    @staticmethod
    def dictionaries():
        import win32com.client

        first = win32com.client.Dispatch("Scripting.Dictionary")
        second = win32com.client.Dispatch("Scripting.Dictionary")
        first.Add("a", 1)
        second.Add("a", 1)
        second.Add("b", 2)
        return first, second

    def test_the_two_dictionaries_share_one_olerepr(self):
        first, second = self.dictionaries()
        assert first._olerepr_ is second._olerepr_

    def test_an_odd_answer_from_one_object_stays_with_it(self):
        import pywintypes
        import win32com.client
        import winerror

        first, second = self.dictionaries()
        real = first._oleobj_

        class AnswersOddly:
            """Answers one property read with E_INVALIDARG, which pywin32
            takes to mean the name is a method."""

            def Invoke(self, *args):
                raise pywintypes.com_error(winerror.E_INVALIDARG, "odd", None, None)

            def __getattr__(self, name):
                return getattr(real, name)

        first.__dict__["_oleobj_"] = AnswersOddly()
        # The odd object is read the way pywin32 has always read it.
        assert callable(first.Count)
        # Without the guard these were <bound method Count>, for every
        # Dictionary from then on.
        assert second.Count == 2
        assert win32com.client.Dispatch("Scripting.Dictionary").Count == 0
        assert second.Item("b") == 2

    def test_an_object_returned_by_a_call_is_wrapped_and_reads(self):
        import os

        import win32com.client

        fso = win32com.client.Dispatch("Scripting.FileSystemObject")
        temp = os.environ.get("TEMP") or os.getcwd()
        first, second = fso.GetFolder(temp), fso.GetFolder(temp)
        assert second._olerepr_ is first._olerepr_
        assert second.Path == first.Path
        assert second.Name == first.Name


class TestWrappingAKnownInterface:
    """utils.com_wrapper.wrap_known_interfaces, against stand-ins for
    win32com.client and its gencache."""

    class Wrapped:
        def __init__(self, ob, olerepr, userName=None, lazydata=None):
            self.ob, self.olerepr, self.lazydata = ob, olerepr, lazydata

    def client(self, shared, generated=()):
        from utils.com_wrapper import wrap_known_interfaces

        calls = []

        class CDispatch(self.Wrapped):
            def _wrap_dispatch_(self, ob, userName=None, returnCLSID=None):
                calls.append(ob)
                return "pywin32's own"

        client = type("Client", (), {"CDispatch": CDispatch})
        gencache = type("Gencache", (), {
            "GetClassForCLSID": staticmethod(lambda clsid: "class" if clsid in generated else None)})
        wrap_known_interfaces(client, shared, gencache, ComError, NULL_IID)
        return client, calls

    class Object:
        def __init__(self, iid, no_type_info=False):
            self.iid, self.no_type_info = iid, no_type_info
            self.asked = []

        def GetTypeInfo(self):
            self.asked.append("GetTypeInfo")
            if self.no_type_info:
                raise ComError("no type info")
            return self

        def GetTypeAttr(self):
            self.asked.append("GetTypeAttr")
            return (self.iid, 0, 0)

        def GetTypeComp(self):
            self.asked.append("GetTypeComp")
            return "typecomp"

    def test_an_interface_already_seen_is_wrapped_with_its_shared_olerepr(self):
        shared = {"{Shape}": "shape olerepr"}
        client, calls = self.client(shared)
        shape = self.Object("{Shape}")
        wrapped = client.CDispatch(None, None)._wrap_dispatch_(shape, "Shape")
        assert wrapped.olerepr == "shape olerepr"
        assert calls == []
        # Two calls into PowerPoint, not four: no ITypeComp until a name needs one.
        assert shape.asked == ["GetTypeInfo", "GetTypeAttr"]
        typeinfo, typecomp = wrapped.lazydata
        assert (typeinfo, typecomp) == (shape, "typecomp")
        assert shape.asked[-1] == "GetTypeComp"

    def test_anything_else_takes_pywin32s_own_path(self):
        shared = {"{Shape}": "olerepr", NULL_IID: "null", "{Made}": "made"}
        client, calls = self.client(shared, generated={"{Made}"})
        wrapper = client.CDispatch(None, None)
        new, null, bare = self.Object("{New}"), self.Object(NULL_IID), self.Object("", no_type_info=True)
        made = self.Object("{Made}")
        for ob in (new, null, bare, made):
            assert wrapper._wrap_dispatch_(ob) == "pywin32's own"
        assert calls == [new, null, bare, made]
