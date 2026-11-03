"""COM names bound once per interface: utils.com_wrapper.share_type_information.

Runs anywhere, against a stand-in for win32com.client.dynamic, since what it
changes is which olerepr pywin32 hands each new object (#264).
"""

import sys

sys.path.insert(0, "src")

from utils.com_wrapper import share_type_information  # noqa: E402


class ComError(Exception):
    pass


class TypeInfo:
    def __init__(self, iid, broken=False):
        self.iid, self.broken = iid, broken

    def GetTypeAttr(self):
        if self.broken:
            raise ComError("no type attributes")
        return (self.iid, 0, 0)


class Dynamic:
    def __init__(self):
        self.built = []

        def MakeOleRepr(dispatch, typeinfo, typecomp):
            self.built.append(dispatch)
            return object()

        self.MakeOleRepr = MakeOleRepr


def test_objects_of_one_interface_share_one_olerepr():
    dynamic = Dynamic()
    share_type_information(dynamic, ComError)
    shape1 = dynamic.MakeOleRepr("shape 1", TypeInfo("{Shape}"), "comp")
    shape2 = dynamic.MakeOleRepr("shape 2", TypeInfo("{Shape}"), "comp")
    frame = dynamic.MakeOleRepr("frame", TypeInfo("{TextFrame}"), "comp")
    assert shape1 is shape2
    assert frame is not shape1
    assert dynamic.built == ["shape 1", "frame"]


def test_an_object_without_type_information_keeps_its_own():
    dynamic = Dynamic()
    share_type_information(dynamic, ComError)
    a = dynamic.MakeOleRepr("a", None, None)
    b = dynamic.MakeOleRepr("b", None, None)
    c = dynamic.MakeOleRepr("c", TypeInfo("{Shape}"), None)
    d = dynamic.MakeOleRepr("d", TypeInfo("{Shape}", broken=True), "comp")
    e = dynamic.MakeOleRepr("e", TypeInfo("{Shape}", broken=True), "comp")
    assert len({id(x) for x in (a, b, c, d, e)}) == 5
    assert dynamic.built == ["a", "b", "c", "d", "e"]
