"""Run the six clipboard tools against a live PowerPoint and print what landed.

The checks that only a hand can make (docs/gvml-design.md section 7): that
PowerPoint accepts the packages, that a second paste does not land inside the
first chart, that the clipboard comes back, that the chart survives a save
and reopen, and that position and z order read back as written.

    PYTHONPATH=src .venv/bin/python scripts/gvml_smoke.py

Makes a throwaway deck in PowerPoint's container, saves it there once to
check the reopen, and deletes it at the end. Not part of the test suite.
"""

import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from appscript import k  # noqa: E402

from backend import pasteboard  # noqa: E402
from backend.mac_ae import EXPORT_STAGING_DIR, ppt  # noqa: E402
from ppt_com.charts import _add_chart_impl, _get_chart_data_impl  # noqa: E402
from ppt_com.freeform import _build_freeform_impl, _get_shape_nodes_impl  # noqa: E402
from ppt_com.groups import (  # noqa: E402
    _get_group_items_impl,
    _group_shapes_impl,
    _ungroup_shapes_impl,
)
from ppt_com.shapes import _add_shape_impl, _add_textbox_impl, _list_shapes_impl  # noqa: E402
from ppt_com.presentation import _save_presentation_as_impl  # noqa: E402
from ppt_com.slides import _add_slide_impl  # noqa: E402

DECK = os.path.join(EXPORT_STAGING_DIR, "gvml-smoke.pptx")


def show(label, result):
    if isinstance(result, str):
        result = json.loads(result)
    print(f"--- {label}")
    print(json.dumps(result, ensure_ascii=False, indent=1)[:1500])
    return result


def run(func, *args):
    return ppt.execute(func, *args)


def main() -> None:
    ppt.start()
    app = run(ppt._connect_impl)
    pres = app.make(new=k.presentation)
    run(ppt._set_target_pres_impl, pres.name())
    run(_add_slide_impl, None, None, "blank")

    # The user's clipboard, to be found intact at the end.
    pasteboard.write("public.utf8-plain-text", "the user's own text".encode())

    # Groups.
    run(_add_shape_impl, 1, 1, 50, 300, 100, 60, "A", *([None] * 16))
    run(_add_shape_impl, 1, 9, 170, 320, 100, 60, "B", *([None] * 16))
    run(_add_textbox_impl, 1, 290, 300, 120, 40, "hello group", *([None] * 7))
    names = [s["name"] for s in run(_list_shapes_impl, 1)["shapes"]]
    print("members:", names)
    group = show("group_shapes", run(_group_shapes_impl, 1, names))
    items = show("get_group_items", run(_get_group_items_impl, 1, group["group_name"]))
    xml_names = [i["name"] for i in items["items"]]
    if xml_names != names:
        # PowerPoint hands a localised default name over Apple Events
        # ("TextBox 3") and stores another in the XML ("テキスト ボックス 3").
        print("NOTE member names differ between Apple Events and the XML:", names, xml_names)
    show("list_shapes after grouping", run(_list_shapes_impl, 1))

    # Freeform.
    nodes = [
        {"seg_int": 0, "et_int": 0, "x1": 300, "y1": 60, "x2": None, "y2": None, "x3": None, "y3": None},
        {"seg_int": 1, "et_int": 0, "x1": 360, "y1": 160, "x2": None, "y2": None, "x3": None, "y3": None},
        {"seg_int": 1, "et_int": 1, "x1": 330, "y1": 230, "x2": 260, "y2": 230, "x3": 220, "y3": 160},
    ]
    free = show("build_freeform", run(_build_freeform_impl, 1, 1, 200, 100, nodes, True, "Blob"))
    show("get_shape_nodes", run(_get_shape_nodes_impl, 1, free["shape_name"], None))

    # Charts, two in a row, so the second cannot have gone inside the first.
    c1 = show("add_chart column", run(_add_chart_impl, 1, "column", 500, 50, 400, 250))
    c2 = show("add_chart pie", run(_add_chart_impl, 1, "pie", 500, 310, 300, 200))
    listing = run(_list_shapes_impl, 1)
    print("shapes now:", [(s["name"], s["type_name"], s["left"], s["top"]) for s in listing["shapes"]])
    assert c2["shape_name"] != c1["shape_name"]
    show("get_chart_data column", run(_get_chart_data_impl, 1, c1["shape_name"]))
    show("get_chart_data pie", run(_get_chart_data_impl, 1, c2["shape_name"]))
    show("add_chart unknown int", run(_add_chart_impl, 1, 9999, 10, 10, 100, 100))

    # The user's clipboard came back.
    print("clipboard now:", pasteboard.read("public.utf8-plain-text"))

    # Ungroup still works on a group made this way.
    show("ungroup", run(_ungroup_shapes_impl, 1, group["group_name"]))

    # Save, reopen, and read the chart again.
    show("save_as", run(_save_presentation_as_impl, DECK, None, None, None))
    run(ppt._get_pres_impl).close(saving=k.no)
    ppt.open_presentation(DECK)
    run(ppt._set_target_pres_impl, os.path.basename(DECK))
    show("get_chart_data after reopen", run(_get_chart_data_impl, 1, c1["shape_name"]))
    show("get_shape_nodes after reopen", run(_get_shape_nodes_impl, 1, "Blob", None))
    reopened = run(ppt._get_pres_impl)
    reopened.close(saving=k.no)
    os.remove(DECK)
    print("deck removed")


if __name__ == "__main__":
    main()
