"""Run the clipboard tools against a live PowerPoint and print what landed.

The checks that only a hand can make (docs/gvml-design.md section 7): that
PowerPoint accepts the packages, that a second paste does not land inside the
first chart, that the clipboard comes back, that the chart survives a save
and reopen, that position and z order read back as written, and that the
nine editors (five for charts, four for freeform nodes) leave the shape where
it was in the z order and say what they cost.

    PYTHONPATH=src .venv/bin/python scripts/gvml_smoke.py

Makes a throwaway deck in PowerPoint's container, saves it there once to
check the reopen, and deletes it at the end. Not part of the test suite.
"""

import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
# The combo chart the suite builds, so the live check and the unit tests
# paste the same one.
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "tests"))

from appscript import k  # noqa: E402

from conftest import combo_chart_xml  # noqa: E402

from backend import pasteboard  # noqa: E402
from backend.mac_ae import EXPORT_STAGING_DIR, ppt, slide_at  # noqa: E402
from backend.mac_enums import MsoShapeType  # noqa: E402
from gvml import charts as gvml_charts  # noqa: E402
from gvml.canvas import emu  # noqa: E402
from ppt_com.constants import msoChart  # noqa: E402
from ppt_mac.gvml_paste import Clipboard, paste_package  # noqa: E402
from ppt_com.animation import _add_animation_impl, _list_animations_impl  # noqa: E402
from ppt_com.charts import (  # noqa: E402
    _add_chart_impl,
    _change_chart_type_impl,
    _format_chart_axis_impl,
    _format_chart_impl,
    _get_chart_data_impl,
    _set_chart_data_impl,
    _set_chart_series_impl,
)
from ppt_com.freeform import (  # noqa: E402
    _build_freeform_impl,
    _delete_node_impl,
    _get_shape_nodes_impl,
    _insert_node_impl,
    _set_node_position_impl,
    _set_segment_type_impl,
)
from ppt_com.groups import (  # noqa: E402
    _get_group_items_impl,
    _group_shapes_impl,
    _ungroup_shapes_impl,
)
from ppt_com.shapes import _add_shape_impl, _add_textbox_impl, _list_shapes_impl  # noqa: E402
from ppt_com.presentation import _save_presentation_as_impl  # noqa: E402
from ppt_com.slides import _add_slide_impl  # noqa: E402

DECK = os.path.join(EXPORT_STAGING_DIR, "gvml-smoke.pptx")


def paste_combo(slide_index, name):
    """A chart with two plot groups and a secondary value axis.

    Nothing in the tools makes one, and PowerPoint's own combo charts were
    what stage two had none of, so it is written here and pasted.
    """
    xml = combo_chart_xml(gvml_charts.chart_xml(51).encode()).decode("utf-8")
    pres = ppt._get_pres_impl()
    slide = slide_at(pres, slide_index)
    raw = gvml_charts.chart_package(name, emu(60), emu(60), emu(480), emu(300), xml)
    clip = Clipboard.take()
    try:
        pasted = paste_package(
            pres, slide, slide_index, raw, clip, "paste_combo",
            MsoShapeType[msoChart], 60, 60, [],
        )
    finally:
        clip.restore()
    return {"shape_name": pasted.name, "warnings": pasted.warnings}


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

    # The chart editors. A shape is added after the chart so the z order has
    # somewhere to go wrong; the chart must read back at the same position.
    run(_add_shape_impl, 1, 1, 50, 500, 60, 40, "after", *([None] * 16))
    z_before = [s["name"] for s in run(_list_shapes_impl, 1)["shapes"]]
    show("set_chart_data", run(_set_chart_data_impl, 1, c1["shape_name"], ["Q1", "Q2", "Q3"], [
        {"name": "Sales", "values": [120, 180, 150]}, {"name": "Cost", "values": [80, 90, 100]},
    ]))
    show("get_chart_data after set", run(_get_chart_data_impl, 1, c1["shape_name"]))
    show("format_chart", run(_format_chart_impl, 1, c1["shape_name"], "売上の推移", True, "top",
                             None, None, None, None, None, None, None))
    show("format_chart refuses chart_style", run(_format_chart_impl, 1, c1["shape_name"], None, None, None,
                                                 5, None, None, None, None, None, None))
    show("format_chart_axis value", run(_format_chart_axis_impl, 1, c1["shape_name"], "value", "円",
                                        0, 200, 50, None, None, None, "cross", None, False, None, None, None, "#,##0"))
    show("format_chart_axis category", run(_format_chart_axis_impl, 1, c1["shape_name"], "category", "四半期",
                                           None, None, None, None, 1, 1, None, "inside", None, None, None, None, None))
    show("set_chart_series", run(_set_chart_series_impl, 1, c1["shape_name"], 2, "#FF0000", True, None))
    show("change_chart_type line", run(_change_chart_type_impl, 1, c1["shape_name"], "line_markers"))
    show("set_chart_series on a line", run(_set_chart_series_impl, 1, c1["shape_name"], 1, "#0000FF", None, 4.5))
    show("change_chart_type pie", run(_change_chart_type_impl, 1, c1["shape_name"], "pie"))
    show("change_chart_type column", run(_change_chart_type_impl, 1, c1["shape_name"], "column"))
    z_after = [s["name"] for s in run(_list_shapes_impl, 1)["shapes"]]
    print("z order kept through the chart edits:", z_before == z_after, z_after)

    # A combo chart, on its own slide. Its series live in two plot groups,
    # one of them over a secondary axis, and a change of kind has to bring
    # all of them into the one plot the new kind has.
    run(_add_slide_impl, None, None, "blank")
    combo = show("paste a combo chart", run(paste_combo, 2, "Combo"))
    show("get_chart_data on the combo", run(_get_chart_data_impl, 2, combo["shape_name"]))
    show("format_chart_axis secondary_value", run(
        _format_chart_axis_impl, 2, combo["shape_name"], "secondary_value", "右軸",
        *([None] * 13)))
    show("change_chart_type on the combo", run(_change_chart_type_impl, 2, combo["shape_name"], "line"))
    show("get_chart_data after the change", run(_get_chart_data_impl, 2, combo["shape_name"]))
    show("format_chart with no argument at all", run(
        _format_chart_impl, 2, combo["shape_name"], *([None] * 10)))

    # The node editors, with an animation on the shape so the loss is counted.
    show("add_animation on Blob", run(_add_animation_impl, 1, "Blob", "fly", "on_click", 1.0, 0.0, False,
                                      *([None] * 13)))
    show("set_node_position", run(_set_node_position_impl, 1, "Blob", None, 2, 320, 80))
    show("insert_node", run(_insert_node_impl, 1, "Blob", None, 2, 0, 0, 340, 140, None, None, None, None))
    show("set_segment_type", run(_set_segment_type_impl, 1, "Blob", None, 1, 1))
    show("delete_node", run(_delete_node_impl, 1, "Blob", None, 3))
    show("get_shape_nodes after edits", run(_get_shape_nodes_impl, 1, "Blob", None))
    show("list_animations after edits", run(_list_animations_impl, 1))
    z_after = [s["name"] for s in run(_list_shapes_impl, 1)["shapes"]]
    print("z order kept through the node edits:", z_before == z_after, z_after)

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
