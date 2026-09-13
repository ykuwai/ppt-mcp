"""Record the GVML fixtures in tests/fixtures/gvml from a live PowerPoint.

Run on a Mac with PowerPoint open and no deck that matters in front:

    PYTHONPATH=src .venv/bin/python scripts/gvml_record.py

It makes a throwaway deck inside PowerPoint's container, puts four shapes on
it, copies each with `copy shape`, and writes the package PowerPoint puts on
the pasteboard to the fixtures directory. The rectangle is one PowerPoint made
itself with `make new shape`; the group, the freeform and the chart are pasted
from packages `gvml` built, then copied back, because the dictionary has no
other way to make them. Either way the bytes are PowerPoint's own output, and
that is what the golden tests in tests/test_gvml.py compare against.

The deck is closed without saving. Nothing is left behind.
"""

import os
import sys
import time

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from appscript import app, k  # noqa: E402

from backend import pasteboard  # noqa: E402
from backend.mac_ae import count  # noqa: E402
from gvml import GVML_UTI, build, validate  # noqa: E402
from gvml import canvas, charts, freeform, shapes  # noqa: E402

FIXTURES = os.path.join(os.path.dirname(__file__), "..", "tests", "fixtures", "gvml")
E = canvas.EMU_PER_PT


def main() -> None:
    os.makedirs(FIXTURES, exist_ok=True)
    pp = app(id="com.microsoft.Powerpoint", terms="sdef")
    pres = pp.make(new=k.presentation)
    pp.make(new=k.slide, at=pres.end, with_properties={k.layout: k.slide_layout_blank})
    slide = pres.slides[1]
    window = pres.document_windows[1]

    def last_shape():
        return slide.shapes[count(slide.shapes)]

    def record(shape, name):
        shape.copy_shape()
        time.sleep(0.2)
        raw = pasteboard.read(GVML_UTI)
        assert raw is not None, f"no GVML package on the pasteboard for {name}"
        validate(raw)
        path = os.path.join(FIXTURES, f"{name}.gvml.zip")
        with open(path, "wb") as fh:
            fh.write(raw)
        print(f"wrote {path} ({len(raw)} bytes)")

    def paste(raw):
        validate(raw)
        before = count(slide.shapes)
        pasteboard.write(GVML_UTI, raw)
        window.selection.unselect()
        window.view.go_to_slide(number=1)
        window.view.paste_object()
        assert count(slide.shapes) == before + 1, "the paste was discarded"
        return last_shape()

    try:
        # 1. A rectangle PowerPoint made itself.
        pp.make(new=k.shape, at=slide.end, with_properties={
            k.auto_shape_type: k.autoshape_rectangle,
            k.left_position: 100, k.top: 100, k.width: 200, k.height: 120,
        })
        rect = last_shape()
        rect.name.set("HandRect")
        record(last_shape(), "rect")

        # 2. A group of two rectangles, pasted then copied back.
        kids = [
            shapes.shape_xml(3, "KidA", 100 * E, 100 * E, 80 * E, 60 * E),
            shapes.shape_xml(4, "KidB", 200 * E, 150 * E, 80 * E, 60 * E),
        ]
        group = shapes.group_xml(2, "HandGroup", 100 * E, 100 * E, 180 * E, 110 * E, kids)
        record(paste(build(canvas.wrap(group, 100 * E, 100 * E, 180 * E, 110 * E))), "group")

        # 3. A freeform: line, auto curve, corner curve, closed.
        nodes = [
            {"seg_int": 0, "et_int": 0, "x1": 200, "y1": 100},
            {"seg_int": 1, "et_int": 0, "x1": 250, "y1": 200},
            {"seg_int": 1, "et_int": 1, "x1": 220, "y1": 260, "x2": 160, "y2": 260, "x3": 120, "y3": 200},
        ]
        geometry, x, y, cx, cy = freeform.build_geometry(100, 150, nodes, True)
        drawing = canvas.wrap(shapes.shape_xml(2, "HandFree", x, y, cx, cy, geometry), x, y, cx, cy)
        record(paste(build(drawing)), "freeform")

        # 4. A clustered column chart with PowerPoint's default data.
        raw = charts.chart_package("HandChart", 50 * E, 50 * E, 500 * E, 350 * E, charts.chart_xml(51))
        record(paste(raw), "chart")
    finally:
        pres.close(saving=k.no)


if __name__ == "__main__":
    main()
