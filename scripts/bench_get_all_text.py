#!/usr/bin/env python3
"""Time ppt_get_all_text and ppt_check_layout on a deck built the same way
everywhere, so numbers from different machines can be compared.

    uv run --with python-pptx python scripts/bench_get_all_text.py [--slides 50] [--runs 2] [--save-copy] [--keep]

It writes the deck as a .pptx with python-pptx, which is not one of the
server's dependencies (hence `--with`), so every machine opens the same file
and nothing is built call by call. Each slide has the same parts: a title with
an accent bar, three bullets with one bold and one italic word, a 4×4 table in
set colours, a picture, a three-stage diagram grouped with its arrows, and six
filled text boxes.

It starts the server's worker in this process, opens that file, times each
tool `--runs` times through the server's own call_tool, as a client's call
goes, then closes it without saving and deletes the file. No other deck is
touched. A PowerPoint the script had to start is quit with that close, as the
server does. On the Mac the files go in PowerPoint's container, where it reads
and writes without asking for access.

`--keep` leaves the deck open afterwards, unsaved, to look at.

`--save-copy` also times another way to read the text: the deck written to a
temporary file with Presentation.SaveCopyAs, and the slides' text read from
its XML. On the Mac that goes through the VBA add-in, so it needs the bridge
on. The deck is left unsaved and keeps its name; the file is deleted after.

Paste the table it prints into the issue or PR you are measuring for.
"""

import argparse
import json
import os
import platform
import re
import struct
import sys
import tempfile
import time
import uuid
import zipfile
import zlib
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

import anyio  # noqa: E402

import server  # noqa: E402
from backend import ppt  # noqa: E402

DRAWING = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


async def _call(name, arguments):
    result = await server.mcp.call_tool(name, arguments)
    blocks = result[0] if isinstance(result, tuple) else getattr(result, "content", result)
    text = blocks[0].text
    try:
        return json.loads(text)
    except ValueError:
        return text


async def run(tool, params, deck=None):
    arguments = {"tool": tool, "params": params}
    if deck:
        arguments["presentation"] = deck
    return await _call("ppt_run", arguments)


ACCENT, DARK, LIGHT = "#C8553D", "#22313F", "#F5EFE6"
# "No Style, No Grid": the cells are coloured one by one instead.
NO_STYLE = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"


def _picture(path):
    """Write a 480×240 PNG, a gradient with a sun and hills, to ``path``.
    Made here so the script needs no network and no file of its own."""
    width, height = 480, 240
    rows = []
    for y in range(height):
        row = bytearray([0])
        for x in range(width):
            t = y / height
            r, g, b = int(40 + 120 * t), int(110 + 90 * t), int(160 + 60 * t)
            if (x - 360) ** 2 + (y - 70) ** 2 < 32 ** 2:
                r, g, b = 250, 214, 120
            hill = 170 + 30 * ((x / 70.0) % 2 - 1) ** 2
            if y > hill:
                r, g, b = 31, 111, 80
            row += bytes((r, g, b))
        rows.append(bytes(row))

    def chunk(kind, data):
        body = kind + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)

    with open(path, "wb") as out:
        out.write(b"\x89PNG\r\n\x1a\n")
        out.write(chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)))
        out.write(chunk(b"IDAT", zlib.compress(b"".join(rows), 9)))
        out.write(chunk(b"IEND", b""))


def _write_deck(path, slides, picture):
    """Write the deck as a .pptx with python-pptx: the same file on every
    machine, in well under a second, rather than building it call by call."""
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Pt

    rgb = lambda hex_: RGBColor.from_string(hex_.lstrip("#"))  # noqa: E731

    def fill(shape, colour):
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb(colour)
        shape.line.fill.background()
        shape.shadow.inherit = False

    deck = Presentation()
    deck.slide_width, deck.slide_height = Pt(960), Pt(540)
    for i in range(1, slides + 1):
        slide = deck.slides.add_slide(deck.slide_layouts[1])  # Title and Content
        title, body = slide.shapes.title, slide.placeholders[1]
        title.left, title.top, title.width, title.height = Pt(40), Pt(28), Pt(880), Pt(68)
        title.text = f"Quarter {i} results"
        title.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
        run = title.text_frame.paragraphs[0].runs[0]
        run.font.size, run.font.bold, run.font.color.rgb = Pt(36), True, rgb(DARK)

        body.left, body.top, body.width, body.height = Pt(40), Pt(110), Pt(440), Pt(170)
        lines = [[("Revenue grew in every region", {})],
                 [("Costs ", {}), ("fell", {"bold": True}), (" ", {}), ("sharply", {"italic": True}),
                  (" after the move", {})],
                 [("Headcount stayed steady", {})]]
        frame = body.text_frame
        for n, parts in enumerate(lines):
            paragraph = frame.paragraphs[0] if n == 0 else frame.add_paragraph()
            for text, style in parts:
                run = paragraph.add_run()
                run.text = text
                run.font.size, run.font.color.rgb = Pt(22), rgb(DARK)
                run.font.bold, run.font.italic = style.get("bold"), style.get("italic")

        fill(slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Pt(34), Pt(14), Pt(56)), ACCENT)

        frame = slide.shapes.add_table(4, 4, Pt(500), Pt(115), Pt(420), Pt(150))
        # "No Style, No Grid", then each cell coloured: not Office's default blue.
        frame._element.graphic.graphicData.tbl.tblPr.find(
            "{http://schemas.openxmlformats.org/drawingml/2006/main}tableStyleId").text = NO_STYLE
        data = [["Region", "Q1", "Q2", "Q3"], ["North", "120", "135", "150"],
                ["South", "98", "104", "117"], ["West", "143", "139", "161"]]
        for r, row in enumerate(data):
            for c, text in enumerate(row):
                cell = frame.table.cell(r, c)
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb((DARK, "#FFFFFF", LIGHT, "#FFFFFF")[r])
                cell.text = text
                run = cell.text_frame.paragraphs[0].runs[0]
                run.font.size, run.font.bold = Pt(16), r == 0
                run.font.color.rgb = rgb("#FFFFFF" if r == 0 else DARK)

        slide.shapes.add_picture(picture, Pt(40), Pt(300), Pt(200), Pt(100))

        group = slide.shapes.add_group_shape()
        for left, text in ((260, "Plan"), (492, "Build"), (724, "Ship")):
            stage = group.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(left), Pt(300), Pt(180), Pt(80))
            fill(stage, ACCENT)
            stage.text = text
            stage.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            run = stage.text_frame.paragraphs[0].runs[0]
            run.font.size, run.font.bold, run.font.color.rgb = Pt(16), True, rgb("#FFFFFF")
        for left in (448, 680):
            fill(group.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Pt(left), Pt(325), Pt(36), Pt(30)), DARK)

        for k in range(6):
            card = slide.shapes.add_textbox(Pt(40 + 150 * k), Pt(420), Pt(140), Pt(70))
            fill(card, LIGHT)
            card.text_frame.text = f"Note {k + 1}\nSlide {i} detail"
            for paragraph in card.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size, run.font.color.rgb = Pt(12), rgb(DARK)
    deck.save(path)


async def timed(tool, params, deck, runs):
    times, answer = [], None
    for _ in range(runs):
        start = time.perf_counter()
        answer = await run(tool, params, deck)
        took = time.perf_counter() - start
        # Each run is checked: a failed one's time is no measurement.
        if isinstance(answer, dict) and answer.get("error"):
            raise SystemExit(f"{tool} failed: {answer['error']}")
        times.append(took)
    return times, answer


def _save_copy(path):
    """Presentation.SaveCopyAs into ``path``: seconds taken, and whether the
    deck's Saved and Name were the same after."""
    def impl(pres):
        saved, name = pres.Saved, pres.Name
        start = time.perf_counter()
        pres.SaveCopyAs(path)
        return time.perf_counter() - start, pres.Saved == saved and pres.Name == name

    if sys.platform == "darwin":
        from backend import vba

        return ppt.execute(vba.run, lambda: impl(vba.current().pres))
    return ppt.execute(lambda: impl(ppt._get_pres_impl()))


def _paragraphs(path):
    with zipfile.ZipFile(path) as package:
        names = [n for n in package.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)]
        return sum(
            1 for n in names for p in ET.fromstring(package.read(n)).iter(DRAWING + "p")
            if "".join(t.text or "" for t in p.iter(DRAWING + "t")))


def _copy_folder():
    if sys.platform == "darwin":
        from backend import vba

        # PowerPoint is sandboxed; inside its container it writes without asking.
        folder = vba.INSTALL_DIR / "tmp"
        folder.mkdir(parents=True, exist_ok=True)
        return str(folder)
    return tempfile.gettempdir()


def _fmt(times):
    return " / ".join(f"{t:.2f}" for t in times)


async def main(args):
    folder = _copy_folder()
    # A name of its own each run: a deck left open by --keep under the same
    # name would be the one opened, and measured, instead.
    path = os.path.join(folder, f"ppt-mcp-bench-{uuid.uuid4().hex[:8]}.pptx")
    picture = os.path.join(folder, "ppt-mcp-bench-picture.png")
    start = time.perf_counter()
    _picture(picture)
    try:
        _write_deck(path, args.slides, picture)
    finally:
        os.remove(picture)
    print(f"Wrote {args.slides} slides in {time.perf_counter() - start:.1f} s", flush=True)
    opened = await run("ppt_open_presentation", {"file_path": path})
    if not isinstance(opened, dict) or "name" not in opened:
        os.remove(path)
        raise SystemExit(f"Could not open the deck: {opened}")
    deck = opened.get("full_name") or opened["name"]
    # Asked once PowerPoint is surely running: before, it can answer an error.
    info = await run("ppt_get_app_info", {})
    rows = []
    try:
        times, text = await timed("ppt_get_all_text", {}, deck, args.runs)
        rows.append(("ppt_get_all_text", _fmt(times), f"{len(str(text))} characters"))
        times, _ = await timed("ppt_check_layout", {}, deck, args.runs)
        rows.append(("ppt_check_layout", _fmt(times), ""))
        if args.save_copy:
            copy = os.path.join(folder, "ppt-mcp-bench-copy.pptx")
            try:
                times, same = [], True
                for _ in range(args.runs):
                    start = time.perf_counter()
                    _, unchanged = _save_copy(copy)
                    paragraphs = _paragraphs(copy)
                    times.append(time.perf_counter() - start)
                    same = same and unchanged
                rows.append(("SaveCopyAs + XML", _fmt(times),
                             f"{paragraphs} paragraphs; deck {'unchanged' if same else 'CHANGED'}"))
            except Exception as exc:  # noqa: BLE001 - reported as the row
                rows.append(("SaveCopyAs + XML", "failed", f"{type(exc).__name__}: {exc}"))
            finally:
                if os.path.exists(copy):
                    os.remove(copy)
    finally:
        if args.keep:
            print(f"Left {deck} open (--keep), from {path}.")
        else:
            # By `presentation`, which finds the deck by its full name too:
            # its Name alone can change while it is open.
            closed = await run("ppt_close_presentation", {"save_changes": False}, deck)
            if isinstance(closed, dict) and closed.get("success"):
                os.remove(path)
            else:
                print(f"Could not close {deck}, so it is still open: {closed}")

    bridge = ""
    if sys.platform == "darwin":
        from backend import vba

        bridge = f", VBA bridge {'on' if vba.ENABLED else 'off'}"
    system = f"macOS {platform.mac_ver()[0]}" if sys.platform == "darwin" else f"Windows {platform.version()}"
    print()
    print(f"{system}, PowerPoint {info.get('version') if isinstance(info, dict) else '?'}{bridge}; "
          f"{args.slides} slides, seconds per run")
    print()
    print("| What | Seconds | Notes |")
    print("|---|---|---|")
    for row in rows:
        print("| " + " | ".join(row) + " |")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("--slides", type=int, default=50)
    parser.add_argument("--runs", type=int, default=2)
    parser.add_argument("--save-copy", action="store_true", help="also time SaveCopyAs and reading its XML")
    parser.add_argument("--keep", action="store_true", help="leave the deck open afterwards, to look at")
    ppt.start()
    try:
        anyio.run(main, parser.parse_args())
    finally:
        ppt.stop()
