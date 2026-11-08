#!/usr/bin/env python3
"""Time ppt_get_all_text and ppt_check_layout on a deck built the same way
everywhere, so numbers from different machines can be compared.

    uv run python scripts/bench_get_all_text.py [--slides 50] [--runs 2] [--save-copy]

It starts the server's worker in this process, creates a new deck, and builds
each slide from the same parts: a title with an accent bar, three bullets with
one bold and one italic word, a 4×4 table, a picture, a three-stage diagram
grouped with its arrows, and six filled text boxes. Then
it times each tool `--runs` times through the server's own call_tool, as a
client's call goes, and closes the deck it made without saving. No other deck
is touched. A PowerPoint the script had to start is quit with that close, as
the server does.

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


def _slide_steps(i, picture):
    """One slide's parts. Step numbers matter: later steps name the shapes
    earlier ones made ("$2.placeholder_name")."""
    def stage(left, text):
        return {"tool": "ppt_add_shape", "params": {
            "slide_index": i, "shape_type": "rounded_rectangle", "left": left, "top": 300, "width": 180,
            "height": 80, "text": text, "font_size": 16, "bold": True, "font_color": "#FFFFFF",
            "fill_color": ACCENT, "line_visible": False}}

    def arrow(left):
        return {"tool": "ppt_add_shape", "params": {
            "slide_index": i, "shape_type": "right_arrow", "left": left, "top": 325, "width": 36,
            "height": 30, "fill_color": DARK, "line_visible": False}}

    def card(k):
        return {"tool": "ppt_add_textbox", "params": {
            "slide_index": i, "left": 40 + 150 * k, "top": 420, "width": 140, "height": 70,
            "text": f"Note {k + 1}\nSlide {i} detail", "font_size": 12, "font_color": DARK}}

    return [
        {"tool": "ppt_set_placeholder_text", "params": {                                      # 1
            "slide_index": i, "placeholder_type": "title", "text": f"Quarter {i} results"}},
        {"tool": "ppt_set_placeholder_text", "params": {                                      # 2
            "slide_index": i, "placeholder_type": "body",
            "text": "Revenue grew in every region\nCosts fell sharply after the move\nHeadcount stayed steady"}},
        {"tool": "ppt_update_shape", "params": {                                              # 3
            "slide_index": i, "shape_name": "$2.placeholder_name", "left": 40, "top": 110, "width": 440, "height": 170}},
        {"tool": "ppt_format_text_range", "params": {                                         # 4
            "slide_index": i, "shape_name_or_index": "$2.placeholder_name", "search_text": "fell", "bold": True}},
        {"tool": "ppt_format_text_range", "params": {                                         # 5
            "slide_index": i, "shape_name_or_index": "$2.placeholder_name", "search_text": "sharply", "italic": True}},
        {"tool": "ppt_add_shape", "params": {                                                 # 6
            "slide_index": i, "shape_type": "rectangle", "left": 0, "top": 34, "width": 14, "height": 56,
            "fill_color": ACCENT, "line_visible": False}},
        {"tool": "ppt_add_table", "params": {                                                 # 7
            "slide_index": i, "rows": 4, "cols": 4, "left": 500, "top": 115, "width": 420, "height": 150}},
        {"tool": "ppt_set_table_data", "params": {                                            # 8
            "slide_index": i, "shape_name_or_index": "$7.shape_name",
            "data": [["Region", "Q1", "Q2", "Q3"], ["North", "120", "135", "150"],
                     ["South", "98", "104", "117"], ["West", "143", "139", "161"]]}},
        {"tool": "ppt_add_picture", "params": {                                               # 9
            "slide_index": i, "file_path": picture, "left": 40, "top": 300, "width": 200, "height": 100}},
        stage(260, "Plan"), arrow(448), stage(492, "Build"), arrow(680), stage(724, "Ship"),  # 10 to 14
        {"tool": "ppt_group_shapes", "params": {                                              # 15
            "slide_index": i, "shape_names": [f"${n}.shape_name" for n in range(10, 15)]}},
        *[card(k) for k in range(6)],                                                         # 16 to 21
        *[{"tool": "ppt_set_fill", "params": {                                                # 22 to 27
            "slide_index": i, "shape_name_or_index": f"${16 + k}.shape_name", "color": LIGHT}} for k in range(6)],
        {"tool": "ppt_format_text", "params": {                                               # 28
            "slide_index": i, "shape_name_or_index": "$2.placeholder_name", "font_size": 22, "color": DARK}},
        {"tool": "ppt_update_shape", "params": {                                              # 29
            "slide_index": i, "shape_name": "$1.placeholder_name", "left": 40, "top": 28, "width": 880, "height": 68}},
        {"tool": "ppt_format_text", "params": {                                               # 30
            "slide_index": i, "shape_name_or_index": "$1.placeholder_name", "font_size": 36, "bold": True,
            "color": DARK}},
        # The table in the deck's own colours, not Office's default style.
        {"tool": "ppt_set_table_style", "params": {                                           # 31
            "slide_index": i, "shape_name_or_index": "$7.shape_name", "style_id": NO_STYLE}},
        *[{"tool": "ppt_set_table_cell", "params": {
            "slide_index": i, "shape_name_or_index": "$7.shape_name", "row": row, "col": col,
            "fill_color": fill, "color": "#FFFFFF" if row == 1 else DARK, "bold": row == 1}}
          for row, fill in ((1, DARK), (2, "#FFFFFF"), (3, LIGHT), (4, "#FFFFFF")) for col in range(1, 5)],
        {"tool": "ppt_set_table_borders", "params": {
            "slide_index": i, "shape_name_or_index": "$7.shape_name", "sides": ["top", "left", "bottom", "right"], "visible": False}},
    ]


async def create():
    made = await run("ppt_create_presentation", {})
    if not isinstance(made, dict) or "name" not in made:
        raise SystemExit(f"Could not create a deck: {made}")
    return made["name"]


async def build(deck, slides, picture):
    await run("ppt_add_slide", {"layout": 2, "count": slides}, deck)
    for i in range(1, slides + 1):
        report = await _call("ppt_run_steps", {"steps": _slide_steps(i, picture), "presentation": deck})
        failed = [s for s in report.get("steps", []) if s.get("error")] if isinstance(report, dict) else [report]
        if failed:
            raise SystemExit(f"Building slide {i} failed: {failed[0]}")


async def timed(tool, params, deck, runs):
    times, answer = [], None
    for _ in range(runs):
        start = time.perf_counter()
        answer = await run(tool, params, deck)
        times.append(time.perf_counter() - start)
    if isinstance(answer, dict) and answer.get("error"):
        raise SystemExit(f"{tool} failed: {answer['error']}")
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
    info = await run("ppt_get_app_info", {})
    print("Building the deck...", flush=True)
    start = time.perf_counter()
    deck = await create()
    rows = []
    try:
        picture = os.path.join(_copy_folder(), "ppt-mcp-bench-picture.png")
        _picture(picture)
        try:
            await build(deck, args.slides, picture)
        finally:
            os.remove(picture)
        print(f"Built {args.slides} slides in {time.perf_counter() - start:.0f} s", flush=True)
        times, text = await timed("ppt_get_all_text", {}, deck, args.runs)
        rows.append(("ppt_get_all_text", _fmt(times), f"{len(str(text))} characters"))
        times, _ = await timed("ppt_check_layout", {}, deck, args.runs)
        rows.append(("ppt_check_layout", _fmt(times), ""))
        if args.save_copy:
            path = os.path.join(_copy_folder(), "ppt-mcp-bench-copy.pptx")
            try:
                times, same = [], True
                for _ in range(args.runs):
                    start = time.perf_counter()
                    took, unchanged = _save_copy(path)
                    paragraphs = _paragraphs(path)
                    times.append(time.perf_counter() - start)
                    same = same and unchanged
                rows.append(("SaveCopyAs + XML", _fmt(times),
                             f"{paragraphs} paragraphs; deck {'unchanged' if same else 'CHANGED'}"))
            except Exception as exc:  # noqa: BLE001 - reported as the row
                rows.append(("SaveCopyAs + XML", "failed", f"{type(exc).__name__}: {exc}"))
            finally:
                if os.path.exists(path):
                    os.remove(path)
    finally:
        if args.keep:
            print(f"Left {deck} open (--keep).")
        else:
            await run("ppt_close_presentation", {"save_changes": False, "presentation_name": deck}, deck)

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
