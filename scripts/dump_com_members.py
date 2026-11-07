#!/usr/bin/env python3
"""Write src/backend/vba/com_members.json from PowerPoint for Windows' type libraries.

Runs on Windows, with PowerPoint for Windows installed:

    uv run python scripts/dump_com_members.py --excel-tlb "Microsoft Excel.tlb"

``--excel-tlb`` is the Excel type library PowerPoint for Mac ships, copied
from ``/Applications/Microsoft PowerPoint.app/Contents/SharedSupport/Type
Libraries/Microsoft Excel.tlb``: ppt_set_chart_data writes a chart's data
through its Excel workbook, and that copy says what PowerPoint for Mac's VBA
sees. Only the interfaces that code reaches are taken from it
(``EXCEL_INTERFACES``); Excel also has a Shape, a Chart and an Application of
its own, which must not stand in for PowerPoint's.

The macOS backend runs the Windows implementations through a VBA bridge, which
calls members by name. Whether a name is a property to read at once or a
method to hand back for calling is in the type library, which PowerPoint for
Mac does not expose to Python; this file carries it. For each interface of
the PowerPoint and Office libraries: its default member, and for each member
its invoke kinds (1 method, 2 get, 4 put, 8 putref), its parameter names and
how many of them are optional.
"""

import argparse
import json
import pathlib
import sys

if sys.platform != "win32":
    sys.exit("Reading PowerPoint's type library needs Windows; run this on Windows.")

import pythoncom  # noqa: E402
import win32com.client  # noqa: E402

ROOT = pathlib.Path(__file__).resolve().parents[1]
OUT = ROOT / "src" / "backend" / "vba" / "com_members.json"
OFFICE = ("{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", 2, 8)
# What ppt_com.charts' Excel path reaches: the chart's workbook, its sheets and
# their ranges.
EXCEL_INTERFACES = {"Workbook", "_Workbook", "Worksheet", "_Worksheet", "Sheets", "Range"}


PARAMFLAG_FOPT, PARAMFLAG_FRETVAL, PARAMFLAG_FHASDEFAULT = 0x10, 0x08, 0x20


def optional(fd) -> int:
    """How many of a member's parameters a caller may leave out.

    Each optional parameter carries PARAMFLAG_FOPT (or a default); counted
    once. cParamsOpt counts some of the same parameters again, so adding the
    two made a required Range(Cell1) look optional, and the bridge read it as
    a property with no argument.
    """
    flagged = sum(1 for a in fd.args
                  if a[1] & (PARAMFLAG_FOPT | PARAMFLAG_FHASDEFAULT) and not a[1] & PARAMFLAG_FRETVAL)
    return max(flagged, fd.cParamsOpt)


def read(tl, out, only=None):
    """Add a type library's interfaces to ``out``, or those named in ``only``."""
    for i in range(tl.GetTypeInfoCount()):
        ti = tl.GetTypeInfo(i)
        attr = ti.GetTypeAttr()
        if attr.typekind not in (pythoncom.TKIND_DISPATCH, pythoncom.TKIND_INTERFACE):
            continue
        name = tl.GetDocumentation(i)[0]
        if only is not None and name not in only:
            continue
        entry = out.setdefault(name, {"default": None, "members": {}})
        for f in range(attr.cFuncs):
            fd = ti.GetFuncDesc(f)
            names = ti.GetNames(fd.memid)
            # The dual-interface view also names a property's return value
            # (RHS); it is not an argument anyone passes.
            params = [p for p in names[1:] if p != "RHS"]
            kinds, old_params, nopt = entry["members"].get(names[0], [[], [], 0])
            entry["members"][names[0]] = [
                sorted(set(kinds) | {fd.invkind}),
                params if len(params) > len(old_params) else old_params,
                max(nopt, optional(fd)),
            ]
            if fd.memid == 0:
                entry["default"] = names[0]
        for v in range(attr.cVars):
            vname = ti.GetNames(ti.GetVarDesc(v).memid)[0]
            entry["members"].setdefault(vname, [[2, 4], [], 0])


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--excel-tlb", required=True,
                        help="PowerPoint for Mac's Microsoft Excel.tlb (see the module docstring)")
    options = parser.parse_args()
    pythoncom.CoInitialize()
    try:
        win32com.client.GetActiveObject("PowerPoint.Application")
        started = False
    except Exception:  # noqa: BLE001 - not running
        started = True
    app = win32com.client.Dispatch("PowerPoint.Application")
    libs = (app._oleobj_.GetTypeInfo().GetContainingTypeLib()[0],
            pythoncom.LoadRegTypeLib(OFFICE[0], OFFICE[1], OFFICE[2], 0))
    out = {}
    for tl in libs:
        read(tl, out)
    read(pythoncom.LoadTypeLib(str(pathlib.Path(options.excel_tlb).resolve())), out, only=EXCEL_INTERFACES)
    if started and app.Presentations.Count == 0:
        app.Quit()
    lines = ["{"]
    items = sorted(out.items())
    for n, (name, entry) in enumerate(items):
        comma = "," if n < len(items) - 1 else ""
        lines.append(f"{json.dumps(name)}: {json.dumps(entry, sort_keys=True, separators=(',', ':'))}{comma}")
    lines.append("}")
    OUT.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"wrote {OUT.relative_to(ROOT)}: {len(out)} interfaces, "
          f"{sum(len(e['members']) for e in out.values())} members")


if __name__ == "__main__":
    main()
