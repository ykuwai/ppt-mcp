#!/usr/bin/env python3
"""Build src/backend/vba/ppt-mcp-bridge.ppam, the add-in the macOS backend loads.

Runs on Windows, with PowerPoint for Windows installed:

    uv run python scripts/build_mac_vba_addin.py

PowerPoint for Mac loads an add-in without asking only when it carries
compiled code, and only PowerPoint for Windows can compile one from a script.
So the .bas sources next to the add-in are imported into a new presentation
and saved as an add-in, which compiles them; a source that does not compile
fails the save.

Importing modules needs "Trust access to the VBA project object model"
(AccessVBOM). It is turned on for the build only and put back exactly as it
was. PowerPoint writes its security settings back when it quits, so the
setting is restored only after the PowerPoint this script started has gone.
"""

import os
import pathlib
import subprocess
import sys
import tempfile
import time

if sys.platform != "win32":
    sys.exit("Building the add-in needs PowerPoint for Windows; run this on Windows.")

import pythoncom  # noqa: E402
import win32com.client  # noqa: E402
import winreg  # noqa: E402

ROOT = pathlib.Path(__file__).resolve().parents[1]
VBA = ROOT / "src" / "backend" / "vba"
MODULES = ("PptMcpHandles", "PptMcpBridge")
OUT = VBA / "ppt-mcp-bridge.ppam"
KEY = r"Software\Microsoft\Office\16.0\PowerPoint\Security"
ppSaveAsOpenXMLAddin = 30


def read_setting():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, KEY) as k:
            return winreg.QueryValueEx(k, "AccessVBOM")[0]
    except FileNotFoundError:
        return None


def write_setting(value):
    if value is None:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, KEY, 0, winreg.KEY_SET_VALUE) as k:
                winreg.DeleteValue(k, "AccessVBOM")
        except FileNotFoundError:
            pass
        return
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, KEY) as k:
        winreg.SetValueEx(k, "AccessVBOM", 0, winreg.REG_DWORD, value)


def powerpoint_running():
    return b"POWERPNT" in subprocess.run(["tasklist"], capture_output=True).stdout


def build(work: pathlib.Path) -> bytes:
    pythoncom.CoInitialize()
    if powerpoint_running():
        sys.exit("Close PowerPoint first: the build starts its own and quits it.")
    app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        pres = app.Presentations.Add(0)
        project = pres.VBProject
        for name in MODULES:
            project.VBComponents.Import(str(VBA / f"{name}.bas"))
        target = work / OUT.name
        pres.SaveAs(str(target), ppSaveAsOpenXMLAddin)  # compiles the project
        pres.Close()
        return target.read_bytes()
    finally:
        try:
            while app.Presentations.Count:
                p = app.Presentations(1)
                p.Saved = True
                p.Close()
            app.Quit()
        except Exception:  # noqa: BLE001 - quitting is best effort
            pass
        app = None
        for _ in range(60):
            if not powerpoint_running():
                break
            time.sleep(1)


def main():
    before = read_setting()
    write_setting(1)
    try:
        with tempfile.TemporaryDirectory() as tmp:
            data = build(pathlib.Path(tmp))
    finally:
        write_setting(before)
        time.sleep(5)
        if read_setting() != before:  # PowerPoint wrote it back on its way out
            write_setting(before)
    OUT.write_bytes(data)
    print(f"wrote {OUT.relative_to(ROOT)} ({len(data)} bytes); AccessVBOM back to {before!r}")


if __name__ == "__main__":
    main()
