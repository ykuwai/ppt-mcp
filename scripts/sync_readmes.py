"""Write the tool counts into the tool list and every README, from the code.

The READMEs are kept by hand in several languages, and the numbers in them
went stale whenever a tool was added in one pull request and the translations
were not touched. This script owns those numbers, so nobody has to count:

- each row of the table in TOOLS.md, which is in English only, gets the
  number of tools in its category, and the total row gets the sum;
- every ``<!--tools-->N<!--/tools-->`` and ``<!--categories-->N<!--/categories-->``
  in the prose gets the total and the number of categories, and every
  ``<!--mac-refusals-->N<!--/mac-refusals-->`` the number of tools that always
  refuse on macOS, from MACOS_PORT's headline (which a test pins to the code).

Run it after adding, removing or moving a tool::

    uv run python scripts/sync_readmes.py          # rewrite the READMEs
    uv run python scripts/sync_readmes.py --check  # say what is stale, change nothing

``tests/test_readmes.py`` runs the check, so CI fails on a stale number and
names this script.
"""

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# Every README, English first. The others are translations of it.
READMES = {
    "en": "README.md",
    "ja": "docs/README.ja.md",
    "zh-CN": "docs/README.zh-CN.md",
    "es": "docs/README.es.md",
    "pt-BR": "docs/README.pt-BR.md",
    "he": "docs/README.he.md",
}
TOOL_LIST = "TOOLS.md"

# The tool table's rows, in order, and the module each one counts.
CATEGORIES = [
    ("App", "server"),
    ("Presentation", "presentation"),
    ("Slides", "slides"),
    ("Shapes", "shapes"),
    ("Text", "text"),
    ("Placeholders", "placeholders"),
    ("Formatting", "formatting"),
    ("Tables", "tables"),
    ("Export", "export"),
    ("Slideshow", "slideshow"),
    ("Charts", "charts"),
    ("Animation", "animation"),
    ("Themes", "themes"),
    ("Groups", "groups"),
    ("Connectors", "connectors"),
    ("Hyperlinks", "hyperlinks"),
    ("Sections", "sections"),
    ("Properties", "properties"),
    ("Media", "media"),
    ("SmartArt", "smartart"),
    ("Edit Operations", "edit_ops"),
    ("Layout", "layout"),
    ("Effects", "effects"),
    ("Comments", "comments"),
    ("Advanced", "advanced_ops"),
    ("Freeform", "freeform"),
]

# Tools listed under another row than their module's.
MOVED = {
    "ppt_get_slide_preview": "Export",
    "ppt_check_layout": "Layout",
    "ppt_batch_apply_formatting": "Advanced",
}

# The three tools catalog mode lists instead of the rest; not counted.
CATALOG = {"ppt_describe", "ppt_run", "ppt_run_steps"}

TOOL_NAME = re.compile(r'name="(ppt_\w+)"')
ROW = re.compile(r"^(\| \*\*[^|]+\*\* \| )(\d+)( \|)", re.MULTILINE)
TOTAL = re.compile(r"^(\| \| \*\*)(\d+)(\*\* \| \|)", re.MULTILINE)
MARKER = re.compile(r"<!--(tools|categories|mac-refusals)-->\d+<!--/\1-->")
HEADLINE = re.compile(r"\*\*(\d+) tools\. (\d+) do the job\. (\d+) always refuse\.\*\*")


def counts() -> list:
    """Each category's tool count, in table order."""
    module_row = {module: name for name, module in CATEGORIES}
    tally = {name: 0 for name, _ in CATEGORIES}
    files = [ROOT / "src" / "server.py", *sorted((ROOT / "src" / "ppt_com").glob("*.py"))]
    for path in files:
        for tool in TOOL_NAME.findall(path.read_text(encoding="utf-8")):
            if tool in CATALOG:
                continue
            row = MOVED.get(tool) or module_row.get(path.stem)
            if row is None:
                raise SystemExit(
                    f"{tool} in {path.name} belongs to no row of the tool table. "
                    "Add its module to CATEGORIES in scripts/sync_readmes.py, "
                    f"and a row to the table in {TOOL_LIST}."
                )
            tally[row] += 1
    return [tally[name] for name, _ in CATEGORIES]


def mac_refusals() -> int:
    """The number of tools that always refuse on macOS, as MACOS_PORT states it."""
    stated = HEADLINE.search((ROOT / "MACOS_PORT.md").read_text(encoding="utf-8"))
    if not stated:
        raise SystemExit("MACOS_PORT.md no longer states its headline count.")
    return int(stated.group(3))


def synced(body: str, numbers: list, name: str, refusals: int) -> str:
    """``body`` with its numbers brought up to date."""
    if name == TOOL_LIST:
        rows = ROW.findall(body)
        if len(rows) != len(numbers):
            raise SystemExit(
                f"{name}'s tool table has {len(rows)} rows, and scripts/sync_readmes.py "
                f"has {len(numbers)} categories. Keep them in the same order."
            )
        it = iter(numbers)
        body = ROW.sub(lambda m: f"{m.group(1)}{next(it)}{m.group(3)}", body)
        body, found = TOTAL.subn(lambda m: f"{m.group(1)}{sum(numbers)}{m.group(3)}", body)
        if found != 1:
            raise SystemExit(f"{name}'s tool table has no total row.")
    values = {"tools": sum(numbers), "categories": len(numbers), "mac-refusals": refusals}
    return MARKER.sub(lambda m: f"<!--{m.group(1)}-->{values[m.group(1)]}<!--/{m.group(1)}-->", body)


def main(argv) -> int:
    check = "--check" in argv
    numbers = counts()
    refusals = mac_refusals()
    stale = []
    for name in [*READMES.values(), TOOL_LIST]:
        path = ROOT / name
        body = path.read_text(encoding="utf-8")
        new = synced(body, numbers, name, refusals)
        if new != body:
            stale.append(name)
            if not check:
                path.write_text(new, encoding="utf-8")
    if check and stale:
        print("Stale tool counts in " + ", ".join(stale)
              + ". Run: uv run python scripts/sync_readmes.py")
        return 1
    print(f"{sum(numbers)} tools in {len(numbers)} categories"
          + ("; updated " + ", ".join(stale) if stale and not check else "; all up to date"))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
