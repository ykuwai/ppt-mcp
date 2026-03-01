"""Codebase consistency tests.

These tests verify structural invariants of the codebase — tool naming
conventions, annotation completeness, and documentation accuracy —
WITHOUT importing the server (which requires a live COM connection).
"""

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC_DIR = ROOT / "src"
SERVER_PY = SRC_DIR / "server.py"
PPT_COM_DIR = SRC_DIR / "ppt_com"
README_MD = ROOT / "README.md"
SCRIPTS_DIR = ROOT / "scripts"

TOOL_NAME_RE = re.compile(r'name="(ppt_\w+)"')

REQUIRED_ANNOTATION_KEYS = {
    "title",
    "readOnlyHint",
    "destructiveHint",
    "idempotentHint",
    "openWorldHint",
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _collect_source_files() -> list[Path]:
    """Return all Python source files that may contain tool definitions."""
    files = [SERVER_PY]
    files.extend(sorted(PPT_COM_DIR.glob("*.py")))
    return files


def _extract_tool_names(path: Path) -> list[str]:
    """Extract all tool names from a single source file."""
    text = path.read_text(encoding="utf-8")
    return TOOL_NAME_RE.findall(text)


def _all_tool_names() -> list[str]:
    """Return every tool name found across all source files."""
    names: list[str] = []
    for path in _collect_source_files():
        names.extend(_extract_tool_names(path))
    return names


def _extract_annotation_blocks(path: Path) -> list[dict[str, list[str]]]:
    """Extract annotation dicts from a source file.

    Returns a list of dicts, each mapping to the set of keys found in one
    ``annotations={...}`` block.
    """
    text = path.read_text(encoding="utf-8")
    # Match multi-line annotations={...} blocks
    block_re = re.compile(
        r'annotations=\{\s*\n(.*?)\}',
        re.DOTALL,
    )
    key_re = re.compile(r'"(\w+)":\s*')
    blocks = []
    for match in block_re.finditer(text):
        keys = key_re.findall(match.group(1))
        blocks.append(keys)
    return blocks


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------


class TestToolNaming:
    """All tool names must start with 'ppt_'."""

    def test_all_tool_names_start_with_ppt(self):
        names = _all_tool_names()
        assert names, "No tool names found — regex may be broken"
        bad = [n for n in names if not n.startswith("ppt_")]
        assert bad == [], f"Tool names not starting with 'ppt_': {bad}"


class TestNoDuplicateTools:
    """Every tool name must be unique across the entire codebase."""

    def test_no_duplicate_tool_names(self):
        names = _all_tool_names()
        seen: dict[str, int] = {}
        for name in names:
            seen[name] = seen.get(name, 0) + 1
        duplicates = {n: c for n, c in seen.items() if c > 1}
        assert duplicates == {}, f"Duplicate tool names: {duplicates}"


class TestToolCountMatchesReadme:
    """The total tool count in source must match the count in README.md."""

    def test_tool_count_matches_readme(self):
        source_count = len(_all_tool_names())

        readme_text = README_MD.read_text(encoding="utf-8")
        # README contains a summary row like: | | **146** | |
        count_match = re.search(r'\|\s*\*\*(\d+)\*\*\s*\|', readme_text)
        assert count_match, "Could not find tool count in README.md"
        readme_count = int(count_match.group(1))

        assert source_count == readme_count, (
            f"Source files define {source_count} tools, "
            f"but README.md claims {readme_count}"
        )


class TestAnnotationKeys:
    """Every @mcp.tool annotation must include all 5 required keys."""

    def test_all_annotations_have_required_keys(self):
        missing_report: list[str] = []
        for path in _collect_source_files():
            blocks = _extract_annotation_blocks(path)
            for i, keys in enumerate(blocks, 1):
                key_set = set(keys)
                missing = REQUIRED_ANNOTATION_KEYS - key_set
                if missing:
                    missing_report.append(
                        f"{path.name} block #{i}: missing {missing}"
                    )
        assert missing_report == [], (
            "Annotation blocks with missing keys:\n"
            + "\n".join(missing_report)
        )


class TestScriptsExist:
    """Required scripts must exist."""

    def test_count_tools_script_exists(self):
        script = SCRIPTS_DIR / "count_tools.sh"
        assert script.exists(), f"{script} does not exist"
