"""The READMEs, English and its translations, say the same things.

They are kept by hand, and a change made in one language is easily left out
of the others: the Japanese one once said 21 tools refused on the Mac when the
English one said 12. The numbers are now written by ``scripts/sync_readmes.py``
and checked here. The rest pins what can be compared without reading the
prose: the settings, the catalog tools, the language bar, the sections, and
that the Hebrew one reads right to left.
"""

import importlib.util
import re
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
_spec = importlib.util.spec_from_file_location("sync_readmes", ROOT / "scripts" / "sync_readmes.py")
sync_readmes = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(sync_readmes)
READMES = sync_readmes.READMES
TOOL_NAME_RE = re.compile(r'name="(ppt_\w+)"')
CATALOG = {"ppt_describe", "ppt_run", "ppt_run_steps"}


def text(lang):
    return (ROOT / READMES[lang]).read_text(encoding="utf-8")


def tool_count():
    files = [ROOT / "src" / "server.py", *sorted((ROOT / "src" / "ppt_com").glob("*.py"))]
    names = [n for f in files for n in TOOL_NAME_RE.findall(f.read_text(encoding="utf-8"))]
    return len([n for n in names if n not in CATALOG])


@pytest.mark.parametrize("lang", sorted(READMES))
def test_the_tool_table_adds_up_to_the_real_count(lang):
    totals = re.findall(r"\|\s*\*\*(\d+)\*\*\s*\|", text(lang))
    assert totals, f"{READMES[lang]} has no total row in its tool table"
    assert int(totals[-1]) == tool_count()


@pytest.mark.parametrize("lang", sorted(READMES))
def test_the_rows_of_the_tool_table_add_up_to_its_total(lang):
    body = text(lang)
    start = body.index("## 🛠️")
    table = body[start:body.index("\n## ", start + 1)]
    rows = re.findall(r"^\| \*\*[^|]+\*\* \| (\d+) \|", table, re.MULTILINE)
    total = int(re.findall(r"\|\s*\*\*(\d+)\*\*\s*\|", body)[-1])
    assert sum(int(r) for r in rows) == total


@pytest.mark.parametrize("lang", [lang for lang in READMES if lang != "en"])
def test_every_setting_in_the_english_readme_is_in_the_others(lang):
    english = set(re.findall(r"\bPPT_[A-Z_]+\b", text("en")))
    assert english, "the English README names no settings"
    assert english <= set(re.findall(r"\bPPT_[A-Z_]+\b", text(lang)))


@pytest.mark.parametrize("lang", sorted(READMES))
def test_each_describes_the_catalog_tools_and_how_to_list_every_tool(lang):
    body = text(lang)
    for name in CATALOG:
        assert f"`{name}`" in body, f"{READMES[lang]} does not name {name}"
    assert '"PPT_TOOL_MODE": "catalog"' in body
    assert "PPT_TOOL_MODE=catalog" in body
    assert "PPT_TOOL_MODE=full" in body
    assert "$prev.shape_name" in body


@pytest.mark.parametrize("lang", sorted(READMES))
def test_each_has_the_same_language_bar(lang):
    """Every README opens with the same list of languages, each written in its
    own script with its code, linking every other README and naming itself
    in bold."""
    body = text(lang)
    for other, name in READMES.items():
        if other == lang:
            assert "<strong>" in body and f"({other})</strong>" in body, f"{READMES[lang]} does not name itself in its language bar"
        else:
            assert f'href="{name}"' in body, f"{READMES[lang]} does not link to {name}"
            assert f"({other})</a>" in body, f"{READMES[lang]} does not give {name}'s code in its language bar"


def test_the_numbers_are_up_to_date():
    """The tool counts in every README match the code. If this fails, run
    ``uv run python scripts/sync_readmes.py``."""
    numbers = sync_readmes.counts()
    refusals = sync_readmes.mac_refusals()
    stale = [name for name in READMES.values()
             if sync_readmes.synced((ROOT / name).read_text(encoding="utf-8"), numbers, name, refusals)
             != (ROOT / name).read_text(encoding="utf-8")]
    assert not stale, f"Stale numbers in {stale}. Run: uv run python scripts/sync_readmes.py"


@pytest.mark.parametrize("lang", sorted(READMES))
def test_each_marks_its_counts(lang):
    """Each README states the tool count and the Mac's refusals through the
    markers the sync script fills, so a number cannot be typed in by hand."""
    body = text(lang)
    assert "<!--tools-->" in body and "<!--mac-refusals-->" in body


@pytest.mark.parametrize("lang", sorted(READMES))
def test_each_has_the_same_sections(lang):
    count = len(re.findall(r"^## ", text(lang), re.MULTILINE))
    assert count == len(re.findall(r"^## ", text("en"), re.MULTILINE))


@pytest.mark.parametrize("lang", sorted(READMES))
def test_the_same_pictures_and_examples_are_shown(lang):
    pictures = sorted(re.findall(r'src="(assets/[^"]+)"', text(lang)))
    assert pictures == sorted(re.findall(r'src="(assets/[^"]+)"', text("en")))


@pytest.mark.parametrize("lang", sorted(READMES))
def test_code_blocks_are_closed(lang):
    assert len(re.findall(r"^```", text(lang), re.MULTILINE)) % 2 == 0


def test_the_hebrew_readme_reads_right_to_left():
    body = text("he")
    assert body.count('<div dir="rtl">') == 1
    assert body.rstrip().endswith("</div>")
    # Markdown inside an HTML block needs a blank line on each side.
    assert '<div dir="rtl">\n\n' in body and "\n\n</div>" in body


def test_the_hebrew_readme_links_resolve_to_its_own_anchors():
    body = text("he")
    for anchor in re.findall(r"\]\(#([^)]+)\)", body):
        assert f'id="{anchor}"' in body, f"README.he.md links to #{anchor}, which it does not have"
