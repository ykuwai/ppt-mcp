"""ppt_mac.layout_check._package_facts_impl: a deck with unsaved changes is
judged from a copy the VBA add-in saves aside, and as before without one."""

import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

pytestmark = pytest.mark.skipif(sys.platform != "darwin", reason="the Apple Event impl")


@pytest.fixture
def unsaved(monkeypatch):
    from ppt_mac import layout_check as mac

    deck = SimpleNamespace(full_name=lambda: "Presentation1", saved=lambda: False)
    monkeypatch.setattr(mac.ppt, "_get_pres_impl", lambda: deck)
    monkeypatch.setattr(mac, "_charts_by_clipboard", lambda pres, indices: pytest.fail("no clipboard"))
    return mac


def test_an_unsaved_deck_is_read_from_a_copy_and_the_copy_removed(unsaved, monkeypatch, tmp_path):
    from ppt_com import text as com_text
    from test_layout_check import deck_file, linked_box

    folder = tmp_path / "copy"
    folder.mkdir()
    path = deck_file(folder, [linked_box("Body", [[("click here", "rId2")]])], {})
    monkeypatch.setattr(com_text, "_copy_deck_impl", lambda: {"path": str(path), "folder": str(folder), "slides": 1})
    facts = unsaved._package_facts_impl([1], charts=True)
    assert facts["links"][1] == {"Body": ["click here"]}
    assert not folder.exists()


def test_without_the_add_in_it_is_judged_as_before(unsaved, monkeypatch):
    from ppt_com import text as com_text

    monkeypatch.setattr(com_text, "_copy_deck_impl", lambda: None)
    assert unsaved._package_facts_impl([1], charts=False) == {}


def test_a_copy_that_fails_is_judged_as_before(unsaved, monkeypatch):
    from ppt_com import text as com_text

    def refuse():
        raise RuntimeError("the add-in is busy")
    monkeypatch.setattr(com_text, "_copy_deck_impl", refuse)
    assert unsaved._package_facts_impl([1], charts=False) == {}
