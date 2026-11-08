"""On Windows a new deck's Name follows its first slide's title once it has
one ("Presentation1" becomes "Quarter 1 results"), and its FullName keeps the
name ppt_create_presentation gave the caller. That name still finds it (#371).
"""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from ppt_com import presentation as presentation_mod, slides as slides_mod  # noqa: E402

pytestmark = pytest.mark.skipif(
    sys.platform != "win32", reason="on macOS the Apple Event impl takes the place of the COM one")


class Deck:
    def __init__(self, name, full_name):
        self.Name, self.FullName = name, full_name


class Presentations:
    def __init__(self, *decks):
        self.decks = decks
        self.Count = len(decks)

    def __call__(self, index):
        return self.decks[index - 1]


def app_with(*decks):
    return type("App", (), {"Presentations": Presentations(*decks)})()


@pytest.fixture(params=[presentation_mod, slides_mod], ids=["presentation", "slides"])
def resolve(request):
    return request.param._resolve_presentation


def test_the_name_it_was_created_with_still_finds_it(resolve):
    deck = Deck("Quarter 1 results", "Presentation1")
    assert resolve(app_with(Deck("Other", "Other"), deck), presentation_name="Presentation1") is deck


def test_its_name_now_finds_it_too(resolve):
    deck = Deck("Quarter 1 results", "Presentation1")
    assert resolve(app_with(deck), presentation_name="Quarter 1 results") is deck


def test_a_name_comes_before_a_full_name(resolve):
    named = Deck("Presentation1", r"C:\decks\Presentation1.pptx")
    assert resolve(app_with(Deck("Title", "Presentation1"), named), presentation_name="Presentation1") is named


def test_a_name_nothing_has_is_still_an_error(resolve):
    with pytest.raises(ValueError, match="No presentation named 'Nope'"):
        resolve(app_with(Deck("A", "A")), presentation_name="Nope")
