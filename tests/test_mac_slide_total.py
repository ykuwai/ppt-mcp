"""slide_at and the slide count on macOS ask the deck for its count in one
event rather than reading every slide to count them (about 80 ms on a deck of
50, on nearly every call); a reference that no longer reaches a deck still
gets the message that says so."""

import sys

import pytest

sys.path.insert(0, "src")

pytestmark = pytest.mark.skipif(
    sys.platform != "darwin", reason="the Apple Event backend needs appscript"
)


class _Slides:
    def __init__(self, n):
        self.n, self.read = n, 0

    def get(self):
        self.read += 1
        return [f"slide {i}" for i in range(1, self.n + 1)]

    def __getitem__(self, i):
        return f"slide {i}"


class _Deck:
    def __init__(self, n, counts=True):
        self.slides, self.counts, self.asked = _Slides(n), counts, []

    def count(self, each):
        self.asked.append(each)
        if not self.counts:
            from appscript.reference import CommandError

            raise CommandError(None, None, Exception("-1728"), None)
        return self.slides.n

    def name(self):
        return "Deck.pptx"


def test_the_deck_counts_its_own_slides():
    from appscript import k

    from backend.mac_ae import slide_at, slide_total

    deck = _Deck(50)
    assert slide_total(deck) == 50 and deck.asked == [k.slide] and deck.slides.read == 0
    assert slide_at(deck, 50) == "slide 50" and deck.slides.read == 0
    with pytest.raises(ValueError, match="out of range. The presentation has 50 slides"):
        slide_at(deck, 51)


def test_when_the_deck_will_not_count_the_slides_are_read():
    from backend.mac_ae import slide_total

    deck = _Deck(3, counts=False)
    assert slide_total(deck) == 3 and deck.slides.read == 1
