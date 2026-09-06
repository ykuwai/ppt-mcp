"""Tests for the section tools (ppt_add_section / ppt_list_sections).

Pure Python tests -- a fake COM SectionProperties object mimics PowerPoint's
behaviour so the slide-index handling can be verified without PowerPoint.

Background: SectionProperties.AddSection takes a SECTION index as its first
argument (valid 1..Count+1). Passing the slide index there only works for the
very first section; every later one fails with "Integer out of range". The
implementation must use AddBeforeSlide, which takes the slide index.
"""

import sys
from unittest.mock import patch

import pytest

sys.path.insert(0, "src")

from ppt_com import sections  # noqa: E402


# --- Fake COM object graph -------------------------------------------------

class _FakeSlides:
    def __init__(self, count):
        self.Count = count


class _FakeSectionProperties:
    """Mimics PowerPoint's SectionProperties for a fixed number of slides.

    Sections are stored as [name, first_slide]; an empty section has
    first_slide -1, exactly like PowerPoint reports it.
    """

    def __init__(self, slide_count, sections=None):
        self._slide_count = slide_count
        # list of [name, first_slide]
        self._sections = [list(s) for s in (sections or [])]
        self.add_section_calls = []
        self.add_before_slide_calls = []

    # -- read side -----------------------------------------------------------
    @property
    def Count(self):
        return len(self._sections)

    def Name(self, i):
        return self._sections[i - 1][0]

    def FirstSlide(self, i):
        return self._sections[i - 1][1]

    def SlidesCount(self, i):
        first = self._sections[i - 1][1]
        if first < 1:
            return 0
        later = [s[1] for s in self._sections[i:] if s[1] >= 1]
        end = min(later) - 1 if later else self._slide_count
        return end - first + 1

    # -- write side ----------------------------------------------------------
    def AddSection(self, section_index, name):
        self.add_section_calls.append((section_index, name))
        if not 1 <= section_index <= self.Count + 1:
            raise Exception(
                "SectionProperties.AddSection : Integer out of range. "
                f"{section_index} is not in sectionIndex's valid range of "
                f"1 to {self.Count + 1}."
            )
        self._sections.insert(section_index - 1, [name, -1])
        return section_index

    def AddBeforeSlide(self, slide_index, name):
        self.add_before_slide_calls.append((slide_index, name))
        if not 1 <= slide_index <= self._slide_count:
            raise Exception(
                "SectionProperties.AddBeforeSlide : Integer out of range. "
                f"{slide_index} is not in SlideIndex's valid range of "
                f"1 to {self._slide_count}."
            )
        # A section already starting here is kept (now empty) at its index;
        # the new section goes right after it.
        for pos, sec in enumerate(self._sections):
            if sec[1] == slide_index:
                sec[1] = -1
                self._sections.insert(pos + 1, [name, slide_index])
                return pos + 2
        # Otherwise insert after the last section that starts before this
        # slide (that section is split).
        pos = 0
        for idx, sec in enumerate(self._sections):
            if 1 <= sec[1] < slide_index:
                pos = idx + 1
        self._sections.insert(pos, [name, slide_index])
        return pos + 1


class _FakePres:
    def __init__(self, slide_count, sections=None):
        self.Slides = _FakeSlides(slide_count)
        self.SectionProperties = _FakeSectionProperties(slide_count, sections)


def _with_pres(pres):
    """Route the module's COM accessors to the fake presentation."""
    return (
        patch.object(sections.ppt, "_get_app_impl", return_value=object()),
        patch.object(sections.ppt, "_get_pres_impl", return_value=pres),
    )


def _ranges(listing):
    return [(s["first_slide"], s["last_slide"]) for s in listing["sections"]]


# --- _add_section_impl -----------------------------------------------------

def test_add_three_sections_uses_add_before_slide():
    pres = _FakePres(slide_count=6)
    sp = pres.SectionProperties
    p_app, p_pres = _with_pres(pres)
    with p_app, p_pres:
        r1 = sections._add_section_impl("Intro", 1)
        r2 = sections._add_section_impl("Body", 3)
        r3 = sections._add_section_impl("Close", 5)
        listing = sections._list_sections_impl()

    assert (r1["section_index"], r2["section_index"], r3["section_index"]) == (1, 2, 3)
    assert all(r["success"] for r in (r1, r2, r3))
    assert "warning" not in r2 and "warning" not in r3
    assert sp.add_before_slide_calls == [(1, "Intro"), (3, "Body"), (5, "Close")]
    assert sp.add_section_calls == []
    assert _ranges(listing) == [(1, 2), (3, 4), (5, 6)]
    assert r3["slides_count"] == 2
    assert r3["sections_count"] == 3


def test_add_section_slide_index_out_of_range_is_clear_error():
    pres = _FakePres(slide_count=6, sections=[("Intro", 1)])
    p_app, p_pres = _with_pres(pres)
    with p_app, p_pres, pytest.raises(ValueError) as exc:
        sections._add_section_impl("Too far", 7)
    msg = str(exc.value)
    assert "out of range" in msg
    assert "6 slide" in msg
    assert pres.SectionProperties.add_before_slide_calls == []


def test_add_section_without_slides_is_clear_error():
    pres = _FakePres(slide_count=0)
    p_app, p_pres = _with_pres(pres)
    with p_app, p_pres, pytest.raises(ValueError) as exc:
        sections._add_section_impl("Intro", 1)
    assert "no slides" in str(exc.value)


def test_add_section_at_existing_start_warns_about_emptied_section():
    pres = _FakePres(slide_count=6, sections=[("Intro", 1), ("Body", 3), ("Close", 5)])
    p_app, p_pres = _with_pres(pres)
    with p_app, p_pres:
        result = sections._add_section_impl("Twice", 3)
        listing = sections._list_sections_impl()

    assert result["success"] is True
    assert result["section_index"] == 3
    assert "Body" in result["warning"]
    assert "index 2" in result["warning"]
    assert [s["name"] for s in listing["sections"]] == ["Intro", "Body", "Twice", "Close"]
    assert listing["sections"][1]["empty"] is True


# --- _list_sections_impl ---------------------------------------------------

def test_list_sections_reports_empty_section_and_last_slide():
    pres = _FakePres(slide_count=6, sections=[("Intro", 1), ("Gone", -1), ("Rest", 3)])
    p_app, p_pres = _with_pres(pres)
    with p_app, p_pres:
        listing = sections._list_sections_impl()

    assert listing["success"] is True
    assert listing["sections_count"] == 3
    assert listing["slides_total"] == 6

    intro, gone, rest = listing["sections"]
    assert (intro["first_slide"], intro["last_slide"], intro["slides_count"]) == (1, 2, 2)
    assert intro["empty"] is False

    assert gone["first_slide"] is None
    assert gone["last_slide"] is None
    assert gone["slides_count"] == 0
    assert gone["empty"] is True

    assert (rest["first_slide"], rest["last_slide"]) == (3, 6)
    assert rest["last_slide"] == rest["first_slide"] + rest["slides_count"] - 1


def test_list_sections_empty_presentation():
    pres = _FakePres(slide_count=0)
    p_app, p_pres = _with_pres(pres)
    with p_app, p_pres:
        listing = sections._list_sections_impl()
    assert listing == {
        "success": True,
        "sections_count": 0,
        "slides_total": 0,
        "sections": [],
    }
