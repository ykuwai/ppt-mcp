"""ppt_replace_font on Windows, against a fake COM Fonts collection.

Fonts.Replace answers "Object required" (0x800A01A8) for a font the deck does
not use. The Mac walks the deck and replaces nothing, so both now say that
nothing was changed instead of one of them failing with a raw COM error.
"""

import sys
from unittest.mock import patch

import pytest

sys.path.insert(0, "src")

from ppt_com import advanced_ops  # noqa: E402

# On macOS this impl is replaced by the Apple Event one, which walks the deck;
# tests/test_mac_advanced.py covers that side.
pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="drives COM Fonts directly")


class _Font:
    def __init__(self, name):
        self.Name = name


class _Fonts:
    def __init__(self, names):
        self.names = list(names)
        self.replaced = []

    @property
    def Count(self):
        return len(self.names)

    def __call__(self, i):
        return _Font(self.names[i - 1])

    def Replace(self, original, replacement):
        if original not in self.names:
            raise Exception(-2147352567, "Exception occurred.",
                            (0, None, None, None, 0, -2146827864), None)
        self.replaced.append((original, replacement))


class _Pres:
    def __init__(self, names):
        self.Fonts = _Fonts(names)


def _run(pres, original, replacement):
    with patch.object(advanced_ops.ppt, "_get_app_impl", lambda: None), \
         patch.object(advanced_ops.ppt, "_get_pres_impl", lambda: pres):
        return advanced_ops._replace_font_impl(original, replacement)


def test_a_font_the_deck_uses_is_replaced():
    pres = _Pres(["Arial", "Aptos"])
    result = _run(pres, "Arial", "Segoe UI")
    assert result["replaced"] is True
    assert pres.Fonts.replaced == [("Arial", "Segoe UI")]


def test_a_font_the_deck_does_not_use_changes_nothing_and_says_so():
    pres = _Pres(["Arial", "Aptos"])
    result = _run(pres, "Nonexistent Font", "Arial")
    assert result["success"] is True and result["replaced"] is False
    assert "not used in this presentation" in result["note"]
    assert "['Arial', 'Aptos']" in result["note"]
    assert pres.Fonts.replaced == []


def test_the_name_is_matched_in_any_case():
    pres = _Pres(["Arial"])
    assert _run(pres, "arial", "Aptos")["replaced"] is True
    assert pres.Fonts.replaced == [("Arial", "Aptos")]
