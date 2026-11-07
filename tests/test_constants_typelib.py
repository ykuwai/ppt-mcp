"""constants.py against PowerPoint for Windows' own type library.

Some numbers in constants.py were typed in by hand and were wrong on both
platforms: the date formats from 8 on, the media types, text orientation 4 and
6, the transition speeds (fast and slow swapped) and the 16:9 and 16:10 slide
sizes, which also put SLIDE_SIZE_MAP's "a3", "16:9" and "16:10" on the wrong
sizes. The fixture holds the enumerations as PowerPoint for Windows declares
them, so every constant is checked against the source rather than a list.

PowerPoint for Mac's type library declares the same numbers for all of these;
only PpSaveAsFileType differs there, and the Mac backend pairs those by name.

Pure Python; nothing here needs PowerPoint.
"""

import json
import pathlib
import re
import sys

import pytest

sys.path.insert(0, "src")

from ppt_com import constants  # noqa: E402

ROOT = pathlib.Path(__file__).resolve().parents[1]
FIXTURE = ROOT / "tests" / "fixtures" / "typelib" / "windows_enums.json"
MAC_ENUMS = ROOT / "src" / "backend" / "mac_enums.py"


def _typelib():
    """Every member the fixture carries, as name -> value."""
    data = json.loads(FIXTURE.read_text(encoding="utf-8"))
    members = {}
    for lib, enums in data.items():
        if lib.startswith("_"):
            continue
        for enum in enums.values():
            members.update(enum)
    return members


TYPELIB = _typelib()


def _constants():
    """The module's integer constants that the type library also names."""
    return {
        name: value
        for name, value in vars(constants).items()
        if isinstance(value, int) and not isinstance(value, bool) and name in TYPELIB
    }


def test_the_fixture_covers_the_constants():
    # Guards the comparison below against passing because nothing matched.
    assert len(_constants()) > 300


def test_every_constant_has_the_type_library_value():
    wrong = {
        name: (value, TYPELIB[name])
        for name, value in _constants().items()
        if value != TYPELIB[name]
    }
    assert wrong == {}, "constant: (ours, PowerPoint's)"


@pytest.mark.parametrize(
    "name, value",
    [
        ("ppTransitionSpeedSlow", 1),
        ("ppTransitionSpeedFast", 3),
        ("ppMediaTypeOther", 1),
        ("ppMediaTypeSound", 2),
        ("ppDateTimeHmm", 10),
        ("ppDateTimehmmssAMPM", 13),
        ("msoTextOrientationVerticalFarEast", 4),
        ("ppSlideSizeA3Paper", 9),
        ("ppSlideSizeOnScreen16x9", 15),
        ("ppSlideSizeOnScreen16x10", 16),
    ],
)
def test_the_corrected_values(name, value):
    assert getattr(constants, name) == value


def test_names_powerpoint_does_not_have_are_gone():
    # Neither type library declares these; Windows has no widescreen size
    # and no plain Mddyy date format.
    assert not hasattr(constants, "ppSlideSizeWidescreen")
    assert not hasattr(constants, "ppDateTimeMddyy")


def test_slide_size_names_point_at_the_right_sizes():
    m = constants.SLIDE_SIZE_MAP
    assert m["4:3"] == TYPELIB["ppSlideSizeOnScreen"]
    assert m["letter"] == TYPELIB["ppSlideSizeLetterPaper"]
    assert m["a4"] == TYPELIB["ppSlideSizeA4Paper"]
    assert m["a3"] == TYPELIB["ppSlideSizeA3Paper"]
    assert m["16:9"] == TYPELIB["ppSlideSizeOnScreen16x9"]
    assert m["16:10"] == TYPELIB["ppSlideSizeOnScreen16x10"]
    assert m["widescreen"] == TYPELIB["ppSlideSizeOnScreen16x9"]


def test_get_slide_size_names_the_preset_from_the_size():
    # ppt_get_slide_size names the size by a reverse lookup of SLIDE_SIZE_MAP;
    # A3 used to come back as "16:9".
    reverse = {}
    for name, value in constants.SLIDE_SIZE_MAP.items():
        reverse.setdefault(value, name)
    assert reverse[9] == "a3"
    assert reverse[15] == "16:9"
    assert reverse[16] == "16:10"
    assert 8 not in reverse  # ledger paper has no preset name


def test_mac_enum_keys_are_the_windows_numbers():
    # Each entry of the macOS table is keyed by the Windows number of the
    # constant named in its comment, so a corrected constant must move its key.
    entry = re.compile(r"^\s+(-?\d+): k\.\w+,\s+# (\w+)$")
    wrong = {}
    for line in MAC_ENUMS.read_text(encoding="utf-8").splitlines():
        match = entry.match(line)
        if match and match.group(2) in TYPELIB:
            key, name = int(match.group(1)), match.group(2)
            if key != TYPELIB[name]:
                wrong[name] = (key, TYPELIB[name])
    assert wrong == {}, "constant: (key in mac_enums, PowerPoint's)"
