#!/usr/bin/env python3
"""Generate src/backend/mac_enums.py from PowerPoint's own AppleScript dictionary.

Windows COM takes numeric constants (``msoShapeRectangle`` is 1). macOS takes
named enumerators (``autoshape rectangle``). The public MCP vocabulary is built
on the Windows numbers in ``constants.py``, so the macOS backend needs a bridge.

The obvious shortcut does not work. Enumerator codes in the sdef do encode a
number, but it is macOS's own numbering, and the two platforms only sometimes
agree. ``ppSaveAsPNG`` is 18 on Windows and ``save as PNG`` is 24 on macOS,
while ``ppLayoutBlank`` is 12 on both. Assuming they match silently produces a
table that is right in most places and wrong in a few, which is the worst kind
of table to own.

So the pairing is done **by name**. Each banner section of ``constants.py`` is
matched to an sdef enumeration, the prefix every member of a group shares is
stripped from both sides, what remains is normalised, and equal names are
paired. Whatever fails to pair is reported rather than guessed at, because a
Windows constant with no macOS counterpart is a real finding, not a bug in the
generator.

    uv run python scripts/gen_mac_enums.py

Reads /Applications/Microsoft PowerPoint.app plus src/ppt_com/constants.py and
writes src/backend/mac_enums.py.
"""

import keyword as kwmod
import pathlib
import re
import subprocess
import sys
import xml.etree.ElementTree as ET

ROOT = pathlib.Path(__file__).resolve().parents[1]
APP = "/Applications/Microsoft PowerPoint.app"
CONSTANTS = ROOT / "src" / "ppt_com" / "constants.py"
OUT = ROOT / "src" / "backend" / "mac_enums.py"

# Windows spells small numbers as digits, macOS spells them as words.
NUMBER_WORDS = {
    "one": "1", "two": "2", "three": "3", "four": "4", "five": "5",
    "six": "6", "seven": "7", "eight": "8", "nine": "9", "ten": "10",
    "eleven": "11", "twelve": "12", "sixteen": "16", "twenty": "20",
    "twentyfour": "24", "thirtytwo": "32",
}


def split_words(name):
    """Break an identifier into lowercase words, however it was written."""
    name = name.replace("_", " ")
    name = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", " ", name)
    name = re.sub(r"(?<=[A-Z])(?=[A-Z][a-z])", " ", name)
    return [w.lower() for w in name.split() if w]


def normalise(words):
    """Reduce a word list to the form used for matching across platforms."""
    joined = "".join(NUMBER_WORDS.get(w, w) for w in words)
    return re.sub(r"[^a-z0-9]", "", joined)


# Enumerations the two platforms file under different names.
ENUM_ALIASES = {
    "PpParagraphAlignment": "MsoParagraphAlignment",
}

# Enumerations constants.py records as a friendly name map rather than as a
# banner section of named constants. The tool arguments for these are words the
# caller types, so the Windows side never needed the constant names, but the
# numbers behind them are the enumeration all the same and the macOS side has
# to translate them like any other.
NAME_MAPS = {
    "MsoAnimDirection": "ANIM_DIRECTION_MAP",
}

# Pairs the automatic matcher cannot reach, because the two sides genuinely use
# different naming schemes rather than different spelling. Windows numbers its
# theme colours as a suffix and macOS names them as an ordinal, and Windows
# writes star points as digits where macOS spells multi-word numbers. Each name
# here is checked against the shipped dictionary, so a PowerPoint update that
# renames one fails the generator instead of silently dropping the entry.
OVERRIDES = {
    "MsoThemeColorIndex": {
        "msoThemeColorDark1": "first dark theme color",
        "msoThemeColorLight1": "first light theme color",
        "msoThemeColorDark2": "second dark theme color",
        "msoThemeColorLight2": "second light theme color",
        "msoThemeColorAccent1": "first accent theme color",
        "msoThemeColorAccent2": "second accent theme color",
        "msoThemeColorAccent3": "third accent theme color",
        "msoThemeColorAccent4": "fourth accent theme color",
        "msoThemeColorAccent5": "fifth accent theme color",
        "msoThemeColorAccent6": "sixth accent theme color",
    },
    # ppAutoSizeTextToFitShape has no entry here on purpose. constants.py files
    # it under PpAutoSize, but the value belongs to MsoAutoSize (the comment
    # beside it says so), and macOS puts `text to fit shape` in MsoAutoSize
    # too. Forcing it into the wrong enumeration is what the override check
    # caught, so the caller reads MsoAutoSize for a text frame's auto size.
    "MsoLineDashStyle": {
        # Windows msoLineDot is a square dot; macOS spells that out, so the
        # name matcher cannot see that they are the same thing.
        "msoLineDot": "line dash style square dot",
    },
    "MsoShapeType": {
        # Windows says msoSmartArt and macOS says "shape type smartart
        # graphic", which the name matcher reads as three words against one.
        # Without this a genuine SmartArt shape reports no numeric type at all.
        "msoSmartArt": "shape type smartart graphic",
    },
    "MsoArrowheadStyle": {
        # Windows says msoArrowheadNone, macOS says "no arrowhead". Every other
        # member of this enumeration pairs on its own.
        "msoArrowheadNone": "no arrowhead",
    },
    "PpSlideShowRangeType": {
        # Windows numbers this one ppShowSlideRange and macOS spells it out as
        # "slide show range", which the name matcher cannot see through.
        "ppShowSlideRange": "slide show range",
    },
    "MsoAnimDirection": {
        # The two words neither platform spells the same way. Everything else
        # in this enumeration matches on its own.
        "none": "no direction",
        "in": "inward",
    },
    "MsoAutoShapeType": {
        "msoShape4pointStar": "autoshape four point star",
        "msoShape5pointStar": "autoshape five point star",
        "msoShape8pointStar": "autoshape eight point star",
        "msoShape16pointStar": "autoshape sixteen point star",
        "msoShape24pointStar": "autoshape twenty four point star",
        "msoShape32pointStar": "autoshape thirty two point star",
    },
}

# Words that carry no meaning inside an enumeration and that the two platforms
# sprinkle differently. Windows says msoBringToFront where macOS says "bring
# shape to front"; neither "shape" tells you anything the enumeration did not.
FILLER = {"shape", "type", "style", "cmd", "index", "mso", "pp", "xl", "e"}


def variants(name, enum_words, group_drop, strict_only=False):
    """Every spelling of a member worth trying when matching across platforms.

    The two platforms prefix and pad names differently and neither is
    consistent, so rather than guess one rule, generate the few plausible
    readings and let an exact hit on any of them decide. Deterministic, and it
    never invents a match that is not a real word-for-word agreement.
    """
    words = split_words(name)
    forms = {tuple(words), tuple(words[group_drop:])}
    if strict_only:
        return {normalise(list(f)) for f in forms if f}
    for base in list(forms):
        forms.add(tuple(w for w in base if w not in enum_words))
        forms.add(tuple(w for w in base if w not in FILLER))
        forms.add(tuple(w for w in base if w not in enum_words and w not in FILLER))
    return {normalise(list(f)) for f in forms if f}


def common_prefix(word_lists):
    """How many leading words every member of a group shares.

    Windows writes ``msoShapeRectangle`` and macOS writes ``autoshape
    rectangle``. Neither prefix carries meaning once the enumeration is known,
    and dropping both is what makes the two sides comparable. Never drops so
    much that a name disappears entirely.
    """
    if len(word_lists) < 2:
        return 0
    shortest = min(len(w) for w in word_lists)
    depth = 0
    while depth < shortest - 1:
        first = word_lists[0][depth]
        if any(w[depth] != first for w in word_lists):
            break
        depth += 1
    return depth


def keyword_for(name):
    """The attribute appscript exposes an enumerator under."""
    ident = name.replace(" ", "_")
    if kwmod.iskeyword(ident) or ident in ("None", "True", "False"):
        ident += "_"
    return ident


def is_plain_identifier(ident):
    return ident.isidentifier() and not kwmod.iskeyword(ident)


def read_windows_constants():
    """Parse constants.py into {enumeration name: {constant: value}}.

    The file is organised in banner comment sections named after the
    enumeration they hold, which is what makes this possible without pywin32.
    """
    sections = {}
    current = None
    banner = False
    for line in CONSTANTS.read_text().splitlines():
        if line.startswith("# ====="):
            # Titles sit between two rules, so only the line after the opening
            # one is a title; the closing rule just re-arms the flag.
            banner = not banner
            continue
        if banner:
            title = re.sub(r"\s*\(.*\)$", "", line.lstrip("# ").strip())
            if re.fullmatch(r"[A-Za-z][A-Za-z0-9]*", title):
                current = title
                sections.setdefault(current, {})
            else:
                current = None
            continue
        match = re.match(r"^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(-?\d+)\s*$", line)
        if match and current:
            sections[current][match.group(1)] = int(match.group(2))

    sections.update(read_name_maps())
    return sections


def read_name_maps():
    """Read the NAME_MAPS entries into the same shape as a banner section.

    Evaluated rather than parsed, because these are ordinary dictionaries and
    a regular expression over a multi line literal would be the fragile way to
    read one. Only the names in NAME_MAPS are taken, and only if they are
    dictionaries of str to int.
    """
    namespace = {}
    exec(compile(CONSTANTS.read_text(), str(CONSTANTS), "exec"), namespace)
    sections = {}
    for enum_name, map_name in NAME_MAPS.items():
        table = namespace.get(map_name)
        if not isinstance(table, dict):
            raise SystemExit(
                "{} is named in NAME_MAPS but constants.py no longer has it"
                .format(map_name)
            )
        sections[enum_name] = {
            str(name): int(value) for name, value in table.items()
        }
    return sections


def read_mac_enumerations():
    """Parse the sdef into {enumeration name: [enumerator names]}."""
    sdef = subprocess.run(["sdef", APP], capture_output=True, check=True).stdout
    root = ET.fromstring(sdef)
    return {
        e.get("name"): [x.get("name") for x in e.findall("enumerator")]
        for e in root.iter("enumeration")
    }


def pair_enumerations(windows, mac):
    """Match a Windows enumeration name to its macOS one.

    macOS prefixes PowerPoint's own enumerations with E (``PpSaveAsFileType``
    becomes ``EPPSaveAsFileType``) and leaves the shared Office ones alone.
    """
    by_key = {}
    for name in mac:
        by_key.setdefault(re.sub(r"[^a-z0-9]", "", name.lower()), name)
    pairs = []
    for win_name in windows:
        target = ENUM_ALIASES.get(win_name, win_name)
        key = re.sub(r"[^a-z0-9]", "", target.lower())
        for candidate in (key, "e" + key):
            if candidate in by_key:
                pairs.append((win_name, by_key[candidate]))
                break
    return pairs


def main():
    windows = read_windows_constants()
    mac = read_mac_enumerations()

    blocks = []
    unmatched_report = []
    total_pairs = 0

    for win_name, mac_name in sorted(pair_enumerations(windows, mac)):
        win_members = windows[win_name]
        mac_members = mac[mac_name]
        if not win_members or not mac_members:
            continue

        enum_words = set(split_words(win_name)) | set(split_words(mac_name))
        win_drop = common_prefix([split_words(n) for n in win_members])
        mac_drop = common_prefix([split_words(n) for n in mac_members])

        # Two separate indexes, not one merged one. A strict lookup has to be
        # answered only by strict readings; otherwise `line dash style dash
        # dot` reduced loosely to "dot" answers msoLineDot's exact "dot" and
        # the strict pass stops being strict.
        mac_strict, mac_loose = {}, {}
        for name in mac_members:
            for form in variants(name, enum_words, mac_drop, True):
                mac_strict.setdefault(form, name)
        for name in mac_members:
            for form in variants(name, enum_words, mac_drop, False):
                mac_loose.setdefault(form, name)

        overrides = OVERRIDES.get(win_name, {})
        for win_const, mac_const in overrides.items():
            if mac_const not in mac_members:
                raise SystemExit(
                    "override for {} points at {!r}, which PowerPoint's "
                    "dictionary no longer has. Check what it was renamed to."
                    .format(win_const, mac_const)
                )

        entries = []
        missing = []
        # Strict first, loose second. Dropping the enumeration's own words is
        # what lets msoLineDash reach `line dash style dash`, but those same
        # words are meaningful inside some members, so `line dash style dash
        # dot` also reduces to "dot" and msoLineDot would claim it before
        # msoLineDashDot ever got to ask. Letting every constant have its exact
        # reading before anyone falls back to a loose one keeps that from
        # happening, and leaves msoLineDot correctly unmatched, because macOS
        # has square dot and round dot and no plain dot.
        matched = {}
        taken = set()
        for strict_only in (True, False):
            for name, value in sorted(win_members.items(), key=lambda kv: kv[1]):
                if name in matched:
                    continue
                match = overrides.get(name)
                if match is None:
                    index = mac_strict if strict_only else mac_loose
                    for form in sorted(
                        variants(name, enum_words, win_drop, strict_only)
                    ):
                        candidate = index.get(form)
                        # One macOS enumerator cannot stand for two Windows
                        # constants.
                        if candidate is not None and candidate not in taken:
                            match = candidate
                            break
                if match is None:
                    continue
                taken.add(match)
                matched[name] = (value, name, keyword_for(match))

        entries = sorted(matched.values())
        missing = [n for n in win_members if n not in matched]

        if entries:
            total_pairs += len(entries)
            body = "\n".join(
                (f"    {v}: k.{kw},  # {n}" if is_plain_identifier(kw)
                 else f"    {v}: getattr(k, {kw!r}),  # {n}")
                for v, n, kw in entries
            )
            blocks.append(f"{win_name}: dict = {{\n{body}\n}}\n")
        if missing:
            unmatched_report.append((win_name, missing))

    unmatched_block = "\n".join(
        "    {}: {}".format(win, ", ".join(names))
        for win, names in sorted(unmatched_report)
    )

    header = '''"""Windows numeric constants mapped to macOS named enumerators.

Generated by ``scripts/gen_mac_enums.py`` from PowerPoint's own AppleScript
dictionary. Do not edit by hand; regenerate instead, which is also how this
stays correct when Microsoft ships a new PowerPoint.

The pairing is done by name rather than by number, because the two platforms
only sometimes agree on the number. ``ppSaveAsPNG`` is 18 on Windows while
``save as PNG`` is 24 on macOS, and ``ppLayoutBlank`` is 12 on both.

{total} constants across {count} enumerations, read from PowerPoint {version}.

Windows constants with no macOS counterpart, left out on purpose so that asking
for one raises rather than quietly resolving to something close:

{unmatched}
"""

from appscript import k


def to_keyword(table, value, what):
    """Translate a Windows constant, or say plainly that macOS has no word for it.

    A missing entry is not a bug in the table. Windows genuinely has constants
    PowerPoint for Mac never learned, and the useful thing is to name which one
    rather than to fall back to something close.
    """
    try:
        return table[value]
    except KeyError:
        raise ValueError(
            "PowerPoint for Mac has no %s matching the Windows constant %r"
            % (what, value)
        ) from None
'''.format(
        total=total_pairs,
        count=len(blocks),
        version=_version(),
        unmatched=unmatched_block or "    (none)",
    )

    OUT.write_text(header + "\n\n" + "\n\n".join(blocks))
    print("wrote {}".format(OUT))
    print("  {} constants paired across {} enumerations".format(total_pairs, len(blocks)))
    for win, names in sorted(unmatched_report):
        print("  unmatched in {}: {} ({}{})".format(
            win, len(names), ", ".join(names[:6]), ", ..." if len(names) > 6 else ""))
    return 0


def _version():
    try:
        return subprocess.run(
            ["osascript", "-e",
             'tell application "Microsoft PowerPoint" to return version'],
            capture_output=True, text=True, timeout=15,
        ).stdout.strip() or "unknown"
    except Exception:
        return "unknown"


if __name__ == "__main__":
    sys.exit(main())
