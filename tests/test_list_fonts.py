"""ppt_list_fonts: the fonts the deck uses, or with installed the ones here (#194).

A caller reached for it to learn which fonts the machine has, so as to pick
one, and got the deck's. The description now says first what it lists, and
`installed: true` lists this computer's font families from the reader the
layout check already uses to name a font that is missing.
"""

import json
import os
import sys
import types
from unittest.mock import patch

import anyio
import pytest

sys.path.insert(0, "src")

from ppt_com import advanced_ops, layout_check  # noqa: E402


def _installed():
    return layout_check.FontNames(
        {"Yu Gothic", "游ゴシック", "Arial", "@Yu Gothic", "Arial Bold"},
        listed={"Yu Gothic", "游ゴシック", "Arial", "arial black"})


def _run(fn, *args, **kwargs):
    if fn is layout_check._installed_fonts_impl:
        return _installed()
    if fn is advanced_ops._list_fonts_impl:
        return {"success": True, "fonts_count": 1, "fonts": ["Aptos"]}
    raise AssertionError(f"unexpected call to {fn}")


class TestTheNamesKeptForTheList:
    def test_the_list_keeps_each_name_as_written_and_sorts_without_case(self):
        assert _installed().listed == ["Arial", "arial black", "Yu Gothic", "游ゴシック"]

    def test_without_a_shorter_list_every_name_is_listed(self):
        names = layout_check.FontNames({"Meiryo", "Arial", ""})
        assert names.listed == ["Arial", "Meiryo"]
        assert "meiryo" in names


class TestTheTool:
    def test_installed_lists_this_computers_fonts(self):
        with patch.object(advanced_ops.ppt, "execute", _run):
            said = json.loads(advanced_ops.list_fonts(advanced_ops.ListFontsInput(installed=True)))
        assert said == {"success": True, "installed": True, "fonts_count": 4,
                        "fonts": ["Arial", "arial black", "Yu Gothic", "游ゴシック"]}

    @pytest.mark.parametrize("params", [None, advanced_ops.ListFontsInput()])
    def test_by_default_it_lists_the_decks_fonts(self, params):
        with patch.object(advanced_ops.ppt, "execute", _run):
            said = json.loads(advanced_ops.list_fonts(params))
        assert said["fonts"] == ["Aptos"]

    def test_a_reader_that_fails_says_so(self):
        def fail(*_):
            raise RuntimeError("no font manager")

        with patch.object(advanced_ops.ppt, "execute", fail):
            said = json.loads(advanced_ops.list_fonts(advanced_ops.ListFontsInput(installed=True)))
        assert said == {"error": "Failed to list fonts: no font manager"}

    def test_the_description_says_first_that_it_lists_the_decks_fonts(self):
        import server

        tool = server.mcp._registered_tool("ppt_list_fonts")
        assert server._summary(tool.description).startswith(
            "List the fonts this presentation uses")


class TestThroughTheServer:
    """Still callable with no params at all, as before, and with the option."""

    @pytest.mark.parametrize("arguments, fonts", [
        ({}, ["Aptos"]),
        ({"params": {}}, ["Aptos"]),
        ({"params": None}, ["Aptos"]),
        ({"params": {"installed": True}}, ["Arial", "arial black", "Yu Gothic", "游ゴシック"]),
    ])
    def test_a_call(self, arguments, fonts):
        import server

        with patch.object(advanced_ops.ppt, "execute", _run):
            result = anyio.run(server.mcp.call_tool, "ppt_list_fonts", arguments)
        [block] = server._content_of(result)
        assert json.loads(block.text)["fonts"] == fonts

    def test_an_unknown_option_is_refused(self):
        import server

        with pytest.raises(server.ToolError, match="installed"):
            anyio.run(server.mcp.call_tool, "ppt_list_fonts",
                      {"params": {"instaled": True}})


@pytest.mark.skipif(sys.platform != "darwin", reason="the macOS reader")
class TestTheMacReader:
    def test_the_list_holds_families_and_the_check_holds_their_styles_too(self, monkeypatch):
        from ppt_mac import layout_check as mac_check
        from utils import fonts

        class Manager:
            def availableFontFamilies(self):  # noqa: N802 - AppKit's name
                return ["Helvetica Neue", "Hiragino Sans"]

            def availableMembersOfFontFamily_(self, family):  # noqa: N802
                return [["", "Bold"], ["", "Light"]]

        appkit = types.ModuleType("AppKit")
        appkit.NSFontManager = types.SimpleNamespace(sharedFontManager=Manager)
        monkeypatch.setitem(sys.modules, "AppKit", appkit)
        monkeypatch.setattr(fonts, "names_in_folder",
                            lambda folder: {"Calibri"} if folder.endswith("DFonts") else set())
        monkeypatch.setattr(mac_check, "_FONTS", {})
        # Nor the Office cloud fonts this machine has downloaded.
        isdir = os.path.isdir
        monkeypatch.setattr(os.path, "isdir",
                            lambda path: "CloudFonts" not in path and isdir(path))

        names = mac_check._installed_fonts_impl()

        assert names.listed == ["Calibri", "Helvetica Neue", "Hiragino Sans"]
        assert "helvetica neue bold" in names
