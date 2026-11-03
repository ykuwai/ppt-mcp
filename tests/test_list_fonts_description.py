"""ppt_list_fonts says on its first line what it lists (#194, item 2).

A caller reaching for it wanted the fonts on the computer, to pick one, and
read only after the answer that it lists the deck's. The catalog shows a tool
by its first sentence, so that sentence has to say it.
"""

import sys

sys.path.insert(0, "src")


def _description():
    import server

    return server.mcp._tool_manager.get_tool("ppt_list_fonts").description


def test_the_first_line_says_the_deck_and_not_the_computer():
    first = _description().strip().splitlines()[0]
    assert "presentation uses" in first
    assert "not those installed" in first


def test_it_says_where_a_missing_font_is_reported():
    assert "font_missing" in _description()
