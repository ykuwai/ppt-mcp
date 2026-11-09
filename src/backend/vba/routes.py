"""Which macOS tool implementations take the VBA route, and when.

Each entry names a ``ppt_com`` module and one of its ``_*_impl`` functions.
The value says, from the call's arguments, whether this call takes the route:
``ALWAYS`` for a tool whose Apple Event implementation refuses or goes
through the clipboard, or a test of the arguments for a tool that refuses
only some of them on the Mac. Everything else keeps its Apple Event
implementation, and so does every routed tool whenever the add-in cannot be
loaded.

Each route was checked on PowerPoint for Mac 16.114 by reading the result
back from the saved file, and for the edits, that the shape kept its id and
its animations, which the clipboard route loses.
"""


def ALWAYS(arguments) -> bool:  # noqa: N802 - reads as a constant in the table
    return True


def given(*names):
    """The call passes at least one of these arguments."""
    def condition(arguments):
        return any(arguments.get(n) is not None for n in names)
    condition.__name__ = "given(" + ", ".join(names) + ")"
    return condition


def text_target(arguments) -> bool:
    return arguments.get("target") == "text"


def reflection_details(arguments) -> bool:
    return text_target(arguments) or given("blur", "offset", "size", "transparency")(arguments)


def media_levels(arguments) -> bool:
    # Volume, mute, trimming and fades: MediaFormat, which the Mac's VBA has
    # and its dictionary does not. Looping and hiding the frame alone stay on
    # the Apple Event route, which sets them.
    return given("volume", "muted", "start_point", "end_point", "fade_in", "fade_out")(arguments)


def shrink_to_fit(arguments) -> bool:
    return arguments.get("auto_size") == "shrink_to_fit"


def no_bubble_sizes(arguments) -> bool:
    # Through Excel, which writes names and values only; the Apple Event route
    # sets bubble sizes in the chart's XML.
    return not any(isinstance(s, dict) and "sizes" in s for s in arguments.get("series") or [])


def keeps_layout(arguments) -> bool:
    """A SmartArt edit that does not pick a layout.

    Layouts are picked from ppt_list_smartart_layouts, which lists the 162
    PowerPoint for Mac installs, numbered in its own order, with no deck open.
    Its VBA knows 159 of them, in another order, and none until a deck is
    open; so adding a SmartArt and changing its layout stay where the list
    is. Colour schemes and styles come in the same order on both sides.
    """
    return arguments.get("action") != "change_layout"


def group_member(arguments) -> bool:
    """The text tool's target is inside a group.

    A ``"Group/Member"`` path is; so is a name that no top-level shape on the
    slide has, since the tools reach a group's member by its own name too. A
    top-level shape stays on the Apple Event route, which edits it in place
    already; only a member goes through the clipboard there. So does a call
    that sets a highlight: the Windows code for one needs pywin32 and
    ExecuteMso, neither of which the Mac has, and it would fail after the
    first write.
    """
    if arguments.get("highlight_color") is not None or any(
        isinstance(run, dict) and run.get("highlight_color") is not None
        for run in arguments.get("runs") or []
    ):
        return False
    target = arguments.get("shape_name_or_index")
    if not isinstance(target, str):
        return False
    if "/" in target:
        return True
    from backend import ppt

    try:
        names = ppt._get_pres_impl().slides[arguments["slide_index"]].shapes.name.get()
    except Exception:  # noqa: BLE001 - let the Apple Event route answer
        return False
    return target not in names


ROUTES = {
    # Refused on the Mac: nothing in the dictionary reaches them.
    ("ppt_com.advanced_ops", "_select_shapes_impl"): ALWAYS,
    ("ppt_com.advanced_ops", "_set_tag_impl"): ALWAYS,
    ("ppt_com.advanced_ops", "_get_tags_impl"): ALWAYS,
    ("ppt_com.advanced_ops", "_copy_animation_impl"): ALWAYS,
    ("ppt_com.animation", "_remove_animation_impl"): ALWAYS,
    ("ppt_com.tables", "_set_table_style_impl"): ALWAYS,
    # Through the clipboard on the Mac, which recreates the shape.
    ("ppt_com.charts", "_format_chart_impl"): ALWAYS,
    ("ppt_com.charts", "_format_chart_axis_impl"): ALWAYS,
    ("ppt_com.charts", "_set_chart_series_impl"): ALWAYS,
    ("ppt_com.charts", "_change_chart_type_impl"): ALWAYS,
    ("ppt_com.charts", "_get_chart_data_impl"): ALWAYS,
    ("ppt_com.charts", "_set_chart_data_impl"): no_bubble_sizes,
    ("ppt_com.freeform", "_get_shape_nodes_impl"): ALWAYS,
    ("ppt_com.freeform", "_set_node_editing_type_impl"): ALWAYS,
    ("ppt_com.freeform", "_set_node_position_impl"): ALWAYS,
    ("ppt_com.freeform", "_insert_node_impl"): ALWAYS,
    ("ppt_com.freeform", "_delete_node_impl"): ALWAYS,
    ("ppt_com.freeform", "_set_segment_type_impl"): ALWAYS,
    ("ppt_com.freeform", "_build_freeform_impl"): ALWAYS,
    ("ppt_com.groups", "_group_shapes_impl"): ALWAYS,
    ("ppt_com.groups", "_get_group_items_impl"): ALWAYS,
    ("ppt_com.slides", "_duplicate_slide_impl"): ALWAYS,
    # Modern comments, the only kind PowerPoint makes now, are not in the
    # dictionary: a slide's `comment` elements answer none of them, and
    # `make` is refused. VBA's Comments has them, with author and date.
    ("ppt_com.comments", "_add_comment_impl"): ALWAYS,
    ("ppt_com.comments", "_list_comments_impl"): ALWAYS,
    ("ppt_com.comments", "_delete_comment_impl"): ALWAYS,
    # Editing a SmartArt, but not changing its layout (see keeps_layout).
    ("ppt_com.smartart", "_modify_smartart_impl"): keeps_layout,
    ("ppt_com.slides", "_copy_slide_impl"): ALWAYS,
    # Through the clipboard on the Mac only for a group's member.
    ("ppt_com.text", "_copy_deck_impl"): ALWAYS,
    ("ppt_com.text", "_set_text_impl"): group_member,
    ("ppt_com.text", "_format_text_impl"): group_member,
    # A template's own layout by name: the Mac reads names from the saved file
    # and pastes a slide on the layout; VBA finds the layout itself.
    ("ppt_com.slides", "_add_slide_impl"): given("layout_name"),
    # Arguments the Mac refuses; the rest of these tools' calls stay as they are.
    ("ppt_com.hyperlinks", "_add_hyperlink_impl"): given("screen_tip"),
    ("ppt_com.tables", "_add_table_row_impl"): given("position"),
    ("ppt_com.tables", "_add_table_column_impl"): given("position"),
    ("ppt_com.animation", "_add_animation_impl"): given("trigger_shape"),
    ("ppt_com.effects", "_set_reflection_impl"): reflection_details,
    ("ppt_com.effects", "_set_glow_impl"): text_target,
    ("ppt_com.formatting", "_set_shadow_impl"): text_target,
    ("ppt_com.text", "_set_textframe_impl"): shrink_to_fit,
    ("ppt_com.media", "_set_media_settings_impl"): media_levels,
}

# Left on the Apple Event route, and why. MACOS_PORT.md says the same.
NOT_ROUTED = {
    "ppt_merge_shapes": "PowerPoint for Mac's VBA has no MergeShapes either",
    "ppt_execute_mso": "CommandBars.ExecuteMso and GetEnabledMso do not work in PowerPoint for Mac's VBA",
    "ppt_start_undo_entry": "PowerPoint for Mac's VBA has no StartNewUndoEntry",
    "ppt_set_window_state": "Application.WindowState changes nothing on the Mac",
    "ppt_add_chart": "PowerPoint for Mac's VBA has no AddChart2",
    "bubble sizes on ppt_set_chart_data": "through VBA the data goes through Excel, which writes names and values only",
    "ppt_list_smartart_layouts, ppt_add_smartart, changing a SmartArt's layout":
        "PowerPoint for Mac's VBA knows 159 layouts in its own order, and none with no deck open; "
        "the Mac's list holds 162, needs no deck, and numbers the layouts these take",
    "link_to_file on ppt_add_video / ppt_add_audio": "a linked file outside PowerPoint's container raises a macOS sandbox prompt",
}


def route_for(module_name, impl_name):
    return ROUTES.get((module_name, impl_name))
