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
    # Volume, mute and fades. Trimming stays where it is: the Windows code for
    # it needs a fix of its own first (StartPoint and EndPoint).
    return (given("volume", "muted", "fade_in", "fade_out")(arguments)
            and arguments.get("start_point") is None and arguments.get("end_point") is None)


def shrink_to_fit(arguments) -> bool:
    return arguments.get("auto_size") == "shrink_to_fit"


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
    ("ppt_com.freeform", "_get_shape_nodes_impl"): ALWAYS,
    ("ppt_com.freeform", "_set_node_position_impl"): ALWAYS,
    ("ppt_com.freeform", "_insert_node_impl"): ALWAYS,
    ("ppt_com.freeform", "_delete_node_impl"): ALWAYS,
    ("ppt_com.freeform", "_set_segment_type_impl"): ALWAYS,
    ("ppt_com.freeform", "_build_freeform_impl"): ALWAYS,
    ("ppt_com.groups", "_group_shapes_impl"): ALWAYS,
    ("ppt_com.groups", "_get_group_items_impl"): ALWAYS,
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
    "ppt_set_node_editing_type": "through PowerPoint for Mac's VBA the same call made a straight path curved once and a curved one straight another time",
    "ppt_set_chart_data": "through VBA it needs Excel's own object model and a refresh the Windows code does not do",
    "ppt_add_smartart / ppt_modify_smartart": "PowerPoint for Mac's VBA answers 'Permission denied' from SmartArt now and then",
    "link_to_file on ppt_add_video / ppt_add_audio": "a linked file outside PowerPoint's container raises a macOS sandbox prompt",
}


def route_for(module_name, impl_name):
    return ROUTES.get((module_name, impl_name))
