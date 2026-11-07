"""Media operations for PowerPoint COM automation.

Handles adding video/audio files to slides and configuring media
playback settings such as volume, looping, and fade effects.
"""

import json
import logging
import os
from typing import Optional, Union

from pydantic import BaseModel, Field, ConfigDict, model_validator

from utils.offload import run_offloaded
from backend import ppt
from utils.navigation import goto_slide
from ppt_com.constants import msoTrue, msoFalse
from ppt_com.shape_lookup import resolve_shape as _get_shape

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class AddVideoInput(BaseModel):
    """Input for adding a video to a slide."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    file_path: str = Field(..., description="Path to the video file")
    left: float = Field(default=100.0, description="Left position in points")
    top: float = Field(default=100.0, description="Top position in points")
    width: Optional[float] = Field(default=None, description="Width in points (omit to use native size)")
    height: Optional[float] = Field(default=None, description="Height in points (omit to use native size)")
    link_to_file: bool = Field(
        default=False,
        description="If true, link to file instead of embedding. Linked files are not saved in the presentation.",
    )


class AddAudioInput(BaseModel):
    """Input for adding an audio file to a slide."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    file_path: str = Field(..., description="Path to the audio file")
    left: float = Field(default=100.0, description="Left position in points")
    top: float = Field(default=100.0, description="Top position in points")
    width: Optional[float] = Field(default=None, description="Width in points (omit to use default icon size)")
    height: Optional[float] = Field(default=None, description="Height in points (omit to use default icon size)")
    link_to_file: bool = Field(
        default=False,
        description="If true, link to file instead of embedding. Linked files are not saved in the presentation.",
    )


class SetMediaSettingsInput(BaseModel):
    """Input for configuring media playback settings."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Media shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    volume: Optional[float] = Field(
        default=None, ge=0.0, le=1.0,
        description="Playback volume from 0.0 (mute) to 1.0 (full)",
    )
    muted: Optional[bool] = Field(default=None, description="Mute the media")
    start_point: Optional[int] = Field(
        default=None, ge=0,
        description="Trim start, in milliseconds from the start of the clip. Must be before the end of the trim.",
    )
    end_point: Optional[int] = Field(
        default=None, ge=0,
        description="Trim end, in milliseconds from the start of the clip. At most the clip's length.",
    )
    fade_in: Optional[int] = Field(
        default=None, ge=0,
        description="Fade-in duration in milliseconds",
    )
    fade_out: Optional[int] = Field(
        default=None, ge=0,
        description="Fade-out duration in milliseconds",
    )
    loop: Optional[bool] = Field(default=None, description="Loop playback until stopped")
    hide_while_not_playing: Optional[bool] = Field(
        default=None, description="Hide the media icon/frame while not playing",
    )

    @model_validator(mode="after")
    def _trim_in_order(self):
        if (self.start_point is not None and self.end_point is not None
                and self.start_point >= self.end_point):
            raise ValueError(
                f"start_point ({self.start_point} ms) must be before end_point "
                f"({self.end_point} ms)."
            )
        return self


# ---------------------------------------------------------------------------
# COM implementation functions (run on COM thread via ppt.execute)
# ---------------------------------------------------------------------------
def _add_video_impl(slide_index, file_path, left, top, width, height, link_to_file):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)

    abs_path = os.path.abspath(file_path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"Video file not found: {abs_path}")

    link_flag = msoTrue if link_to_file else msoFalse
    save_flag = msoFalse if link_to_file else msoTrue

    # AddMediaObject2 positional args: FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height
    if width is not None and height is not None:
        shape = slide.Shapes.AddMediaObject2(abs_path, link_flag, save_flag, left, top, width, height)
    else:
        shape = slide.Shapes.AddMediaObject2(abs_path, link_flag, save_flag, left, top)

    return {
        "success": True,
        "shape_name": shape.Name,
        "file_path": abs_path,
    }


def _add_audio_impl(slide_index, file_path, left, top, width, height, link_to_file):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)

    abs_path = os.path.abspath(file_path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"Audio file not found: {abs_path}")

    link_flag = msoTrue if link_to_file else msoFalse
    save_flag = msoFalse if link_to_file else msoTrue

    # AddMediaObject2 positional args: FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height
    if width is not None and height is not None:
        shape = slide.Shapes.AddMediaObject2(abs_path, link_flag, save_flag, left, top, width, height)
    else:
        shape = slide.Shapes.AddMediaObject2(abs_path, link_flag, save_flag, left, top)

    return {
        "success": True,
        "shape_name": shape.Name,
        "file_path": abs_path,
    }


def _set_media_settings_impl(
    slide_index, shape_name_or_index,
    volume, muted, start_point, end_point,
    fade_in, fade_out, loop, hide_while_not_playing,
):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    # MediaFormat settings
    media = shape.MediaFormat
    # The trim is checked before anything is written, so a trim PowerPoint
    # would refuse leaves the volume and the rest as they were.
    trim = None
    if start_point is not None or end_point is not None:
        trim = _trim_steps(media, start_point, end_point)
    if volume is not None:
        media.Volume = volume
    if muted is not None:
        media.Muted = muted
    if trim is not None:
        for name, value in trim:
            setattr(media, name, value)
    if fade_in is not None:
        media.FadeInDuration = fade_in
    if fade_out is not None:
        media.FadeOutDuration = fade_out

    # PlaySettings (may not be available on all media types)
    try:
        play = shape.AnimationSettings.PlaySettings
        if loop is not None:
            play.LoopUntilStopped = msoTrue if loop else msoFalse
        if hide_while_not_playing is not None:
            play.HideWhileNotPlaying = msoTrue if hide_while_not_playing else msoFalse
    except Exception:
        pass

    result = {
        "success": True,
        "shape_name": shape.Name,
    }
    if trim is not None:
        wanted = dict(trim)
        got = {"StartPoint": media.StartPoint, "EndPoint": media.EndPoint}
        if got != wanted:
            raise RuntimeError(
                f"PowerPoint kept the trim at {got['StartPoint']} to {got['EndPoint']} ms, "
                f"not {wanted['StartPoint']} to {wanted['EndPoint']} ms."
            )
        result.update(start_point=got["StartPoint"], end_point=got["EndPoint"],
                      length=media.Length)
    return result


def _trim_steps(media, start_point, end_point):
    """The trim to write, as (property, value) pairs in an order PowerPoint takes.

    PowerPoint refuses a StartPoint at or past the current EndPoint, and an
    EndPoint at or before the current StartPoint, or past the clip's length.
    So the whole trim is checked here against the clip first, and the end is
    written first when the new start is at or past the current end.
    """
    length = media.Length
    current_start, current_end = media.StartPoint, media.EndPoint
    start = current_start if start_point is None else start_point
    end = current_end if end_point is None else end_point
    if end > length:
        raise ValueError(
            f"end_point {end} ms is past the end of the clip, which is {length} ms long."
        )
    if start >= end:
        raise ValueError(
            f"The trim would start at {start} ms and end at {end} ms; the start must "
            f"be before the end (the clip is {length} ms long, currently trimmed to "
            f"{current_start} to {current_end} ms)."
        )
    if start >= current_end:
        return [("EndPoint", end), ("StartPoint", start)]
    return [("StartPoint", start), ("EndPoint", end)]


# ---------------------------------------------------------------------------
# MCP tool functions (sync wrappers that delegate to COM thread)
# ---------------------------------------------------------------------------
def add_video(params: AddVideoInput) -> str:
    """Add a video to a slide.

    Args:
        params: Video parameters including file path, position, and size.

    Returns:
        JSON with shape name and resolved file path.
    """
    try:
        result = ppt.execute(
            _add_video_impl,
            params.slide_index, params.file_path,
            params.left, params.top, params.width, params.height,
            params.link_to_file,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to add video: {str(e)}"})


def add_audio(params: AddAudioInput) -> str:
    """Add an audio file to a slide.

    Args:
        params: Audio parameters including file path, position, and size.

    Returns:
        JSON with shape name and resolved file path.
    """
    try:
        result = ppt.execute(
            _add_audio_impl,
            params.slide_index, params.file_path,
            params.left, params.top, params.width, params.height,
            params.link_to_file,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to add audio: {str(e)}"})


def set_media_settings(params: SetMediaSettingsInput) -> str:
    """Configure media playback settings.

    Args:
        params: Media shape identifier and playback settings to apply.

    Returns:
        JSON confirming the settings update.
    """
    try:
        result = ppt.execute(
            _set_media_settings_impl,
            params.slide_index, params.shape_name_or_index,
            params.volume, params.muted, params.start_point, params.end_point,
            params.fade_in, params.fade_out, params.loop, params.hide_while_not_playing,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to set media settings: {str(e)}"})


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register all media tools with the MCP server."""

    @mcp.tool(
        name="ppt_add_video",
        annotations={
            "title": "Add Video",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_add_video(params: AddVideoInput) -> str:
        """Add a video file to a slide.

        Inserts a video using AddMediaObject2. The file path is resolved to
        an absolute Windows path. Supports embedding or linking.
        All positions and sizes are in points (72 points = 1 inch).
        """
        return await run_offloaded(add_video, params)

    @mcp.tool(
        name="ppt_add_audio",
        annotations={
            "title": "Add Audio",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_add_audio(params: AddAudioInput) -> str:
        """Add an audio file to a slide.

        Inserts an audio file using AddMediaObject2. The file path is resolved
        to an absolute Windows path. Supports embedding or linking.
        All positions and sizes are in points (72 points = 1 inch).
        """
        return await run_offloaded(add_audio, params)

    @mcp.tool(
        name="ppt_set_media_settings",
        annotations={
            "title": "Set Media Settings",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_set_media_settings(params: SetMediaSettingsInput) -> str:
        """Configure playback settings for a media shape (video or audio).

        Adjust volume, mute state, fade-in/out duration, and looping.
        Identify the media shape by name or 1-based shape index.
        """
        return await run_offloaded(set_media_settings, params)


# ---------------------------------------------------------------------------
# macOS
# ---------------------------------------------------------------------------
# The implementations above walk COM. Their Apple Event counterparts have the
# same names and signatures, so on macOS they simply take their place; nothing
# else in this module changes.
from backend import IS_MACOS, use_mac_impls  # noqa: E402

if IS_MACOS:  # pragma: no cover - platform specific
    from ppt_mac import media as _mac_media

    use_mac_impls(globals(), _mac_media)
