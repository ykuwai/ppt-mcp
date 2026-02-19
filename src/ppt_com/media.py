"""Media operations for PowerPoint COM automation.

Handles adding video/audio files to slides and configuring media
playback settings such as volume, looping, and fade effects.
"""

import json
import logging
import os
from typing import Optional, Union

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from utils.navigation import goto_slide
from ppt_com.constants import msoTrue, msoFalse

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class AddVideoInput(BaseModel):
    """Input for adding a video to a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

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
    model_config = ConfigDict(str_strip_whitespace=True)

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
    model_config = ConfigDict(str_strip_whitespace=True)

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
        description="Trim start point in milliseconds",
    )
    end_point: Optional[int] = Field(
        default=None, ge=0,
        description="Trim end point in milliseconds",
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


# ---------------------------------------------------------------------------
# Helper: find a shape on a slide
# ---------------------------------------------------------------------------
def _get_shape(slide, name_or_index: Union[str, int]):
    """Find a shape on a slide by name or 1-based index.

    Args:
        slide: Slide COM object
        name_or_index: Shape name (str) or 1-based index (int)

    Returns:
        Shape COM object

    Raises:
        ValueError: If shape not found
    """
    if isinstance(name_or_index, int):
        if name_or_index < 1 or name_or_index > slide.Shapes.Count:
            raise ValueError(
                f"Shape index {name_or_index} out of range "
                f"(1-{slide.Shapes.Count})"
            )
        return slide.Shapes(name_or_index)
    else:
        for i in range(1, slide.Shapes.Count + 1):
            if slide.Shapes(i).Name == name_or_index:
                return slide.Shapes(i)
        raise ValueError(f"Shape '{name_or_index}' not found on slide")


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
    if volume is not None:
        media.Volume = volume
    if muted is not None:
        media.Muted = muted
    if start_point is not None:
        media.SetDisplayPictureFromPoint(start_point)
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

    return {
        "success": True,
        "shape_name": shape.Name,
    }


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
        return add_video(params)

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
        return add_audio(params)

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
        return set_media_settings(params)
