# Module: SlideShow, Export & Media

## Overview

This module handles three related areas: (1) SlideShow control -- starting, stopping, and navigating slide shows, pointer control; (2) Media operations -- inserting video/audio, configuring playback settings, trimming; (3) Animation -- adding animation effects to shapes via the Timeline API, configuring triggers and timing. It also covers hyperlinks and action settings.

## Dependencies

- **Internal**: `utils.com_wrapper` (PowerPointCOMWrapper, safe_com_call), `utils.units` (inches_to_points, cm_to_points), `utils.color` (rgb_to_int, int_to_rgb, hex_to_int, int_to_hex), `ppt_com.constants` (all slideshow/media/animation constants)
- **External**: `pywin32` (`win32com.client`, `pywintypes`)
- **Standard library**: `logging`, `os`, `time`

### Importing from Core Module

```python
from utils.com_wrapper import PowerPointCOMWrapper, safe_com_call
from utils.units import inches_to_points, points_to_inches
from utils.color import rgb_to_int, int_to_rgb, hex_to_int, int_to_hex
from ppt_com.constants import (
    # MsoTriState
    msoTrue, msoFalse,
    # SlideShow types
    ppShowTypeSpeaker, ppShowTypeWindow, ppShowTypeKiosk,
    # SlideShow advance modes
    ppSlideShowManualAdvance, ppSlideShowUseSlideTimings,
    # SlideShow range types
    ppShowAll, ppShowSlideRange, ppShowNamedSlideShow,
    # SlideShow states
    ppSlideShowRunning, ppSlideShowPaused,
    ppSlideShowBlackScreen, ppSlideShowWhiteScreen, ppSlideShowDone,
    # Pointer types
    ppSlideShowPointerNone, ppSlideShowPointerArrow,
    ppSlideShowPointerPen, ppSlideShowPointerAlwaysHidden,
    ppSlideShowPointerAutoArrow, ppSlideShowPointerEraser,
    # Animation effects (commonly used)
    msoAnimEffectAppear, msoAnimEffectFade, msoAnimEffectFly,
    msoAnimEffectBlinds, msoAnimEffectBox, msoAnimEffectCheckerboard,
    msoAnimEffectCircle, msoAnimEffectDiamond, msoAnimEffectDissolve,
    msoAnimEffectWipe, msoAnimEffectZoom, msoAnimEffectSplit,
    msoAnimEffectStretch, msoAnimEffectSpiral, msoAnimEffectBounce,
    msoAnimEffectFloat, msoAnimEffectGrowAndTurn,
    msoAnimEffectSpin, msoAnimEffectChangeFillColor,
    msoAnimEffectChangeFontColor, msoAnimEffectTransparency,
    msoAnimEffectPathDown, msoAnimEffectPathUp,
    msoAnimEffectPathLeft, msoAnimEffectPathRight,
    # Animation triggers
    msoAnimTriggerNone, msoAnimTriggerOnPageClick,
    msoAnimTriggerWithPrevious, msoAnimTriggerAfterPrevious,
    msoAnimTriggerOnShapeClick,
    # Action types
    ppActionNone, ppActionNextSlide, ppActionPreviousSlide,
    ppActionFirstSlide, ppActionLastSlide, ppActionHyperlink,
    ppActionEndShow,
    # Mouse click / mouse over
    ppMouseClick, ppMouseOver,
)
import logging
import os
import time

logger = logging.getLogger(__name__)
```

## File Structure

```
ppt_com_mcp/
  ppt_com/
    slideshow.py    # SlideShow control
    media.py        # Media (video/audio) operations
    animations.py   # Animation effects
```

---

## Constants Needed

These constants MUST be defined in `ppt_com/constants.py` (from the core module).

### PpSlideShowType

| Name | Value | Description |
|------|-------|-------------|
| `ppShowTypeSpeaker` | 1 | Full screen (speaker) |
| `ppShowTypeWindow` | 2 | Window mode |
| `ppShowTypeKiosk` | 3 | Kiosk (auto full screen, loop) |

### PpSlideShowAdvanceMode

| Name | Value | Description |
|------|-------|-------------|
| `ppSlideShowManualAdvance` | 1 | Manual advance |
| `ppSlideShowUseSlideTimings` | 2 | Use slide timings |
| `ppSlideShowRehearseNewTimings` | 3 | Rehearsal mode |

### PpSlideShowRangeType

| Name | Value | Description |
|------|-------|-------------|
| `ppShowAll` | 1 | All slides |
| `ppShowSlideRange` | 2 | Slide range |
| `ppShowNamedSlideShow` | 3 | Named slide show |

### PpSlideShowState

| Name | Value | Description |
|------|-------|-------------|
| `ppSlideShowRunning` | 1 | Running |
| `ppSlideShowPaused` | 2 | Paused |
| `ppSlideShowBlackScreen` | 3 | Black screen |
| `ppSlideShowWhiteScreen` | 4 | White screen |
| `ppSlideShowDone` | 5 | Done |

### PpSlideShowPointerType

| Name | Value | Description |
|------|-------|-------------|
| `ppSlideShowPointerNone` | 0 | No pointer |
| `ppSlideShowPointerArrow` | 1 | Arrow |
| `ppSlideShowPointerPen` | 2 | Pen |
| `ppSlideShowPointerAlwaysHidden` | 3 | Always hidden |
| `ppSlideShowPointerAutoArrow` | 4 | Auto arrow |
| `ppSlideShowPointerEraser` | 5 | Eraser |

### MsoAnimEffect (commonly used)

| Name | Value | Description | Category |
|------|-------|-------------|----------|
| `msoAnimEffectAppear` | 1 | Appear | Entrance |
| `msoAnimEffectFly` | 2 | Fly in | Entrance |
| `msoAnimEffectBlinds` | 3 | Blinds | Entrance |
| `msoAnimEffectBox` | 4 | Box | Entrance |
| `msoAnimEffectCheckerboard` | 5 | Checkerboard | Entrance |
| `msoAnimEffectCircle` | 6 | Circle | Entrance |
| `msoAnimEffectDiamond` | 8 | Diamond | Entrance |
| `msoAnimEffectDissolve` | 9 | Dissolve | Entrance |
| `msoAnimEffectFade` | 10 | Fade | Entrance |
| `msoAnimEffectSplit` | 16 | Split | Entrance |
| `msoAnimEffectWipe` | 22 | Wipe | Entrance |
| `msoAnimEffectZoom` | 23 | Zoom | Entrance |
| `msoAnimEffectBounce` | 26 | Bounce | Entrance |
| `msoAnimEffectFloat` | 56 | Float | Entrance |
| `msoAnimEffectGrowAndTurn` | 57 | Grow and Turn | Entrance |
| `msoAnimEffectSpin` | 61 | Spin | Emphasis |
| `msoAnimEffectChangeFillColor` | 54 | Change fill color | Emphasis |
| `msoAnimEffectChangeFontColor` | 58 | Change font color | Emphasis |
| `msoAnimEffectTransparency` | 62 | Transparency | Emphasis |
| `msoAnimEffectPathDown` | 64 | Path down | Motion path |
| `msoAnimEffectPathUp` | 65 | Path up | Motion path |
| `msoAnimEffectPathLeft` | 66 | Path left | Motion path |
| `msoAnimEffectPathRight` | 67 | Path right | Motion path |

### MsoAnimTriggerType

| Name | Value | Description |
|------|-------|-------------|
| `msoAnimTriggerNone` | 0 | No trigger |
| `msoAnimTriggerOnPageClick` | 1 | On page click |
| `msoAnimTriggerWithPrevious` | 2 | With previous |
| `msoAnimTriggerAfterPrevious` | 3 | After previous |
| `msoAnimTriggerOnShapeClick` | 4 | On shape click |

### PpActionType

| Name | Value | Description |
|------|-------|-------------|
| `ppActionNone` | 0 | No action |
| `ppActionNextSlide` | 1 | Next slide |
| `ppActionPreviousSlide` | 2 | Previous slide |
| `ppActionFirstSlide` | 3 | First slide |
| `ppActionLastSlide` | 4 | Last slide |
| `ppActionLastSlideViewed` | 5 | Last viewed slide |
| `ppActionEndShow` | 6 | End show |
| `ppActionHyperlink` | 7 | Hyperlink |
| `ppActionRunProgram` | 9 | Run program |

### ActionSettings Index

| Name | Value | Description |
|------|-------|-------------|
| `ppMouseClick` | 1 | Click action |
| `ppMouseOver` | 2 | Mouse-over action |

---

## File: `ppt_com/slideshow.py` - SlideShow Control

### Purpose

Provide MCP tools for controlling PowerPoint slide shows: starting, stopping, navigating, and controlling the pointer.

---

### Tool: `start_slideshow`

- **Description**: Start a slide show presentation
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `show_type` | int | No | 1=speaker (fullscreen), 2=window, 3=kiosk. Default: 1 |
  | `start_slide` | int | No | Starting slide index. Default: 1 |
  | `end_slide` | int | No | Ending slide index. Default: last slide |
  | `advance_mode` | int | No | 1=manual, 2=use timings. Default: 1 |
  | `loop` | bool | No | Loop continuously. Default: false |
  | `show_narration` | bool | No | Play narration. Default: true |
  | `show_animation` | bool | No | Play animations. Default: true |
- **Returns**:
  ```json
  {
    "status": "success",
    "show_type": 1,
    "current_slide": 1,
    "total_slides": 10
  }
  ```
- **COM Implementation**:
  ```python
  def start_slideshow(app, show_type=1, start_slide=1, end_slide=None,
                      advance_mode=1, loop=False,
                      show_narration=True, show_animation=True):
      pres = app.ActivePresentation
      settings = pres.SlideShowSettings

      if end_slide is None:
          end_slide = pres.Slides.Count

      settings.ShowType = show_type
      settings.AdvanceMode = advance_mode
      settings.LoopUntilStopped = msoTrue if loop else msoFalse
      settings.ShowWithNarration = msoTrue if show_narration else msoFalse
      settings.ShowWithAnimation = msoTrue if show_animation else msoFalse

      # Set slide range
      settings.RangeType = ppShowSlideRange  # 2
      settings.StartingSlide = start_slide
      settings.EndingSlide = end_slide

      # Start the show
      ssw = settings.Run()  # Returns SlideShowWindow

      # Give it a moment to start
      time.sleep(0.5)

      view = ssw.View
      return {
          "status": "success",
          "show_type": show_type,
          "current_slide": view.CurrentShowPosition,
          "total_slides": pres.Slides.Count,
      }
  ```
- **Error Cases**:
  - No presentation open
  - Invalid slide range

---

### Tool: `stop_slideshow`

- **Description**: End the currently running slide show
- **Parameters**: None
- **Returns**:
  ```json
  {
    "status": "success",
    "message": "Slide show ended"
  }
  ```
- **COM Implementation**:
  ```python
  def stop_slideshow(app):
      try:
          ssw = app.SlideShowWindows(1)
          ssw.View.Exit()
      except Exception:
          # No slide show running
          pass

      return {
          "status": "success",
          "message": "Slide show ended",
      }
  ```
- **Error Cases**:
  - No slide show running (handled gracefully)

---

### Tool: `navigate_slideshow`

- **Description**: Navigate within a running slide show
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `action` | string | Yes | "next", "previous", "first", "last", "goto" |
  | `slide_index` | int | No | Target slide (required for "goto") |
- **Returns**:
  ```json
  {
    "status": "success",
    "current_slide": 3,
    "state": 1
  }
  ```
- **COM Implementation**:
  ```python
  def navigate_slideshow(app, action, slide_index=None):
      ssw = app.SlideShowWindows(1)
      view = ssw.View

      if action == "next":
          view.Next()
      elif action == "previous":
          view.Previous()
      elif action == "first":
          view.First()
      elif action == "last":
          view.Last()
      elif action == "goto":
          if slide_index is None:
              raise ValueError("slide_index is required for 'goto' action")
          view.GotoSlide(slide_index)
      else:
          raise ValueError(f"Unknown action: {action}")

      return {
          "status": "success",
          "current_slide": view.CurrentShowPosition,
          "state": view.State,
      }
  ```
- **Error Cases**:
  - No slide show running
  - Invalid action
  - slide_index out of range for "goto"

---

### Tool: `set_slideshow_pointer`

- **Description**: Change the slide show pointer type and color
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `pointer_type` | int | No | 0=none, 1=arrow, 2=pen, 3=always hidden, 4=auto arrow, 5=eraser |
  | `pen_color` | string | No | Pen color as "#RRGGBB" |
- **Returns**:
  ```json
  {
    "status": "success",
    "pointer_type": 2
  }
  ```
- **COM Implementation**:
  ```python
  def set_slideshow_pointer(app, pointer_type=None, pen_color=None):
      ssw = app.SlideShowWindows(1)
      view = ssw.View

      if pointer_type is not None:
          view.PointerType = pointer_type

      if pen_color is not None:
          view.PointerColor.RGB = hex_to_int(pen_color)

      return {
          "status": "success",
          "pointer_type": view.PointerType,
      }
  ```
- **Error Cases**:
  - No slide show running

---

### Tool: `get_slideshow_state`

- **Description**: Get the current state of the running slide show
- **Parameters**: None
- **Returns**:
  ```json
  {
    "status": "success",
    "is_running": true,
    "current_slide": 3,
    "state": 1,
    "state_name": "Running",
    "elapsed_time": 45.2,
    "pointer_type": 1
  }
  ```
- **COM Implementation**:
  ```python
  STATE_NAMES = {
      1: "Running",
      2: "Paused",
      3: "BlackScreen",
      4: "WhiteScreen",
      5: "Done",
  }

  def get_slideshow_state(app):
      try:
          if app.SlideShowWindows.Count == 0:
              return {"status": "success", "is_running": False}

          ssw = app.SlideShowWindows(1)
          view = ssw.View

          return {
              "status": "success",
              "is_running": True,
              "current_slide": view.CurrentShowPosition,
              "state": view.State,
              "state_name": STATE_NAMES.get(view.State, f"Unknown({view.State})"),
              "elapsed_time": view.PresentationElapsedTime,
              "pointer_type": view.PointerType,
          }
      except Exception:
          return {"status": "success", "is_running": False}
  ```
- **Error Cases**:
  - No slide show running (returns is_running: false)

---

## File: `ppt_com/media.py` - Media Operations

### Purpose

Provide MCP tools for inserting and configuring video and audio media in presentations.

---

### Tool: `insert_media`

- **Description**: Insert a video or audio file into a slide
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `file_path` | string | Yes | Absolute path to media file (MP4, WMV, AVI, MP3, WAV, etc.) |
  | `left` | float | No | Left position in points. Default: 100 |
  | `top` | float | No | Top position in points. Default: 100 |
  | `width` | float | No | Width in points. Default: -1 (auto) |
  | `height` | float | No | Height in points. Default: -1 (auto) |
  | `link_to_file` | bool | No | Link instead of embed. Default: false |
- **Returns**:
  ```json
  {
    "status": "success",
    "shape_name": "Video 1",
    "media_type": "Video",
    "is_embedded": true,
    "duration_ms": 30000,
    "left": 100.0,
    "top": 100.0,
    "width": 480.0,
    "height": 270.0
  }
  ```
- **COM Implementation**:
  ```python
  def insert_media(app, slide_index, file_path,
                   left=100, top=100, width=-1, height=-1,
                   link_to_file=False):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)

      if not os.path.isabs(file_path):
          raise ValueError("file_path must be an absolute path")
      if not os.path.exists(file_path):
          raise FileNotFoundError(f"Media file not found: {file_path}")

      # AddMediaObject2 is the modern method
      # IMPORTANT: LinkToFile and SaveWithDocument cannot both be False
      shape = slide.Shapes.AddMediaObject2(
          FileName=file_path,
          LinkToFile=msoTrue if link_to_file else msoFalse,
          SaveWithDocument=msoFalse if link_to_file else msoTrue,
          Left=left,
          Top=top,
          Width=width,
          Height=height,
      )

      media = shape.MediaFormat
      # PpMediaType: 1=Other, 2=Sound, 3=Movie
      media_type_names = {1: "Other", 2: "Audio", 3: "Video"}

      return {
          "status": "success",
          "shape_name": shape.Name,
          "media_type": media_type_names.get(media.MediaType, "Unknown"),
          "is_embedded": media.IsEmbedded,
          "duration_ms": media.Length,
          "left": shape.Left,
          "top": shape.Top,
          "width": shape.Width,
          "height": shape.Height,
      }
  ```
- **Error Cases**:
  - File not found
  - Unsupported media format
  - File path is not absolute

---

### Tool: `set_media_properties`

- **Description**: Configure media playback properties
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Media shape name or index |
  | `volume` | float | No | Volume 0.0 (silent) to 1.0 (max) |
  | `muted` | bool | No | Mute on/off |
  | `start_point` | int | No | Trim start in milliseconds |
  | `end_point` | int | No | Trim end in milliseconds |
  | `fade_in` | int | No | Fade-in duration in milliseconds |
  | `fade_out` | int | No | Fade-out duration in milliseconds |
  | `play_on_entry` | bool | No | Auto-play when slide is shown |
  | `loop` | bool | No | Loop until stopped |
  | `hide_while_not_playing` | bool | No | Hide media when not playing |
  | `rewind_after` | bool | No | Rewind to beginning after playback |
- **Returns**:
  ```json
  {
    "status": "success",
    "shape_name": "Video 1",
    "volume": 0.5,
    "duration_ms": 30000,
    "start_point": 5000,
    "end_point": 25000
  }
  ```
- **COM Implementation**:
  ```python
  def set_media_properties(app, slide_index, shape_name_or_index,
                           volume=None, muted=None,
                           start_point=None, end_point=None,
                           fade_in=None, fade_out=None,
                           play_on_entry=None, loop=None,
                           hide_while_not_playing=None,
                           rewind_after=None):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = _get_media_shape(slide, shape_name_or_index)
      media = shape.MediaFormat

      if volume is not None:
          media.Volume = volume
      if muted is not None:
          media.Muted = muted
      if start_point is not None:
          media.StartPoint = start_point
      if end_point is not None:
          media.EndPoint = end_point
      if fade_in is not None:
          media.FadeInDuration = fade_in
      if fade_out is not None:
          media.FadeOutDuration = fade_out

      # PlaySettings are accessed via AnimationSettings
      play = shape.AnimationSettings.PlaySettings
      if play_on_entry is not None:
          play.PlayOnEntry = msoTrue if play_on_entry else msoFalse
          if play_on_entry:
              shape.AnimationSettings.Animate = msoTrue
      if loop is not None:
          play.LoopUntilStopped = msoTrue if loop else msoFalse
      if hide_while_not_playing is not None:
          play.HideWhileNotPlaying = msoTrue if hide_while_not_playing else msoFalse
      if rewind_after is not None:
          play.RewindMovie = msoTrue if rewind_after else msoFalse

      return {
          "status": "success",
          "shape_name": shape.Name,
          "volume": media.Volume,
          "duration_ms": media.Length,
          "start_point": media.StartPoint,
          "end_point": media.EndPoint,
      }
  ```
- **Error Cases**:
  - Shape is not a media object
  - start_point/end_point out of media duration range

---

## File: `ppt_com/animations.py` - Animation Effects

### Purpose

Provide MCP tools for adding and configuring animation effects on shapes via the Timeline API, and for managing hyperlinks/action settings.

---

### Tool: `add_animation`

- **Description**: Add an animation effect to a shape
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Target shape name or index |
  | `effect_id` | int | Yes | MsoAnimEffect value (e.g., 1=Appear, 10=Fade, 2=Fly) |
  | `trigger` | int | No | 1=on click, 2=with previous, 3=after previous. Default: 1 |
  | `duration` | float | No | Duration in seconds. Default: 0.5 |
  | `delay` | float | No | Delay before start in seconds. Default: 0 |
  | `is_exit` | bool | No | If true, this is an exit (disappear) animation. Default: false |
- **Returns**:
  ```json
  {
    "status": "success",
    "shape_name": "Rectangle 1",
    "effect_id": 10,
    "trigger": 1,
    "duration": 0.5,
    "effect_index": 1
  }
  ```
- **COM Implementation**:
  ```python
  def add_animation(app, slide_index, shape_name_or_index,
                    effect_id, trigger=1, duration=0.5,
                    delay=0, is_exit=False):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = _get_shape(slide, shape_name_or_index)

      # Use Timeline.MainSequence for standard animations
      effect = slide.TimeLine.MainSequence.AddEffect(
          Shape=shape,
          effectId=effect_id,
          trigger=trigger,
      )

      # Configure timing
      effect.Timing.Duration = duration
      if delay > 0:
          effect.Timing.TriggerDelayTime = delay

      # Set as exit animation if requested
      if is_exit:
          effect.Exit = msoTrue

      # Get effect index in sequence
      effect_index = 0
      for i in range(1, slide.TimeLine.MainSequence.Count + 1):
          if slide.TimeLine.MainSequence(i) is effect:
              effect_index = i
              break

      return {
          "status": "success",
          "shape_name": shape.Name,
          "effect_id": effect_id,
          "trigger": trigger,
          "duration": duration,
          "effect_index": effect_index,
      }
  ```
- **Error Cases**:
  - Shape not found
  - Invalid effect_id

---

### Tool: `list_animations`

- **Description**: List all animations on a slide
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
- **Returns**:
  ```json
  {
    "status": "success",
    "slide_index": 1,
    "animation_count": 3,
    "animations": [
      {
        "index": 1,
        "shape_name": "Title 1",
        "effect_id": 10,
        "is_exit": false,
        "trigger": 1,
        "duration": 0.5,
        "delay": 0.0
      }
    ]
  }
  ```
- **COM Implementation**:
  ```python
  def list_animations(app, slide_index):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      seq = slide.TimeLine.MainSequence

      animations = []
      for i in range(1, seq.Count + 1):
          effect = seq(i)
          animations.append({
              "index": i,
              "shape_name": effect.Shape.Name,
              "effect_id": effect.EffectType,
              "is_exit": effect.Exit == msoTrue,
              "trigger": effect.Timing.TriggerType,
              "duration": effect.Timing.Duration,
              "delay": effect.Timing.TriggerDelayTime,
          })

      return {
          "status": "success",
          "slide_index": slide_index,
          "animation_count": seq.Count,
          "animations": animations,
      }
  ```
- **Error Cases**:
  - Invalid slide_index

---

### Tool: `remove_animation`

- **Description**: Remove an animation effect from a slide
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `effect_index` | int | No | 1-based effect index in the sequence. If omitted, removes all |
- **Returns**:
  ```json
  {
    "status": "success",
    "removed": 1,
    "remaining": 2
  }
  ```
- **COM Implementation**:
  ```python
  def remove_animation(app, slide_index, effect_index=None):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      seq = slide.TimeLine.MainSequence

      if effect_index is not None:
          seq(effect_index).Delete()
          removed = 1
      else:
          # Remove all animations
          count = seq.Count
          # Delete from end to start to avoid index shifting
          for i in range(count, 0, -1):
              seq(i).Delete()
          removed = count

      return {
          "status": "success",
          "removed": removed,
          "remaining": seq.Count,
      }
  ```
- **Error Cases**:
  - Invalid slide_index
  - effect_index out of range

---

### Tool: `set_hyperlink`

- **Description**: Add or modify a hyperlink on a shape
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
  | `shape_name_or_index` | string or int | Yes | Shape name or index |
  | `address` | string | No | URL, file path, or email (e.g., "https://...", "mailto:...") |
  | `sub_address` | string | No | Intra-document target (e.g., ",3," for slide 3) |
  | `screen_tip` | string | No | Tooltip on hover |
  | `action_type` | int | No | PpActionType (0=none, 1=next, 2=previous, 6=end show, 7=hyperlink) |
  | `event_type` | int | No | 1=click, 2=mouse over. Default: 1 |
- **Returns**:
  ```json
  {
    "status": "success",
    "shape_name": "Button 1",
    "action_type": 7,
    "address": "https://www.example.com"
  }
  ```
- **COM Implementation**:
  ```python
  def set_hyperlink(app, slide_index, shape_name_or_index,
                    address=None, sub_address=None, screen_tip=None,
                    action_type=None, event_type=1):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      shape = _get_shape(slide, shape_name_or_index)

      action = shape.ActionSettings(event_type)  # 1=ppMouseClick, 2=ppMouseOver

      if action_type is not None:
          action.Action = action_type

      # For hyperlinks (action_type=7=ppActionHyperlink)
      if address is not None:
          if action.Action != 7:
              action.Action = 7  # ppActionHyperlink
          action.Hyperlink.Address = address

      if sub_address is not None:
          if action.Action != 7:
              action.Action = 7
          action.Hyperlink.SubAddress = sub_address

      if screen_tip is not None:
          action.Hyperlink.ScreenTip = screen_tip

      return {
          "status": "success",
          "shape_name": shape.Name,
          "action_type": action.Action,
          "address": action.Hyperlink.Address if action.Action == 7 else None,
      }
  ```
- **Error Cases**:
  - Shape not found

### SubAddress Format for Internal Slide Links

To link to a specific slide within the same presentation, use the SubAddress format: `",SlideIndex,"` (SlideID, SlideIndex, SlideTitle -- any field can be empty):

```python
# Link to slide 5
action.Hyperlink.Address = ""  # Empty = same document
action.Hyperlink.SubAddress = ",5,"

# Link to a slide by ID and index
action.Hyperlink.SubAddress = "256,3,My Slide Title"
```

---

### Tool: `get_hyperlinks`

- **Description**: List all hyperlinks on a slide
- **Parameters**:
  | Name | Type | Required | Description |
  |------|------|----------|-------------|
  | `slide_index` | int | Yes | 1-based slide index |
- **Returns**:
  ```json
  {
    "status": "success",
    "slide_index": 1,
    "hyperlink_count": 2,
    "hyperlinks": [
      {
        "index": 1,
        "address": "https://www.example.com",
        "sub_address": "",
        "type": 1
      }
    ]
  }
  ```
- **COM Implementation**:
  ```python
  def get_hyperlinks(app, slide_index):
      pres = app.ActivePresentation
      slide = pres.Slides(slide_index)
      hls = slide.Hyperlinks

      hyperlinks = []
      for i in range(1, hls.Count + 1):
          hl = hls(i)
          hyperlinks.append({
              "index": i,
              "address": hl.Address,
              "sub_address": hl.SubAddress,
              "type": hl.Type,
          })

      return {
          "status": "success",
          "slide_index": slide_index,
          "hyperlink_count": hls.Count,
          "hyperlinks": hyperlinks,
      }
  ```
- **Error Cases**:
  - Invalid slide_index

---

### Helper Functions

```python
def _get_shape(slide, name_or_index):
    """Get a shape by name or 1-based index."""
    if isinstance(name_or_index, int):
        if name_or_index < 1 or name_or_index > slide.Shapes.Count:
            raise ValueError(
                f"Shape index {name_or_index} out of range (1-{slide.Shapes.Count})"
            )
        return slide.Shapes(name_or_index)
    else:
        for i in range(1, slide.Shapes.Count + 1):
            if slide.Shapes(i).Name == name_or_index:
                return slide.Shapes(i)
        raise ValueError(f"Shape '{name_or_index}' not found")


def _get_media_shape(slide, name_or_index):
    """Get a media shape and verify it contains media."""
    shape = _get_shape(slide, name_or_index)
    # msoMedia = 16
    if shape.Type != 16:
        raise ValueError(f"Shape '{shape.Name}' is not a media object (type={shape.Type})")
    return shape
```

---

## Implementation Notes

### 1. SlideShowWindows Collection

A running slide show is accessed via `app.SlideShowWindows`. If no show is running, `SlideShowWindows.Count` is 0. Always check before accessing:

```python
if app.SlideShowWindows.Count > 0:
    view = app.SlideShowWindows(1).View
```

### 2. SlideShowSettings vs SlideShowView

- `SlideShowSettings` (pres.SlideShowSettings): Configure BEFORE starting the show
- `SlideShowView` (ssw.View): Control DURING the running show

Do not confuse the two. Settings are set before `Run()`, View is used after.

### 3. Media PlaySettings Access Path

Media playback settings are accessed through a surprising path:

```
shape.AnimationSettings.PlaySettings
```

NOT `shape.MediaFormat.PlaySettings` (which does not exist).

To enable auto-play, you may also need to set `shape.AnimationSettings.Animate = msoTrue`.

### 4. AddMediaObject2 LinkToFile/SaveWithDocument Rules

- `LinkToFile=False, SaveWithDocument=True`: Embed the media (default, recommended)
- `LinkToFile=True, SaveWithDocument=False`: Link only (smaller file, requires original media path)
- `LinkToFile=True, SaveWithDocument=True`: Link AND embed a copy
- `LinkToFile=False, SaveWithDocument=False`: **ERROR** -- at least one must be True

### 5. Supported Media Formats

- **Video**: MP4, WMV, AVI, M4V, MOV (some formats may require codecs)
- **Audio**: MP3, WAV, WMA, M4A, AAC, MIDI

### 6. Timeline API: MainSequence vs InteractiveSequences

- `slide.TimeLine.MainSequence`: Standard animations triggered by page click or timing
- `slide.TimeLine.InteractiveSequences`: Animations triggered by clicking a specific shape

For interactive (shape-click) triggers, use:
```python
interactive_seq = slide.TimeLine.InteractiveSequences.Add()
effect = interactive_seq.AddEffect(
    Shape=target_shape,
    effectId=1,
    trigger=4,  # msoAnimTriggerOnShapeClick
)
effect.Timing.TriggerShape = trigger_shape
```

### 7. Exit Animations

To make an animation an "exit" (disappear) effect, set `effect.Exit = msoTrue` after creating it:

```python
effect = seq.AddEffect(Shape=shape, effectId=10, trigger=1)
effect.Exit = msoTrue  # Now this is a "Fade Out" instead of "Fade In"
```

### 8. Deleting Animations in Reverse Order

When removing multiple animations, delete from the end of the sequence to the beginning to avoid index shifting:

```python
for i in range(seq.Count, 0, -1):
    seq(i).Delete()
```

### 9. Hyperlink via ActionSettings

The most reliable way to set hyperlinks is through ActionSettings, not directly via the Hyperlinks collection. Always set `Action = ppActionHyperlink (7)` before setting the Hyperlink.Address:

```python
action = shape.ActionSettings(1)  # ppMouseClick
action.Action = 7  # ppActionHyperlink
action.Hyperlink.Address = "https://example.com"
```

### 10. SlideShow Timing

After calling `settings.Run()`, give the show a brief moment (e.g., `time.sleep(0.5)`) before querying the SlideShowView, as the show may not be fully initialized immediately. This is especially important in automated scenarios.

### 11. All Indices Are 1-Based

Consistent with all PowerPoint COM collections:
- `SlideShowWindows(1)` = first (usually only) slide show window
- `MainSequence(1)` = first animation effect
- `ActionSettings(1)` = ppMouseClick, `ActionSettings(2)` = ppMouseOver
