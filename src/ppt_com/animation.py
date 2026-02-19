"""Animation and transition operations for PowerPoint COM automation.

Handles slide transitions, adding/listing/removing shape animations,
and clearing all animations from a slide.
"""

import json
import logging
from typing import Optional, Union

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from utils.navigation import goto_slide
from ppt_com.constants import (
    msoTrue, msoFalse,
    msoAnimEffectAppear, msoAnimEffectFade, msoAnimEffectFly,
    msoAnimEffectWipe, msoAnimEffectZoom,
    msoAnimTriggerOnPageClick, msoAnimTriggerWithPrevious,
    msoAnimTriggerAfterPrevious,
    ppEffectNone, ppEffectFade, ppEffectPush, ppEffectWipe, ppEffectSplit,
    ANIMATION_EFFECT_NAMES, ANIMATION_TRIGGER_NAMES,
)

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Friendly name maps
# ---------------------------------------------------------------------------
ANIMATION_EFFECT_MAP: dict[str, int] = {
    "appear": 1, "fly": 2, "blinds": 3, "box": 4,
    "checkerboard": 5, "circle": 6, "diamond": 8,
    "dissolve": 9, "fade": 10, "split": 16, "wipe": 22,
    "zoom": 23, "bounce": 26, "float": 56,
    "grow_and_turn": 57, "spin": 61, "transparency": 62,
}

TRIGGER_MAP: dict[str, int] = {
    "on_click": 1, "with_previous": 2, "after_previous": 3, "on_shape_click": 4,
}

TRANSITION_EFFECT_MAP: dict[str, int] = {
    "none": 0, "cut": 257, "fade": 3844, "push": 3845,
    "wipe": 3846, "split": 3847, "reveal": 3848,
    "random": 513, "blinds_horizontal": 769, "blinds_vertical": 770, "dissolve": 1537,
}


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class SetSlideTransitionInput(BaseModel):
    """Input for setting a slide transition effect."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    effect: Union[int, str] = Field(
        default="fade",
        description=(
            "Transition effect: friendly name ('fade', 'push', 'wipe', 'split', "
            "'cut', 'reveal', 'random', 'blinds_horizontal', 'blinds_vertical', "
            "'dissolve', 'none') or PpEntryEffect integer"
        ),
    )
    duration: Optional[float] = Field(
        default=None, description="Transition duration in seconds"
    )
    advance_on_click: Optional[bool] = Field(
        default=None, description="Advance slide on mouse click"
    )
    advance_on_time: Optional[bool] = Field(
        default=None, description="Advance slide automatically after time"
    )
    advance_time: Optional[float] = Field(
        default=None, description="Auto-advance time in seconds"
    )


class AddAnimationInput(BaseModel):
    """Input for adding an animation effect to a shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    effect: Union[int, str] = Field(
        default="appear",
        description=(
            "Animation effect: friendly name ('appear', 'fade', 'fly', 'wipe', "
            "'zoom', 'bounce', 'spin', etc.) or MsoAnimEffect integer"
        ),
    )
    trigger: str = Field(
        default="on_click",
        description=(
            "Trigger type: 'on_click', 'with_previous', 'after_previous', "
            "or 'on_shape_click'"
        ),
    )
    duration: Optional[float] = Field(
        default=None, description="Animation duration in seconds"
    )
    delay: Optional[float] = Field(
        default=None, description="Delay before animation starts in seconds"
    )


class ListAnimationsInput(BaseModel):
    """Input for listing animations on a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")


class RemoveAnimationInput(BaseModel):
    """Input for removing a single animation from a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    animation_index: int = Field(..., ge=1, description="1-based animation index in the main sequence")


class ClearAnimationsInput(BaseModel):
    """Input for clearing all animations from a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    clear_transitions: bool = Field(
        default=False,
        description="Also clear the slide transition effect",
    )


# ---------------------------------------------------------------------------
# Helper: find a shape by name or index
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
def _set_slide_transition_impl(
    slide_index, effect, duration, advance_on_click, advance_on_time, advance_time,
):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    transition = slide.SlideShowTransition
    effect_int = TRANSITION_EFFECT_MAP.get(effect, effect) if isinstance(effect, str) else effect
    transition.EntryEffect = effect_int

    if duration is not None:
        transition.Duration = duration
    if advance_on_click is not None:
        transition.AdvanceOnClick = msoTrue if advance_on_click else msoFalse
    if advance_on_time is not None:
        transition.AdvanceOnTime = msoTrue if advance_on_time else msoFalse
    if advance_time is not None:
        transition.AdvanceTime = advance_time

    return {
        "success": True,
        "slide_index": slide_index,
        "effect": effect_int,
    }


def _add_animation_impl(
    slide_index, shape_name_or_index, effect, trigger, duration, delay,
):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    effect_int = ANIMATION_EFFECT_MAP.get(effect, effect) if isinstance(effect, str) else effect
    trigger_int = TRIGGER_MAP.get(trigger, 1)

    # AddEffect uses positional args: Shape, effectId, level, trigger, index
    effect_obj = slide.TimeLine.MainSequence.AddEffect(shape, effect_int, 0, trigger_int)

    if duration is not None:
        effect_obj.Timing.Duration = duration
    if delay is not None:
        effect_obj.Timing.TriggerDelayTime = delay

    return {
        "success": True,
        "shape_name": shape.Name,
        "effect": effect_int,
        "animation_index": effect_obj.Index,
    }


def _list_animations_impl(slide_index):
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    seq = slide.TimeLine.MainSequence
    animations = []
    for i in range(1, seq.Count + 1):
        eff = seq(i)
        effect_type = eff.EffectType
        trigger_type = eff.Timing.TriggerType
        animations.append({
            "index": eff.Index,
            "shape_name": eff.Shape.Name,
            "effect_type": effect_type,
            "effect_name": ANIMATION_EFFECT_NAMES.get(effect_type, f"Unknown({effect_type})"),
            "trigger_type": trigger_type,
            "trigger_name": ANIMATION_TRIGGER_NAMES.get(trigger_type, f"Unknown({trigger_type})"),
            "duration": eff.Timing.Duration,
        })

    return {
        "success": True,
        "slide_index": slide_index,
        "count": seq.Count,
        "animations": animations,
    }


def _remove_animation_impl(slide_index, animation_index):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    seq = slide.TimeLine.MainSequence
    if animation_index < 1 or animation_index > seq.Count:
        raise ValueError(
            f"Animation index {animation_index} out of range (1-{seq.Count})"
        )
    seq(animation_index).Delete()

    return {
        "success": True,
        "slide_index": slide_index,
        "remaining_count": seq.Count,
    }


def _clear_animations_impl(slide_index, clear_transitions):
    app = ppt._get_app_impl()
    goto_slide(app, slide_index)
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    seq = slide.TimeLine.MainSequence
    cleared_count = seq.Count

    # Delete in reverse order to avoid index shifting issues
    for i in range(seq.Count, 0, -1):
        seq(i).Delete()

    if clear_transitions:
        slide.SlideShowTransition.EntryEffect = 0  # ppEffectNone

    return {
        "success": True,
        "slide_index": slide_index,
        "cleared_count": cleared_count,
    }


# ---------------------------------------------------------------------------
# MCP tool functions (sync wrappers that delegate to COM thread)
# ---------------------------------------------------------------------------
def set_slide_transition(params: SetSlideTransitionInput) -> str:
    """Set the transition effect for a slide.

    Args:
        params: Slide index and transition properties.

    Returns:
        JSON confirming the transition was set.
    """
    try:
        result = ppt.execute(
            _set_slide_transition_impl,
            params.slide_index, params.effect, params.duration,
            params.advance_on_click, params.advance_on_time, params.advance_time,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to set slide transition: {str(e)}"})


def add_animation(params: AddAnimationInput) -> str:
    """Add an animation effect to a shape.

    Args:
        params: Slide index, shape identifier, and animation properties.

    Returns:
        JSON with shape name, effect, and animation index.
    """
    try:
        result = ppt.execute(
            _add_animation_impl,
            params.slide_index, params.shape_name_or_index,
            params.effect, params.trigger, params.duration, params.delay,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to add animation: {str(e)}"})


def list_animations(params: ListAnimationsInput) -> str:
    """List all animations in the main sequence of a slide.

    Args:
        params: Slide index.

    Returns:
        JSON with animation count and details for each animation.
    """
    try:
        result = ppt.execute(
            _list_animations_impl,
            params.slide_index,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to list animations: {str(e)}"})


def remove_animation(params: RemoveAnimationInput) -> str:
    """Remove a single animation from a slide's main sequence.

    Args:
        params: Slide index and 1-based animation index.

    Returns:
        JSON confirming removal and remaining count.
    """
    try:
        result = ppt.execute(
            _remove_animation_impl,
            params.slide_index, params.animation_index,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to remove animation: {str(e)}"})


def clear_animations(params: ClearAnimationsInput) -> str:
    """Clear all animations from a slide.

    Args:
        params: Slide index and whether to also clear transitions.

    Returns:
        JSON confirming how many animations were cleared.
    """
    try:
        result = ppt.execute(
            _clear_animations_impl,
            params.slide_index, params.clear_transitions,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": f"Failed to clear animations: {str(e)}"})


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register all animation tools with the MCP server."""

    @mcp.tool(
        name="ppt_set_slide_transition",
        annotations={
            "title": "Set Slide Transition",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_set_slide_transition(params: SetSlideTransitionInput) -> str:
        """Set the transition effect for a slide.

        Use a friendly name ('fade', 'push', 'wipe', 'split', 'cut', 'reveal',
        'random', 'blinds_horizontal', 'blinds_vertical', 'dissolve', 'none')
        or a PpEntryEffect integer. Optionally set duration, advance-on-click,
        and auto-advance timing.
        """
        return set_slide_transition(params)

    @mcp.tool(
        name="ppt_add_animation",
        annotations={
            "title": "Add Animation",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_add_animation(params: AddAnimationInput) -> str:
        """Add an animation effect to a shape on a slide.

        Specify the shape by name or 1-based index. Use a friendly effect name
        ('appear', 'fade', 'fly', 'wipe', 'zoom', 'bounce', 'spin', etc.)
        or an MsoAnimEffect integer. Set trigger, duration, and delay.
        """
        return add_animation(params)

    @mcp.tool(
        name="ppt_list_animations",
        annotations={
            "title": "List Animations",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_list_animations(params: ListAnimationsInput) -> str:
        """List all animations in the main sequence of a slide.

        Returns each animation's index, target shape name, effect type,
        trigger type, and duration.
        """
        return list_animations(params)

    @mcp.tool(
        name="ppt_remove_animation",
        annotations={
            "title": "Remove Animation",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_remove_animation(params: RemoveAnimationInput) -> str:
        """Remove a single animation from a slide's main sequence.

        Specify the 1-based animation index. Remaining animations re-index
        automatically. Use ppt_list_animations to find the correct index.
        """
        return remove_animation(params)

    @mcp.tool(
        name="ppt_clear_animations",
        annotations={
            "title": "Clear Animations",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_clear_animations(params: ClearAnimationsInput) -> str:
        """Clear all animations from a slide's main sequence.

        Optionally also clears the slide transition effect by setting
        clear_transitions=true.
        """
        return clear_animations(params)
