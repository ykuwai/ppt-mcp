"""Handing a chart's animations to the copy that replaces it (#343).

The stand-in timeline behaves as PowerPoint 16.0.20527 was measured to on
Windows: ApplyAnimation appends copies of the main sequence effects, in their
order, and then each interactive effect as a click effect with no trigger;
deleting a shape deletes every effect on it.
"""

import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from ppt_com import animation_carry as carry  # noqa: E402

FADE, APPEAR, WIPE, SPIN, PATH_DOWN = 10, 1, 22, 61, 127
ON_CLICK, AFTER_PREVIOUS, ON_SHAPE_CLICK, ON_BOOKMARK = 1, 3, 4, 5


class Behaviors:
    def __init__(self, items):
        self.items = items

    @property
    def Count(self):
        return len(self.items)

    def __call__(self, i):
        return self.items[i - 1]


class Parameters:
    """Setting the direction puts the effect's duration back to its default,
    as PowerPoint does."""

    def __init__(self, effect, direction):
        self._effect = effect
        self._direction = direction

    @property
    def Direction(self):
        return self._direction

    @Direction.setter
    def Direction(self, value):
        self._direction = value
        self._effect.Timing.Duration = 0.5


class Effect:
    def __init__(self, sequence, shape, kind, *, exit=0, trigger=ON_CLICK, duration=0.5,
                 delay=0.0, trigger_shape=None, path=None, direction=None):
        self.sequence = sequence
        self.Shape = shape
        self.EffectType = kind
        self.Exit = exit
        self.Timing = SimpleNamespace(TriggerType=trigger, Duration=duration, TriggerDelayTime=delay,
                                      TriggerShape=trigger_shape, RepeatCount=1, RewindAtEnd=0,
                                      AutoReverse=0, Accelerate=0.0, Decelerate=0.0)
        moves = [SimpleNamespace(Type=carry.msoAnimTypeMotion,
                                 MotionEffect=SimpleNamespace(Path=path))] if path else []
        self.Behaviors = Behaviors(moves)
        self.EffectInformation = SimpleNamespace(BuildByLevelEffect=0, AfterEffect=0,
                                                 Dim=SimpleNamespace(RGB=0))
        self.EffectParameters = Parameters(self, direction)

    def copy_for(self, sequence, shape, **changes):
        effect = Effect(sequence, shape, self.EffectType, exit=self.Exit,
                        trigger=self.Timing.TriggerType, duration=self.Timing.Duration,
                        delay=self.Timing.TriggerDelayTime, trigger_shape=self.Timing.TriggerShape,
                        direction=self.EffectParameters.Direction)
        effect.Behaviors = self.Behaviors
        for key in ("RepeatCount", "RewindAtEnd", "AutoReverse", "Accelerate", "Decelerate"):
            setattr(effect.Timing, key, getattr(self.Timing, key))
        effect.EffectInformation = SimpleNamespace(**vars(self.EffectInformation))
        for key, value in changes.items():
            setattr(effect.Timing, key, value)
        return effect

    def Delete(self):
        self.sequence.effects.remove(self)

    def MoveTo(self, index):
        self.sequence.effects.remove(self)
        self.sequence.effects.insert(index - 1, self)


class Sequence:
    def __init__(self):
        self.effects = []

    @property
    def Count(self):
        return len(self.effects)

    def __call__(self, i):
        return self.effects[i - 1]

    def add(self, shape, kind, **kw):
        effect = Effect(self, shape, kind, **kw)
        self.effects.append(effect)
        return effect

    def AddTriggerEffect(self, shape, kind, trigger, trigger_shape):
        return self.add(shape, kind, trigger=trigger, trigger_shape=trigger_shape, duration=0.5)

    def ConvertToAfterEffect(self, effect, after, dim=None):
        effect.EffectInformation.AfterEffect = after
        if dim is not None:
            effect.EffectInformation.Dim = SimpleNamespace(RGB=dim)
        return effect


class Slide:
    def __init__(self):
        self.main = Sequence()
        self.interactive = []
        slide = self
        self.TimeLine = SimpleNamespace(
            MainSequence=self.main,
            InteractiveSequences=type("I", (), {
                "Count": property(lambda s: len(slide.interactive)),
                "__call__": lambda s, i: slide.interactive[i - 1],
                "Add": lambda s: slide.new_interactive()})())
        self.picked = None

    def new_interactive(self):
        sequence = Sequence()
        self.interactive.append(sequence)
        return sequence

    def shape(self, shape_id):
        slide = self
        shape = SimpleNamespace(Id=shape_id)
        shape.PickupAnimation = lambda: setattr(slide, "picked", shape)
        shape.ApplyAnimation = lambda: slide.apply(shape)
        return shape

    def apply(self, target):
        source = self.picked
        own = [e for e in self.main.effects if e.Shape is source]
        own += [e for s in self.interactive for e in s.effects if e.Shape is source]
        for effect in own:
            interactive = effect.sequence is not self.main
            if interactive and effect.Timing.TriggerShape is source:
                # Started by clicking the shape itself: a sequence of its own,
                # started by clicking the target.
                sequence = self.new_interactive()
                sequence.effects.append(effect.copy_for(sequence, target, TriggerShape=target))
                continue
            self.main.effects.append(effect.copy_for(
                self.main, target,
                **({"TriggerType": ON_CLICK, "TriggerShape": None} if interactive else {})))

    def delete(self, shape):
        for sequence in [self.main] + self.interactive:
            sequence.effects = [e for e in sequence.effects if e.Shape is not shape]
        self.interactive = [s for s in self.interactive if s.effects]


def connector(begin, begin_site, end, end_site):
    ends = SimpleNamespace(BeginConnected=-1, BeginConnectedShape=begin, BeginConnectionSite=begin_site,
                           EndConnected=-1, EndConnectedShape=end, EndConnectionSite=end_site)

    def connect(which):
        def attach(shape, site):
            setattr(ends, f"{which}ConnectedShape", shape)
            setattr(ends, f"{which}ConnectionSite", site)
        return attach

    ends.BeginConnect = connect("Begin")
    ends.EndConnect = connect("End")
    return SimpleNamespace(ConnectorFormat=ends)


@pytest.fixture
def deck():
    """The slide win-tester built: S, chart G, R, and a connector from G to R."""
    slide = Slide()
    g, r, s = slide.shape(2), slide.shape(3), slide.shape(4)
    main = slide.main
    main.add(s, FADE)
    main.add(g, FADE, trigger=AFTER_PREVIOUS, duration=0.7, delay=0.3)
    main.add(s, APPEAR, duration=0.001)
    main.add(g, PATH_DOWN, duration=2.0, path="M 0 0 L 0 0.25 E")
    main.add(g, SPIN, duration=2.0)
    main.add(g, FADE, exit=-1)
    slide.new_interactive().add(g, WIPE, trigger=ON_SHAPE_CLICK, trigger_shape=r, duration=1.1,
                                direction=4)
    slide.new_interactive().add(r, APPEAR, trigger=ON_SHAPE_CLICK, trigger_shape=g)
    link = connector(g, 4, r, 2)
    return slide, g, r, link


def test_the_copy_ends_up_where_the_chart_was(deck):
    slide, g, r, link = deck
    before = carry.timeline(slide, g.Id)
    n = slide.shape(6)

    carry.copy_animations(slide, g, n)
    carry.move_links(slide, g, n, [link], [])
    assert carry.timeline(slide, n.Id, leave_out=g.Id) == before

    slide.delete(g)
    assert carry.timeline(slide, n.Id) == before
    assert link.ConnectorFormat.BeginConnectedShape is n
    assert link.ConnectorFormat.BeginConnectionSite == 4
    assert slide.interactive[0](1).Timing.TriggerShape is r
    assert slide.interactive[1](1).Timing.TriggerShape is n


def test_effects_a_paste_brought_along_are_removed_first(deck):
    slide, g, _, link = deck
    before = carry.timeline(slide, g.Id)
    n = slide.shape(6)
    slide.main.add(n, WIPE)  # as a paste turns an interactive effect into
    carry.copy_animations(slide, g, n)
    carry.move_links(slide, g, n, [link], [])
    assert carry.timeline(slide, n.Id, leave_out=g.Id) == before


def test_a_chart_with_no_animations_is_left_alone():
    slide = Slide()
    g, n = slide.shape(2), slide.shape(6)
    carry.copy_animations(slide, g, n)
    assert slide.picked is None


def test_a_copy_that_does_not_match_is_seen(deck):
    slide, g, _, link = deck
    before = carry.timeline(slide, g.Id)
    n = slide.shape(6)
    carry.copy_animations(slide, g, n)
    carry.move_links(slide, g, n, [link], [])
    assert carry.timeline(slide, n.Id, leave_out=g.Id) == before
    copy = slide.main.effects[2]
    assert copy.Shape is n and copy.EffectType == FADE
    copy.Timing.Duration = 0.9
    assert carry.timeline(slide, n.Id, leave_out=g.Id) != before


def test_more_or_fewer_copies_than_expected_are_refused(deck, monkeypatch):
    slide, g, _, _ = deck
    n = slide.shape(6)
    n.ApplyAnimation = lambda: None
    with pytest.raises(carry.NotCarried, match="made 0 effects, not 5"):
        carry.copy_animations(slide, g, n)


def test_an_effect_started_by_a_bookmark_is_refused_before_anything_changes(deck):
    slide, g, _, _ = deck
    slide.interactive[0](1).Timing.TriggerType = ON_BOOKMARK
    n = slide.shape(6)
    with pytest.raises(carry.NotCarried, match="something other than a click"):
        carry.copy_animations(slide, g, n)
    assert slide.picked is None


def test_links_that_cannot_all_move_are_put_back(deck):
    slide, g, r, link = deck
    n = slide.shape(6)
    ends = link.ConnectorFormat

    def refuse(shape, site):
        raise RuntimeError("cannot connect")

    ends.BeginConnect = refuse
    undo = []
    with pytest.raises(RuntimeError, match="cannot connect"):
        carry.move_links(slide, g, n, [link], undo)
    assert slide.interactive[1](1).Timing.TriggerShape is n, "moved before the connector failed"
    assert carry.put_back(undo) == []
    assert slide.interactive[1](1).Timing.TriggerShape is g, "and pointed back"


def test_put_back_says_what_did_not_go_back():
    def fails():
        raise RuntimeError("gone")

    assert carry.put_back([lambda: None, fails]) == ["gone"]


def test_animated_sees_effects_and_triggers(deck):
    slide, g, r, _ = deck
    assert carry.animated(slide, g.Id) and carry.animated(slide, r.Id)
    assert not carry.animated(slide, 99)


def test_a_chart_its_own_click_animates_is_carried():
    """Clicking G spins G: the copy comes over in a sequence of its own,
    started by clicking the copy, and G's goes when G does."""
    slide = Slide()
    g, r = slide.shape(2), slide.shape(3)
    slide.main.add(g, FADE)
    slide.new_interactive().add(g, SPIN, trigger=ON_SHAPE_CLICK, trigger_shape=g, duration=2.0)
    slide.new_interactive().add(r, APPEAR, trigger=ON_SHAPE_CLICK, trigger_shape=g)
    before = carry.timeline(slide, g.Id)
    n = slide.shape(6)

    carry.copy_animations(slide, g, n)
    carry.move_links(slide, g, n, [], [])
    assert carry.timeline(slide, n.Id, leave_out=g.Id) == before
    slide.delete(g)
    assert carry.timeline(slide, n.Id) == before


def test_the_direction_is_set_before_the_duration(deck):
    """Set after, it would put the rebuilt click effect's 1.1 s back to 0.5."""
    slide, g, _, link = deck
    n = slide.shape(6)
    carry.copy_animations(slide, g, n)
    rebuilt = [e for e in slide.interactive[0].effects if e.Shape is n]
    assert [(e.Timing.Duration, e.EffectParameters.Direction) for e in rebuilt] == [(1.1, 4)]


def test_a_rebuilt_click_effect_keeps_its_timing_and_after_effect(deck):
    slide, g, _, link = deck
    c1 = slide.interactive[0](1)
    c1.Timing.Accelerate, c1.Timing.RepeatCount, c1.Timing.RewindAtEnd = 0.5, 2, -1
    c1.EffectInformation.AfterEffect, c1.EffectInformation.Dim = 1, SimpleNamespace(RGB=0x336699)
    before = carry.timeline(slide, g.Id)
    n = slide.shape(6)
    carry.copy_animations(slide, g, n)
    carry.move_links(slide, g, n, [link], [])
    assert carry.timeline(slide, n.Id, leave_out=g.Id) == before


def test_an_after_effect_the_copy_lacks_is_seen(deck):
    slide, g, _, link = deck
    before = carry.timeline(slide, g.Id)
    n = slide.shape(6)
    carry.copy_animations(slide, g, n)
    carry.move_links(slide, g, n, [link], [])
    copy = [e for e in slide.interactive[0].effects if e.Shape is n][0]
    copy.EffectInformation.AfterEffect = 2  # hide
    assert carry.timeline(slide, n.Id, leave_out=g.Id) != before


class PathEffect(Effect):
    """A motion path effect made by AddTriggerEffect: its default path, and an
    Exit that refuses any setting, as PowerPoint's does."""

    @property
    def Exit(self):
        return 0

    @Exit.setter
    def Exit(self, value):
        if hasattr(self, "_made"):
            raise RuntimeError("Effect.Exit : failed")
        self._made = True


def test_a_click_effect_on_a_path_is_made_again_with_its_path(deck):
    slide, g, r, link = deck
    sequence = slide.interactive[0]
    sequence.effects = []
    sequence.add(g, PATH_DOWN, trigger=ON_SHAPE_CLICK, trigger_shape=r, duration=2.0,
                 path="M 0 0 L 0.2 0.1 L 0 0.25 E")

    def add_path(shape, kind, trigger, trigger_shape):
        effect = PathEffect(sequence, shape, kind, trigger=trigger, trigger_shape=trigger_shape,
                            path="M 0 0 L 0 0.25 E")
        sequence.effects.append(effect)
        return effect

    sequence.AddTriggerEffect = add_path
    before = carry.timeline(slide, g.Id)
    n = slide.shape(6)
    carry.copy_animations(slide, g, n)
    carry.move_links(slide, g, n, [link], [])
    assert carry.timeline(slide, n.Id, leave_out=g.Id) == before


def test_a_click_effect_that_cannot_be_made_says_so(deck):
    slide, g, _, _ = deck

    def refuse(*args):
        raise RuntimeError("Sequence.AddTriggerEffect : failed")

    slide.interactive[0].AddTriggerEffect = refuse
    with pytest.raises(carry.NotCarried, match="could not be made again .*AddTriggerEffect"):
        carry.copy_animations(slide, g, slide.shape(6))
