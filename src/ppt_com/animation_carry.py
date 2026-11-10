"""Hand a shape's animations, triggers and connectors to its replacement (#343).

ppt_set_chart_data swaps a chart's data by pasting a copy of it and deleting
the original (#299). Everything on the slide that points at the original by
its id would go with it, so an animated chart used to go through Excel
instead, whose embedding Excel crashes below ppt-mcp (#323, #343).

Measured on Windows (PowerPoint 16.0.20527):

- Setting ``Effect.Shape`` to the copy does not hold. The slide's build list
  still ties the effect to the original, and deleting the original deletes
  the effect too.
- ``PickupAnimation`` and ``ApplyAnimation`` copy every main sequence effect
  with its settings (timing, motion path, amount), and put the copies at the
  end of the main sequence, in their order. An interactive effect of the
  original is not copied as one: it comes after them as a click effect with
  no trigger, so it is deleted and made again with ``AddTriggerEffect``. One
  started by clicking the original itself is the exception: it comes over
  as one started by clicking the copy, in a sequence of its own.
- Setting an effect's direction puts its duration back to the default, so
  the direction is set first.
- ``Effect.MoveTo`` puts a copy back beside the effect it copies.
- ``Timing.TriggerShape`` and ``ConnectorFormat.BeginConnect``/``EndConnect``
  can be pointed at the copy, and stay pointed at it.

Whatever is done is then checked: the slide's sequences, read with the copy
standing for the original, have to be what they were. Anything short of that
raises ``NotCarried``, and the caller removes the copy, which takes every
effect made for it along, and leaves the original as it was.
"""

msoAnimTriggerOnShapeClick = 4
msoAnimTypeMotion = 1
msoAnimAfterEffectNone = 0
msoAnimAfterEffectDim = 1

# Timing settings an effect made again takes from the one it copies.
_TIMING = ("Duration", "TriggerDelayTime", "RepeatCount", "RewindAtEnd", "AutoReverse",
           "Accelerate", "Decelerate")


class NotCarried(Exception):
    """The animations could not be handed over, and nothing of the original changed."""


def _sequences(slide):
    timeline = slide.TimeLine
    interactive = timeline.InteractiveSequences
    return [timeline.MainSequence] + [interactive(i) for i in range(1, interactive.Count + 1)]


def _effects(sequence):
    return [sequence(i) for i in range(1, sequence.Count + 1)]


def _trigger_id(effect):
    try:
        trigger = effect.Timing.TriggerShape
    except Exception:  # noqa: BLE001 - an effect with no trigger shape
        return None
    return None if trigger is None else trigger.Id


def _read(read, default=None):
    try:
        return read()
    except Exception:  # noqa: BLE001 - not every effect has every setting
        return default


def effect_facts(effect) -> tuple:
    """What an effect does, as far as COM says, to compare a copy against."""
    timing = effect.Timing
    behaviors = effect.Behaviors
    moves = []
    for i in range(1, behaviors.Count + 1):
        behavior = behaviors(i)
        kind = _read(lambda: behavior.Type)
        moves.append((kind, _read(lambda: behavior.MotionEffect.Path)
                      if kind == msoAnimTypeMotion else None))
    return (
        effect.EffectType, effect.Exit, timing.TriggerType,
        round(timing.Duration, 3), round(timing.TriggerDelayTime, 3),
        _read(lambda: timing.RepeatCount), _read(lambda: timing.RewindAtEnd),
        _read(lambda: timing.AutoReverse),
        _read(lambda: round(timing.Accelerate, 3)), _read(lambda: round(timing.Decelerate, 3)),
        _read(lambda: effect.EffectInformation.BuildByLevelEffect),
        _read(lambda: effect.EffectInformation.AfterEffect),
        _read(lambda: effect.EffectInformation.Dim.RGB)
        if _read(lambda: effect.EffectInformation.AfterEffect) == msoAnimAfterEffectDim else None,
        _read(lambda: effect.EffectParameters.Direction),
        tuple(moves),
    )


def timeline(slide, chart_id, leave_out=None) -> list:
    """The slide's sequences, an effect each as (whom, trigger, facts), with
    the shape ``chart_id`` named "chart". Effects on ``leave_out`` are skipped,
    and so is a sequence left with none."""
    def who(shape_id):
        return "chart" if shape_id == chart_id else shape_id

    out = []
    for sequence in _sequences(slide):
        entries = []
        for effect in _effects(sequence):
            shape_id = effect.Shape.Id
            if leave_out is not None and shape_id == leave_out:
                continue
            entries.append((who(shape_id), who(_trigger_id(effect)), effect_facts(effect)))
        out.append(entries)
    # The main sequence in its order; the interactive ones, each started by
    # its own click, in any. A copy's effects started by clicking the copy
    # come in a sequence of their own, after the original's, which is left
    # empty here and goes when the original does.
    return [out[0]] + sorted((entries for entries in out[1:] if entries), key=repr)


def animated(slide, shape_id) -> bool:
    """Whether anything on the slide's timeline is about the shape."""
    return any(effect.Shape.Id == shape_id or _trigger_id(effect) == shape_id
               for sequence in _sequences(slide) for effect in _effects(sequence))


def _remove_effects_of(sequence, shape_id) -> None:
    for i in range(sequence.Count, 0, -1):
        if sequence(i).Shape.Id == shape_id:
            sequence(i).Delete()


def _indexes(sequence, shape_id) -> list:
    return [i for i in range(1, sequence.Count + 1) if sequence(i).Shape.Id == shape_id]


def copy_animations(slide, old, new) -> None:
    """Give ``new`` the effects ``old`` has, each beside the one it copies."""
    old_id, new_id = old.Id, new.Id
    sequences = _sequences(slide)
    main = sequences[0]
    # A paste can bring effects of its own along, the interactive ones turned
    # into click effects; they are not the ones made below.
    for sequence in sequences:
        _remove_effects_of(sequence, new_id)

    own_main = len(_indexes(main, old_id))
    own_interactive = []  # started by another shape: made again below
    own_self = 0  # started by clicking the chart itself: copied as they are
    for number, sequence in enumerate(sequences[1:], start=2):
        for index in _indexes(sequence, old_id):
            effect = sequence(index)
            trigger = _trigger_id(effect)
            if effect.Timing.TriggerType != msoAnimTriggerOnShapeClick or trigger is None:
                raise NotCarried("one of its animations starts on something other than a click on a shape")
            if trigger == old_id:
                own_self += 1
            else:
                own_interactive.append((number, index))
    if not own_main and not own_interactive and not own_self:
        return

    old.PickupAnimation()
    new.ApplyAnimation()
    # An effect started by clicking the chart itself comes over as one
    # started by clicking the copy, in a sequence of its own; one started
    # by another shape comes into the main sequence as a click effect.
    added = _indexes(main, new_id)
    if len(added) != own_main + len(own_interactive):
        raise NotCarried(f"copying its animations made {len(added)} effects, "
                         f"not {own_main + len(own_interactive)}")
    copied_self = sum(len(_indexes(sequence, new_id)) for sequence in _sequences(slide)[1:])
    if copied_self != own_self:
        raise NotCarried(f"copying its animations made {copied_self} effects started by "
                         f"clicking it, not {own_self}")
    # The interactive ones come last, as click effects with no trigger.
    for index in reversed(added[own_main:]):
        main(index).Delete()
    for k in range(own_main):
        # Each copy just after the effect it copies. The copies are at the
        # end, so moving one up never moves an original past it.
        originals, copies = _indexes(main, old_id), _indexes(main, new_id)
        if copies[k] != originals[k] + 1:
            main(copies[k]).MoveTo(originals[k] + 1)

    for number, _ in own_interactive:
        sequence = _sequences(slide)[number - 1]
        originals = _indexes(sequence, old_id)
        made = len(_indexes(sequence, new_id))
        original = sequence(originals[made])
        try:
            _make_again(sequence, new, original)
        except Exception as e:
            raise NotCarried(f"one of its click animations could not be made again ({e})") from e
        index = _indexes(sequence, new_id)[-1]
        target = _indexes(sequence, old_id)[made] + 1
        if index != target:
            sequence(index).MoveTo(target)


def _make_again(sequence, new, original) -> None:
    """Add to ``sequence`` an effect on ``new`` started as ``original`` is,
    with its settings. What does not take shows in the timeline comparison."""
    effect = sequence.AddTriggerEffect(new, original.EffectType, msoAnimTriggerOnShapeClick,
                                       original.Timing.TriggerShape)
    # The direction first: setting it puts the duration back to the effect's
    # default.
    direction = _read(lambda: original.EffectParameters.Direction)
    if direction is not None:
        _read(lambda: setattr(effect.EffectParameters, "Direction", direction))
    # Only when it differs: a motion path refuses even the value it has.
    if effect.Exit != original.Exit:
        _read(lambda: setattr(effect, "Exit", original.Exit))
    for name in _TIMING:
        value = _read(lambda name=name: getattr(original.Timing, name))
        if value is not None:
            _read(lambda name=name, value=value: setattr(effect.Timing, name, value))
    # A motion path is made with the effect's default path.
    behaviors, made = original.Behaviors, effect.Behaviors
    for i in range(1, min(behaviors.Count, made.Count) + 1):
        if _read(lambda i=i: behaviors(i).Type) == msoAnimTypeMotion:
            path = _read(lambda i=i: behaviors(i).MotionEffect.Path)
            if path is not None:
                _read(lambda i=i, path=path: setattr(made(i).MotionEffect, "Path", path))
    after = _read(lambda: original.EffectInformation.AfterEffect, msoAnimAfterEffectNone)
    if after != msoAnimAfterEffectNone:
        # Set through the sequence, which answers the effect it made.
        if after == msoAnimAfterEffectDim:
            sequence.ConvertToAfterEffect(effect, after, original.EffectInformation.Dim.RGB)
        else:
            sequence.ConvertToAfterEffect(effect, after)


def move_links(slide, old, new, connectors, undo) -> None:
    """Point what starts on ``old`` and the connectors on it at ``new``.

    Each step that is taken adds to ``undo`` how to take it back, so a swap
    given up afterwards, or halfway through this, can point them back with
    ``put_back``.
    """
    old_id = old.Id
    for sequence in _sequences(slide):
        for effect in _effects(sequence):
            if _trigger_id(effect) == old_id:
                timing = effect.Timing
                timing.TriggerShape = new
                undo.append(lambda timing=timing: setattr(timing, "TriggerShape", old))
    for connector in connectors:
        ends = connector.ConnectorFormat
        for end in ("Begin", "End"):
            if (getattr(ends, f"{end}Connected") == -1
                    and getattr(ends, f"{end}ConnectedShape").Id == old_id):
                site = getattr(ends, f"{end}ConnectionSite")
                getattr(ends, f"{end}Connect")(new, site)
                undo.append(lambda ends=ends, end=end, site=site:
                            getattr(ends, f"{end}Connect")(old, site))
                if getattr(ends, f"{end}ConnectedShape").Id != new.Id:
                    raise NotCarried("a connector on it could not be moved to the copy")


def put_back(undo) -> list:
    """Undo ``move_links``; answers what could not be put back."""
    failed = []
    for step in reversed(undo):
        try:
            step()
        except Exception as e:  # noqa: BLE001 - reported by the caller
            failed.append(str(e))
    return failed
