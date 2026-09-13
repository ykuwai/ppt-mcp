# Bringing ppt-mcp to macOS

Status: design study, now partly built. The measurements below stand; where
building it changed the answer, the section says so.

What has shipped so far, on the `macos-support` branch. The package installs
and 563 tests run on macOS (#185), the Apple Event backend and the generated
enumeration tables (#186), and the app and export tools (#187).

The one place building it beat the study: **slide images do work**, just not the
way Windows does them. Section 6 said there is no export command, which is true,
and section 5 said whole deck PNG export reports success and writes nothing,
which is also true. The route that does work is to export the deck to PDF and
render the pages with Quartz, which ships with macOS. That returns a vector
rendering at any size asked for rather than a screen capture, and it is fast:
0.18s for the PDF, 0.10s per slide after that, 0.08s for a preview.

ppt-mcp drives a live PowerPoint through Windows COM. macOS has no COM, but
PowerPoint for Mac ships an Apple Event object model that is the same object
model wearing different names. This document records what was measured on a
real machine, where the two platforms genuinely diverge, and how to carry the
port without turning the codebase into two codebases.

Measured on macOS 26.6.2 (25G83, arm64), Microsoft PowerPoint 16.97.2,
Python 3.14.5. Every timing and every pass or fail below came from running the
thing, not from reading documentation. Where community documentation and the
live machine disagreed, the machine won and the disagreement is noted.

---

## 1. The short version

**It works, and it is fast.** The full `ppt_add_shape` path (create the shape,
solid fill, fill colour, line weight, line colour, text, Latin font, East Asian
font, size, bold, text colour, paragraph alignment, vertical anchor, corner
radius) is **15 Apple Events in 0.081 seconds**. Reading every property of every
shape on a slide is one event, 0.011 seconds.

The bridge is [appscript](https://github.com/hhas/appscript) with one argument
that appears in no documentation.

```python
from appscript import app, k, mactypes
PP = app(id='com.microsoft.Powerpoint', terms='sdef')
```

Without `terms='sdef'` appscript asks PowerPoint for terminology the old AETE
way, gets nothing, silently falls back to 14 built-in words, and every attribute
raises `AttributeError: Unknown property, element or command: 'presentations'`.
That single missing argument is very likely why the received wisdom says Python
cannot drive Office on a Mac. It can.

appscript 1.4.0 (October 2025) publishes universal2 wheels for CPython 3.10
through 3.14, so there is no build step and no compiler on the user's machine.

Three things are absent from the dictionary. Charts, SmartArt and freeform
path building have no words in it at all. Two of them come back another way:
the clipboard carries a DrawingML package for any shape, PowerPoint pastes one
back whatever it holds, and every chart, freeform and group tool goes through
it (`docs/gvml-design.md`). Creating and reading is a paste or a copy; editing
an existing chart or path is a copy, a rewrite and a paste back, after which
the original is deleted and the new shape put back in its place. That makes
it a new object with the old name, and the tool says so. Section 6 says what
is still missing.

One risk outranks every missing feature, and it is section 5.

---

## 2. How the two object models line up

PowerPoint's Apple Event dictionary is the VBA object model with spaces in the
names. Same tree, same verbs, different spelling. This is why a port is
tractable at all.

| Windows COM | macOS Apple Event | Note |
|---|---|---|
| `Application.ActivePresentation` | `active presentation` | raises -1728 when the start gallery is frontmost |
| `Presentations(i)` | `presentations[i]` | 1 based on both sides |
| `Presentation.Slides(i)` | `presentation.slides[i]` | |
| `Slide.Shapes(i)` | `slide.shapes[i]` | |
| `Shapes.AddShape(Type, L, T, W, H)` | `make new shape at end of slide N with properties {…}` | see 2.1 |
| `Shape.Left` | `left position` | never `top position` for the other one |
| `Shape.Top` | `top` | |
| `Shape.TextFrame.TextRange.Text` | `text frame`'s `text range`'s `content` | |
| `Font.Name` | `font name`, reached through `font of text range` | going via `text range` directly gives -10006 |
| `Font.NameFarEast` | `east asian name` | |
| `Font.Size` | `font size` | |
| `Font.Color.RGB` | `font color` | value shape differs, see 2.2 |
| `Shape.Fill` | `fill format` | |
| `Fill.ForeColor.RGB` | `fill format`'s `fore color` | |
| `Fill.Solid()` | `solid fill format of …` | a command, not a method |
| `Fill.TwoColorGradient(s, v)` | `two color gradient … style … variant …` | |
| `Shape.Fill.UserPicture(path)` | `user picture (fill format of …) picture file …` | verified, gives `fill picture` |
| `Shape.Line` | `line format` | |
| `Shape.Adjustments(1)` | `adjustment 1`'s `adjustment_value` | **not** `value` |
| `Shape.ZOrder(cmd)` | `z order … z order position …` | |
| `ParagraphFormat.Alignment` | `paragraph format`'s `alignment` | |
| `TextFrame.VerticalAnchor` | `text frame`'s `vertical anchor` | |
| `Window.View.GotoSlide(n)` | `document window 1`'s `view`'s `go to slide number n` | the nicety survives, 4 ms |
| `Presentation.SlideMaster` | `slide master`, whose class is `master` | |
| `Slide.Shapes.Placeholders(i)` | `place holder i`, two words | |
| `PageSetup.SlideWidth` | `page setup`'s `slide width` | |
| `PageSetup.SlideHeight` | **missing** | use `slide master`'s `height` |

### 2.1 The insertion location trap

The natural reading is wrong, and the error does not say so.

```applescript
-- fails, "cannot create class slide" (-2710)
make new slide at end of slides of active presentation
-- fails, parameter error (-50)
make new shape at end of shapes of slide 1 of pres

-- correct
make new slide at end of active presentation with properties {layout:slide layout blank}
make new shape at end of slide 1 of pres with properties {auto shape type:autoshape rectangle, …}
```

Through appscript the same rule applies. `at=pres.end` works, `at=pres.slides.end`
raises -1708. Write it down once in the driver and never think about it again.

### 2.2 Colours

Windows COM packs a colour into one BGR integer, `R + G*256 + B*65536`. The Mac
side takes and returns a three element list of 0 to 255 integers. Passing 65535
gets clamped to 255, so the scale really is 8 bit.

`utils/color.py` already funnels everything through `hex_to_int` and
`int_to_hex`, so the whole conversion is two new functions next to them and one
backend level choice of which pair to use. The MCP surface keeps taking
`#RRGGBB` and nothing user facing changes.

### 2.3 Enumerations

`constants.py` is 866 lines of `mso*`, `pp*` and `xl*` integers. The Mac side
does not take integers, it takes named enumerators, and the names are its own.

| Windows constant | macOS enumerator |
|---|---|
| `msoShapeRectangle` = 1 | `autoshape rectangle` |
| `msoShapeRoundedRectangle` = 5 | `autoshape rounded rectangle` |
| `ppLayoutBlank` = 12 | `slide layout blank` |
| `msoBringToFront` = 0 | `bring shape to front` |
| `ppAlignCenter` = 2 | `paragraph align center` |
| `msoAnchorMiddle` = 3 | `anchor middle` |
| `msoThemeColorAccent1` = 5 | `first accent theme color` |
| `ppSaveAsPNG` = 18 | `save as PNG` |

Note the prefixes. It is `autoshape rounded rectangle`, not `rounded rectangle`.
184 auto shape types are available, which is more than `SHAPE_NAME_MAP` exposes
today, so nothing is lost here. That was written as an expectation and it took a
second pass to make true. The generator reads named constants out of banner
sections in `constants.py`, and three of the public vocabularies do not live
there at all, they are friendly name maps in `animation.py` and `shapes.py` with
no named constants behind them. So the generator had nothing to compare and its
unmatched list could not mention them, while 57 of the 125 names those three
tools list as valid failed on macOS with a message saying the words did not
exist. It reads those maps now. The work is mechanical but it is the single
largest translation in the project, and it has to be right because these names
are the public MCP vocabulary's only anchor.

---

## 3. Where the bridge itself misbehaves

These are appscript and PowerPoint quirks rather than design problems, but each
one produces a wrong answer rather than an error, so the driver absorbs all of
them in one place.

| Symptom | Cause | What the driver does |
|---|---|---|
| `ref.count()` raises -1708 | PowerPoint's dictionary declares no Standard Suite commands at all, only the `window` class | `len(ref.get())` |
| `ref.get()` raises -1728 on an empty collection | PowerPoint errors instead of returning an empty list | catch and return 0 |
| `PP.open(...)` returns `None` and opens nothing | same missing Standard Suite | shell out to `osascript` with `open POSIX file "…"`, which does work, verified |
| `ref.duplicate()` raises `unpack requires a buffer of 4 bytes` | appscript falls back to a code PowerPoint does not answer | same osascript fallback |
| `shape.rotation` raises `AttributeError` | the dictionary defines both an enumerator and a property named `rotation`, and appscript resolves the collision toward the enumerator | reach the property by raw code. Verified working, set 33 and read 33.0 back |
| `line_format.dash_style` raises `AttributeError` | same collision class, the `line dash style …` enumerators take the name | raw code `b'LFds'` |
| PowerPoint stops answering | a modal sheet, or an operation that wedged it | every call carries `timeout=`, which appscript supports on every command |

The raw code route is small enough to state in full.

```python
from appscript.reference import Reference
def raw(ref, code):
    return Reference(ref.AS_appdata, ref.AS_aemreference.property(code))

raw(shape, b'ShRt').set(33)                 # rotation
raw(shape.line_format, b'LFds').get()       # dash style
```

The last row deserves emphasis. PowerPoint wedged repeatedly during this study,
several times past 12 seconds and three times fatally with -609. The Windows
wrapper already has the right shape for this in `com_wrapper.py:107-135` (retry
on busy, optional ESC to dismiss a dialog); the Mac side needs the same policy
with a different trigger.

`active presentation` raising -1728 whenever the start gallery is frontmost is
worth calling out separately, because it is the Mac echo of a problem the
project already solved. `_target_pres_full_name` is exactly the right design
here too, and `presentations[1]` is the right fallback.

---

## 4. What is confirmed working

Every line below was executed. Timings are wall clock including the Apple Event
round trip.

| Capability | Result |
|---|---|
| create a presentation | 0.705 s |
| create a slide with an explicit layout | 0.123 s |
| **`ppt_add_shape` equivalent, 15 chained operations** | **0.081 s** |
| `ppt_add_textbox` equivalent, paragraphs split on CR | 0.006 s |
| insert a picture from a path, verified as `shape type picture` at the requested size | 0.021 s |
| picture as a shape fill, verified as `fill picture` | 0.17 s |
| navigate the window to the slide being edited | 0.004 s |
| read every property of every shape on a slide, one event | 0.011 s |
| read names and geometry of every shape in bulk | 0.007 s |
| placeholders, their type and their text | works |
| custom layouts of the slide master, 11 found | works |
| create a table with `make new shape table`, addressed with `get cell from` | works |
| slide dimensions | width from `page setup`, height from `slide master` |
| sections count through `section properties` | works |
| **export the deck to PDF, from the container** | **0.19 s, file verified** |
| **export one shape to PNG, into the container** | **0.15 s, 28 KB verified** |
| **insert a movie or a sound from a path, verified as `shape type media`** | **`make new media2 object`, embedded in `ppt/media/` at the file's own byte size** |
| a chart already in the deck, seen as a shape | name, type and geometry all readable |
| **a chart, a freeform or a group, pasted from a DrawingML package we wrote** | **`paste object` takes `com.microsoft.Art--GVML-ClipFormat` in 5 to 65 ms; the shape reports its own type; the clipboard is put back afterwards** |
| **a chart's numbers or a freeform's points, read from the package `copy shape` writes** | **`chart1.xml` carries the caches with no workbook; `a:custGeom` carries every point** |

Per operation cost on the live object model settles around **0.5 ms for a simple
read** and **2 to 3 ms for a four level chained write**.

The comparison that matters for the architecture. Twenty shapes created through
twenty separate `osascript` invocations cost 178 ms each. The same twenty inside
one batched `tell` block cost 9 ms each. Speaking Apple Events from inside the
Python process, as appscript does, removes that fixed cost entirely. The driver
must not shell out per operation.

---

## 5. The risk that outranks every missing feature

**PowerPoint reports success and does nothing.** For an MCP server this is worse
than an unsupported tool, because the model believes the deck now contains
something it does not, and keeps building on top of the belief.

Four structurally different cases, all observed.

| Operation | Reported | Actually |
|---|---|---|
| `save <pres> in <path> as <fmt>` on a deck that has **no file path yet** | no error, or a 40 second hang | nothing written. Seen to succeed once at 22.8 s and to hang past 40 s another time |
| `save <pres> in <dir> as save as PNG` on a deck **with** a path, into the container | returned `ok` in 0.15 s | **no folder, no files** |
| `save as picture` to a path outside what the sandbox allows | no error | nothing written |
| `make new shape` given a `file name` | no error | an empty 25 by 25 autoshape, and a nonexistent path also reports success |

So the rule for the port is that **no mutation is trusted because it did not
raise**. Verify.

- After any save or export, stat the file and check the size.
- After setting a picture fill, read back `fill format type` and require
  `fill picture`.
- After a save that is meant to change the document's identity, read back
  `full name`.
- After creating a shape, count shapes before and after, and reject a result
  that comes back as `shape type auto` at 25 by 25.

Every one of these is cheap, and every one of them is detectable.

### 5.1 One way of asking takes PowerPoint down

Found by running the port, not by reading the dictionary, and the sharpest
lesson in it. **Asking a slide's animation timeline for its whole effects
collection kills PowerPoint.**

```python
slide.timeline.main_sequence.effects.get()     # -609, the application dies,
                                               # and the open deck with it
```

It happens on a slide with no animations at all, reproducibly. But it is a
problem with the way of asking, not with the timeline, and everything else in
that area is fine.

```python
slide.timeline.main_sequence.effects[1].shape.name()   # works
slide.timeline.sequences[1].effects[1]                 # works
count(slide.timeline.sequences)                        # works
shape.animation_settings.animate()                     # works, and is settable
```

So the rule is to never materialise that collection. There is a cheap safe way
to count it and a slow safe one, and the difference between the cheap one and
the crasher is one word.

```python
seq.effects.count()          # -609, the application dies. Same as .get()
seq.count(each=k.effect)     # correct, one round trip, safe
probe_count(seq.effects)     # correct, N round trips, safe
```

Asking the effects collection how many elements it has kills PowerPoint. Asking
the sequence how many effects it holds does not. `backend.mac_ae.count_of` sends
the second, and `probe_count` walks the collection one element at a time for
anything that will not answer it. And where the question is only whether one
shape is animated, the older per shape API answers instantly and is what
`ppt_get_shape_info` uses.

There is a second one, found later and just as fatal. **Asking PowerPoint to
make a table column before an existing one kills it**, on a plain three by three
table, reproduced twice on its own. The row equivalent only answers -1708, so
inserting a row is merely impossible rather than dangerous; both are refused
now, and appending works.

```python
app.make(new=k.column, at=table.columns[2].before)   # -609, the application dies
app.make(new=k.row, at=table.rows[2].before)         # -1708, harmless
app.make(new=k.column, at=table.end)                 # works
```

Behind all three is one habit. **A reference PowerPoint hands back is not to be
trusted.** Asking for the elements of a collection is unreliable in general,
and a slide holding a text box and an autoshape answers `text_boxes[1]` and
`shapes[2]`, and the second does not resolve. So is the reference a command
returns: `get cell from` answers with something that renders as
`rows[1].cells[1]` and then -1728 on every cell of a fresh table, while the
identical path built here reads and writes all nine. So is `make`'s return
value, for a shape, a row or a column alike.

The rule the port follows is to count, then index, and to build every reference
itself. `positional()` does that for shapes, `probe_count()` for effects, where
even counting has to be done one at a time. Animation effects and table columns
are the two places where getting it wrong costs the document rather than an
error message.

### 5.2 Writes that quietly rewrite the whole slide

The same shape of danger as 5.1, without the crash, and harder to notice because
the call succeeds and returns nothing.

`shape.animation settings` is the old per shape animation API. It can only hold
one entrance per shape, and writing to any of it appears to force the slide back
into that model. Three shapes with one entrance each:

```python
shape1.animation_settings.dim_color.set([255, 0, 0])
# all three effects are now a plain appear

shape2.animation_settings.animate.set(False)
# shape 2's effect is gone, as asked, and shape 1's is now an appear

# and with an exit animation on shape 3, that one disappears outright
```

`text unit effect` and `animate text in reverse` behave the same way.
`animate background` is the only one of the five that leaves the slide alone.

So `ppt_add_animation` and `ppt_update_animation` never write four of them and
say so in `warnings`, and `ppt_remove_animation` refuses outright, because
deleting one effect answers -50 and changes nothing, and clearing its shape
costs the rest of the slide. `ppt_clear_animations` is the exception that is safe, because it is
emptying the slide anyway. It was checked against a slide holding two effects on
one shape and two exit animations, and it left nothing behind.

The same container holds `play settings`, and the two media playback settings
on it turned out not to be alike, which is worth knowing before assuming that
everything hanging off `animation settings` is equally dangerous. Measured on a
slide holding a fly in, a bounce and an exit fade:

```python
media.animation_settings.animation_play_settings.loop_until_stopped.set(True)
# all three effects untouched, entrance types and exit flag included,
# through three writes

media.animation_settings.animation_play_settings.hide_while_not_playing.set(True)
# the exit fade is gone outright and the bounce is now a plain appear.
# Writing False over a value that was already False did it too
```

So `ppt_set_media_settings` writes `loop` freely and writes
`hide_while_not_playing` only on a slide whose main sequence is empty, refusing
that one argument by name anywhere else.

Counting the timeline's own `sequence` elements as a second gate looked free and
is not, which is worth writing down before someone tries it again.

```python
count_of(slide.timeline, k.sequence)   # 0 on a fresh slide, 0 after a text box,
                                       # 0 with one effect in the main sequence,
                                       # and 1 as soon as a movie is inserted
count(slide.timeline.sequences)        # 0 through all four, so the two disagree
```

A media shape brings a sequence of its own for its playback, so a gate that
counted those would refuse the write on every slide the tool is ever called
about. The main sequence is the only count that means what it looks like, and
an animation triggered by clicking a shape is the case it cannot see, which the
successful write says in `warnings`.

### 5.3 The sandbox, and where exports have to go

PowerPoint for Mac is sandboxed. It carries `com.apple.security.app-sandbox` and
`files.user-selected.read-write`, and it has no entitlement for Desktop,
Documents or Downloads.

The consequences were reproduced three separate times. Exporting to a directory
PowerPoint has not written to before blocks for tens of seconds and then kills
the application with **-609**. Exporting to a directory it already has a grant
for completes in under 0.2 seconds.

The reliable answer, and the one Microsoft's own guidance points at, is to stage
through the application's container.

```
~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/
```

Verified there, with the deck itself opened from that folder. PDF export works
(14 KB in 0.19 s) and per shape PNG export works (28 KB in 0.15 s). Whole deck
PNG export is the silent no-op in the table above and must not be relied on, so
slide previews should come from the PDF or from per shape PNG rather than from
`save as save as PNG`.

An unsandboxed Python process can read and write that folder freely, so the
pattern is export into the container, then move the file out with Python.

One more path trap. An HFS colon path is treated as a **literal filename**,
producing a file called `Macintosh HD:Users:…png` inside the container root.
Always use POSIX paths.

### 5.4 One request at a time, and what a caller in the queue used to get back

Everything reaches PowerPoint through a single worker thread, because several
Apple Events in flight at once are not reliable. A caller that fires several
tools in the same turn therefore puts them in a queue.

That queue was where three separate complaints came from, all one bug. The
caller's clock started when the work was **queued**, not when it **ran**, so a
call waiting behind two slow ones blew its budget and reported failure. The
failure it reported was `concurrent.futures.TimeoutError`, whose `str()` is the
empty string, so it arrived as `Failed to add picture: ` with nothing after it.
And the job was still in the queue, so the worker ran it a minute later anyway
and the picture appeared on the slide. A caller who believed the error and tried
again got two.

Fixed by splitting the wait in two. The caller waits for its turn, and that wait
is its to abandon: a job taken back before it begins is skipped by the worker
rather than run late. Only once the work starts does the per call budget begin.
Both waits now end in a sentence rather than in an empty string.

What this does **not** fix is ordering. Parallel tool calls arrive in whatever
order the transport hands them over, so "add a slide" and "read slide 4" sent
together can still run the wrong way round, and the read answers honestly about
a deck that has not grown yet. Nothing inside the server can put that right, so
the instructions say to call these tools one after another on macOS.

### 5.5 PowerPoint cannot read an SVG, and does not say so

`ppt_add_svg_icon` handed PowerPoint the SVG it downloaded. PowerPoint answered
that the picture was made and left a 25 by 25 empty box on the slide instead,
which is the silent no-op of section 5 wearing a different hat. Every icon
failed and each failure left litter behind.

`sips` rasterises the SVG first, and everything it needs was measured rather
than assumed. It honours `viewBox` and renders at the size asked for rather than
scaling up the 48 by 48 the file declares, it keeps the alpha channel so the
icon sits on any background, and it keeps the fill colour substituted in for
`currentColor`. SVG support in `sips` arrived with macOS 13, so it is probed
once and the tool refuses with that reason where it is absent.

`_place_picture` now also deletes the empty box before it reports the failure.

### 5.6 Automation consent

Two independent gates exist and neither substitutes for the other. Automation
consent (TCC) governs the calling process talking to PowerPoint. The App Sandbox
governs PowerPoint talking to the filesystem. Full Disk Access on the caller
does nothing for Apple Events.

Consent is attributed to the **responsible parent process**, not to `osascript`
or to Python. In practice that is the terminal or editor that launched the MCP
server, so the same server behaves differently under Terminal, iTerm, VS Code
and Claude Desktop, and each needs its own grant. Refusal gives **-1743**, and
the prompt cannot be answered headless or over ssh. The server has to recognise
-1743 and say something useful rather than reporting a generic failure.

---

## 6. What is genuinely missing

This is the honest part. These are not workarounds waiting to be found.

| Area | Windows | macOS | Lines affected |
|---|---|---|---|
| Charts | full `Chart` object model, drives a live Excel for the data sheet | **no `chart` class**. A chart is added by writing its XML and pasting it through the clipboard, and its data is read the same way (`docs/gvml-design.md`). An existing chart is edited by copying it, rewriting its `chart1.xml` and pasting it back over the original, which keeps name, position and z order and loses animations, said in `warnings` | `charts.py` 1,126 |
| SmartArt | `SmartArt` object model | **no class, no command**. `shape type smartart graphic` exists, so an existing graphic is a shape like any other | `smartart.py` 810 |
| Freeform paths | `Shapes.BuildFreeform` | **no builder**. A path is written as `a:custGeom` and pasted through the clipboard, and read back the same way. Nodes are moved, inserted, deleted and switched between line and curve by rewriting the path and pasting it back over the original. The editing type (corner, smooth, symmetric) is not in the XML either and stays refused | `freeform.py` 765 |
| Grouping | `ShapeRange.Group` | **no `select` command, so no shape range**. The members are copied off the slide one by one, wrapped in one `a:grpSp` and pasted back as a group; the originals are deleted only once the group is verified | `groups.py` |
| Slide image export | `Slide.Export(path, "PNG")` | no `export` command exists in the dictionary at all. Solved another way, by exporting the deck to PDF and rendering pages with Quartz, which is what shipped | `export.py` 848 |
| Line visibility | `Shape.Line.Visible = False` | `line format` has **no `visible` property**. Weight 0 and transparency 1.0 both apply cleanly and are the practical stand-ins | every tool taking `line_visible` |
| Screen redraw suppression | `LockWindowUpdate` on `PPTFrameClass` | no equivalent. Less needed, because a whole tool call is 80 ms rather than a visible sequence, but the flicker fix from #164 does not transfer | `utils/redraw.py` 91 |
| Shape naming | `Rectangle 1` | `Shape_0` | anything addressing shapes by name across platforms |

The dictionary has gained no class and no command in ten years, and it lost
ground in one place: `paragraph format` no longer carries `first line indent`,
`left indent` or `right indent`. `indent level` did not disappear, it moved to
`text range`, and tab stops are still there as a `tab stop` class reached
through a text style's `ruler`.

### 6.1 Every tool that refuses, by name

The table above is by area. This is the list a user actually wants, and it is
checked against the code by a test, so it cannot quietly go stale.

**155 tools. 143 do the job. 12 always refuse.**

| Why | Tools |
|---|---|
| No `nodes`, and a node's editing type (corner, smooth, symmetric) is not stored in the XML the clipboard carries either; PowerPoint reads it off the handle geometry. Moving the handles with the node position tool is the way to change it | `ppt_set_node_editing_type` |
| No `smart art` class | `ppt_add_smartart`, `ppt_modify_smartart`, `ppt_list_smartart_options` |
| No `select` command, so no shape range | `ppt_select_shapes` |
| No tags anywhere in the dictionary | `ppt_set_tag`, `ppt_get_tags` |
| Deleting one effect answers -50 and clearing its shape costs the slide | `ppt_remove_animation`, `ppt_copy_animation` |
| No table style, only the text direction | `ppt_set_table_style` |
| No ExecuteMso and no StartNewUndoEntry | `ppt_execute_mso`, `ppt_start_undo_entry` |

A further **18 tools work and refuse one argument**, with `error` naming
the argument rather than the tool, so dropping it and calling again works.
`ppt_add_hyperlink` cannot take a `screen_tip`, `ppt_add_table_row` and
`ppt_add_table_column` cannot insert at a `position`, `ppt_add_animation`
cannot take a `trigger_shape`, `ppt_set_reflection` cannot take the four
numeric arguments, `ppt_add_video` and `ppt_add_audio` cannot take
`link_to_file`, `ppt_set_media_settings` cannot take volume, mute, trim or
fade, `ppt_format_chart` cannot take `chart_style`, `legend_font_size`, the
legend and title coordinates or an 8-direction `legend_position` (they are
computed from the chart's rendered size, which the XML does not carry),
`ppt_format_chart_axis` cannot take `tick_label_font_size`, and the rest are
checks that report a write which did not land rather than a capability that
is missing.

The nine editors that go through the clipboard, five for charts and four for
freeform nodes, succeed with a `warnings` line every time, because what they
did is not what the name says: the shape was copied, rewritten and pasted
back, the original deleted, and the new shape given the old name, position
and z order. Animations do not come with it, and the count of the effects
that were lost is in the same list. The first such warning in a server's life
is a paragraph; the rest are one line.

Everything else that differs comes back in `warnings` beside a success, which
is where to look for the smaller gaps: a glow with no transparency, a line
hidden by weight and transparency rather than by a flag, a shadow whose state
PowerPoint will not read back.

### 6.2 The VBA escape hatch, and its ceiling

PowerPoint for Mac supports VBA, and `run VB macro` is in the dictionary.

```applescript
run VB macro macro name "MyMacro" list of parameters {"a", "b"}
```

Reading the type libraries inside `Microsoft PowerPoint.app` shows
`BuildFreeform`, `AddNodes`, `Vertices`, `Export`, `AddPicture`, `AddTextbox`,
`AddTable`, `AddSmartArt` and `AddSection` present, and `AddChart` and
`AddChart2` **absent**. That absence has a clean control, because `Excel.tlb`
shipped in the same bundle and read the same way does yield `AddChart2`.

Treat that absence carefully. Symbols are stored with type signature suffixes
(`AddPictureWW`, `AddNodesP-`), so **presence is reliable evidence and absence is
not**. A first pass with exact string matching wrongly reported `AddTable` and
`BuildFreeform` missing for exactly that reason, and `AddShape` does not appear
either. So the honest statement is that `AddChart` was **not found by a method
whose negatives are unreliable**, which makes charts through VBA unlikely rather
than settled.

What is solid is the other direction. SmartArt, freeform paths and slide image
export are all present in the library and become recoverable if `run VB macro`
can be made to execute at all.

Whether it executes is unproven. With no macro present it returns -18, which is
equally consistent with "macro not found" and "handler is a stub". One line
closes the question. Put `Sub Ping()` in a `.pptm`, open it, and call it.

Three constraints to weigh before building on it. Macro security consent is a
real user facing cost, and an MCP server that asks people to lower it is a hard
sell. `run VB macro` takes only a list of text and returns only an integer, so
anything richer has to come back out of band. And it can only run a macro that
is already in the file. The old `do visual basic`, which compiled a string, is
gone from the dictionary entirely, so there is no route that does not involve
shipping a `.pptm` with the macro already in it.

---

## 7. Design

The constraint that matters most is the author's. Windows and macOS should feel
the same, and the code must not fork into two projects.

### 7.1 The seam is not where it looks

Every one of the 27 modules under `src/ppt_com/` imports `ppt` from
`utils/com_wrapper.py`, which looks like one clean seam. It is not. `ppt.execute`
is a **thread and lifecycle seam only**. It hands a closure a raw COM object,
and each of the roughly 150 `_impl` functions then walks the live object graph
itself. There are 932 distinct member access chains and 1,934 references to
them. One lifecycle seam, zero object model seam.

So the real choice is where to cut, and the answer is at the object model.

```
src/
  backend/
    __init__.py      picks by sys.platform, exposes one `ppt` object
    base.py          the accessor protocol
    win_com.py       today's com_wrapper, behaviour unchanged
    mac_ae.py        appscript, terms='sdef'
    names.py         COM member <-> Apple Event term
    enums.py         numeric constant <-> named enumerator
  ppt_com/           unchanged name for now, no longer Windows only
```

The Windows implementation is a pass through to pywin32, so it costs almost
nothing and changes almost nothing. Cutting here rather than at the tool level
is what keeps everything above the object model single sourced, and there is a
lot of it, including the pseudo Markdown exporter (244 lines), the typography
checker (180 lines), the icon search and SVG insertion, the batch dispatcher, every
layout helper, all of `color.py` and `units.py`, and all the pydantic input
models. Cutting at the tool level would duplicate every one of them.

### 7.2 Keep the niceties

The user visible care in this codebase is the point of it, and most of it
survives.

- **Show the slide you are editing.** `goto_slide` maps directly and costs 4 ms.
  It stays in `utils/navigation.py` with a backend call behind it.
- **Freeze the redraw while building a shape.** No Mac equivalent. `FrozenRedraw`
  already degrades to a no-op when win32 is missing, so the Mac path needs no new
  code, only an honest note that the flicker fix is Windows only.
- **Do not launch PowerPoint on a read.** `allow_launch=False` is platform
  neutral and stays as it is.
- **Lock to one presentation.** `_target_pres_full_name` matters more on the Mac,
  because `active presentation` fails in a state users hit every day.

### 7.3 Degrade honestly, do not pretend

For the missing areas the wrong answer is a silent no-op and the second wrong
answer is hiding the tools, because then the model cannot see the capability
exists and burns turns looking for it. The tool stays listed and returns a
structured refusal naming the platform, the reason and a route to take instead.

```json
{"error": "ppt_add_smartart is not available on macOS",
 "reason": "PowerPoint for Mac's Apple Event dictionary has no `smart art` class and no command that makes one. ...",
 "alternatives": ["ppt_add_shape and ppt_add_connector, for a diagram drawn by hand"]}
```

The tools that go through the clipboard follow the same rule from the other
side. `paste object` reports nothing, so every paste is checked by counting
the slide's shapes and reading the new shape's type; none landed is a
refusal, the wrong thing landed is removed and then refused, and a paste that
went inside a selected chart is prevented by clearing the selection first.
The editors go one step further: the new shape is copied straight back and
the change looked for in the copy, and a paste that reads back without it is
removed before the refusal, with the original untouched. Only after the copy
carries the change is the original deleted.

The server `instructions` string should carry a short platform note too, since
that is what the model reads before planning a deck.

### 7.4 Packaging

`pyproject.toml:43` declares `pywin32>=306` unconditionally, which is a hard
install failure on macOS before a single line of the port matters.

```toml
dependencies = [
  "mcp[cli]>=1.0.0,<3.0.0",
  "pydantic>=2.0.0",
  "pywin32>=306; sys_platform == 'win32'",
  "appscript>=1.4.0; sys_platform == 'darwin'",
]
```

The classifiers need macOS added, and the README badge stops saying Windows.

---

## 8. Sequence

**Phase 0, make it installable and testable on a Mac.** Platform markers in
`pyproject.toml`. Make `com_wrapper.py` import pywin32 lazily, the way
`redraw.py` already imports win32gui lazily, and guard the `winreg` imports in
`onedrive.py:11` and `presentation.py:10` and the `ctypes.windll` use at module
scope in `export.py`. This alone takes the suite from 86 runnable tests on macOS
to 539 of 555, which is a real regression net for the whole schema layer before
any Apple Event code is written.

**Phase 1, the driver.** `backend/mac_ae.py`, `names.py`, `enums.py`, the quirk
absorption from section 3, the verify-every-mutation policy from section 5, and
a timeout policy mirroring the existing busy retry loop.

**Phase 2, the core tools.** app, presentation, slides, shapes, text,
formatting, placeholders, layout, tables, export. This is where most real use
lives.

**Phase 3, the rest, and the honest refusals.**

**Phase 4, the clipboard.** Charts, freeforms and groups through
`com.microsoft.Art--GVML-ClipFormat`, in the order `docs/gvml-design.md`
section 6 gives: the six tools that create or read first, then the five that
edit a chart, then the four that edit a path. All three tiers are in.

**Spike, in parallel and not on the critical path.** Whether a shipped `.ppam`
plus `run VB macro` can reach the VBA object model. If it can, SmartArt comes
back that way. Charts and freeforms no longer depend on it.

---

## 9. What is still unknown

Written down so nobody re-derives it.

1. Whether `run VB macro` actually executes. Section 6.1 has the one line test.
   It matters for SmartArt now; charts are made through the clipboard instead.
2. ~~Table cell addressing.~~ Settled. `get cell from` is the wrong route and
   `table.rows[r].cells[c]` is the right one; see section 5.1. Every table tool
   now runs live.
3. ~~Theme colours.~~ Settled, and the earlier note was wrong. All twelve read
   and write through `slide master`'s `theme`'s `theme color scheme`. What does
   not work is the Windows walk, because a deck's `designs` collection counts
   zero here and `designs[1]` answers -1728.
4. Sections. The command answered through appscript and failed with -1708
   through AppleScript, which should not both be true and needs one clean run.
5. Align, distribute, group and ungroup all take a `shape range`, and the only
   route to one appears to be through the window's selection. That changes the
   shape of those four tools.
6. ~~Whether passing a file reference rather than a path string hands PowerPoint
   a sandbox extension.~~ Settled, and the answer is half of each. A file
   reference is necessary, `save in: "<path>"` as text writes nothing anywhere
   at all, including inside the container, and reports success. It is not
   sufficient. `save in: file "<path>"` outside the container hangs PowerPoint
   until the call times out, on Documents, Desktop and Downloads alike. So
   staging is still the answer, and the file reference plus an explicit format
   is what makes the staged save land.
