# Hybrid COM + UI Automation Patterns for PowerPoint

> Research document covering practical patterns for combining COM automation with
> UI automation in PowerPoint from Python. Covers use-case analysis, shape
> alignment/distribution via COM, architecture patterns, and Python library assessment.

---

## Table of Contents

1. [Practical Use Cases: What Needs UI Automation?](#1-practical-use-cases-what-needs-ui-automation)
2. [Shape Alignment and Distribution via COM](#2-shape-alignment-and-distribution-via-com)
3. [Architecture Patterns for Hybrid COM + UI Systems](#3-architecture-patterns-for-hybrid-com--ui-systems)
4. [Python Library Assessment for UI Automation](#4-python-library-assessment-for-ui-automation)
5. [Implementation Recommendations for ppt-com-mcp](#5-implementation-recommendations-for-ppt-com-mcp)

---

## 1. Practical Use Cases: What Needs UI Automation?

### Summary Matrix

| Operation | COM Can Do It? | COM Method | UI Automation Needed? |
|---|---|---|---|
| Format Painter (copy formatting) | **YES** | `Shape.PickUp` / `Shape.Apply` | No |
| Align shapes | **YES** | `ShapeRange.Align` | No |
| Distribute shapes | **YES** | `ShapeRange.Distribute` | No |
| Smart Guides behavior | **NO** | N/A (visual-only feature) | Not automatable |
| Insert from content library | **NO** | N/A | Yes (or file-based workaround) |
| Animation preview/playback | **Partial** | `SlideShowSettings.Run` | Partial |
| Undo/Redo | **Partial** | `ExecuteMso("Undo")` | Alternative |
| Copy/Paste with formatting | **YES** | `Shapes.Paste` / `Shapes.PasteSpecial` | No |
| Copy animation between shapes | **YES** | `Shape.PickupAnimation` / `Shape.ApplyAnimation` | No |
| Drag to reorder slides | **YES** | `Slide.MoveTo` | No |
| Morph transition preview | **NO** | Can set transition type, but preview requires slideshow | Partial |
| Designer/Design Ideas | **NO** | N/A | Yes |
| Insert Icons / 3D Models from gallery | **NO** | N/A (online gallery) | Yes |

### 1.1 Format Painter (Copy Formatting)

**COM can do this.** PowerPoint exposes `Shape.PickUp` and `Shape.Apply` methods.

- `Shape.PickUp()` copies the formatting of the specified shape into an internal buffer.
- `Shape.Apply()` applies the previously picked-up formatting to the target shape.

```python
def _format_painter_impl(slide_index, source_name, target_name):
    """Copy formatting from source shape to target shape via COM."""
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    source = slide.Shapes(source_name)
    target = slide.Shapes(target_name)

    source.PickUp()
    target.Apply()

    return {"success": True, "source": source_name, "target": target_name}
```

**What PickUp/Apply copies:**
- Fill (solid, gradient, pattern, picture)
- Line (color, weight, dash style)
- Shadow
- 3D effects
- Text formatting is NOT copied by PickUp/Apply (it copies shape-level formatting only)

**For text formatting transfer**, you must manually read and apply font properties:
```python
def _copy_text_format(source_range, target_range):
    """Manually copy text formatting between TextRange objects."""
    target_range.Font.Name = source_range.Font.Name
    target_range.Font.Size = source_range.Font.Size
    target_range.Font.Bold = source_range.Font.Bold
    target_range.Font.Italic = source_range.Font.Italic
    target_range.Font.Color.RGB = source_range.Font.Color.RGB
    # ... etc for each property
```

**Verdict: No UI automation needed.**

### 1.2 Align / Distribute Multiple Shapes

**COM can do this.** See [Section 2](#2-shape-alignment-and-distribution-via-com) for full details.

`ShapeRange.Align(AlignCmd, RelativeTo)` and `ShapeRange.Distribute(DistributeCmd, RelativeTo)` are fully supported in the COM object model.

**Verdict: No UI automation needed.**

### 1.3 Smart Guides Behavior

**COM cannot do this.** Smart Guides are a purely visual, interactive feature that shows alignment guides as you drag shapes in the PowerPoint UI. They have no COM API representation.

However, Smart Guides are not something you would want to automate. They are a visual aid for manual editing. The equivalent programmatic approach is to use `ShapeRange.Align` or calculate positions manually.

**Workaround:** Calculate alignment positions mathematically and set `Shape.Left` / `Shape.Top` directly.

**Verdict: Not automatable and not needed -- use COM alignment methods instead.**

### 1.4 Insert from Content Library

**COM cannot directly access the content library UI.** The content library (organizational assets in Microsoft 365) is accessed through the ribbon UI.

**Workarounds:**
1. If the content library items are stored as files on SharePoint/OneDrive, download them first, then insert via `Shapes.AddPicture` or open as a presentation.
2. Use `ExecuteMso` to open the content library panel, then use UI automation to navigate it.

```python
# Open the content library panel via ExecuteMso (if available)
app.CommandBars.ExecuteMso("ContentLibrary")
# Then UI automation would be needed to select and insert items
```

**Verdict: Requires UI automation for gallery interaction, but file-based workaround is preferred.**

### 1.5 Animation Preview / Playback Control

**COM can partially do this.**

- **Setting up animations**: Fully supported via `TimeLine.MainSequence.AddEffect()`
- **Starting a slideshow**: `SlideShowSettings.Run()` launches the slideshow
- **Controlling playback**: `SlideShowView.Next()`, `.Previous()`, `.GotoSlide()`
- **Preview a single animation**: Not directly available via COM

For previewing a single animation effect (like clicking the Preview button in the Animation tab), use `ExecuteMso`:

```python
# Preview animations on the current slide
app.CommandBars.ExecuteMso("AnimationPreview")
```

**Verdict: Mostly COM, with ExecuteMso for animation preview.**

### 1.6 Undo / Redo Operations

**COM has limited support.** There is no `Application.Undo()` method in the PowerPoint object model (unlike Word/Excel). However, you can invoke it through `ExecuteMso`:

```python
# Undo the last action
app.CommandBars.ExecuteMso("Undo")

# Redo the last undone action
app.CommandBars.ExecuteMso("Redo")
```

**Caveats:**
- `ExecuteMso("Undo")` undoes one action. There is no way to undo N actions at once.
- You cannot inspect the undo stack or get a list of undoable operations.
- COM operations themselves are placed on the undo stack, so they can be undone by the user with Ctrl+Z.
- `CommandBars.GetEnabledMso("Undo")` returns `True` if Undo is available.

**Alternative UI automation approach:** Send `Ctrl+Z` / `Ctrl+Y` keystrokes:
```python
import pyautogui
pyautogui.hotkey('ctrl', 'z')  # Undo
pyautogui.hotkey('ctrl', 'y')  # Redo
```

**Verdict: Use `ExecuteMso` -- no full UI automation needed.**

### 1.7 Copy / Paste with Formatting

**COM can do this fully.** Multiple clipboard methods are available:

```python
# Copy a shape to clipboard
slide.Shapes("MyShape").Copy()

# Paste shapes onto a slide
new_shapes = slide.Shapes.Paste()

# Paste with specific format
# PpPasteDataType values:
#   0 = ppPasteDefault
#   1 = ppPasteBitmap
#   2 = ppPasteEnhancedMetafile
#   3 = ppPasteMetafilePicture
#   7 = ppPasteText
#   8 = ppPasteHTML
#   9 = ppPasteRTF
#  10 = ppPasteOLEObject
#  11 = ppPasteShape
new_shapes = slide.Shapes.PasteSpecial(DataType=2)  # Enhanced Metafile

# Copy a slide
slide.Copy()

# Paste slide at a specific position
pres.Slides.Paste(Index=3)
```

Additional clipboard methods:
- `View.PasteSpecial` -- pastes into the current view context
- `TextRange.PasteSpecial` -- pastes into a text range
- `Shapes.PasteSpecial` -- pastes with data type control

**Verdict: No UI automation needed.**

### 1.8 Copy Animation Between Shapes

**COM can do this** (PowerPoint 2010+):

```python
def _copy_animation_impl(slide_index, source_name, target_name):
    """Copy all animations from one shape to another."""
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    source = slide.Shapes(source_name)
    target = slide.Shapes(target_name)

    source.PickupAnimation()
    target.ApplyAnimation()

    return {"success": True}
```

**Verdict: No UI automation needed.**

### 1.9 Drag to Reorder Slides

**COM can do this.** The `Slide.MoveTo` method reorders slides programmatically:

```python
# Move slide 3 to position 1
pres.Slides(3).MoveTo(1)

# Move the last slide to position 2
last = pres.Slides.Count
pres.Slides(last).MoveTo(2)
```

**Verdict: No UI automation needed.**

### 1.10 Designer / Design Ideas

**COM cannot trigger Design Ideas.** This is a cloud-based AI feature accessible only through the ribbon UI.

```python
# This *might* work depending on PowerPoint version:
app.CommandBars.ExecuteMso("DesignIdeas")
```

But even if the panel opens, the suggestions are displayed as UI thumbnails that require clicking to apply. Full UI automation would be needed to select a design.

**Verdict: Requires UI automation (if implementable at all).**

### 1.11 Insert Icons / 3D Models from Online Gallery

**COM cannot access the online icon/3D model gallery.** These are fetched from Microsoft's online service through the ribbon UI.

**Workaround:** Download SVG/PNG files separately and insert via `Shapes.AddPicture`.

**Verdict: Use file-based workaround; UI automation is fragile for online galleries.**

---

## 2. Shape Alignment and Distribution via COM

### 2.1 ShapeRange.Align Method

PowerPoint COM fully supports shape alignment through the `ShapeRange.Align` method.

**Syntax:**
```
ShapeRange.Align(AlignCmd, RelativeTo)
```

**Parameters:**

| Parameter | Type | Required | Description |
|---|---|---|---|
| `AlignCmd` | `MsoAlignCmd` | Yes | How to align the shapes |
| `RelativeTo` | `MsoTriState` | Yes | If `msoTrue`, align relative to slide edges. If `msoFalse`, align relative to each other. |

### 2.2 MsoAlignCmd Enumeration

| Constant | Value | Description |
|---|---|---|
| `msoAlignLefts` | 0 | Align left edges |
| `msoAlignCenters` | 1 | Align horizontal centers |
| `msoAlignRights` | 2 | Align right edges |
| `msoAlignTops` | 3 | Align top edges |
| `msoAlignMiddles` | 4 | Align vertical middles |
| `msoAlignBottoms` | 5 | Align bottom edges |

### 2.3 ShapeRange.Distribute Method

**Syntax:**
```
ShapeRange.Distribute(DistributeCmd, RelativeTo)
```

**Parameters:**

| Parameter | Type | Required | Description |
|---|---|---|---|
| `DistributeCmd` | `MsoDistributeCmd` | Yes | Distribution direction |
| `RelativeTo` | `MsoTriState` | Yes | If `msoTrue`, distribute over entire slide. If `msoFalse`, distribute over the space the shapes currently occupy. |

### 2.4 MsoDistributeCmd Enumeration

| Constant | Value | Description |
|---|---|---|
| `msoDistributeHorizontally` | 0 | Distribute evenly horizontally |
| `msoDistributeVertically` | 1 | Distribute evenly vertically |

### 2.5 Python Implementation Examples

#### Creating a ShapeRange from shape names

```python
def _align_shapes_impl(slide_index, shape_names, align_cmd, relative_to_slide):
    """Align multiple shapes using ShapeRange.Align."""
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    # Create a ShapeRange from an array of shape names
    # Note: Shapes.Range() accepts an array of names or indices
    shape_range = slide.Shapes.Range(shape_names)

    # MsoTriState: msoTrue = -1, msoFalse = 0
    relative = -1 if relative_to_slide else 0
    shape_range.Align(align_cmd, relative)

    return {
        "success": True,
        "aligned_count": shape_range.Count,
        "align_cmd": align_cmd,
        "relative_to_slide": relative_to_slide,
    }
```

#### Aligning shapes to slide center

```python
def _center_shapes_on_slide_impl(slide_index, shape_names):
    """Center shapes both horizontally and vertically on the slide."""
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape_range = slide.Shapes.Range(shape_names)

    # msoAlignCenters=1, msoAlignMiddles=4, msoTrue=-1
    shape_range.Align(1, -1)  # Center horizontally relative to slide
    shape_range.Align(4, -1)  # Center vertically relative to slide

    return {"success": True}
```

#### Distributing shapes evenly

```python
def _distribute_shapes_impl(slide_index, shape_names, direction, relative_to_slide):
    """Distribute shapes evenly."""
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)
    shape_range = slide.Shapes.Range(shape_names)

    # msoDistributeHorizontally=0, msoDistributeVertically=1
    distribute_cmd = 0 if direction == "horizontal" else 1
    relative = -1 if relative_to_slide else 0

    shape_range.Distribute(distribute_cmd, relative)

    return {
        "success": True,
        "distributed_count": shape_range.Count,
        "direction": direction,
    }
```

#### Creating ShapeRange from indices

```python
def _make_shape_range_from_indices(slide, indices):
    """Create a ShapeRange from 1-based shape indices.

    Note: Shapes.Range() accepts a list/tuple of names or a list/tuple
    of indices, but mixing types is not supported.
    """
    # Convert indices to a tuple for COM
    return slide.Shapes.Range(tuple(indices))
```

#### Complete alignment + distribution workflow

```python
def _layout_shapes_grid_impl(slide_index, shape_names, columns):
    """Arrange shapes in a grid layout.

    This demonstrates combining alignment and distribution.
    """
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    shapes = shape_names
    total = len(shapes)
    rows_count = (total + columns - 1) // columns

    # First, align all shapes to the top-left as starting position
    all_range = slide.Shapes.Range(shapes)

    # Distribute horizontally across the slide
    all_range.Distribute(0, -1)  # msoDistributeHorizontally, msoTrue

    # For multi-row layouts, manually set Top positions
    slide_height = pres.PageSetup.SlideHeight
    row_height = slide_height / (rows_count + 1)

    for i, name in enumerate(shapes):
        row = i // columns
        shape = slide.Shapes(name)
        shape.Top = row_height * (row + 0.5) - shape.Height / 2

    return {"success": True, "rows": rows_count, "columns": columns}
```

### 2.6 Passing Arrays to COM from Python (pywin32)

A critical implementation detail: when calling `Shapes.Range()` with multiple names, you must pass a Python tuple or list. pywin32 converts these to COM SAFEARRAY:

```python
# These all work:
shape_range = slide.Shapes.Range(("Shape1", "Shape2", "Shape3"))
shape_range = slide.Shapes.Range(["Shape1", "Shape2", "Shape3"])
shape_range = slide.Shapes.Range((1, 2, 3))  # By index

# This does NOT work (single string, not array):
shape_range = slide.Shapes.Range("Shape1")  # Returns single ShapeRange with 1 shape
# But this is still valid -- it creates a ShapeRange of one shape
```

### 2.7 Alternative: ExecuteMso for Alignment

If COM ShapeRange methods prove unreliable, alignment can also be triggered via `ExecuteMso` with the shapes selected in the UI:

```python
# These require the shapes to be selected in the PowerPoint UI first
app.CommandBars.ExecuteMso("ObjectsAlignLeftSmart")
app.CommandBars.ExecuteMso("ObjectsAlignCenterHorizontalSmart")
app.CommandBars.ExecuteMso("ObjectsAlignRightSmart")
app.CommandBars.ExecuteMso("ObjectsAlignTopSmart")
app.CommandBars.ExecuteMso("ObjectsAlignMiddleVerticalSmart")
app.CommandBars.ExecuteMso("ObjectsAlignBottomSmart")
app.CommandBars.ExecuteMso("ObjectsDistributeHorizontalSmart")
app.CommandBars.ExecuteMso("ObjectsDistributeVerticalSmart")
```

However, this requires shapes to be selected in the UI first, making the COM `ShapeRange.Align` approach strongly preferred for programmatic use.

---

## 3. Architecture Patterns for Hybrid COM + UI Systems

### 3.1 Decision Framework: When to Use COM vs UI Automation

```
Is there a COM method for this operation?
  YES --> Use COM (always preferred)
    |
  NO --> Is there an ExecuteMso command?
           YES --> Use ExecuteMso via COM
             |
           NO --> Does it require visual/interactive UI?
                    YES --> Use UI automation
                      |
                    NO --> Implement manually (calculate positions, etc.)
```

**Priority order:**
1. **Direct COM API** -- Most reliable, fastest, no UI dependency
2. **ExecuteMso via COM** -- Reliable for ribbon commands, no window focus needed
3. **SendKeys / keyboard shortcuts** -- Requires window focus, moderately reliable
4. **UI Automation (pywinauto)** -- Element-based interaction, reasonably reliable
5. **Pixel-based automation (pyautogui)** -- Fragile, last resort

### 3.2 Thread Safety Architecture

The ppt-com-mcp server already uses a dedicated STA thread for COM operations. When adding UI automation, the threading model must be carefully considered.

#### Current Architecture

```
MCP Server (async, main thread)
    |
    v
ppt.execute(func, *args) --> Queue --> COM Worker Thread (STA)
                                            |
                                            v
                                       PowerPoint COM
```

#### Extended Architecture with UI Automation

```
MCP Server (async, main thread)
    |
    +--> ppt.execute(func) --> Queue --> COM Worker Thread (STA)
    |                                        |
    |                                        v
    |                                   PowerPoint COM
    |
    +--> ui.execute(func) --> Queue --> UI Worker Thread (MTA or separate)
                                            |
                                            v
                                       pywinauto / SendKeys
```

**Key considerations:**

1. **COM operations MUST stay on the STA thread.** The existing `PowerPointCOMWrapper` handles this correctly.

2. **UI Automation (pywinauto with UIA backend) uses COM internally** and since pywinauto 0.6.5+ defaults to MTA, it can run on a separate thread without `CoInitialize()`.

3. **Never mix COM and UI automation on the same thread** if possible. COM operations may trigger UI changes that UI automation is waiting for, causing deadlocks.

4. **Serialization between COM and UI operations is critical.** If a COM operation changes the slide, wait for it to complete before inspecting the UI tree.

#### Proposed UIAutomationWrapper

```python
import threading
from concurrent.futures import Future
from queue import Queue
from typing import Any, Callable

class UIAutomationWrapper:
    """Manages UI automation operations on a dedicated thread.

    Similar pattern to PowerPointCOMWrapper but for pywinauto operations.
    Keeps UI automation isolated from the COM STA thread.
    """

    def __init__(self):
        self._thread = None
        self._queue = Queue()
        self._running = False

    def start(self):
        if self._running:
            return
        self._running = True
        self._thread = threading.Thread(
            target=self._worker, daemon=True, name="UI-Worker"
        )
        self._thread.start()

    def stop(self):
        if not self._running:
            return
        self._running = False
        self._queue.put(None)
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=5.0)

    def _worker(self):
        """Worker thread for UI automation operations."""
        # pywinauto 0.6.5+ uses MTA by default, no CoInitialize needed
        while self._running:
            item = self._queue.get()
            if item is None:
                break
            func, args, kwargs, future = item
            try:
                result = func(*args, **kwargs)
                future.set_result(result)
            except Exception as e:
                future.set_exception(e)

    def execute(self, func: Callable, *args: Any, **kwargs: Any) -> Any:
        """Execute a UI automation function on the UI worker thread."""
        future = Future()
        self._queue.put((func, args, kwargs, future))
        return future.result(timeout=30.0)
```

### 3.3 Ensuring PowerPoint Window State Before UI Operations

UI automation requires the PowerPoint window to be in the correct state. Before performing any UI automation:

```python
import time

def ensure_window_ready(app, view_type=None):
    """Ensure PowerPoint window is visible, focused, and in the right view.

    Must be called on the COM thread before handing off to UI automation.
    """
    # 1. Ensure PowerPoint is visible
    app.Visible = True  # msoTrue

    # 2. Ensure window is not minimized
    # ppWindowNormal=1, ppWindowMaximized=3
    if app.WindowState == 2:  # ppWindowMinimized
        app.WindowState = 1   # ppWindowNormal

    # 3. Activate the window (bring to foreground)
    app.Activate()

    # 4. Set the view type if needed
    if view_type is not None:
        app.ActiveWindow.ViewType = view_type

    # 5. Small delay for UI to settle
    time.sleep(0.3)

    return True
```

### 3.4 Coordinating COM and UI Operations

When an operation needs both COM and UI automation, use a coordinator pattern:

```python
async def hybrid_operation(com_wrapper, ui_wrapper):
    """Example: Operation that needs both COM and UI automation.

    Pattern: COM setup -> Wait -> UI interaction -> COM verification
    """
    # Step 1: COM setup (on COM thread)
    com_wrapper.execute(ensure_window_ready, app)

    # Step 2: COM operation to prepare state
    com_wrapper.execute(select_shapes_for_alignment)

    # Step 3: Brief pause for UI to reflect COM changes
    await asyncio.sleep(0.5)

    # Step 4: UI automation interaction (on UI thread)
    ui_wrapper.execute(click_ribbon_button, "Design Ideas")

    # Step 5: Wait for UI to respond
    await asyncio.sleep(1.0)

    # Step 6: COM verification of results
    result = com_wrapper.execute(verify_operation_result)

    return result
```

### 3.5 Error Handling and Recovery

#### COM Error Recovery (already implemented in ppt-com-mcp)

```python
def safe_com_operation(func, *args, max_retries=3):
    """Execute a COM operation with retry logic."""
    for attempt in range(max_retries):
        try:
            return ppt.execute(func, *args)
        except pywintypes.com_error as e:
            if attempt < max_retries - 1:
                # Reconnect and retry
                ppt.execute(ppt._connect_impl)
                time.sleep(0.5)
            else:
                raise
```

#### UI Automation Error Recovery

```python
def safe_ui_operation(func, *args, max_retries=3, retry_delay=1.0):
    """Execute a UI automation operation with retry logic.

    UI operations are inherently more fragile than COM, so we need
    more sophisticated retry logic.
    """
    last_error = None
    for attempt in range(max_retries):
        try:
            return func(*args)
        except Exception as e:
            last_error = e
            if attempt < max_retries - 1:
                # Ensure window is still in correct state
                ppt.execute(ensure_window_ready, ppt._get_app_impl())
                time.sleep(retry_delay * (attempt + 1))  # Exponential backoff

    raise RuntimeError(
        f"UI operation failed after {max_retries} attempts: {last_error}"
    )
```

#### Wait-for-UI-State Pattern

```python
import time

def wait_for_condition(check_func, timeout=10.0, interval=0.5):
    """Wait until a condition is met or timeout expires.

    Args:
        check_func: Callable that returns True when condition is met
        timeout: Maximum seconds to wait
        interval: Seconds between checks

    Returns:
        True if condition was met, False if timeout expired
    """
    start = time.time()
    while time.time() - start < timeout:
        try:
            if check_func():
                return True
        except Exception:
            pass
        time.sleep(interval)
    return False


# Usage example: wait for a dialog to appear
def wait_for_dialog(app_uia, title, timeout=10.0):
    """Wait for a dialog window to appear in the UI tree."""
    def check():
        try:
            dlg = app_uia.window(title=title)
            return dlg.exists()
        except Exception:
            return False
    return wait_for_condition(check, timeout=timeout)
```

### 3.6 Selection-Based Operations via COM

Some operations require shapes to be selected in the UI. COM can programmatically select shapes:

```python
def _select_shapes_impl(slide_index, shape_names):
    """Select specific shapes in the PowerPoint UI via COM.

    This is useful when an operation requires UI selection state.
    """
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    window = app.ActiveWindow

    # Navigate to the correct slide
    window.View.GotoSlide(slide_index)

    slide = pres.Slides(slide_index)

    # Select the first shape (replace selection)
    slide.Shapes(shape_names[0]).Select()  # msoTrue = replace

    # Add remaining shapes to selection
    for name in shape_names[1:]:
        slide.Shapes(name).Select(0)  # msoFalse = add to selection

    return {"selected": len(shape_names)}
```

---

## 4. Python Library Assessment for UI Automation

### 4.1 pywinauto

**Maturity:** High. Active since 2006, latest release January 2025 (v0.6.8+).

**Architecture:**
- Two backends: `win32` (Win32 API, default) and `uia` (MS UI Automation)
- UIA backend is essential for modern Office ribbon controls
- Since v0.6.5, uses MTA COM model by default (better for threading)

**PowerPoint support:**
- Can interact with the ribbon, dialog boxes, and task panes via UIA backend
- Element identification by name, control type, automation ID
- Supports `click_input()` (mouse), `type_keys()` (keyboard), and programmatic `click()` (UIA invoke)

**Example: Interacting with PowerPoint ribbon:**
```python
from pywinauto.application import Application

# Connect to running PowerPoint
app = Application(backend="uia").connect(
    class_name="PPTFrameClass",
    title_re=".*PowerPoint.*"
)

# Access the main window
main_window = app.window(class_name="PPTFrameClass")

# Click Design Ideas button on the ribbon
# (element names depend on PowerPoint version and language)
ribbon = main_window.child_window(control_type="ToolBar", title="Ribbon")
design_tab = ribbon.child_window(title="Design", control_type="TabItem")
design_tab.click_input()

# Find and click Design Ideas button
design_ideas = ribbon.child_window(title="Design Ideas", control_type="Button")
design_ideas.click_input()
```

**Strengths:**
- Robust element identification via accessibility tree
- Good documentation and community
- Both programmatic (UIA Invoke) and input-simulation methods
- Native integration with Windows accessibility APIs

**Weaknesses:**
- Windows-only
- Ribbon element names change between PowerPoint versions/languages
- Complex UI trees can be slow to traverse
- Some modern UI elements may not be well-exposed in the accessibility tree

### 4.2 pyautogui

**Maturity:** High. Cross-platform. Well-maintained.

**Architecture:**
- Image/coordinate-based automation
- Screenshot + image matching for element location
- Direct mouse/keyboard control

**PowerPoint support:**
- Can send keyboard shortcuts (Ctrl+Z, Ctrl+C, Ctrl+V, etc.)
- Can click at specific screen coordinates
- Can locate UI elements via screenshot matching

**Example:**
```python
import pyautogui

# Send keyboard shortcut for Undo
pyautogui.hotkey('ctrl', 'z')

# Click at a known ribbon position (fragile!)
pyautogui.click(x=150, y=85)

# Use image matching (more robust but still fragile)
location = pyautogui.locateOnScreen('design_ideas_button.png')
if location:
    pyautogui.click(location)
```

**Strengths:**
- Cross-platform (Windows, macOS, Linux)
- Simple API, easy to learn
- Screenshot capability for debugging
- Works with any application (no accessibility tree needed)

**Weaknesses:**
- Extremely fragile (coordinate/image-based)
- Breaks with resolution changes, DPI scaling, theme changes
- No understanding of UI structure
- Slow (screenshot analysis)
- Cannot read text from controls
- Requires visual access to screen (fails in headless/RDP scenarios)

### 4.3 keyboard + mouse Libraries

**Libraries:** `keyboard`, `mouse` (by boppreh)

**Architecture:**
- Low-level input hooks and injection
- Global hotkey registration
- Event-based

```python
import keyboard
import mouse

# Send keyboard shortcut
keyboard.send('ctrl+z')

# Type text
keyboard.write('Hello World')

# Click at position
mouse.click('left')
mouse.move(100, 200)
```

**Strengths:**
- Very lightweight
- Low-level control (scan codes, raw input)
- Global hotkey support
- Fast

**Weaknesses:**
- No UI element awareness
- No accessibility tree integration
- Coordinate-based (same fragility as pyautogui)
- `keyboard` library requires root/admin on some systems

### 4.4 comtypes with UIAutomation

**Architecture:**
- Direct Python bindings to Windows UI Automation COM interfaces
- Lower-level than pywinauto but more control

```python
import comtypes
from comtypes.client import CreateObject

# Create UIAutomation object
uia = CreateObject("{ff48dba4-60ef-4201-aa87-54103eef594e}",
                    interface=comtypes.gen.UIAutomationClient.IUIAutomation)

# Get root element
root = uia.GetRootElement()

# Find PowerPoint window
condition = uia.CreatePropertyCondition(
    30005,  # UIA_NamePropertyId
    "PowerPoint"
)
ppt_window = root.FindFirst(1, condition)  # TreeScope_Children
```

**Strengths:**
- Direct access to UIAutomation COM interfaces
- Maximum control and flexibility
- No abstraction overhead
- Can access any UIAutomation property and pattern

**Weaknesses:**
- Very verbose and low-level
- Steep learning curve
- Manual memory management of COM objects
- No convenience wrappers

### 4.5 Python-UIAutomation-for-Windows (uiautomation)

**Architecture:**
- Higher-level wrapper around UIAutomation COM interfaces
- Pure Python with comtypes dependency
- Fluent API for element finding

```python
import uiautomation as auto

# Find PowerPoint window
ppt_window = auto.WindowControl(
    searchDepth=1,
    ClassName="PPTFrameClass"
)

# Navigate to ribbon
ribbon = ppt_window.ToolBarControl(Name="Ribbon")

# Click a button
button = ribbon.ButtonControl(Name="Design Ideas")
button.Click()
```

**Strengths:**
- Cleaner API than raw comtypes
- Good for quick scripts
- Active maintenance (as of 2025)
- Better Performance than pywinauto for some operations

**Weaknesses:**
- Less documentation than pywinauto
- Smaller community
- Some edge cases in COM object lifetime management

### 4.6 win32gui + win32api Direct Approach

**Architecture:**
- Direct Win32 API calls via pywin32
- Window enumeration, message sending, input simulation

```python
import win32gui
import win32api
import win32con

# Find PowerPoint window
hwnd = win32gui.FindWindow("PPTFrameClass", None)

# Bring to foreground
win32gui.SetForegroundWindow(hwnd)

# Send message (e.g., WM_COMMAND)
win32gui.SendMessage(hwnd, win32con.WM_COMMAND, command_id, 0)

# Simulate key press
win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
win32api.keybd_event(ord('Z'), 0, 0, 0)
win32api.keybd_event(ord('Z'), 0, win32con.KEYEVENTF_KEYUP, 0)
win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
```

**Strengths:**
- No additional dependencies (pywin32 is already required for COM)
- Maximum control at OS level
- Can find/manipulate windows by class name
- Useful for window management (activate, minimize, etc.)

**Weaknesses:**
- Very low-level
- No accessibility tree (cannot inspect ribbon buttons, etc.)
- WM_COMMAND IDs are internal and undocumented for Office
- Modern ribbon controls are not traditional Win32 controls

### 4.7 Accessibility Insights for Windows (Discovery Tool)

Not a Python library, but an essential **discovery tool** for UI automation development.

**Purpose:** Inspect the UI Automation tree of running applications to find:
- Control names, types, and automation IDs
- Parent-child relationships in the element tree
- Supported control patterns (Invoke, Toggle, Selection, etc.)
- Property values for building selectors

**Usage for PowerPoint automation:**
1. Open Accessibility Insights for Windows
2. Hover over PowerPoint ribbon buttons
3. Note the `Name`, `ControlType`, `AutomationId` properties
4. Use these properties to build pywinauto selectors

**Key discovery information for PowerPoint:**
- PowerPoint main window class: `PPTFrameClass`
- Ribbon control type: `ToolBar`
- Ribbon tabs: `TabItem` controls
- Ribbon buttons: `Button` or `SplitButton` controls
- Slide thumbnails: Custom controls in the slide panel
- Slide editing area: `Pane` control

### 4.8 Recommendation Matrix

| Use Case | Recommended Library | Reason |
|---|---|---|
| Keyboard shortcuts (Ctrl+Z, etc.) | **win32api** or **pyautogui** | Lightweight, reliable for hotkeys |
| Click ribbon buttons | **pywinauto (UIA)** | Accessibility-based, robust |
| Interact with dialogs | **pywinauto (UIA)** | Element identification by name/type |
| Navigate task panes | **pywinauto (UIA)** | Good tree traversal |
| Read UI text/values | **pywinauto (UIA)** | Can read control properties |
| Window management | **win32gui** | Already available via pywin32 |
| Image-based fallback | **pyautogui** | Cross-platform, screenshot matching |
| Low-level element inspection | **comtypes + UIAutomation** | Maximum control |
| Quick prototyping | **uiautomation** | Clean API, fast to write |
| Discovery/development | **Accessibility Insights** | Visual inspection tool |

**Overall recommendation for ppt-com-mcp:**

1. **Primary: COM API** (already implemented) -- handles 90%+ of operations
2. **Secondary: `ExecuteMso` via COM** -- handles ribbon commands without UI dependency
3. **Tertiary: pywinauto (UIA backend)** -- for the rare cases requiring true UI interaction
4. **Utility: win32gui/win32api** -- for window management (activate, foreground)
5. **Avoid: pyautogui** -- too fragile for production use with Office

---

## 5. Implementation Recommendations for ppt-com-mcp

### 5.1 New COM Tools to Add (No UI Automation Needed)

Based on this research, these high-value operations are fully supported by COM but not yet in ppt-com-mcp:

#### Shape Alignment and Distribution

```python
# Constants to add to constants.py:
msoAlignLefts = 0
msoAlignCenters = 1
msoAlignRights = 2
msoAlignTops = 3
msoAlignMiddles = 4
msoAlignBottoms = 5
msoDistributeHorizontally = 0
msoDistributeVertically = 1
```

New tools:
- `ppt_align_shapes` -- Align multiple shapes (left, center, right, top, middle, bottom)
- `ppt_distribute_shapes` -- Distribute shapes evenly (horizontal, vertical)

#### Format Painter (PickUp / Apply)

New tools:
- `ppt_copy_formatting` -- Copy formatting from one shape to another via PickUp/Apply
- `ppt_copy_animation` -- Copy animation from one shape to another via PickupAnimation/ApplyAnimation

#### Clipboard Operations

New tools:
- `ppt_copy_shape` -- Copy shape to clipboard
- `ppt_paste_shapes` -- Paste shapes from clipboard
- `ppt_paste_special` -- Paste with specific data type

### 5.2 ExecuteMso-Based Tools

These use `CommandBars.ExecuteMso` for operations without direct COM methods:

```python
# Useful ExecuteMso commands for PowerPoint:
EXECUTEMSO_COMMANDS = {
    # Undo/Redo
    "Undo": "Undo",
    "Redo": "Redo",

    # Clipboard (alternative to COM methods)
    "Copy": "Copy",
    "Cut": "Cut",
    "Paste": "Paste",
    "PasteSourceFormatting": "PasteSourceFormatting",
    "PasteDestinationFormatting": "PasteDestinationFormatting",
    "PastePicture": "PastePicture",

    # Format Painter
    "FormatPainter": "FormatPainter",

    # Alignment (requires selection)
    "ObjectsAlignLeftSmart": "ObjectsAlignLeftSmart",
    "ObjectsAlignCenterHorizontalSmart": "ObjectsAlignCenterHorizontalSmart",
    "ObjectsAlignRightSmart": "ObjectsAlignRightSmart",
    "ObjectsAlignTopSmart": "ObjectsAlignTopSmart",
    "ObjectsAlignMiddleVerticalSmart": "ObjectsAlignMiddleVerticalSmart",
    "ObjectsAlignBottomSmart": "ObjectsAlignBottomSmart",

    # Distribution (requires selection)
    "ObjectsDistributeHorizontalSmart": "ObjectsDistributeHorizontalSmart",
    "ObjectsDistributeVerticalSmart": "ObjectsDistributeVerticalSmart",

    # Animation
    "AnimationPreview": "AnimationPreview",

    # View
    "SlideShowFromBeginning": "SlideShowFromBeginning",
    "SlideShowFromCurrent": "SlideShowFromCurrent",

    # Selection
    "SelectAll": "SelectAll",
}
```

Helper to check if a command is available:
```python
def _is_command_enabled(command_id):
    """Check if an ExecuteMso command is currently available."""
    app = ppt._get_app_impl()
    return bool(app.CommandBars.GetEnabledMso(command_id))

def _execute_mso_impl(command_id):
    """Execute a ribbon command via ExecuteMso."""
    app = ppt._get_app_impl()
    if not app.CommandBars.GetEnabledMso(command_id):
        return {"error": f"Command '{command_id}' is not currently enabled"}
    app.CommandBars.ExecuteMso(command_id)
    return {"success": True, "command": command_id}
```

### 5.3 When to Consider UI Automation (Future)

UI automation should only be added if there is clear user demand for operations that cannot be done via COM or ExecuteMso. Potential candidates:

1. **Designer / Design Ideas** -- Triggering and selecting AI-generated designs
2. **Online content galleries** -- Icons, stock images, 3D models
3. **Accessibility Checker** -- Running and reading accessibility check results
4. **Translation features** -- Using the built-in translation pane

For these, the recommended approach would be:
1. Add pywinauto as an optional dependency
2. Create a `UIAutomationWrapper` similar to the existing `PowerPointCOMWrapper`
3. Use `win32gui` for window management
4. Use Accessibility Insights during development to map the UI tree

### 5.4 Dependencies Impact

| Library | Install Size | Already Required | Risk |
|---|---|---|---|
| pywin32 | ~30MB | YES (COM core) | None |
| pywinauto | ~5MB | No | Low (well-maintained) |
| pyautogui | ~2MB | No | Low (but fragile results) |
| comtypes | ~1MB | No (but pywinauto needs it) | Low |
| keyboard | ~100KB | No | Low |

**Recommendation:** Add pywinauto as an optional dependency only if/when UI automation features are actually needed. All immediate feature gaps can be filled with COM API methods and ExecuteMso.

---

## References

### Microsoft Documentation
- [ShapeRange.Align method (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shaperange.align)
- [ShapeRange.Distribute method (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shaperange.distribute)
- [MsoAlignCmd enumeration](https://learn.microsoft.com/en-us/office/vba/api/office.msoaligncmd)
- [MsoDistributeCmd enumeration](https://learn.microsoft.com/en-us/office/vba/api/office.msodistributecmd)
- [Shape.PickUp method](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.pickup)
- [Shape.Apply method](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.apply)
- [Shape.PickupAnimation method](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.pickupanimation)
- [Shape.ApplyAnimation method](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.applyanimation)
- [CommandBars.ExecuteMso method](https://learn.microsoft.com/en-us/office/vba/api/office.commandbars.executemso)
- [CommandBars.GetEnabledMso method](https://learn.microsoft.com/en-us/office/vba/api/Office.CommandBars.GetEnabledMso)
- [Slide.MoveTo method](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.moveto)
- [Shapes.PasteSpecial method](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shapes.pastespecial)
- [View.PasteSpecial method](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.View.PasteSpecial)
- [PpPasteDataType enumeration](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.pppastedatatype)

### Python Libraries
- [pywinauto GitHub](https://github.com/pywinauto/pywinauto)
- [pywinauto Documentation](https://pywinauto.readthedocs.io/en/latest/)
- [pywinauto UIA Threading Recipe](https://github.com/pywinauto/pywinauto/wiki/UIA-threading-recipe)
- [Python-UIAutomation-for-Windows](https://github.com/yinkaisheng/Python-UIAutomation-for-Windows)
- [PyAutoGUI Documentation](https://pyautogui.readthedocs.io/)

### Tools
- [Accessibility Insights for Windows](https://accessibilityinsights.io/docs/windows/overview/)
- [idMso Control List for PowerPoint](http://youpresent.co.uk/idmso-control-list-powerpoint-2013-2010/)
- [Office Fluent UI Command Identifiers (GitHub)](https://github.com/OfficeDev/office-fluent-ui-command-identifiers)

### Community Resources
- [Windows UI Automation Tools Comparison](https://medium.com/@jagsmehra.092019/windows-ui-automation-tools-comparison-ae9f8a143f27)
- [PyAutoGUI vs pywinauto Comparison](https://slashdot.org/software/comparison/PyAutoGUI-vs-pywinauto/)
