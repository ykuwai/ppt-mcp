# PowerPoint UI Automation: Shape Selection, SendKeys, and Ribbon Commands

> Research Date: 2026-02-17
> Purpose: Investigate UI-level automation capabilities for the PowerPoint COM MCP server,
> covering programmatic shape selection, keyboard simulation, and ribbon command execution.

---

## Table of Contents

1. [Shape Selection via COM](#1-shape-selection-via-com)
2. [SendKeys via pywin32](#2-sendkeys-via-pywin32)
3. [Ribbon / Menu Operations via COM](#3-ribbon--menu-operations-via-com)
4. [Comparison and Recommendations](#4-comparison-and-recommendations)

---

## 1. Shape Selection via COM

### 1.1 Shape.Select() Method

The `Shape.Select()` method selects a specified shape on a slide. It takes an optional
`Replace` parameter of type `MsoTriState`.

**Signature (VBA):**
```vb
expression.Select(Replace)
```

**Parameters:**

| Name    | Required/Optional | Data Type      | Description                                     |
|---------|-------------------|----------------|-------------------------------------------------|
| Replace | Optional          | MsoTriState    | Whether the selection replaces any previous one. |

**MsoTriState values for Replace:**

| Constant   | Value | Description                                        |
|------------|-------|----------------------------------------------------|
| msoTrue    | -1    | (Default) Selection replaces any previous selection |
| msoFalse   | 0     | Selection is added to the previous selection        |

**Python (pywin32) example -- single shape selection:**
```python
from utils.com_wrapper import ppt
from ppt_com.constants import msoTrue, msoFalse

def _select_shape(slide_index: int, shape_index: int):
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    # Navigate to the slide first (required for Select to work)
    app.ActiveWindow.View.GotoSlide(slide_index)

    # Select the shape (replaces any existing selection)
    slide.Shapes(shape_index).Select()
```

**Python example -- multi-select (additive selection):**
```python
def _select_multiple_shapes(slide_index: int, shape_indices: list):
    """Select multiple shapes on a slide by adding to selection."""
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    # Must navigate to the target slide first
    app.ActiveWindow.View.GotoSlide(slide_index)

    # Select first shape (replace mode)
    slide.Shapes(shape_indices[0]).Select(msoTrue)

    # Add remaining shapes to selection (additive mode)
    for idx in shape_indices[1:]:
        slide.Shapes(idx).Select(msoFalse)
```

### 1.2 Selecting Multiple Shapes via Shapes.Range()

A cleaner approach to multi-selection uses `Shapes.Range()` which accepts an array
of indices or shape names and returns a `ShapeRange` object.

**Shapes.Range() Parameters:**

| Name  | Required/Optional | Data Type | Description                                              |
|-------|-------------------|-----------|----------------------------------------------------------|
| Index | Optional          | Variant   | Integer index, String name, or Array of integers/strings |

**Python example -- select by index array:**
```python
def _select_shapes_by_indices(slide_index: int, shape_indices: list):
    """Select multiple shapes using Shapes.Range with an index array."""
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    app.ActiveWindow.View.GotoSlide(slide_index)

    # Pass a Python list -- pywin32 converts it to a VARIANT array
    shape_range = slide.Shapes.Range(shape_indices)
    shape_range.Select()
```

**Python example -- select by name array:**
```python
def _select_shapes_by_names(slide_index: int, shape_names: list):
    """Select multiple shapes using Shapes.Range with a name array."""
    app = ppt._get_app_impl()
    pres = app.ActivePresentation
    slide = pres.Slides(slide_index)

    app.ActiveWindow.View.GotoSlide(slide_index)

    # Names can be passed as a list of strings
    shape_range = slide.Shapes.Range(shape_names)
    shape_range.Select()
```

**Important notes on Shapes.Range():**
- When passing a Python list, pywin32 automatically marshals it to a COM VARIANT array.
- Index values are 1-based (consistent with all COM collections).
- If you pass a single integer or string, it returns a ShapeRange with one shape.
- Omitting Index returns all shapes on the slide.

### 1.3 ShapeRange.Select() Method

The `ShapeRange.Select()` method has the same `Replace` parameter as `Shape.Select()`.

```python
def _add_range_to_selection(slide_index: int, additional_indices: list):
    """Add a ShapeRange to an existing selection without replacing it."""
    app = ppt._get_app_impl()
    slide = app.ActivePresentation.Slides(slide_index)

    app.ActiveWindow.View.GotoSlide(slide_index)

    # Select first group normally
    slide.Shapes.Range([1, 2]).Select()

    # Add another group to the existing selection
    slide.Shapes.Range([4, 5]).Select(msoFalse)
```

### 1.4 ActiveWindow.Selection Object

The `Selection` object represents the current selection in the specified document window.
It is accessed via `ActiveWindow.Selection`.

**Key properties:**

| Property        | Type        | Description                                             |
|-----------------|-------------|---------------------------------------------------------|
| Type            | Long        | PpSelectionType: 0=None, 1=Slides, 2=Shapes, 3=Text    |
| ShapeRange      | ShapeRange  | All selected shapes (read-only, valid when Type=2 or 3) |
| SlideRange      | SlideRange  | Selected slides (read-only, valid when Type=1)          |
| TextRange       | TextRange   | Selected text (read-only, valid when Type=3)            |
| TextRange2      | TextRange2  | Extended text range with additional formatting access    |
| HasChildShapeRange | Boolean  | Whether selection contains child shapes within a group  |
| ChildShapeRange | ShapeRange  | Child shapes within a selected group shape              |

**Key methods:**

| Method   | Description                                          |
|----------|------------------------------------------------------|
| Copy     | Copies the selection to the Clipboard                |
| Cut      | Cuts the selection to the Clipboard                  |
| Delete   | Deletes the selection                                |
| Unselect | Deselects the current selection                      |

**Important behavior:**
- The Selection object is **reset whenever you change slides** -- `Type` becomes `ppSelectionNone`.
- You must check `Type` before accessing `ShapeRange`, `SlideRange`, or `TextRange`, because
  accessing the wrong property for the current selection type raises a COM error.

**Python example -- reading the current selection:**
```python
from ppt_com.constants import (
    ppSelectionNone, ppSelectionSlides,
    ppSelectionShapes, ppSelectionText,
)

def _get_selection_info():
    """Get detailed information about the current selection."""
    app = ppt._get_app_impl()
    if app.Windows.Count == 0:
        return {"error": "No windows open"}

    sel = app.ActiveWindow.Selection
    sel_type = sel.Type
    result = {"type": sel_type}

    if sel_type == ppSelectionNone:
        result["description"] = "Nothing selected"

    elif sel_type == ppSelectionSlides:
        sr = sel.SlideRange
        result["slides"] = [sr(i).SlideIndex for i in range(1, sr.Count + 1)]

    elif sel_type == ppSelectionShapes:
        shapes = []
        for i in range(1, sel.ShapeRange.Count + 1):
            s = sel.ShapeRange(i)
            shapes.append({
                "name": s.Name,
                "id": s.Id,
                "type": s.Type,
                "left": s.Left,
                "top": s.Top,
                "width": s.Width,
                "height": s.Height,
            })
        result["shapes"] = shapes

    elif sel_type == ppSelectionText:
        result["text"] = sel.TextRange.Text
        result["parent_shape"] = sel.ShapeRange(1).Name

    return result
```

### 1.5 Working with Selection.ShapeRange

The `ShapeRange` returned by `Selection.ShapeRange` supports all the same operations
as a regular `ShapeRange` -- you can set properties in bulk.

**Python example -- bulk formatting of selected shapes:**
```python
def _format_selected_shapes(fill_color_bgr: int):
    """Apply a fill color to all currently selected shapes."""
    app = ppt._get_app_impl()
    sel = app.ActiveWindow.Selection

    if sel.Type not in (ppSelectionShapes, ppSelectionText):
        raise ValueError("No shapes are selected")

    sr = sel.ShapeRange
    sr.Fill.Visible = True
    sr.Fill.Solid()
    sr.Fill.ForeColor.RGB = fill_color_bgr
```

**Python example -- align and distribute selected shapes:**
```python
# MsoAlignCmd constants
msoAlignLefts = 0
msoAlignCenters = 1
msoAlignRights = 2
msoAlignTops = 3
msoAlignMiddles = 4
msoAlignBottoms = 5

# MsoDistributeCmd constants
msoDistributeHorizontally = 0
msoDistributeVertically = 1

def _align_selected_shapes(align_cmd: int, relative_to_slide: bool = False):
    """Align the currently selected shapes."""
    app = ppt._get_app_impl()
    sel = app.ActiveWindow.Selection
    if sel.Type != ppSelectionShapes:
        raise ValueError("No shapes selected")

    # Align: second param True = relative to slide, False = relative to each other
    sel.ShapeRange.Align(align_cmd, relative_to_slide)

def _distribute_selected_shapes(distribute_cmd: int, relative_to_slide: bool = False):
    """Distribute the currently selected shapes evenly."""
    app = ppt._get_app_impl()
    sel = app.ActiveWindow.Selection
    if sel.Type != ppSelectionShapes:
        raise ValueError("No shapes selected")

    sel.ShapeRange.Distribute(distribute_cmd, relative_to_slide)
```

### 1.6 ActiveWindow.View.GotoSlide

Navigates to a specific slide in the active window.

**Signature:**
```
View.GotoSlide(Index)
```

| Name  | Required | Type | Description                        |
|-------|----------|------|------------------------------------|
| Index | Yes      | Long | 1-based slide number to switch to  |

**Python example:**
```python
def _goto_slide(slide_index: int):
    """Navigate to a specific slide in the active window."""
    app = ppt._get_app_impl()
    app.ActiveWindow.View.GotoSlide(slide_index)
```

**Critical requirement:** You **must** call `GotoSlide` before calling `Shape.Select()`
on shapes of that slide. Selecting a shape on a slide that is not currently displayed
raises a COM error. The view must be in Normal or Slide view (`ppViewNormal` or `ppViewSlide`).

**SlideShowView.GotoSlide** has an additional optional parameter:

| Name       | Required | Type         | Description                                    |
|------------|----------|--------------|------------------------------------------------|
| Index      | Yes      | Long         | Slide number                                   |
| ResetSlide | Optional | MsoTriState  | msoTrue (default): restart animations; msoFalse: resume where left off |

### 1.7 Threading / STA Considerations for Select()

All COM calls in this project are routed through a dedicated STA (Single-Threaded Apartment)
worker thread (see `com_wrapper.py`). This architecture has specific implications for
`Shape.Select()` and related UI operations:

**Key considerations:**

1. **Select() modifies UI state.** Unlike pure data operations (getting/setting properties),
   `Select()` interacts with the PowerPoint window's visual state. This means:
   - PowerPoint must have a visible window (`app.Visible = True`).
   - The window must be in a compatible view (Normal / Slide view).
   - The correct slide must be active via `GotoSlide()`.

2. **All calls from our STA thread are safe.** Since `com_wrapper.py` initializes with
   `pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)` and all COM calls
   are serialized through the single worker thread's queue, there are no cross-apartment
   marshalling issues.

3. **Cannot pass COM objects between threads.** Never return a `Selection`, `ShapeRange`,
   or `Shape` object from `ppt.execute()` to the async caller. Always extract primitive
   data (strings, numbers, dicts) on the COM thread and return those.

4. **UI operations may be slower.** `Select()` triggers screen redraws. Consider using
   `Application.ScreenUpdating = False` around batch operations, though note that this
   is less reliable in PowerPoint than in Excel.

5. **Modal dialogs block the COM thread.** If `Select()` or any operation triggers a
   modal dialog (e.g., a security prompt), the COM thread will hang until the dialog
   is dismissed. The `ppt.execute()` call has a 30-second timeout for this reason.

**Python pattern for safe selection from async code:**
```python
def _select_and_get_info(slide_index: int, shape_indices: list):
    """Select shapes and return info -- all on the COM thread."""
    app = ppt._get_app_impl()
    app.ActiveWindow.View.GotoSlide(slide_index)

    slide = app.ActivePresentation.Slides(slide_index)
    slide.Shapes.Range(shape_indices).Select()

    # Extract data before returning
    sel = app.ActiveWindow.Selection
    shapes = []
    for i in range(1, sel.ShapeRange.Count + 1):
        s = sel.ShapeRange(i)
        shapes.append({"name": s.Name, "id": s.Id})
    return shapes

# Called from async MCP tool:
result = ppt.execute(_select_and_get_info, slide_index, shape_indices)
```

---

## 2. SendKeys via pywin32

### 2.1 WScript.Shell SendKeys

The `WScript.Shell` COM object provides a `SendKeys` method that simulates keyboard input
to the currently active (foreground) window.

**Basic usage:**
```python
import win32com.client
import time

shell = win32com.client.Dispatch("WScript.Shell")

# Send simple text
shell.SendKeys("Hello World")

# Send with modifier keys
shell.SendKeys("^c")       # Ctrl+C (copy)
shell.SendKeys("^v")       # Ctrl+V (paste)
shell.SendKeys("%{F4}")    # Alt+F4 (close)
shell.SendKeys("+{TAB}")   # Shift+Tab
```

**Modifier key symbols:**

| Symbol | Key   | Example       | Meaning      |
|--------|-------|---------------|--------------|
| `+`    | SHIFT | `+{TAB}`      | Shift+Tab    |
| `^`    | CTRL  | `^c`          | Ctrl+C       |
| `%`    | ALT   | `%{F4}`       | Alt+F4       |
| `~`    | ENTER | `~`           | Enter key    |

**Special key codes (curly brace syntax):**

| Key Code          | Key              |
|-------------------|------------------|
| `{BACKSPACE}` / `{BS}` / `{BKSP}` | Backspace |
| `{DELETE}` / `{DEL}`  | Delete      |
| `{ENTER}`         | Enter             |
| `{ESC}`           | Escape            |
| `{TAB}`           | Tab               |
| `{INSERT}` / `{INS}` | Insert        |
| `{UP}`            | Up Arrow          |
| `{DOWN}`          | Down Arrow        |
| `{LEFT}`          | Left Arrow        |
| `{RIGHT}`         | Right Arrow       |
| `{HOME}`          | Home              |
| `{END}`           | End               |
| `{PGUP}`          | Page Up           |
| `{PGDN}`          | Page Down         |
| `{F1}` - `{F16}`  | Function keys    |
| `{CAPSLOCK}`      | Caps Lock         |
| `{NUMLOCK}`       | Num Lock          |
| `{SCROLLLOCK}`    | Scroll Lock       |
| `{PRTSC}`         | Print Screen      |
| `{BREAK}`         | Break             |
| `{HELP}`          | Help              |

**Repeating keys:** `{LEFT 10}` sends Left arrow 10 times.

**Literal braces/special chars:** `{{}` sends `{`, `{}}` sends `}`, `{+}` sends literal `+`.

**Second parameter:** `shell.SendKeys("^c", 0)` -- the second argument controls wait behavior:
- `0` (or omit): Do not wait
- `1`: Wait for keys to be processed before returning

### 2.2 win32api.keybd_event()

Lower-level keyboard simulation using the Windows API directly.

```python
import win32api
import win32con
import time

# Virtual key codes
VK_CONTROL = 0x11
VK_SHIFT = 0x10
VK_ALT = 0x12  # VK_MENU
VK_RETURN = 0x0D
VK_TAB = 0x09
VK_ESCAPE = 0x1B
VK_F5 = 0x74

def press_key(vk_code):
    """Press and release a single key."""
    win32api.keybd_event(vk_code, 0, 0, 0)                          # Key down
    time.sleep(0.05)
    win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)   # Key up

def press_combo(*vk_codes):
    """Press a key combination (e.g., Ctrl+C)."""
    # Press all keys down
    for vk in vk_codes:
        win32api.keybd_event(vk, 0, 0, 0)
        time.sleep(0.02)

    # Release all keys in reverse order
    for vk in reversed(vk_codes):
        win32api.keybd_event(vk, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.02)

# Examples:
press_key(VK_RETURN)                       # Press Enter
press_combo(VK_CONTROL, ord('C'))          # Ctrl+C
press_combo(VK_CONTROL, VK_SHIFT, ord('V'))  # Ctrl+Shift+V
press_combo(VK_ALT, VK_F5)                # Alt+F5
```

**Note:** `win32api.keybd_event()` is considered legacy. The modern replacement is
`SendInput()` via ctypes, but `keybd_event` still works reliably on all current
Windows versions.

### 2.3 ctypes SendInput (Modern Approach)

```python
import ctypes
import time

user32 = ctypes.windll.user32

# Constants
INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_SCANCODE = 0x0008

class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", ctypes.c_ushort),
        ("wScan", ctypes.c_ushort),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]

class INPUT(ctypes.Structure):
    class _INPUT(ctypes.Union):
        _fields_ = [("ki", KEYBDINPUT)]
    _fields_ = [
        ("type", ctypes.c_ulong),
        ("_input", _INPUT),
    ]

def send_key(vk_code):
    """Send a key press and release using SendInput."""
    inputs = (INPUT * 2)()

    # Key down
    inputs[0].type = INPUT_KEYBOARD
    inputs[0]._input.ki.wVk = vk_code

    # Key up
    inputs[1].type = INPUT_KEYBOARD
    inputs[1]._input.ki.wVk = vk_code
    inputs[1]._input.ki.dwFlags = KEYEVENTF_KEYUP

    user32.SendInput(2, ctypes.pointer(inputs[0]), ctypes.sizeof(INPUT))

def send_combo(vk_codes):
    """Send a key combination using SendInput."""
    n = len(vk_codes)
    inputs = (INPUT * (n * 2))()

    # Key down events
    for i, vk in enumerate(vk_codes):
        inputs[i].type = INPUT_KEYBOARD
        inputs[i]._input.ki.wVk = vk

    # Key up events (reverse order)
    for i, vk in enumerate(reversed(vk_codes)):
        inputs[n + i].type = INPUT_KEYBOARD
        inputs[n + i]._input.ki.wVk = vk
        inputs[n + i]._input.ki.dwFlags = KEYEVENTF_KEYUP

    user32.SendInput(n * 2, ctypes.pointer(inputs[0]), ctypes.sizeof(INPUT))
```

### 2.4 pyautogui for Keyboard/Mouse Simulation

`pyautogui` is a cross-platform library for GUI automation. It provides a simpler API
than raw win32api calls.

```python
import pyautogui
import time

# Set global pause between actions (default: 0.1 seconds)
pyautogui.PAUSE = 0.1

# Type text
pyautogui.typewrite("Hello World", interval=0.05)

# Press a single key
pyautogui.press("enter")
pyautogui.press("tab")
pyautogui.press("f5")

# Key combinations (hotkeys)
pyautogui.hotkey("ctrl", "c")         # Ctrl+C
pyautogui.hotkey("ctrl", "shift", "v")  # Ctrl+Shift+V
pyautogui.hotkey("alt", "f4")         # Alt+F4

# Hold a key down
pyautogui.keyDown("shift")
pyautogui.press("tab")
pyautogui.press("tab")
pyautogui.keyUp("shift")

# Mouse click at position
pyautogui.click(x=100, y=200)
pyautogui.doubleClick(x=100, y=200)
pyautogui.rightClick(x=100, y=200)
```

**Advantages over SendKeys:**
- Cross-platform (works on Linux/macOS too, though not relevant for this project).
- Built-in safety feature: moving mouse to corner of screen raises `FailSafeException`.
- Built-in `PAUSE` delay prevents overwhelming the target application.
- Simpler API for complex key combinations.

**Disadvantages:**
- Requires pip install (`pip install pyautogui`).
- Mouse operations can interfere with user activity.
- Cannot target specific windows -- always sends to the foreground window.
- Not reliable on multi-monitor setups for mouse operations.
- Can be "too fast" for some applications to process.

### 2.5 Can SendKeys Trigger Ribbon Commands?

Yes, but with significant limitations:

**Keyboard shortcuts work:**
```python
shell = win32com.client.Dispatch("WScript.Shell")

# Ctrl+G -- Group selected objects (if supported shortcut)
shell.SendKeys("^g")

# F5 -- Start slideshow
shell.SendKeys("{F5}")

# Alt key sequences for ribbon navigation (KeyTips)
shell.SendKeys("%")           # Activate ribbon KeyTips
time.sleep(0.3)
shell.SendKeys("H")           # Home tab
time.sleep(0.3)
shell.SendKeys("GA")          # Arrange > Align
```

**Alt+Key ribbon navigation (KeyTips):**
PowerPoint supports sequential Alt-key access to ribbon commands. Pressing Alt activates
"KeyTips" (letter overlays on the ribbon), and then you type the letters to navigate.

```python
def trigger_ribbon_via_keytips(shell, *keys):
    """Navigate the ribbon using KeyTip sequences."""
    shell.SendKeys("%")  # Activate KeyTips
    time.sleep(0.5)      # Wait for KeyTips to appear
    for key in keys:
        shell.SendKeys(key)
        time.sleep(0.3)  # Wait between each key
```

**Limitations of SendKeys for ribbon commands:**
- KeyTip sequences vary by language/locale.
- Sequences can change between Office versions.
- Timing is critical -- too fast and keys are missed, too slow and KeyTips timeout.
- If a dialog opens, the remaining keys are sent to the dialog instead of the ribbon.
- Not reliable for automated testing or production use.

### 2.6 Ensuring PowerPoint Window is Focused

Before sending keys, you must ensure PowerPoint is the foreground window.

**Method 1: Using WScript.Shell.AppActivate:**
```python
import win32com.client
import time

shell = win32com.client.Dispatch("WScript.Shell")

# AppActivate by window title (partial match)
shell.AppActivate("PowerPoint")
time.sleep(0.5)  # Wait for focus
shell.SendKeys("^c")
```

**Method 2: Using win32gui.SetForegroundWindow:**
```python
import win32gui
import win32con
import time

def focus_powerpoint():
    """Bring the PowerPoint window to the foreground."""
    hwnd = win32gui.FindWindow("PPTFrameClass", None)
    if hwnd == 0:
        raise RuntimeError("PowerPoint window not found")

    # If window is minimized, restore it first
    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

    # SetForegroundWindow may fail if our process is not the foreground.
    # Workaround: simulate a key press first to "unlock" foreground access.
    import win32api
    win32api.keybd_event(0, 0, 0, 0)  # Dummy key event
    time.sleep(0.05)

    win32gui.SetForegroundWindow(hwnd)
    time.sleep(0.3)  # Wait for window to come to front
```

**Method 3: Using PowerPoint COM Application.Activate:**
```python
def _focus_powerpoint_via_com():
    """Use the COM Application object to activate PowerPoint."""
    app = ppt._get_app_impl()
    app.Activate()
    # Additionally activate the first window if available
    if app.Windows.Count > 0:
        app.ActiveWindow.Activate()
```

**Important caveats about SetForegroundWindow:**
- Windows restricts `SetForegroundWindow()` -- only the foreground process can set another
  window to the foreground. A background process attempting this will cause the taskbar
  button to flash instead.
- Workaround: Use `win32api.keybd_event()` with a dummy key press before calling
  `SetForegroundWindow()`. This tricks Windows into thinking the calling process has
  keyboard focus.
- Another workaround: Use `AttachThreadInput()` to attach to the foreground thread first.

### 2.7 Timing Issues and Reliability Concerns

**Summary of reliability issues across all SendKeys approaches:**

| Issue                        | WScript.Shell | keybd_event | ctypes SendInput | pyautogui |
|------------------------------|:-------------:|:-----------:|:----------------:|:---------:|
| Requires foreground window   | Yes           | Yes         | Yes              | Yes       |
| Timing-sensitive             | High          | Medium      | Medium           | Medium    |
| Interferes with user input   | Yes           | Yes         | Yes              | Yes       |
| Language/locale dependent    | For Alt-keys  | No          | No               | No        |
| Can target background window | No            | No          | No               | No        |
| Modal dialog blocking        | Yes           | Yes         | Yes              | Yes       |

**Best practices for reliability:**
1. Always add `time.sleep()` delays between key events (0.1-0.5 seconds).
2. Focus the target window before sending keys.
3. Verify the expected state after sending keys (check selection, etc.).
4. Use COM object model directly when possible -- SendKeys should be a last resort.
5. Consider disabling screen saver and other UI interruptions during automation.

---

## 3. Ribbon / Menu Operations via COM

### 3.1 CommandBars.ExecuteMso()

The `ExecuteMso` method executes a built-in ribbon command identified by its `idMso` string.
This is the **recommended approach** for triggering ribbon commands programmatically when
there is no direct object model equivalent.

**Signature:**
```
Application.CommandBars.ExecuteMso(idMso)
```

| Name  | Required | Type   | Description                    |
|-------|----------|--------|--------------------------------|
| idMso | Yes      | String | Identifier for the control     |

**Return value:** None. On failure, returns `E_InvalidArg` for an invalid idMso,
or `E_Fail` for controls that are not enabled or not visible.

**Python (pywin32) example:**
```python
def _execute_mso_command(command_id: str):
    """Execute a ribbon command by its idMso identifier."""
    app = ppt._get_app_impl()
    try:
        app.CommandBars.ExecuteMso(command_id)
    except Exception as e:
        raise RuntimeError(
            f"ExecuteMso failed for '{command_id}': {e}. "
            "The command may be invalid, disabled, or requires a specific selection state."
        )
```

**Usage examples:**
```python
# Copy/Cut/Paste
ppt.execute(_execute_mso_command, "Copy")
ppt.execute(_execute_mso_command, "Cut")
ppt.execute(_execute_mso_command, "Paste")
ppt.execute(_execute_mso_command, "PasteSpecialDialog")

# Format Painter
ppt.execute(_execute_mso_command, "FormatPainter")

# Grouping
ppt.execute(_execute_mso_command, "ObjectsGroup")
ppt.execute(_execute_mso_command, "ObjectsUngroup")
ppt.execute(_execute_mso_command, "ObjectsRegroup")

# Z-Order
ppt.execute(_execute_mso_command, "ObjectBringToFront")
ppt.execute(_execute_mso_command, "ObjectSendToBack")
ppt.execute(_execute_mso_command, "ObjectBringForward")
ppt.execute(_execute_mso_command, "ObjectSendBackward")

# Alignment (requires shapes to be selected)
ppt.execute(_execute_mso_command, "ObjectsAlignLeftSmart")
ppt.execute(_execute_mso_command, "ObjectsAlignRightSmart")
ppt.execute(_execute_mso_command, "ObjectsAlignTopSmart")
ppt.execute(_execute_mso_command, "ObjectsAlignBottomSmart")
ppt.execute(_execute_mso_command, "ObjectsAlignCenterHorizontalSmart")
ppt.execute(_execute_mso_command, "ObjectsAlignMiddleVerticalSmart")

# Distribution
ppt.execute(_execute_mso_command, "AlignDistributeHorizontally")
ppt.execute(_execute_mso_command, "AlignDistributeVertically")

# Rotation
ppt.execute(_execute_mso_command, "ObjectRotateRight90")
ppt.execute(_execute_mso_command, "ObjectRotateLeft90")
ppt.execute(_execute_mso_command, "ObjectFlipHorizontal")
ppt.execute(_execute_mso_command, "ObjectFlipVertical")

# Selection pane
ppt.execute(_execute_mso_command, "SelectionPane")

# Slide operations
ppt.execute(_execute_mso_command, "SlideNew")
ppt.execute(_execute_mso_command, "SlideDelete")
ppt.execute(_execute_mso_command, "DuplicateSelectedSlides")

# View operations
ppt.execute(_execute_mso_command, "ViewSlideShowView")
ppt.execute(_execute_mso_command, "ViewSlideSorterView")
ppt.execute(_execute_mso_command, "SlideShowFromBeginning")
ppt.execute(_execute_mso_command, "SlideShowFromCurrent")
```

### 3.2 Comprehensive idMso Reference for PowerPoint

Below is a categorized list of commonly useful idMso command identifiers.

#### Clipboard & Editing

| idMso                    | Type         | Description                  |
|--------------------------|--------------|------------------------------|
| `Copy`                   | button       | Copy                         |
| `Cut`                    | button       | Cut                          |
| `Paste`                  | button       | Paste                        |
| `PasteSpecialDialog`     | button       | Paste Special...             |
| `PasteAsHyperlink`       | button       | Paste as Hyperlink           |
| `PasteDuplicate`         | button       | Duplicate                    |
| `SelectAll`              | button       | Select All                   |
| `Undo`                   | gallery      | Undo                         |
| `Redo`                   | gallery      | Redo                         |
| `FindDialog`             | button       | Find...                      |
| `ReplaceDialog`          | button       | Replace...                   |
| `ClearFormatting`        | button       | Clear Formatting             |
| `ShowClipboard`          | button       | Office Clipboard...          |

#### Formatting & Style

| idMso                          | Type         | Description               |
|--------------------------------|--------------|---------------------------|
| `FormatPainter`                | toggleButton | Format Painter            |
| `PickUpStyle`                  | button       | Pick Up Style             |
| `PasteApplyStyle`              | button       | Apply Style               |
| `Bold`                         | toggleButton | Bold                      |
| `Italic`                       | toggleButton | Italic                    |
| `Underline`                    | toggleButton | Underline                 |
| `Strikethrough`                | toggleButton | Strikethrough             |
| `Superscript`                  | toggleButton | Superscript               |
| `Subscript`                    | toggleButton | Subscript                 |
| `AlignLeft`                    | toggleButton | Align Left                |
| `AlignCenter`                  | toggleButton | Center                    |
| `AlignRight`                   | toggleButton | Align Right               |
| `AlignJustify`                 | toggleButton | Justify                   |
| `FontSizeIncrease`             | button       | Increase Font Size        |
| `FontSizeDecrease`             | button       | Decrease Font Size        |
| `IndentIncrease`               | button       | Increase Indent           |
| `IndentDecrease`               | button       | Decrease Indent           |
| `ChangeCaseToggle`             | button       | Toggle Case               |
| `CharacterFormattingReset`     | button       | Reset Character Formatting|

#### Shape Arrangement

| idMso                                    | Type         | Description                 |
|------------------------------------------|--------------|-----------------------------|
| `ObjectsGroup`                           | button       | Group                       |
| `ObjectsUngroup`                         | button       | Ungroup                     |
| `ObjectsRegroup`                         | button       | Regroup                     |
| `ObjectBringToFront`                     | button       | Bring to Front              |
| `ObjectBringForward`                     | button       | Bring Forward               |
| `ObjectSendToBack`                       | button       | Send to Back                |
| `ObjectSendBackward`                     | button       | Send Backward               |
| `ObjectRotateRight90`                    | button       | Rotate Right 90 degrees     |
| `ObjectRotateLeft90`                     | button       | Rotate Left 90 degrees      |
| `ObjectFlipHorizontal`                   | button       | Flip Horizontal             |
| `ObjectFlipVertical`                     | button       | Flip Vertical               |
| `ObjectRotateFree`                       | button       | Free Rotate                 |
| `ObjectsAlignLeftSmart`                  | button       | Align Left                  |
| `ObjectsAlignRightSmart`                 | button       | Align Right                 |
| `ObjectsAlignTopSmart`                   | button       | Align Top                   |
| `ObjectsAlignBottomSmart`                | button       | Align Bottom                |
| `ObjectsAlignCenterHorizontalSmart`      | button       | Align Center                |
| `ObjectsAlignMiddleVerticalSmart`        | button       | Align Middle                |
| `AlignDistributeHorizontally`            | button       | Distribute Horizontally     |
| `AlignDistributeVertically`              | button       | Distribute Vertically       |
| `ObjectsAlignSelectedSmart`              | toggleButton | Align Selected Objects      |
| `ObjectsAlignRelativeToContainerSmart`   | toggleButton | Align to Slide              |
| `ObjectNudgeUp`                          | button       | Nudge Up                    |
| `ObjectNudgeDown`                        | button       | Nudge Down                  |
| `ObjectNudgeLeft`                        | button       | Nudge Left                  |
| `ObjectNudgeRight`                       | button       | Nudge Right                 |
| `ObjectSizeAndPositionDialog`            | button       | Size and Position...        |

#### Selection & Navigation

| idMso                    | Type         | Description                  |
|--------------------------|--------------|------------------------------|
| `SelectionPane`          | toggleButton | Selection Pane...            |
| `ObjectsSelect`          | toggleButton | Select Objects               |
| `ObjectsMultiSelect`     | button       | Select Multiple Objects      |

#### Insert Operations

| idMso                              | Type         | Description            |
|------------------------------------|--------------|------------------------|
| `SlideNew`                         | button       | New Slide              |
| `TextBoxInsert`                    | toggleButton | Text Box               |
| `TableInsert`                      | button       | Insert Table...        |
| `ChartInsert`                      | button       | Chart...               |
| `SmartArtInsert`                   | button       | SmartArt...            |
| `PictureInsertFromFilePowerPoint`  | button       | Picture...             |
| `HyperlinkInsert`                  | button       | Hyperlink...           |
| `SymbolInsert`                     | button       | Symbol...              |
| `ActionInsert`                     | button       | Action                 |
| `WordArtInsertGallery`             | gallery      | WordArt                |
| `MovieFromFileInsert`              | button       | Movie from File...     |
| `SoundInsertFromFile`              | button       | Sound from File...     |

#### Slide & Presentation

| idMso                         | Type         | Description                   |
|-------------------------------|--------------|-------------------------------|
| `SlideDelete`                 | button       | Delete Slide                  |
| `SlideHide`                   | toggleButton | Hide Slide                    |
| `SlideReset`                  | button       | Reset Slide                   |
| `DuplicateSelectedSlides`     | button       | Duplicate Selected Slides     |
| `SlideBackgroundFormatDialog` | button       | Format Background...          |
| `SlideShowFromBeginning`      | button       | From Beginning                |
| `SlideShowFromCurrent`        | button       | From Current Slide            |
| `SlideShowSetUpDialog`        | button       | Set Up Slide Show...          |
| `SlideShowRehearseTimings`    | button       | Rehearse Timings              |
| `HeaderFooterInsert`          | button       | Header & Footer...            |

#### File Operations

| idMso                    | Type         | Description               |
|--------------------------|--------------|---------------------------|
| `FileSave`               | button       | Save                      |
| `FileSaveAs`             | button       | Save As                   |
| `FileOpen`               | button       | Open                      |
| `FileNew`                | button       | New                       |
| `FileClose`              | button       | Close                     |
| `FilePrint`              | button       | Print                     |
| `FileSaveAsPdfOrXps`     | button       | Publish as PDF or XPS     |

#### View & Window

| idMso                    | Type         | Description               |
|--------------------------|--------------|---------------------------|
| `ViewNormalViewPowerPoint` | toggleButton | Normal                  |
| `ViewSlideSorterView`   | toggleButton | Slide Sorter              |
| `ViewNotesPageView`     | toggleButton | Notes Page                |
| `ViewSlideShowView`     | button       | Slide Show                |
| `ViewSlideMasterView`   | toggleButton | Slide Master              |
| `ViewRulerPowerPoint`   | checkBox     | Ruler                     |
| `ViewGridlinesPowerPoint` | checkBox   | Gridlines                 |
| `ZoomDialog`             | button       | Zoom...                   |
| `ZoomFitToWindow`        | button       | Fit to Window             |

### 3.3 How to Discover idMso Command IDs

**Method 1: Customize Ribbon dialog**
In PowerPoint, go to File > Options > Customize Ribbon. Hover over any command to see
its ScreenTip, which often corresponds to the idMso. You can also right-click any ribbon
button and select "Add to Quick Access Toolbar" -- the tooltip shows the command name.

**Method 2: Microsoft's official control identifier lists**
- [PowerPoint 2007 idMso list (MS-CUSTOMUI)](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui/f2a8e3c0-14cb-4ad3-88cd-a8b5b1b9a8a0)
- [Office Fluent UI Command Identifiers (GitHub)](https://github.com/OfficeDev/office-fluent-ui-command-identifiers)
- Microsoft downloads for Office 2010/2013/2016 contain Excel files with all control IDs.

**Method 3: Trial and error with VBA Immediate Window**
```vb
' In PowerPoint VBA Immediate Window:
Application.CommandBars.ExecuteMso "FormatPainter"
```
If the command name is invalid, you get a runtime error. If it requires a selection or
context that is not present, you get a different error.

**Method 4: Enumerate CommandBars programmatically (limited)**
```python
def _list_command_bars():
    """List all available CommandBars (legacy menus/toolbars)."""
    app = ppt._get_app_impl()
    bars = []
    for i in range(1, app.CommandBars.Count + 1):
        bar = app.CommandBars(i)
        bars.append({
            "index": i,
            "name": bar.Name,
            "type": bar.Type,
            "visible": bool(bar.Visible),
            "enabled": bool(bar.Enabled),
        })
    return bars
```

Note: This lists legacy CommandBar objects (toolbars and context menus), not Ribbon
idMso values. The Ribbon command IDs are not enumerable through the object model.

### 3.4 ExecuteMso Practical Examples

**Select shapes then align them:**
```python
def _select_and_align_left(slide_index: int, shape_indices: list):
    """Select multiple shapes and align them to the left."""
    app = ppt._get_app_impl()

    # Navigate and select
    app.ActiveWindow.View.GotoSlide(slide_index)
    slide = app.ActivePresentation.Slides(slide_index)
    slide.Shapes.Range(shape_indices).Select()

    # Execute alignment command
    app.CommandBars.ExecuteMso("ObjectsAlignLeftSmart")
```

**Format Painter workflow:**
```python
def _apply_format_painter(slide_index: int, source_shape: int, target_shape: int):
    """Use Format Painter to copy formatting from one shape to another."""
    app = ppt._get_app_impl()
    slide = app.ActivePresentation.Slides(slide_index)

    app.ActiveWindow.View.GotoSlide(slide_index)

    # Select source shape
    slide.Shapes(source_shape).Select()

    # Activate Format Painter (it becomes a toggle)
    app.CommandBars.ExecuteMso("FormatPainter")

    # Note: Format Painter via ExecuteMso activates the mode,
    # but you need to click the target shape to apply it.
    # This means Format Painter via COM is not fully automatable
    # without also using SendKeys or mouse simulation to click
    # the target shape. For programmatic formatting transfer,
    # prefer copying individual formatting properties directly.
```

**Group/Ungroup workflow:**
```python
def _group_shapes(slide_index: int, shape_indices: list):
    """Group selected shapes using ExecuteMso."""
    app = ppt._get_app_impl()
    slide = app.ActivePresentation.Slides(slide_index)

    app.ActiveWindow.View.GotoSlide(slide_index)
    slide.Shapes.Range(shape_indices).Select()

    # Group the selection
    app.CommandBars.ExecuteMso("ObjectsGroup")

    # Note: For programmatic grouping, prefer the object model:
    # slide.Shapes.Range(shape_indices).Group()
    # which returns the new Group shape and doesn't require selection.
```

### 3.5 CommandBars.FindControl() Approach (Legacy)

The `FindControl` method searches for a command bar control by type and/or ID.
This is the legacy approach from pre-Ribbon Office versions.

```python
def _find_and_execute_control(control_id: int):
    """Find a legacy CommandBar control by its ID and execute it."""
    app = ppt._get_app_impl()

    # FindControl(Type, Id, Tag, Visible, Recursive)
    ctrl = app.CommandBars.FindControl(Id=control_id)
    if ctrl is not None:
        ctrl.Execute()
    else:
        raise RuntimeError(f"Control with ID {control_id} not found")
```

**Limitations:**
- Most Ribbon controls do not have legacy CommandBar equivalents.
- Control IDs differ between Office versions.
- `ExecuteMso` is the preferred approach for any Ribbon-era command.
- `FindControl` is mainly useful for context menu items and older toolbar commands.

### 3.6 ExecuteMso vs Direct Object Model

| Capability                  | ExecuteMso                     | Direct Object Model               |
|-----------------------------|--------------------------------|------------------------------------|
| Group shapes                | Requires selection first       | `ShapeRange.Group()` -- no selection needed |
| Align shapes                | Requires selection first       | `ShapeRange.Align()` -- no selection needed |
| Distribute shapes           | Requires selection first       | `ShapeRange.Distribute()` -- no selection needed |
| Format Painter              | Activates mode only            | Copy properties individually       |
| Rotate 90 degrees           | Works on selection             | `Shape.Rotation += 90`            |
| Flip shapes                 | Works on selection             | `Shape.Flip(msoFlipHorizontal)`   |
| Paste Special               | Opens dialog (blocking)        | `View.PasteSpecial()` with params |
| Selection Pane              | Opens/closes pane              | No object model equivalent         |
| Z-order                     | Requires selection             | `Shape.ZOrder(mso*)`              |
| Undo/Redo                   | Executes undo/redo             | No object model equivalent         |

**Recommendation:** Use the direct object model whenever possible. `ExecuteMso` is best
reserved for operations that have **no object model equivalent**, such as:
- `SelectionPane` (toggle the selection pane)
- `Undo` / `Redo`
- `FormatPainter` (though limited without mouse interaction)
- `FileSaveAsPdfOrXps` (though `Presentation.ExportAsFixedFormat` exists)
- Certain dialog boxes (`PasteSpecialDialog`, `SlideBackgroundFormatDialog`)

---

## 4. Comparison and Recommendations

### 4.1 When to Use Each Approach

| Approach                     | Use Case                                              | Reliability |
|------------------------------|-------------------------------------------------------|-------------|
| **Direct COM Object Model**  | Any operation with a documented API                   | Highest     |
| **ExecuteMso**               | Ribbon commands without object model equivalent       | High        |
| **Shape.Select + ShapeRange**| When you need selection state for subsequent ExecuteMso| High        |
| **SendKeys (WScript.Shell)** | Last resort for keyboard-shortcut-only operations     | Low         |
| **keybd_event / SendInput**  | Alternative to SendKeys, slightly more control        | Low         |
| **pyautogui**                | Mouse simulation, screenshot-based automation         | Low         |

### 4.2 Recommended Architecture for MCP Tools

For the ppt-com-mcp server, the recommended approach is:

1. **Primary: Use the COM object model directly** for all operations that have API methods
   (shapes, text, formatting, slides, charts, etc.).

2. **Secondary: Use ExecuteMso** for operations that are only available through the ribbon
   (undo/redo, selection pane, certain dialogs).

3. **Tertiary: Use Shape.Select() + ExecuteMso** for operations that require selection
   context (alignment/distribution when you want the "smart" behavior that considers
   the slide boundaries).

4. **Avoid SendKeys** in production MCP tools. It is unreliable, requires foreground focus,
   and interferes with user activity. If absolutely necessary, document it clearly and
   add appropriate error handling.

### 4.3 Implementation Pattern for Selection-Based Operations

```python
def _select_and_execute(
    slide_index: int,
    shape_indices: list,
    mso_command: str,
) -> dict:
    """Generic pattern: select shapes, execute a ribbon command, return result."""
    app = ppt._get_app_impl()

    # 1. Navigate to slide
    app.ActiveWindow.View.GotoSlide(slide_index)

    # 2. Select shapes
    slide = app.ActivePresentation.Slides(slide_index)
    shape_range = slide.Shapes.Range(shape_indices)
    shape_range.Select()

    # 3. Execute ribbon command
    app.CommandBars.ExecuteMso(mso_command)

    # 4. Return updated state
    result_shapes = []
    for idx in shape_indices:
        s = slide.Shapes(idx)
        result_shapes.append({
            "name": s.Name,
            "left": s.Left,
            "top": s.Top,
            "width": s.Width,
            "height": s.Height,
            "rotation": s.Rotation,
        })
    return {"success": True, "shapes": result_shapes}
```

---

## References

- [Shape.Select method (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.select)
- [ShapeRange.Select method (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shaperange.select)
- [Selection object (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.selection)
- [Selection.ShapeRange property (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Selection.ShapeRange)
- [Shapes.Range method (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shapes.range)
- [View.GotoSlide method (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.view.gotoslide)
- [CommandBars.ExecuteMso method (Office) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/office.commandbars.executemso)
- [PowerPoint 2007 idMso list (MS-CUSTOMUI) - Microsoft Learn](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui/f2a8e3c0-14cb-4ad3-88cd-a8b5b1b9a8a0)
- [PowerPoint idMso Control List for 2013 and 2010 - YOUpresent](http://youpresent.co.uk/idmso-control-list-powerpoint-2013-2010/)
- [PowerPoint Button idMsos - RibbonCreator](https://www.ribboncreator2010.de/Onlinehelp/EN/_39c0kjx66.htm)
- [Office Fluent UI Command Identifiers - GitHub](https://github.com/OfficeDev/office-fluent-ui-command-identifiers)
- [Controlling Applications via SendKeys - win32com.goermezer.de](https://win32com.goermezer.de/microsoft/windows/controlling-applications-via-sendkeys.html)
- [Python win32api keybd_event example - GitHub Gist](https://gist.github.com/chriskiehl/2906125)
- [PyAutoGUI documentation](https://pyautogui.readthedocs.io/en/latest/quickstart.html)
- [PyWin32 - How to Bring a Window to Front](https://www.blog.pythonlibrary.org/2014/10/20/pywin32-how-to-bring-a-window-to-front/)
- [keybd_event function (Win32) - Microsoft Learn](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-keybd_event)
