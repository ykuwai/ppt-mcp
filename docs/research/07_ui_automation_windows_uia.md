# Windows UI Automation for PowerPoint Control
# UI Automation / pywinauto / Accessibility API / Low-Level Input

> Research Date: 2026-02-17
> Purpose: Evaluate Windows UI Automation approaches for controlling PowerPoint beyond COM object model

---

## Table of Contents

1. [Windows UI Automation Framework](#1-windows-ui-automation-framework)
2. [pywinauto for PowerPoint](#2-pywinauto-for-powerpoint)
3. [uiautomation Library](#3-uiautomation-library)
4. [Accessibility API Approach (MSAA vs UIA)](#4-accessibility-api-approach-msaa-vs-uia)
5. [ctypes / win32gui Direct Approach](#5-ctypes--win32gui-direct-approach)
6. [Low-Level Input Simulation](#6-low-level-input-simulation)
7. [Practical Recipes](#7-practical-recipes)
8. [Comparison and Recommendations](#8-comparison-and-recommendations)

---

## 1. Windows UI Automation Framework

### 1.1 Overview

Windows UI Automation (UIA) is Microsoft's accessibility and test automation framework, the successor to Microsoft Active Accessibility (MSAA). It provides programmatic access to UI elements in Windows applications through a tree of automation elements.

Key concepts:
- **Automation Elements**: Every UI element (buttons, text fields, menus, etc.) is represented as an automation element in a hierarchical tree rooted at the desktop
- **Control Patterns**: Interfaces that define element capabilities (InvokePattern for clickable elements, ValuePattern for editable text, etc.)
- **Properties**: Name, ClassName, AutomationId, ControlType, BoundingRectangle, etc.
- **Events**: Notifications when UI state changes

### 1.2 Python Libraries for UIA

There are three main Python libraries for accessing UIA:

| Library | Approach | PyPI Package | Notes |
|---------|----------|-------------|-------|
| **pywinauto** | High-level wrapper with UIA backend | `pywinauto` | Most mature, both win32 and UIA backends |
| **uiautomation** | Direct UIA COM wrapper | `uiautomation` | Lighter weight, closer to raw API |
| **comtypes** | Raw COM interface to UIAutomation | `comtypes` | Lowest level, maximum control |

#### comtypes (low-level)

```python
import comtypes
from comtypes.client import CreateObject, GetModule

# Generate UIA type library bindings
GetModule('UIAutomationCore.dll')
from comtypes.gen.UIAutomationClient import CUIAutomation

# Create IUIAutomation instance
uia = CreateObject(CUIAutomation)

# Get root element (Desktop)
root = uia.GetRootElement()
print(f"Root: {root.CurrentName}")

# Find PowerPoint window using condition
condition = uia.CreatePropertyCondition(
    30005,  # UIA_NamePropertyId
    "PowerPoint"
)
ppt_element = root.FindFirst(
    4,  # TreeScope_Descendants
    condition
)
```

#### uiautomation (mid-level)

```python
import uiautomation as auto

# Find PowerPoint main window
ppt_window = auto.WindowControl(
    searchDepth=1,
    SubName='PowerPoint'
)

if ppt_window.Exists(maxSearchSeconds=5):
    print(f"Found: {ppt_window.Name}")
    # Print the full control tree
    ppt_window.GetChildren()
```

#### pywinauto (high-level)

```python
from pywinauto import Application

# Connect to running PowerPoint
app = Application(backend="uia").connect(
    path="POWERPNT.EXE"
)
main_window = app.window(title_re=".*PowerPoint.*")
```

### 1.3 Inspection Tools

Before automating, you need to discover the UI element tree. Use these tools:

| Tool | Source | Notes |
|------|--------|-------|
| **Accessibility Insights for Windows** | Microsoft (free) | Modern replacement for Inspect.exe. Best for exploring UIA trees |
| **Inspect.exe** | Windows SDK | Legacy but functional. Shows both UIA and MSAA properties |
| **Spy++** | Visual Studio | Win32 window hierarchy (class names, styles, messages) |
| **pywinauto `print_control_identifiers()`** | pywinauto | Dumps element tree from Python |
| **uiautomation CLI** | uiautomation package | `python -m uiautomation -t 0 -n` prints active window controls |

### 1.4 Finding PowerPoint UI Elements

PowerPoint's UI hierarchy (as seen by UIA) typically looks like:

```
Desktop
  └── Window "Presentation1 - PowerPoint"  (PPTFrameClass)
        ├── Pane "Ribbon"
        │     ├── Tab "Home"
        │     ├── Tab "Insert"
        │     ├── Tab "Design"
        │     │     └── Group "Themes"
        │     │           └── Button "Theme 1"
        │     ...
        │     └── Group "..."
        │           └── Button "..."
        ├── Pane "Slide Panel"  (mdiViewWnd)
        │     └── Pane "Slide"  (paneClassDC)
        ├── Pane "Slides Panel" (thumbnail strip on left)
        ├── Pane "Notes"
        └── StatusBar
```

#### Discovering the tree with pywinauto

```python
from pywinauto import Application

app = Application(backend="uia").connect(path="POWERPNT.EXE")
main_win = app.window(title_re=".*PowerPoint.*")

# Print full control hierarchy (can be very large)
main_win.print_control_identifiers()

# Print only top-level children (depth=1)
main_win.print_control_identifiers(depth=1)
```

#### Discovering the tree with uiautomation

```python
import uiautomation as auto

ppt = auto.WindowControl(searchDepth=1, SubName='PowerPoint')
children = ppt.GetChildren()
for child in children:
    print(f"  {child.ControlTypeName}: '{child.Name}' "
          f"[{child.ClassName}]")
    for grandchild in child.GetChildren():
        print(f"    {grandchild.ControlTypeName}: "
              f"'{grandchild.Name}' [{grandchild.ClassName}]")
```

### 1.5 Interacting with Ribbon Buttons

The Office ribbon is exposed through UIA as a hierarchy of Panes, Tabs, Groups, and Buttons.

```python
import uiautomation as auto

ppt = auto.WindowControl(searchDepth=1, SubName='PowerPoint')

# Click a ribbon tab
insert_tab = ppt.TabItemControl(Name='Insert')
if insert_tab.Exists(3):
    insert_tab.Click()

# Find and click a button in the ribbon
shapes_button = ppt.ButtonControl(Name='Shapes')
if shapes_button.Exists(3):
    shapes_button.Click()
```

### 1.6 Reading Ribbon State

```python
import uiautomation as auto

ppt = auto.WindowControl(searchDepth=1, SubName='PowerPoint')

# Check which ribbon tab is active
# Active tab typically has SelectionItemPattern with IsSelected=True
tabs = ppt.TabItemControl(searchDepth=10)
# Iterate to find selected tab
for tab in ppt.GetChildren():
    if tab.ControlTypeName == 'TabItem':
        try:
            pattern = tab.GetSelectionItemPattern()
            if pattern.IsSelected:
                print(f"Active tab: {tab.Name}")
        except Exception:
            pass

# Check if a toggle button is pressed
bold_button = ppt.ButtonControl(Name='Bold')
if bold_button.Exists(2):
    try:
        toggle = bold_button.GetTogglePattern()
        print(f"Bold is: {'ON' if toggle.ToggleState else 'OFF'}")
    except Exception:
        print("Bold button does not support TogglePattern")
```

---

## 2. pywinauto for PowerPoint

### 2.1 Backend Selection: "uia" vs "win32"

pywinauto offers two backends for accessing application controls:

| Feature | `backend="win32"` | `backend="uia"` |
|---------|-------------------|-----------------|
| Technology | Win32 API (SendMessage, etc.) | Microsoft UI Automation COM |
| Control discovery | Window class names, control IDs | UIA properties, AutomationId |
| Ribbon support | Limited (owner-drawn controls) | Good (ribbon is UIA-aware) |
| Modern apps | No (Win32 only) | Yes (WPF, UWP, etc.) |
| Performance | Faster for simple Win32 apps | Slower but more reliable |
| PowerPoint recommendation | **Not recommended for ribbon** | **Recommended** |

**For PowerPoint, always use `backend="uia"`** because the Office ribbon is built with custom rendering that does not expose standard Win32 controls, but does expose UIA automation elements.

### 2.2 Connecting to PowerPoint

```python
from pywinauto import Application, Desktop
import subprocess
import time

# Option 1: Start PowerPoint
app = Application(backend="uia").start(
    r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
)

# Option 2: Connect to running PowerPoint by executable path
app = Application(backend="uia").connect(path="POWERPNT.EXE")

# Option 3: Connect by window title regex
app = Application(backend="uia").connect(
    title_re=".*PowerPoint.*"
)

# Option 4: Connect by process ID
app = Application(backend="uia").connect(process=12345)

# Option 5: Desktop-based access (cross-process, no Application needed)
dlg = Desktop(backend="uia").window(title_re=".*PowerPoint.*")
```

### 2.3 Finding and Clicking Ribbon Buttons

```python
from pywinauto import Application

app = Application(backend="uia").connect(path="POWERPNT.EXE")
main_win = app.window(title_re=".*PowerPoint.*")

# --- Switch to the Insert tab ---
insert_tab = main_win.child_window(
    title="Insert",
    control_type="TabItem"
)
insert_tab.click_input()

# --- Click a ribbon button by name ---
shapes_btn = main_win.child_window(
    title="Shapes",
    control_type="Button"
)
shapes_btn.click_input()

# --- Find button within a specific ribbon group ---
# First find the group
clipboard_group = main_win.child_window(
    title="Clipboard",
    control_type="Group"
)
# Then find button within that group
paste_btn = clipboard_group.child_window(
    title="Paste",
    control_type="Button"
)
paste_btn.click_input()

# --- Using best_match (fuzzy name matching) ---
main_win.child_window(best_match="InsertTab").click_input()
```

### 2.4 Working with Dialog Boxes

PowerPoint dialogs (Format Shape, Insert Picture, Save As, etc.) appear as separate top-level windows or child windows.

```python
from pywinauto import Application
import time

app = Application(backend="uia").connect(path="POWERPNT.EXE")
main_win = app.window(title_re=".*PowerPoint.*")

# --- Open Format Shape dialog via right-click context menu ---
# (Assuming a shape is selected)
main_win.type_keys("+{F10}")  # Shift+F10 = right-click
time.sleep(0.5)

# Find and click "Format Shape..." in context menu
format_shape = app.window(title="Format Shape")
format_shape.wait('visible', timeout=5)

# Interact with dialog controls
format_shape.print_control_identifiers()  # Discover controls

# Example: Set width in the Format Shape dialog
width_edit = format_shape.child_window(
    title="Width",
    control_type="Edit"
)
width_edit.set_text("5")

# Close dialog
format_shape.child_window(title="Close", control_type="Button").click()

# --- File Open dialog ---
main_win.type_keys("^o")  # Ctrl+O
time.sleep(1)

# The backstage/open dialog
open_dlg = app.window(title="Open")
open_dlg.wait('visible', timeout=5)

# Navigate to a file
file_name_edit = open_dlg.child_window(
    title="File name:",
    control_type="Edit"
)
file_name_edit.set_text(r"C:\path\to\presentation.pptx")
open_dlg.child_window(title="Open", control_type="Button").click()

# --- Save As dialog ---
main_win.type_keys("{F12}")  # F12 = Save As
time.sleep(1)
save_dlg = app.window(title_re="Save As.*")
save_dlg.wait('visible', timeout=5)
```

### 2.5 Drag and Drop Simulation

pywinauto provides two methods for drag-and-drop:

| Method | Description | Use Case |
|--------|-------------|----------|
| `drag_mouse()` | SendMessage-based, no visible cursor movement | Background automation |
| `drag_mouse_input()` | Real mouse movement simulation | Foreground, most reliable |

```python
from pywinauto import Application
from pywinauto import mouse
import time

app = Application(backend="uia").connect(path="POWERPNT.EXE")
main_win = app.window(title_re=".*PowerPoint.*")

# --- Method 1: Using drag_mouse_input on a control ---
# Drag a shape from one position to another
slide_pane = main_win.child_window(
    control_type="Pane",
    title_re=".*Slide.*"
)

# Get the pane's rectangle for coordinate reference
rect = slide_pane.rectangle()
start_x = rect.left + 200
start_y = rect.top + 150
end_x = rect.left + 400
end_y = rect.top + 300

# Use pywinauto.mouse module directly for precise control
mouse.press(button='left', coords=(start_x, start_y))
time.sleep(0.1)
mouse.move(coords=(end_x, end_y))
time.sleep(0.1)
mouse.release(button='left', coords=(end_x, end_y))

# --- Method 2: Using drag_mouse_input on a wrapper ---
slide_pane.drag_mouse_input(
    dst=(end_x, end_y),
    src=(start_x, start_y),
    button='left'
)
```

### 2.6 Mouse Click Simulation

```python
from pywinauto import Application
from pywinauto import mouse

app = Application(backend="uia").connect(path="POWERPNT.EXE")
main_win = app.window(title_re=".*PowerPoint.*")

# --- Click on an element (UIA Invoke pattern) ---
# This does NOT move the physical mouse cursor
button = main_win.child_window(title="Bold", control_type="Button")
button.click()  # Uses UIA InvokePattern

# --- Click with real mouse input (moves cursor) ---
button.click_input()  # Physically clicks the control

# --- Click at specific screen coordinates ---
mouse.click(button='left', coords=(500, 300))

# --- Double-click ---
mouse.double_click(button='left', coords=(500, 300))

# --- Right-click ---
mouse.right_click(coords=(500, 300))

# --- Click at coordinates relative to a control ---
slide_pane = main_win.child_window(
    control_type="Pane",
    title_re=".*Slide.*"
)
rect = slide_pane.rectangle()
# Click at center of the slide pane
center_x = (rect.left + rect.right) // 2
center_y = (rect.top + rect.bottom) // 2
mouse.click(button='left', coords=(center_x, center_y))

# --- Scroll ---
mouse.scroll(coords=(center_x, center_y), wheel_dist=3)  # scroll up
mouse.scroll(coords=(center_x, center_y), wheel_dist=-3)  # scroll down
```

### 2.7 Keyboard Input

```python
from pywinauto import Application

app = Application(backend="uia").connect(path="POWERPNT.EXE")
main_win = app.window(title_re=".*PowerPoint.*")

# --- Type text ---
main_win.type_keys("Hello World", with_spaces=True)

# --- Special keys ---
main_win.type_keys("{ENTER}")       # Enter
main_win.type_keys("{TAB}")         # Tab
main_win.type_keys("{DELETE}")      # Delete
main_win.type_keys("{BACKSPACE}")   # Backspace
main_win.type_keys("{ESCAPE}")      # Escape

# --- Modifier keys ---
main_win.type_keys("^a")           # Ctrl+A (select all)
main_win.type_keys("^c")           # Ctrl+C (copy)
main_win.type_keys("^v")           # Ctrl+V (paste)
main_win.type_keys("^z")           # Ctrl+Z (undo)
main_win.type_keys("+{F10}")       # Shift+F10 (context menu)
main_win.type_keys("%{F4}")        # Alt+F4 (close)

# --- Arrow keys ---
main_win.type_keys("{UP}")
main_win.type_keys("{DOWN}")
main_win.type_keys("{LEFT}")
main_win.type_keys("{RIGHT}")

# --- Combined key sequences ---
main_win.type_keys("^a{DELETE}New text{ENTER}")
```

### 2.8 COM Threading Considerations

pywinauto's UIA backend uses COM internally. If you are also using COM for PowerPoint automation (like our MCP server does), be careful with threading:

```python
import sys
# MUST be set before importing pywinauto
sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED (STA)

from pywinauto import Application
```

**Important**: Our MCP server uses a dedicated STA thread for COM calls via `com_wrapper.py`. If mixing pywinauto UIA with direct COM automation, ensure they either share the same STA thread or use proper marshaling.

---

## 3. uiautomation Library

### 3.1 Overview

The `uiautomation` package by yinkaisheng is a Python 3 wrapper around Microsoft's `IUIAutomation` COM interface. It is lighter weight than pywinauto and provides more direct access to UIA patterns and properties.

### 3.2 Installation and Setup

```bash
pip install uiautomation
```

**Note**: Running as Administrator may be required on Windows 7+ for full control enumeration.

### 3.3 Finding PowerPoint Window

```python
import uiautomation as auto

# Set global search timeout
auto.uiautomation.SetGlobalSearchTimeout(15)

# Find PowerPoint by partial window name
ppt = auto.WindowControl(searchDepth=1, SubName='PowerPoint')

if not ppt.Exists(maxSearchSeconds=5):
    print("PowerPoint not found")
    exit(1)

print(f"Window: {ppt.Name}")
print(f"Class: {ppt.ClassName}")
print(f"ProcessId: {ppt.ProcessId}")
print(f"Handle: {ppt.NativeWindowHandle}")
```

### 3.4 Inspecting the Element Tree

```python
import uiautomation as auto

ppt = auto.WindowControl(searchDepth=1, SubName='PowerPoint')

def dump_tree(control, indent=0):
    """Recursively print control tree."""
    prefix = "  " * indent
    print(f"{prefix}{control.ControlTypeName}: "
          f"'{control.Name}' [{control.ClassName}] "
          f"AutomationId='{control.AutomationId}'")
    for child in control.GetChildren():
        dump_tree(child, indent + 1)

# Dump first 2 levels
def dump_tree_limited(control, indent=0, max_depth=2):
    if indent >= max_depth:
        return
    prefix = "  " * indent
    print(f"{prefix}{control.ControlTypeName}: "
          f"'{control.Name}' [{control.ClassName}]")
    for child in control.GetChildren():
        dump_tree_limited(child, indent + 1, max_depth)

dump_tree_limited(ppt, max_depth=3)
```

### 3.5 Interacting with Ribbon

```python
import uiautomation as auto
import time

ppt = auto.WindowControl(searchDepth=1, SubName='PowerPoint')

# --- Click a ribbon tab ---
home_tab = ppt.TabItemControl(Name='Home')
if home_tab.Exists(3):
    home_tab.Click()
    time.sleep(0.3)

# --- Click a button in the ribbon ---
bold_btn = ppt.ButtonControl(Name='Bold')
if bold_btn.Exists(3):
    bold_btn.Click()

# --- Access a split button (like Paste) ---
paste_btn = ppt.SplitButtonControl(Name='Paste')
if paste_btn.Exists(3):
    # Click the main button part
    paste_btn.Click()
    # Or expand the dropdown
    paste_btn.GetInvokePattern().Invoke()

# --- Read button state ---
bold_btn = ppt.ButtonControl(Name='Bold')
if bold_btn.Exists(2):
    # Check toggle state if available
    try:
        toggle_pattern = bold_btn.GetTogglePattern()
        state = toggle_pattern.ToggleState
        # 0 = Off, 1 = On, 2 = Indeterminate
        print(f"Bold toggle state: {state}")
    except Exception:
        pass
```

### 3.6 Working with Dialogs

```python
import uiautomation as auto
import time

ppt = auto.WindowControl(searchDepth=1, SubName='PowerPoint')

# Open Insert > Table dialog via ribbon
insert_tab = ppt.TabItemControl(Name='Insert')
insert_tab.Click()
time.sleep(0.5)

table_btn = ppt.ButtonControl(Name='Table')
table_btn.Click()
time.sleep(0.5)

# Find the table size picker or dialog
# This varies by PowerPoint version

# --- Generic dialog interaction ---
dialog = auto.WindowControl(
    searchDepth=1,
    SubName='Insert Table'
)
if dialog.Exists(3):
    # Find edit controls
    rows_edit = dialog.EditControl(Name='Number of rows')
    cols_edit = dialog.EditControl(Name='Number of columns')

    if rows_edit.Exists(2):
        rows_edit.GetValuePattern().SetValue('4')
    if cols_edit.Exists(2):
        cols_edit.GetValuePattern().SetValue('3')

    # Click OK
    ok_btn = dialog.ButtonControl(Name='OK')
    if ok_btn.Exists(2):
        ok_btn.Click()
```

### 3.7 Sending Keys

```python
import uiautomation as auto

ppt = auto.WindowControl(searchDepth=1, SubName='PowerPoint')

# Set focus to PowerPoint
ppt.SetFocus()

# Send keystrokes
auto.SendKeys('Hello World')
auto.SendKeys('{Enter}')
auto.SendKeys('{Ctrl}a')     # Select all
auto.SendKeys('{Ctrl}c')     # Copy
auto.SendKeys('{Ctrl}v')     # Paste
auto.SendKeys('{Delete}')
auto.SendKeys('{Escape}')

# Send keys with modifiers held
auto.SendKeys('{Ctrl}{Shift}s')  # Ctrl+Shift+S
```

### 3.8 Control.Element (Low-Level Access)

Every control wraps an `IUIAutomationElement` COM object:

```python
import uiautomation as auto

ppt = auto.WindowControl(searchDepth=1, SubName='PowerPoint')

# Access low-level COM element
element = ppt.Element  # IUIAutomationElement

# Get bounding rectangle
rect = element.CurrentBoundingRectangle
print(f"Bounds: ({rect.left}, {rect.top}, {rect.right}, {rect.bottom})")

# Get properties directly
name = element.CurrentName
class_name = element.CurrentClassName
control_type = element.CurrentControlType
automation_id = element.CurrentAutomationId
```

---

## 4. Accessibility API Approach (MSAA vs UIA)

### 4.1 MSAA vs UIA Comparison

| Feature | MSAA (Active Accessibility) | UIA (UI Automation) |
|---------|---------------------------|---------------------|
| Age | Windows 95 era (legacy) | Windows Vista+ (modern) |
| Interface | `IAccessible` | `IUIAutomationElement` |
| Object model | Fixed roles (ROLE_SYSTEM_*) | Extensible control patterns |
| Navigation | `IAccessible::accNavigate` | Tree walkers with views |
| Text model | None | Full `TextPattern` support |
| Properties | Small fixed set | Rich, extensible properties |
| Performance | In-process faster, out-of-process slow | Good out-of-process performance |
| Office ribbon | Basic exposure | Full exposure |
| **Recommendation** | **Legacy apps only** | **Use for all new work** |

Microsoft recommends UIA for all new development. MSAA is maintained only for backward compatibility.

### 4.2 PowerPoint Accessibility Tree

PowerPoint exposes its UI through both MSAA and UIA, but the UIA tree is richer:

**What UIA exposes in PowerPoint:**
- Main window and title bar
- Ribbon tabs, groups, and buttons (with states)
- Slide pane (the main editing area)
- Slide thumbnail panel
- Notes pane
- Status bar
- Task panes (Format Shape, Animation, etc.)
- Dialog boxes

**What UIA does NOT reliably expose:**
- Individual shapes on the slide canvas (these are rendered, not native controls)
- Shape resize handles
- Shape connection points
- Ruler and guide positions
- Slide show presentation view content

### 4.3 Accessing Slide Canvas Elements

The slide canvas in PowerPoint is a custom-rendered surface. Shapes on the canvas are NOT individual UIA elements in most cases. The slide area appears as a single Pane control to UIA.

**Workaround approaches:**

```python
import uiautomation as auto

ppt = auto.WindowControl(searchDepth=1, SubName='PowerPoint')

# The slide pane is typically one of these
slide_pane = ppt.PaneControl(ClassName='paneClassDC')
# or
slide_pane = ppt.PaneControl(Name='Slide')

if slide_pane.Exists(3):
    rect = slide_pane.BoundingRectangle
    print(f"Slide pane: {rect}")

    # Shapes on the canvas are NOT UIA children
    children = slide_pane.GetChildren()
    print(f"Number of UIA children: {len(children)}")
    # Typically returns 0 or very few elements

    # To interact with shapes, use COM object model instead:
    # shape = slide.Shapes(1)  # via COM
    # Then get shape position and convert to screen coordinates
```

### 4.4 Shape Selection via COM + Screen Coordinates

Since UIA cannot directly access individual shapes, combine COM object model (to get shape positions) with mouse input (to interact with them):

```python
import win32com.client
from pywinauto import mouse
import uiautomation as auto

# Get shape position via COM
ppt_app = win32com.client.Dispatch("PowerPoint.Application")
slide = ppt_app.ActiveWindow.View.Slide
shape = slide.Shapes(1)

# Shape position in points (72 points = 1 inch)
shape_left = shape.Left    # points from left
shape_top = shape.Top      # points from top
shape_width = shape.Width
shape_height = shape.Height

# Get slide pane screen position via UIA
ppt_win = auto.WindowControl(searchDepth=1, SubName='PowerPoint')
slide_pane = ppt_win.PaneControl(ClassName='paneClassDC')
pane_rect = slide_pane.BoundingRectangle

# Get zoom level from COM
zoom = ppt_app.ActiveWindow.View.Zoom / 100.0  # e.g., 0.68 for 68%

# Convert shape center (points) to screen pixels
# This is approximate - depends on slide offset within pane
slide_width_pts = ppt_app.ActivePresentation.PageSetup.SlideWidth
slide_height_pts = ppt_app.ActivePresentation.PageSetup.SlideHeight

pane_width = pane_rect.right - pane_rect.left
pane_height = pane_rect.bottom - pane_rect.top

# Points to pixels (at current zoom)
# 1 point = 1/72 inch, screen DPI varies (typically 96)
dpi = 96
px_per_point = (dpi / 72.0) * zoom

# Shape center in screen coordinates (approximate)
shape_center_x = int(pane_rect.left + (shape_left + shape_width / 2) * px_per_point)
shape_center_y = int(pane_rect.top + (shape_top + shape_height / 2) * px_per_point)

# Note: There's also a scroll/pan offset to account for.
# This calculation is approximate. The exact mapping depends on
# the current scroll position and the slide's position within the pane.

# Click the shape
mouse.click(button='left', coords=(shape_center_x, shape_center_y))
```

### 4.5 Reading Selection State via COM

For selection state, the COM object model is more reliable than accessibility APIs:

```python
import win32com.client

ppt = win32com.client.Dispatch("PowerPoint.Application")
window = ppt.ActiveWindow

# Get current selection
selection = window.Selection

# Selection type constants
# ppSelectionNone = 0
# ppSelectionSlides = 1
# ppSelectionShapes = 2
# ppSelectionText = 3

sel_type = selection.Type
print(f"Selection type: {sel_type}")

if sel_type == 2:  # ppSelectionShapes
    shapes = selection.ShapeRange
    for i in range(1, shapes.Count + 1):
        shape = shapes.Item(i)
        print(f"  Selected shape: {shape.Name} "
              f"({shape.Left}, {shape.Top}) "
              f"{shape.Width}x{shape.Height}")

elif sel_type == 3:  # ppSelectionText
    text_range = selection.TextRange
    print(f"  Selected text: '{text_range.Text}'")
    print(f"  In shape: {selection.ShapeRange.Item(1).Name}")
```

### 4.6 MSAA via Python (IAccessible)

For completeness, here is how to access MSAA/IAccessible from Python, though UIA is preferred:

```python
import ctypes
from ctypes import wintypes
import win32gui
import win32con

# Load oleacc.dll for accessibility functions
oleacc = ctypes.windll.oleacc

# AccessibleObjectFromWindow
def get_accessible_from_window(hwnd):
    """Get IAccessible interface from a window handle."""
    import comtypes
    from comtypes.automation import IDispatch

    # OBJID_CLIENT = -4 (client area accessible object)
    OBJID_CLIENT = 0xFFFFFFFC
    IID_IAccessible = comtypes.GUID(
        '{618736E0-3C3D-11CF-810C-00AA00389B71}'
    )

    p_acc = ctypes.POINTER(comtypes.IUnknown)()
    hr = oleacc.AccessibleObjectFromWindow(
        hwnd,
        OBJID_CLIENT,
        ctypes.byref(IID_IAccessible),
        ctypes.byref(p_acc)
    )
    if hr == 0:
        return p_acc
    return None

# AccessibleObjectFromWindow with OBJID_NATIVEOM
# to get PowerPoint's native object model from a window
def get_native_om_from_window(hwnd):
    """Get native object model (IDispatch) from PowerPoint window."""
    import comtypes
    from comtypes.automation import IDispatch

    OBJID_NATIVEOM = 0xFFFFFFF0
    IID_IDispatch = comtypes.GUID(
        '{00020400-0000-0000-C000-000000000046}'
    )

    p_dispatch = ctypes.POINTER(comtypes.IUnknown)()
    hr = oleacc.AccessibleObjectFromWindow(
        hwnd,
        OBJID_NATIVEOM,
        ctypes.byref(IID_IDispatch),
        ctypes.byref(p_dispatch)
    )
    if hr == 0:
        return p_dispatch
    return None
```

---

## 5. ctypes / win32gui Direct Approach

### 5.1 Finding PowerPoint Windows

PowerPoint uses specific window class names that vary by version:

| Office Version | Main Window Class |
|---------------|-------------------|
| PowerPoint 2000 | `PP9FrameClass` |
| PowerPoint 2002 (XP) | `PP10FrameClass` |
| PowerPoint 2003 | `PP11FrameClass` |
| PowerPoint 2007+ | `PPTFrameClass` |
| PowerPoint 365 / 2019+ | `PPTFrameClass` |

```python
import win32gui
import win32con
import win32process

# --- Find main PowerPoint window ---
hwnd = win32gui.FindWindow("PPTFrameClass", None)
if hwnd:
    title = win32gui.GetWindowText(hwnd)
    print(f"Found PowerPoint: hwnd={hwnd:#x}, title='{title}'")
else:
    print("PowerPoint not found")

# --- Find by enumerating all windows ---
def find_powerpoint_windows():
    """Find all PowerPoint windows."""
    results = []

    def enum_callback(hwnd, extra):
        class_name = win32gui.GetClassName(hwnd)
        if class_name == "PPTFrameClass":
            title = win32gui.GetWindowText(hwnd)
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            results.append({
                'hwnd': hwnd,
                'title': title,
                'class': class_name,
                'pid': pid
            })

    win32gui.EnumWindows(enum_callback, None)
    return results

windows = find_powerpoint_windows()
for w in windows:
    print(f"  hwnd={w['hwnd']:#x} pid={w['pid']} "
          f"title='{w['title']}'")
```

### 5.2 Enumerating Child Windows

PowerPoint's child window hierarchy typically looks like:

```
PPTFrameClass (main window)
  ├── MsoCommandBarDock (top)
  │     └── MsoCommandBar (ribbon)
  ├── mdiViewWnd (MDI view container)
  │     └── paneClassDC (slide rendering pane)
  ├── MsoCommandBarDock (bottom - status bar)
  └── Various other child windows
```

```python
import win32gui

def enum_child_windows(parent_hwnd, max_depth=3, depth=0):
    """Recursively enumerate child windows."""
    results = []

    def callback(hwnd, _):
        class_name = win32gui.GetClassName(hwnd)
        title = win32gui.GetWindowText(hwnd)
        rect = win32gui.GetWindowRect(hwnd)
        indent = "  " * depth
        print(f"{indent}hwnd={hwnd:#x} class='{class_name}' "
              f"title='{title}' rect={rect}")
        results.append(hwnd)

    win32gui.EnumChildWindows(parent_hwnd, callback, None)
    return results

# Find PowerPoint and dump child windows
ppt_hwnd = win32gui.FindWindow("PPTFrameClass", None)
if ppt_hwnd:
    print(f"PowerPoint: {ppt_hwnd:#x}")
    children = enum_child_windows(ppt_hwnd)
    print(f"Total child windows: {len(children)}")
```

### 5.3 Finding Specific Child Windows

```python
import win32gui

def find_child_by_class(parent_hwnd, class_name):
    """Find a child window by class name."""
    return win32gui.FindWindowEx(parent_hwnd, 0, class_name, None)

ppt_hwnd = win32gui.FindWindow("PPTFrameClass", None)

# Find the MDI view window (contains the slide)
mdi_hwnd = find_child_by_class(ppt_hwnd, "mdiViewWnd")
print(f"MDI View: {mdi_hwnd:#x}")

# Find the slide pane within the MDI view
if mdi_hwnd:
    slide_hwnd = find_child_by_class(mdi_hwnd, "paneClassDC")
    print(f"Slide Pane: {slide_hwnd:#x}")

    # Get the slide pane's screen rectangle
    rect = win32gui.GetWindowRect(slide_hwnd)
    print(f"Slide rect: left={rect[0]}, top={rect[1]}, "
          f"right={rect[2]}, bottom={rect[3]}")

    # Get client area rectangle
    client_rect = win32gui.GetClientRect(slide_hwnd)
    print(f"Client rect: {client_rect}")
```

### 5.4 Window Coordinates

```python
import win32gui

def get_window_screen_rect(hwnd):
    """Get window's client area in screen coordinates."""
    left, top, right, bottom = win32gui.GetClientRect(hwnd)
    # Convert client (0,0) to screen coordinates
    screen_left, screen_top = win32gui.ClientToScreen(hwnd, (left, top))
    screen_right, screen_bottom = win32gui.ClientToScreen(
        hwnd, (right, bottom)
    )
    return (screen_left, screen_top, screen_right, screen_bottom)

def screen_to_client(hwnd, screen_x, screen_y):
    """Convert screen coordinates to window client coordinates."""
    return win32gui.ScreenToClient(hwnd, (screen_x, screen_y))

def client_to_screen(hwnd, client_x, client_y):
    """Convert client coordinates to screen coordinates."""
    return win32gui.ClientToScreen(hwnd, (client_x, client_y))

# Example usage
ppt_hwnd = win32gui.FindWindow("PPTFrameClass", None)
if ppt_hwnd:
    mdi_hwnd = win32gui.FindWindowEx(ppt_hwnd, 0, "mdiViewWnd", None)
    if mdi_hwnd:
        slide_hwnd = win32gui.FindWindowEx(
            mdi_hwnd, 0, "paneClassDC", None
        )
        if slide_hwnd:
            rect = get_window_screen_rect(slide_hwnd)
            print(f"Slide pane screen rect: {rect}")
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]
            print(f"Slide pane size: {width}x{height} pixels")
```

### 5.5 SendMessage / PostMessage

```python
import win32gui
import win32con
import win32api

def make_lparam(x, y):
    """Pack x, y into lParam for mouse messages."""
    return (y << 16) | (x & 0xFFFF)

def send_click_to_window(hwnd, x, y):
    """Send a mouse click to a specific window at client coordinates.

    NOTE: This sends messages directly to the window's message queue.
    The physical mouse cursor does NOT move. This can work for some
    controls but is unreliable for custom-rendered surfaces like
    PowerPoint's slide pane.
    """
    lparam = make_lparam(x, y)
    win32gui.PostMessage(hwnd, win32con.WM_LBUTTONDOWN,
                         win32con.MK_LBUTTON, lparam)
    win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)

def send_key_to_window(hwnd, vk_code):
    """Send a key press to a specific window.

    NOTE: PostMessage key events may not work for all applications.
    Some apps require the window to be in the foreground.
    """
    win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, 0)
    win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, 0)

def send_char_to_window(hwnd, char):
    """Send a character to a specific window."""
    win32gui.PostMessage(hwnd, win32con.WM_CHAR, ord(char), 0)

# Example: Send Escape key to PowerPoint
ppt_hwnd = win32gui.FindWindow("PPTFrameClass", None)
if ppt_hwnd:
    send_key_to_window(ppt_hwnd, win32con.VK_ESCAPE)
```

**Warning**: `SendMessage`/`PostMessage` for mouse events is unreliable with PowerPoint's custom-rendered surfaces. For the slide canvas, use `SendInput` or physical mouse movement instead.

### 5.6 SetForegroundWindow / SetFocus

```python
import win32gui
import win32con
import win32api
import win32process
import ctypes
import time

def bring_window_to_front(hwnd):
    """Bring a window to the foreground reliably.

    Windows has restrictions on which process can set the foreground
    window. This function uses multiple strategies to work around them.
    """
    # Strategy 1: Simple SetForegroundWindow
    try:
        # If window is minimized, restore it first
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

        win32gui.SetForegroundWindow(hwnd)
        return True
    except Exception:
        pass

    # Strategy 2: Attach to foreground thread
    try:
        foreground_hwnd = win32gui.GetForegroundWindow()
        foreground_tid, _ = win32process.GetWindowThreadProcessId(
            foreground_hwnd
        )
        target_tid, _ = win32process.GetWindowThreadProcessId(hwnd)

        # Attach our thread to the foreground thread
        ctypes.windll.user32.AttachThreadInput(
            foreground_tid, target_tid, True
        )

        win32gui.SetForegroundWindow(hwnd)
        win32gui.BringWindowToTop(hwnd)

        # Detach
        ctypes.windll.user32.AttachThreadInput(
            foreground_tid, target_tid, False
        )
        return True
    except Exception:
        pass

    # Strategy 3: Alt key trick
    try:
        # Simulate Alt key press to bypass foreground restriction
        win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
        win32gui.SetForegroundWindow(hwnd)
        win32api.keybd_event(
            win32con.VK_MENU, 0,
            win32con.KEYEVENTF_KEYUP, 0
        )
        return True
    except Exception:
        return False

# Example
ppt_hwnd = win32gui.FindWindow("PPTFrameClass", None)
if ppt_hwnd:
    success = bring_window_to_front(ppt_hwnd)
    print(f"Brought to front: {success}")
```

---

## 6. Low-Level Input Simulation

### 6.1 SendInput with ctypes

`SendInput` is the modern Windows API for simulating keyboard and mouse input at the system level. Unlike `SendMessage`/`PostMessage`, it goes through the full input pipeline and is the most reliable method for input simulation.

```python
import ctypes
from ctypes import wintypes
import time

user32 = ctypes.WinDLL('user32', use_last_error=True)

# Constants
INPUT_MOUSE = 0
INPUT_KEYBOARD = 1

MOUSEEVENTF_MOVE = 0x0001
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
MOUSEEVENTF_RIGHTDOWN = 0x0008
MOUSEEVENTF_RIGHTUP = 0x0010
MOUSEEVENTF_MIDDLEDOWN = 0x0020
MOUSEEVENTF_MIDDLEUP = 0x0040
MOUSEEVENTF_WHEEL = 0x0800
MOUSEEVENTF_ABSOLUTE = 0x8000

KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_UNICODE = 0x0004
KEYEVENTF_SCANCODE = 0x0008
MAPVK_VK_TO_VSC = 0

# Structure definitions
wintypes.ULONG_PTR = wintypes.WPARAM

class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", wintypes.LONG),
        ("dy", wintypes.LONG),
        ("mouseData", wintypes.DWORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", wintypes.ULONG_PTR),
    ]

class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", wintypes.ULONG_PTR),
    ]

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        if not self.dwFlags & KEYEVENTF_UNICODE:
            self.wScan = user32.MapVirtualKeyExW(
                self.wVk, MAPVK_VK_TO_VSC, 0
            )

class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [
        ("uMsg", wintypes.DWORD),
        ("wParamL", wintypes.WORD),
        ("wParamH", wintypes.WORD),
    ]

class INPUT(ctypes.Structure):
    class _INPUT(ctypes.Union):
        _fields_ = [
            ("ki", KEYBDINPUT),
            ("mi", MOUSEINPUT),
            ("hi", HARDWAREINPUT),
        ]

    _anonymous_ = ("_input",)
    _fields_ = [
        ("type", wintypes.DWORD),
        ("_input", _INPUT),
    ]

LPINPUT = ctypes.POINTER(INPUT)

def _check_count(result, func, args):
    if result == 0:
        raise ctypes.WinError(ctypes.get_last_error())
    return args

user32.SendInput.errcheck = _check_count
user32.SendInput.argtypes = (
    wintypes.UINT, LPINPUT, ctypes.c_int
)
```

### 6.2 Mouse Functions Using SendInput

```python
def _screen_to_absolute(x, y):
    """Convert screen coordinates to absolute (0-65535) coordinates."""
    screen_w = user32.GetSystemMetrics(0)  # SM_CXSCREEN
    screen_h = user32.GetSystemMetrics(1)  # SM_CYSCREEN
    abs_x = int(x * 65535 / (screen_w - 1))
    abs_y = int(y * 65535 / (screen_h - 1))
    return abs_x, abs_y

def move_mouse(x, y):
    """Move mouse cursor to screen coordinates (x, y)."""
    abs_x, abs_y = _screen_to_absolute(x, y)
    inp = INPUT(
        type=INPUT_MOUSE,
        mi=MOUSEINPUT(
            dx=abs_x,
            dy=abs_y,
            dwFlags=MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE
        )
    )
    user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))

def click_at(x, y, button='left'):
    """Click at screen coordinates."""
    move_mouse(x, y)
    time.sleep(0.01)

    if button == 'left':
        down_flag = MOUSEEVENTF_LEFTDOWN
        up_flag = MOUSEEVENTF_LEFTUP
    elif button == 'right':
        down_flag = MOUSEEVENTF_RIGHTDOWN
        up_flag = MOUSEEVENTF_RIGHTUP
    elif button == 'middle':
        down_flag = MOUSEEVENTF_MIDDLEDOWN
        up_flag = MOUSEEVENTF_MIDDLEUP
    else:
        raise ValueError(f"Unknown button: {button}")

    inputs = (INPUT * 2)(
        INPUT(type=INPUT_MOUSE, mi=MOUSEINPUT(dwFlags=down_flag)),
        INPUT(type=INPUT_MOUSE, mi=MOUSEINPUT(dwFlags=up_flag)),
    )
    user32.SendInput(2, inputs, ctypes.sizeof(INPUT))

def double_click_at(x, y):
    """Double-click at screen coordinates."""
    click_at(x, y)
    time.sleep(0.05)
    click_at(x, y)

def drag(start_x, start_y, end_x, end_y, steps=20, duration=0.5):
    """Drag from start to end coordinates with smooth movement."""
    move_mouse(start_x, start_y)
    time.sleep(0.05)

    # Press
    inp_down = INPUT(
        type=INPUT_MOUSE,
        mi=MOUSEINPUT(dwFlags=MOUSEEVENTF_LEFTDOWN)
    )
    user32.SendInput(1, ctypes.byref(inp_down), ctypes.sizeof(INPUT))
    time.sleep(0.05)

    # Move in steps
    step_delay = duration / steps
    for i in range(1, steps + 1):
        t = i / steps
        cx = int(start_x + (end_x - start_x) * t)
        cy = int(start_y + (end_y - start_y) * t)
        move_mouse(cx, cy)
        time.sleep(step_delay)

    # Release
    inp_up = INPUT(
        type=INPUT_MOUSE,
        mi=MOUSEINPUT(dwFlags=MOUSEEVENTF_LEFTUP)
    )
    user32.SendInput(1, ctypes.byref(inp_up), ctypes.sizeof(INPUT))

def scroll_at(x, y, clicks):
    """Scroll mouse wheel at coordinates. Positive = up, negative = down."""
    move_mouse(x, y)
    time.sleep(0.01)

    WHEEL_DELTA = 120
    inp = INPUT(
        type=INPUT_MOUSE,
        mi=MOUSEINPUT(
            mouseData=ctypes.c_ulong(clicks * WHEEL_DELTA).value,
            dwFlags=MOUSEEVENTF_WHEEL
        )
    )
    user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))
```

### 6.3 Keyboard Functions Using SendInput

```python
def press_key(vk_code):
    """Press a key down."""
    inp = INPUT(
        type=INPUT_KEYBOARD,
        ki=KEYBDINPUT(wVk=vk_code)
    )
    user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))

def release_key(vk_code):
    """Release a key."""
    inp = INPUT(
        type=INPUT_KEYBOARD,
        ki=KEYBDINPUT(wVk=vk_code, dwFlags=KEYEVENTF_KEYUP)
    )
    user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))

def press_and_release(vk_code):
    """Press and release a key."""
    press_key(vk_code)
    time.sleep(0.01)
    release_key(vk_code)

def type_key_combo(*vk_codes):
    """Press a key combination (e.g., Ctrl+A).

    All keys are pressed in order, then released in reverse order.
    """
    for vk in vk_codes:
        press_key(vk)
        time.sleep(0.01)
    for vk in reversed(vk_codes):
        release_key(vk)
        time.sleep(0.01)

def type_text(text):
    """Type a string using Unicode input events."""
    for char in text:
        # Unicode key press
        inp_down = INPUT(
            type=INPUT_KEYBOARD,
            ki=KEYBDINPUT(
                wVk=0,
                wScan=ord(char),
                dwFlags=KEYEVENTF_UNICODE
            )
        )
        user32.SendInput(
            1, ctypes.byref(inp_down), ctypes.sizeof(INPUT)
        )
        time.sleep(0.01)

        # Unicode key release
        inp_up = INPUT(
            type=INPUT_KEYBOARD,
            ki=KEYBDINPUT(
                wVk=0,
                wScan=ord(char),
                dwFlags=KEYEVENTF_UNICODE | KEYEVENTF_KEYUP
            )
        )
        user32.SendInput(
            1, ctypes.byref(inp_up), ctypes.sizeof(INPUT)
        )
        time.sleep(0.01)

# Virtual key code constants (subset)
VK_RETURN = 0x0D
VK_ESCAPE = 0x1B
VK_TAB = 0x09
VK_BACK = 0x08
VK_DELETE = 0x2E
VK_SHIFT = 0x10
VK_CONTROL = 0x11
VK_MENU = 0x12  # Alt
VK_LEFT = 0x25
VK_UP = 0x26
VK_RIGHT = 0x27
VK_DOWN = 0x28
VK_HOME = 0x24
VK_END = 0x23
VK_PRIOR = 0x21  # Page Up
VK_NEXT = 0x22   # Page Down
VK_F1 = 0x70
VK_F5 = 0x74
VK_F12 = 0x7B
VK_A = 0x41
VK_C = 0x43
VK_V = 0x56
VK_Z = 0x5A

# Example: Ctrl+A (select all)
type_key_combo(VK_CONTROL, VK_A)

# Example: Type text
type_text("Hello, PowerPoint!")
```

### 6.4 Legacy mouse_event / keybd_event

The older `mouse_event` and `keybd_event` functions are simpler but deprecated. They still work and are sometimes more reliable than `SendInput` for certain scenarios:

```python
import win32api
import win32con

# --- Mouse click at screen coordinates ---
def legacy_click(x, y):
    """Click using the legacy mouse_event API."""
    # Move cursor
    win32api.SetCursorPos((x, y))
    time.sleep(0.01)

    # Left click
    win32api.mouse_event(
        win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    )
    time.sleep(0.01)
    win32api.mouse_event(
        win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    )

# --- Right click ---
def legacy_right_click(x, y):
    win32api.SetCursorPos((x, y))
    time.sleep(0.01)
    win32api.mouse_event(
        win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
    )
    time.sleep(0.01)
    win32api.mouse_event(
        win32con.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    )

# --- Key press ---
def legacy_key_press(vk_code):
    """Press and release using legacy keybd_event API."""
    win32api.keybd_event(vk_code, 0, 0, 0)
    time.sleep(0.01)
    win32api.keybd_event(
        vk_code, 0, win32con.KEYEVENTF_KEYUP, 0
    )
```

---

## 7. Practical Recipes

### 7.1 Recipe: Click a Shape on the Slide Canvas

Combining COM (for shape coordinates) with SendInput (for clicking):

```python
import win32com.client
import win32gui
import ctypes
import time

def click_shape_on_canvas(shape_index):
    """Click on a shape in the active slide using its index."""
    # Get shape position via COM
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    window = ppt.ActiveWindow
    slide = window.View.Slide
    shape = slide.Shapes(shape_index)

    shape_center_x_pts = shape.Left + shape.Width / 2
    shape_center_y_pts = shape.Top + shape.Height / 2

    # Get slide pane window handle
    ppt_hwnd = win32gui.FindWindow("PPTFrameClass", None)
    mdi_hwnd = win32gui.FindWindowEx(ppt_hwnd, 0, "mdiViewWnd", None)
    slide_hwnd = win32gui.FindWindowEx(mdi_hwnd, 0, "paneClassDC", None)

    # Get slide pane rectangle in screen coords
    client_rect = win32gui.GetClientRect(slide_hwnd)
    screen_left, screen_top = win32gui.ClientToScreen(
        slide_hwnd, (0, 0)
    )
    pane_width = client_rect[2]
    pane_height = client_rect[3]

    # Get zoom and slide dimensions
    zoom = window.View.Zoom / 100.0
    slide_w = ppt.ActivePresentation.PageSetup.SlideWidth
    slide_h = ppt.ActivePresentation.PageSetup.SlideHeight

    # Calculate rendered slide size in pixels (96 DPI)
    dpi = 96
    px_per_pt = (dpi / 72.0) * zoom
    rendered_w = slide_w * px_per_pt
    rendered_h = slide_h * px_per_pt

    # Slide is centered in the pane
    offset_x = (pane_width - rendered_w) / 2
    offset_y = (pane_height - rendered_h) / 2

    # Convert shape center to screen coordinates
    screen_x = int(screen_left + offset_x
                   + shape_center_x_pts * px_per_pt)
    screen_y = int(screen_top + offset_y
                   + shape_center_y_pts * px_per_pt)

    # Bring PowerPoint to front
    win32gui.SetForegroundWindow(ppt_hwnd)
    time.sleep(0.2)

    # Click using SendInput (see section 6.2)
    click_at(screen_x, screen_y)

# Usage
click_shape_on_canvas(1)  # Click first shape
```

### 7.2 Recipe: Navigate Ribbon via Keyboard

Instead of finding ribbon elements by UIA, use keyboard shortcuts:

```python
import win32com.client
import time

def activate_ribbon_tab_by_key(ppt_hwnd, alt_key):
    """Activate a ribbon tab using Alt+key shortcut.

    Common Alt keys for PowerPoint ribbon:
      Alt+H = Home
      Alt+N = Insert
      Alt+G = Design
      Alt+K = Transitions
      Alt+A = Animations
      Alt+S = Slide Show
      Alt+R = Review
      Alt+W = View
    """
    bring_window_to_front(ppt_hwnd)
    time.sleep(0.2)

    # Press Alt to activate ribbon key tips
    press_and_release(VK_MENU)  # Alt key
    time.sleep(0.3)

    # Press the tab key
    press_and_release(ord(alt_key.upper()))
    time.sleep(0.3)

# Example: Go to Insert tab
ppt_hwnd = win32gui.FindWindow("PPTFrameClass", None)
activate_ribbon_tab_by_key(ppt_hwnd, 'N')
```

### 7.3 Recipe: Interact with a Dialog

```python
from pywinauto import Application
import time

def insert_table_via_dialog():
    """Insert a table using the Insert Table dialog."""
    app = Application(backend="uia").connect(path="POWERPNT.EXE")
    main_win = app.window(title_re=".*PowerPoint.*")

    # Navigate to Insert tab
    main_win.type_keys("%n")  # Alt+N for Insert tab
    time.sleep(0.5)

    # Click Table button
    table_btn = main_win.child_window(
        title="Table",
        control_type="MenuItem"
    )
    table_btn.click_input()
    time.sleep(0.5)

    # Click "Insert Table..." option
    insert_table = main_win.child_window(
        title="Insert Table...",
        control_type="MenuItem"
    )
    insert_table.click_input()
    time.sleep(0.5)

    # Work with the Insert Table dialog
    dialog = app.window(title="Insert Table")
    dialog.wait('visible', timeout=5)

    # Set values
    cols_edit = dialog.child_window(
        title="Number of columns:",
        control_type="Edit"
    )
    cols_edit.set_text("4")

    rows_edit = dialog.child_window(
        title="Number of rows:",
        control_type="Edit"
    )
    rows_edit.set_text("3")

    # Click OK
    dialog.child_window(title="OK", control_type="Button").click()

insert_table_via_dialog()
```

### 7.4 Recipe: Drag a Shape to Move It

```python
import win32com.client
import win32gui
import time

def drag_shape(shape_index, delta_x_pts, delta_y_pts):
    """Drag a shape by a given offset (in points).

    Args:
        shape_index: 1-based shape index on the active slide
        delta_x_pts: Horizontal offset in points
        delta_y_pts: Vertical offset in points
    """
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    window = ppt.ActiveWindow
    slide = window.View.Slide
    shape = slide.Shapes(shape_index)

    # Calculate current and target positions
    center_x = shape.Left + shape.Width / 2
    center_y = shape.Top + shape.Height / 2
    target_x = center_x + delta_x_pts
    target_y = center_y + delta_y_pts

    # Convert to screen coordinates (reuse logic from recipe 7.1)
    ppt_hwnd = win32gui.FindWindow("PPTFrameClass", None)
    mdi_hwnd = win32gui.FindWindowEx(ppt_hwnd, 0, "mdiViewWnd", None)
    slide_hwnd = win32gui.FindWindowEx(
        mdi_hwnd, 0, "paneClassDC", None
    )

    client_rect = win32gui.GetClientRect(slide_hwnd)
    screen_left, screen_top = win32gui.ClientToScreen(
        slide_hwnd, (0, 0)
    )
    pane_width = client_rect[2]
    pane_height = client_rect[3]

    zoom = window.View.Zoom / 100.0
    slide_w = ppt.ActivePresentation.PageSetup.SlideWidth
    slide_h = ppt.ActivePresentation.PageSetup.SlideHeight

    dpi = 96
    px_per_pt = (dpi / 72.0) * zoom
    rendered_w = slide_w * px_per_pt
    rendered_h = slide_h * px_per_pt
    offset_x = (pane_width - rendered_w) / 2
    offset_y = (pane_height - rendered_h) / 2

    start_sx = int(screen_left + offset_x + center_x * px_per_pt)
    start_sy = int(screen_top + offset_y + center_y * px_per_pt)
    end_sx = int(screen_left + offset_x + target_x * px_per_pt)
    end_sy = int(screen_top + offset_y + target_y * px_per_pt)

    # Bring PowerPoint to front
    bring_window_to_front(ppt_hwnd)
    time.sleep(0.2)

    # Perform the drag (see section 6.2)
    drag(start_sx, start_sy, end_sx, end_sy, steps=20, duration=0.3)

# Usage: Move shape 1 right by 72 points (1 inch) and down by 36 points
drag_shape(1, 72, 36)
```

### 7.5 Recipe: Read Slide Thumbnails Panel

```python
from pywinauto import Application

def get_slide_thumbnails():
    """Get information about slides from the thumbnail panel."""
    app = Application(backend="uia").connect(path="POWERPNT.EXE")
    main_win = app.window(title_re=".*PowerPoint.*")

    # The slides panel contains ListItem controls for each slide
    # Find the slides panel (exact name may vary by language/version)
    slides_panel = main_win.child_window(
        title="Slides",
        control_type="Pane"
    )

    if slides_panel.exists():
        # Get all slide thumbnail items
        items = slides_panel.children(control_type="ListItem")
        print(f"Found {len(items)} slide thumbnails")
        for item in items:
            print(f"  Slide: {item.window_text()}")
            rect = item.rectangle()
            print(f"    Position: ({rect.left}, {rect.top})")
    else:
        print("Slides panel not found")

get_slide_thumbnails()
```

### 7.6 Recipe: Monitor PowerPoint State Changes

```python
import uiautomation as auto
import time

def monitor_powerpoint_title(interval=1.0, duration=30):
    """Monitor PowerPoint window title for changes.

    Useful for detecting when a file is saved, renamed, etc.
    """
    ppt = auto.WindowControl(searchDepth=1, SubName='PowerPoint')
    if not ppt.Exists(5):
        print("PowerPoint not found")
        return

    last_title = ppt.Name
    print(f"Initial title: {last_title}")

    start = time.time()
    while time.time() - start < duration:
        ppt.Refind()
        current_title = ppt.Name
        if current_title != last_title:
            print(f"Title changed: '{last_title}' -> '{current_title}'")
            last_title = current_title
        time.sleep(interval)

monitor_powerpoint_title()
```

---

## 8. Comparison and Recommendations

### 8.1 Approach Comparison

| Approach | Ribbon | Dialogs | Slide Canvas | Background | Reliability |
|----------|--------|---------|-------------|------------|-------------|
| **COM Object Model** | No | No | Yes (read/write data) | Yes | Very High |
| **pywinauto (UIA)** | Yes | Yes | Limited | Partial | High |
| **uiautomation** | Yes | Yes | Limited | Partial | High |
| **SendInput (ctypes)** | Via coords | Via coords | Yes (click/drag) | No* | Medium |
| **SendMessage/PostMessage** | Limited | Limited | Unreliable | Yes | Low |
| **Keyboard shortcuts** | Yes | Yes | No | No* | High |

*SendInput and keyboard shortcuts require the window to be in the foreground.

### 8.2 Recommended Strategy for This Project

For the ppt-com-mcp server, the recommended approach is a **layered strategy**:

1. **Primary: COM Object Model** (already implemented)
   - Use for all data operations: reading/writing shapes, text, formatting, etc.
   - This is the most reliable and does not require foreground access
   - Already covers 95%+ of use cases

2. **Secondary: pywinauto/UIA for UI interaction**
   - Use for ribbon commands that have no COM equivalent
   - Use for dialog box interaction (e.g., custom dialogs, file pickers)
   - Use for reading UI state (active tab, panel visibility)

3. **Tertiary: SendInput for canvas interaction**
   - Use for drag-and-drop of shapes (visual repositioning)
   - Use for resize handle manipulation
   - Use for drawing tool operations (freeform, etc.)
   - Combine with COM to get precise shape coordinates

4. **Keyboard shortcuts as fallback**
   - Use when UI elements are hard to locate via UIA
   - Reliable for standard operations (Alt+key ribbon navigation)

### 8.3 Key Limitations and Caveats

1. **Foreground requirement**: SendInput and real mouse clicks require PowerPoint to be in the foreground. This means the MCP server cannot run these operations invisibly in the background.

2. **DPI scaling**: On high-DPI displays, coordinate calculations must account for the system DPI scaling factor. Use `ctypes.windll.user32.GetDpiForWindow(hwnd)` (Windows 10 1607+) or `ctypes.windll.shcore.GetDpiForMonitor()`.

3. **Multi-monitor**: Screen coordinate calculations must account for multi-monitor setups where coordinates can be negative (monitors to the left/above the primary).

4. **Zoom level**: The slide canvas zoom affects the mapping between shape coordinates (in points) and screen pixels. Always read `ActiveWindow.View.Zoom` before coordinate calculations.

5. **COM threading**: If mixing COM automation with UIA (which also uses COM internally), be careful with COM apartment models. Use STA (Single Threaded Apartment) consistently.

6. **Office version differences**: Ribbon layout, window class names, and UIA element trees can differ between Office versions. Test against the target version.

7. **Timing**: UI automation often requires `time.sleep()` delays between operations to wait for UI updates. Use `wait()` methods when available (pywinauto) or explicit delays (SendInput).

### 8.4 Library Installation

```bash
# pywinauto (includes UIA support)
pip install pywinauto

# uiautomation (standalone UIA wrapper)
pip install uiautomation

# pywin32 (already used by this project for COM)
pip install pywin32

# pyautogui (alternative, cross-platform mouse/keyboard)
pip install pyautogui
```

### 8.5 Debugging and Inspection Tools

| Tool | Purpose | How to Get |
|------|---------|-----------|
| **Accessibility Insights** | Inspect UIA element tree | [Microsoft download](https://accessibilityinsights.io/) |
| **Inspect.exe** | Legacy UIA/MSAA inspector | Windows SDK |
| **Spy++** | Win32 window hierarchy | Visual Studio |
| **print_control_identifiers()** | pywinauto element dump | Built into pywinauto |
| **uiautomation CLI** | uiautomation element dump | `python -m uiautomation` |

---

## References

- [pywinauto Documentation](https://pywinauto.readthedocs.io/en/latest/)
- [pywinauto Getting Started](https://pywinauto.readthedocs.io/en/latest/getting_started.html)
- [pywinauto How-To's](https://pywinauto.readthedocs.io/en/latest/HowTo.html)
- [pywinauto Mouse Module](https://pywinauto.readthedocs.io/en/latest/code/pywinauto.mouse.html)
- [pywinauto UIA Controls](https://pywinauto.readthedocs.io/en/latest/code/pywinauto.controls.uia_controls.html)
- [pywinauto GitHub](https://github.com/pywinauto/pywinauto)
- [Python-UIAutomation-for-Windows (uiautomation)](https://github.com/yinkaisheng/Python-UIAutomation-for-Windows)
- [uiautomation on PyPI](https://pypi.org/project/uiautomation/)
- [MSAA vs UIA Comparison (Microsoft Learn)](https://learn.microsoft.com/en-us/windows/win32/winauto/microsoft-active-accessibility-and-ui-automation-compared)
- [UI Automation Overview (Microsoft Learn)](https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-msaa)
- [Accessibility Insights for Windows](https://accessibilityinsights.io/)
- [Inspect.exe (Microsoft Learn)](https://learn.microsoft.com/en-us/windows/win32/winauto/inspect-objects)
- [win32gui.FindWindow (pywin32)](https://mhammond.github.io/pywin32/win32gui__FindWindow_meth.html)
- [SetForegroundWindow (Microsoft Learn)](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setforegroundwindow)
- [mouse_event (Microsoft Learn)](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-mouse_event)
- [SendInput ctypes Example (GitHub Gist)](https://gist.github.com/Aniruddha-Tapas/1627257344780e5429b10bc92eb2f52a)
- [pywinauto Drag-and-Drop Example (GitHub Gist)](https://gist.github.com/vasily-v-ryabov/f6c6f4d94fe313be8236)
