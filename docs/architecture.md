# ppt-com-mcp: Architecture Document

> **Project**: ppt-com-mcp
> **Version**: 1.0
> **Date**: 2026-02-16
> **Status**: Draft

---

## 1. System Architecture Overview

### 1.1 High-Level Architecture

```
+-------------------+       +-------------------+       +-------------------+
|                   |  MCP  |                   |  COM  |                   |
|  Claude / LLM     |<----->|  MCP Server       |<----->|  PowerPoint.exe   |
|  (Client)         |  JSON |  (Python/FastMCP) |  RPC  |  (Running App)    |
|                   |       |                   |       |                   |
+-------------------+       +-------------------+       +-------------------+
                            |                   |
                            |  +--------------+ |
                            |  | ppt_com/     | |
                            |  |  app.py      | |
                            |  |  slides.py   | |
                            |  |  shapes.py   | |
                            |  |  text.py     | |
                            |  |  ...         | |
                            |  +--------------+ |
                            |  | utils/       | |
                            |  |  units.py    | |
                            |  |  color.py    | |
                            |  |  com_wrapper | |
                            |  +--------------+ |
                            +-------------------+
```

### 1.2 Data Flow

```
User Prompt (natural language)
        |
        v
+------------------+
| Claude / LLM     |  -- interprets intent, selects MCP tool, builds args
+------------------+
        |  MCP tool call (JSON)
        v
+------------------+
| FastMCP Server   |  -- routes to registered tool function
| (server.py)      |
+------------------+
        |  Python function call
        v
+------------------+
| ppt_com module   |  -- validates input, converts units/colors
| (e.g. slides.py) |
+------------------+
        |  COM method call (via win32com)
        v
+------------------+
| COM Wrapper      |  -- manages COM lifecycle, error handling
| (com_wrapper.py) |
+------------------+
        |  COM RPC
        v
+------------------+
| PowerPoint.exe   |  -- executes operation, updates UI
+------------------+
        |
        v
Result bubbles back up: COM return value -> Python -> JSON -> MCP response -> LLM
```

### 1.3 Technology Stack

| Layer | Technology | Purpose |
|-------|-----------|---------|
| MCP Protocol | FastMCP (Python) | Expose tools to LLM clients |
| Application Logic | Python 3.10+ | Business logic, validation, orchestration |
| COM Bridge | pywin32 (`win32com.client`) | Bridge to PowerPoint COM API |
| COM Threading | `pythoncom` | COM apartment management |
| Target Application | Microsoft PowerPoint | Presentation engine |
| Logging | Python `logging` | Structured application logging |

---

## 2. Module Design

### 2.1 Directory Structure

```
ppt-com-mcp/
├── src/
│   ├── server.py                # MCP server entry point (FastMCP registration)
│   ├── ppt_com/
│   │   ├── __init__.py          # Package init, re-exports
│   │   ├── app.py               # Application connection & lifecycle
│   │   ├── presentation.py      # Presentation management (new, open, save, close)
│   │   ├── slides.py            # Slide operations (add, delete, duplicate, move, list)
│   │   ├── shapes.py            # Shape creation & manipulation
│   │   ├── text.py              # Text content & formatting operations
│   │   ├── placeholders.py      # Placeholder access & content operations
│   │   ├── tables.py            # Table creation, cell access, formatting
│   │   ├── charts.py            # Chart creation, data, formatting
│   │   ├── media.py             # Video/audio insertion & settings
│   │   ├── slideshow.py         # SlideShow control (start, stop, navigate)
│   │   ├── export.py            # Export operations (PDF, images, video)
│   │   ├── formatting.py        # Fill, Line, Shadow, Glow, Reflection, SoftEdge, 3D
│   │   ├── animation.py         # Animation effects & transitions
│   │   ├── themes.py            # Theme, master, layout, headers/footers
│   │   └── constants.py         # All PowerPoint/Office enum constants
│   └── utils/
│       ├── __init__.py
│       ├── units.py             # Unit conversion (pt, in, cm, emu)
│       ├── color.py             # Color helpers (RGB, hex, theme colors, BGR conversion)
│       └── com_wrapper.py       # COM lifecycle management & error handling
├── tests/
│   ├── __init__.py
│   ├── test_app.py
│   ├── test_presentation.py
│   ├── test_slides.py
│   ├── test_shapes.py
│   ├── test_text.py
│   ├── test_placeholders.py
│   ├── test_tables.py
│   └── ...
├── docs/
│   ├── requirements.md
│   ├── architecture.md
│   └── research/
│       ├── 01_application_window_presentation.md
│       ├── 02_slides_and_shapes.md
│       ├── 03_text_and_formatting.md
│       ├── 04_tables_charts_media_advanced.md
│       └── 05_placeholders_and_masters.md
├── pyproject.toml
└── README.md
```

### 2.2 Module Descriptions

#### `server.py` -- MCP Server Entry Point

- Initializes FastMCP server instance
- Registers all MCP tools from `ppt_com` modules
- Handles server lifecycle (startup, shutdown)
- Ensures COM is properly initialized/uninitialized on the server thread

#### `ppt_com/app.py` -- Application Connection & Lifecycle

- Connects to running PowerPoint via `GetActiveObject` or starts new instance via `Dispatch`
- Manages the singleton `Application` COM reference
- Provides application info (version, name, window state)
- Provides active window and selection info
- Handles reconnection if PowerPoint crashes

#### `ppt_com/presentation.py` -- Presentation Management

- Create new presentation (`Presentations.Add`)
- Open from file (`Presentations.Open`)
- Save / SaveAs / SaveCopyAs
- Close presentation (with save/no-save option)
- List open presentations
- Get presentation properties (name, path, slide count, page setup, document properties)
- Apply template/theme

#### `ppt_com/slides.py` -- Slide Operations

- Add slide (by layout index or layout name)
- Delete slide
- Duplicate slide
- Move slide
- List all slides (with summary info)
- Get slide details
- Navigate to slide (`View.GotoSlide`)
- Copy slides between presentations (`InsertFromFile`, `Copy`/`Paste`)

#### `ppt_com/shapes.py` -- Shape Creation & Manipulation

- Add shape (auto shape by type)
- Add text box
- Add image (from file path)
- Add line
- Delete shape
- Set position/size (Left, Top, Width, Height)
- Set rotation
- Set z-order (front, back, forward, backward)
- Set name, visibility
- Duplicate shape
- Flip shape (horizontal, vertical)
- List all shapes on a slide
- Get shape details
- Group/ungroup shapes
- Add connector shapes

#### `ppt_com/text.py` -- Text & Formatting Operations

- Set/get full text of a shape
- Set font formatting (name, size, bold, italic, underline, color, shadow, etc.)
- Format partial text via `Characters(Start, Length)`
- Insert text (before/after)
- Find and replace text
- Set paragraph formatting (alignment, spacing, indent level)
- Set bullet/numbering formatting
- Access text by words, lines, sentences, paragraphs, runs
- TextFrame properties (margins, word wrap, auto-size, orientation)

#### `ppt_com/placeholders.py` -- Placeholder Operations

- List placeholders on a slide (with type, name, contained type)
- Set placeholder text (by index or by type)
- Get placeholder details
- Find placeholder by type (title, body, subtitle, etc.)
- List available layouts
- Set slide layout

#### `ppt_com/tables.py` -- Table Operations

- Create table (rows, columns, position, size)
- Set/get cell text
- Get entire table data as 2D array
- Format cell (fill, borders, text alignment)
- Add/delete rows and columns
- Set row height, column width
- Merge/split cells
- Apply table style
- Table band settings (first row, last row, etc.)

#### `ppt_com/charts.py` -- Chart Operations

- Create chart (`AddChart2` with chart type)
- Set chart data via embedded Excel workbook
- Set chart title, legend, axes
- Format series colors, data labels
- Change chart type and style

#### `ppt_com/media.py` -- Media Operations

- Insert video (`AddMediaObject2`)
- Insert audio (`AddMediaObject2`)
- Set media properties (volume, mute, trim, fade)
- Set play settings (auto-play, loop, hide when not playing)

#### `ppt_com/slideshow.py` -- SlideShow Control

- Configure slide show settings (range, type, loop, advance mode)
- Start slide show
- Navigate (next, previous, first, last, go to slide)
- Set show state (running, paused, black/white screen)
- Get show status (current slide, elapsed time)
- Exit slide show

#### `ppt_com/export.py` -- Export Operations

- Export to PDF (`ExportAsFixedFormat` and `SaveAs`)
- Export slides as images (PNG, JPG, etc. with custom resolution)
- Export as video (WMV, MP4)
- Print presentation

#### `ppt_com/formatting.py` -- Visual Formatting

- Shape fill (solid, gradient, pattern, texture, picture)
- Shape line/border (color, weight, dash style, arrows)
- Shadow effect
- Glow effect
- Reflection effect
- Soft edge effect
- 3D format (bevel, extrusion, rotation, lighting, material)

#### `ppt_com/animation.py` -- Animation & Transitions

- Set slide transition (effect, duration, advance settings)
- Add animation effect to shape (`TimeLine.MainSequence.AddEffect`)
- Set animation trigger and timing

#### `ppt_com/themes.py` -- Theme & Master Operations

- Apply theme/template
- Get theme color scheme
- List/manage slide masters (via Designs)
- List/manage custom layouts
- Set headers/footers (footer text, date/time, slide number)
- Master background settings
- TextStyles management

#### `ppt_com/constants.py` -- PowerPoint/Office Constants

- All PpSlideLayout values
- All MsoAutoShapeType values
- All PpPlaceholderType values
- All MsoShapeType values
- All color/formatting enumerations
- All PpSaveAsFileType values
- Friendly name-to-value mappings for LLM usability

#### `utils/units.py` -- Unit Conversion

- Points to/from inches
- Points to/from centimeters
- Points to/from EMU
- Parse value strings with unit suffixes (e.g., `"2.5in"`, `"72pt"`, `"5cm"`)

#### `utils/color.py` -- Color Helpers

- RGB tuple `(r, g, b)` to BGR integer (PowerPoint format)
- Hex string `"#RRGGBB"` to BGR integer
- BGR integer to RGB tuple
- Theme color name to `MsoThemeColorIndex` value
- Named colors (e.g., `"red"`, `"blue"`) to RGB values

#### `utils/com_wrapper.py` -- COM Lifecycle Management

- COM initialization (`CoInitializeEx`) and cleanup (`CoUninitialize`)
- COM error handling (translate `pywintypes.com_error` to structured errors)
- Connection management (connect, reconnect, disconnect)
- Reference cleanup helpers

---

## 3. Module Dependency Diagram

```
server.py
    |
    +---> ppt_com/app.py ------------> utils/com_wrapper.py
    |         |
    +---> ppt_com/presentation.py ---> utils/com_wrapper.py
    |         |                    \-> utils/units.py
    |         |
    +---> ppt_com/slides.py ---------> utils/com_wrapper.py
    |         |                    \-> ppt_com/app.py (get active presentation)
    |         |
    +---> ppt_com/shapes.py ---------> utils/com_wrapper.py
    |         |                    \-> utils/units.py
    |         |                    \-> utils/color.py
    |         |
    +---> ppt_com/text.py -----------> utils/com_wrapper.py
    |         |                    \-> utils/color.py
    |         |
    +---> ppt_com/placeholders.py ---> utils/com_wrapper.py
    |         |                    \-> ppt_com/app.py
    |         |
    +---> ppt_com/tables.py ---------> utils/com_wrapper.py
    |         |                    \-> utils/color.py
    |         |
    +---> ppt_com/charts.py ---------> utils/com_wrapper.py
    |         |
    +---> ppt_com/media.py ----------> utils/com_wrapper.py
    |         |
    +---> ppt_com/slideshow.py ------> utils/com_wrapper.py
    |         |                    \-> ppt_com/app.py
    |         |
    +---> ppt_com/export.py ---------> utils/com_wrapper.py
    |         |
    +---> ppt_com/formatting.py -----> utils/com_wrapper.py
    |         |                    \-> utils/color.py
    |         |
    +---> ppt_com/animation.py ------> utils/com_wrapper.py
    |         |
    +---> ppt_com/themes.py ---------> utils/com_wrapper.py
              |                    \-> utils/color.py

All ppt_com modules import from:
  - ppt_com/constants.py (enum values)
  - utils/com_wrapper.py (COM access, error handling)
```

**Key Dependencies**:
- Every `ppt_com/*` module depends on `utils/com_wrapper.py` for COM access
- Modules that deal with position/size depend on `utils/units.py`
- Modules that deal with colors depend on `utils/color.py`
- All modules use `ppt_com/constants.py` for enum values
- `ppt_com/app.py` is the central module that manages the Application COM object; other modules obtain COM references through it

---

## 4. COM Management Pattern

### 4.1 Connection Lifecycle

```python
# Singleton pattern for COM Application reference

class PowerPointApp:
    _instance = None
    _app = None

    @classmethod
    def get_app(cls):
        """Get or create the PowerPoint Application COM object."""
        if cls._app is not None:
            try:
                # Verify connection is still alive
                _ = cls._app.Version
                return cls._app
            except Exception:
                cls._app = None  # Connection lost, will reconnect

        # Try to connect to running instance first
        try:
            cls._app = win32com.client.GetActiveObject("PowerPoint.Application")
        except Exception:
            # No running instance, create new one
            cls._app = win32com.client.Dispatch("PowerPoint.Application")
            cls._app.Visible = True

        return cls._app

    @classmethod
    def release(cls):
        """Release the COM reference."""
        if cls._app is not None:
            try:
                cls._app = None
            except Exception:
                pass
            gc.collect()
```

### 4.2 Thread Safety Approach

```
MCP Request Thread(s)
        |
        v
+-------------------+
| Request Queue     |  -- all tool calls are serialized here
+-------------------+
        |
        v
+-------------------+
| COM Worker Thread |  -- single thread with STA COM initialization
| (pythoncom.       |
|  CoInitializeEx)  |
+-------------------+
        |
        v
+-------------------+
| PowerPoint COM    |
+-------------------+
```

All COM operations run on a single dedicated thread. The FastMCP server dispatches tool calls to this thread via a thread-safe queue. This ensures:

1. COM apartment threading rules are respected
2. No concurrent COM access (which would cause crashes)
3. The MCP server can still accept new requests while a COM operation is running

### 4.3 Error Recovery Patterns

```python
import pywintypes

def safe_com_call(func, *args, **kwargs):
    """Execute a COM call with error handling and retry logic."""
    max_retries = 2
    for attempt in range(max_retries):
        try:
            return func(*args, **kwargs)
        except pywintypes.com_error as e:
            hr = e.hresult
            desc = ""
            if e.excepinfo and len(e.excepinfo) > 2:
                desc = e.excepinfo[2] or ""

            # RPC_E_CALL_REJECTED (-2147418111) -- PowerPoint is busy
            if hr == -2147418111 and attempt < max_retries - 1:
                time.sleep(0.5)
                continue

            # RPC_E_DISCONNECTED (-2147417848) -- PowerPoint closed
            if hr == -2147417848:
                PowerPointApp.release()
                raise ConnectionError(
                    f"PowerPoint connection lost: {desc}"
                ) from e

            # Other COM errors
            raise COMOperationError(
                hresult=hr,
                message=desc or str(e.strerror),
                source=e.excepinfo[1] if e.excepinfo else None,
            ) from e
```

### 4.4 COM Reference Cleanup

```python
def cleanup_com_ref(*refs):
    """Explicitly release COM references."""
    for ref in refs:
        try:
            if ref is not None:
                ref = None
        except Exception:
            pass
    gc.collect()
```

---

## 5. MCP Tool Naming Convention

### 5.1 Pattern

All MCP tools follow the naming convention:

```
ppt_{module}_{action}
```

Where:
- `ppt` is the global prefix (identifies this as a PowerPoint tool)
- `{module}` is the functional area (e.g., `app`, `slide`, `shape`, `text`)
- `{action}` is the operation (e.g., `connect`, `add`, `set_format`, `list`)

### 5.2 Complete Tool Name Registry

#### Application Tools (ppt_app_*)

| Tool Name | Feature ID | Description |
|-----------|-----------|-------------|
| `ppt_app_connect` | F-APP-001 | Connect to running PowerPoint or launch new instance |
| `ppt_app_get_info` | F-APP-002 | Get application info (version, state, presentations) |
| `ppt_app_get_active_window` | F-APP-003 | Get active window, view, and selection info |

#### Presentation Tools (ppt_pres_*)

| Tool Name | Feature ID | Description |
|-----------|-----------|-------------|
| `ppt_pres_create` | F-PRES-001 | Create new blank presentation |
| `ppt_pres_open` | F-PRES-002 | Open presentation from file |
| `ppt_pres_save` | F-PRES-003 | Save presentation (save, save-as, save-copy) |
| `ppt_pres_close` | F-PRES-004 | Close presentation |
| `ppt_pres_list` | F-PRES-005 | List all open presentations |
| `ppt_pres_get_info` | F-PRES-006 | Get presentation details |

#### Slide Tools (ppt_slide_*)

| Tool Name | Feature ID | Description |
|-----------|-----------|-------------|
| `ppt_slide_add` | F-SLIDE-001 | Add new slide with layout |
| `ppt_slide_delete` | F-SLIDE-002 | Delete slide by index |
| `ppt_slide_duplicate` | F-SLIDE-003 | Duplicate slide |
| `ppt_slide_move` | F-SLIDE-004 | Move slide to new position |
| `ppt_slide_list` | F-SLIDE-005 | List all slides with summary |
| `ppt_slide_get_info` | F-SLIDE-006 | Get slide details |
| `ppt_slide_goto` | F-SLIDE-007 | Navigate to specific slide |

#### Shape Tools (ppt_shape_*)

| Tool Name | Feature ID | Description |
|-----------|-----------|-------------|
| `ppt_shape_add_shape` | F-SHAPE-001 | Add auto shape |
| `ppt_shape_add_textbox` | F-SHAPE-002 | Add text box |
| `ppt_shape_add_image` | F-SHAPE-003 | Add image from file |
| `ppt_shape_add_line` | F-SHAPE-004 | Add line |
| `ppt_shape_delete` | F-SHAPE-005 | Delete shape |
| `ppt_shape_set_position` | F-SHAPE-006 | Set shape position and size |
| `ppt_shape_list` | F-SHAPE-007 | List shapes on slide |
| `ppt_shape_get_info` | F-SHAPE-008 | Get shape details |
| `ppt_shape_duplicate` | F-SHAPE-009 | Duplicate shape |
| `ppt_shape_set_zorder` | F-SHAPE-010 | Set shape z-order |

#### Text Tools (ppt_text_*)

| Tool Name | Feature ID | Description |
|-----------|-----------|-------------|
| `ppt_text_set` | F-TEXT-001 | Set shape text content |
| `ppt_text_get` | F-TEXT-002 | Get shape text content |
| `ppt_text_set_font` | F-TEXT-003 | Set font formatting |
| `ppt_text_format_range` | F-TEXT-004 | Format partial text (characters range) |
| `ppt_text_insert` | F-TEXT-005 | Insert text before/after |
| `ppt_text_replace` | F-TEXT-006 | Find and replace text |

#### Placeholder Tools (ppt_placeholder_*)

| Tool Name | Feature ID | Description |
|-----------|-----------|-------------|
| `ppt_placeholder_list` | F-PH-001 | List placeholders on slide |
| `ppt_placeholder_set_text` | F-PH-002 | Set placeholder text |
| `ppt_placeholder_get_info` | F-PH-003 | Get placeholder details |
| `ppt_placeholder_list_layouts` | F-PH-004 | List available slide layouts |

#### Table Tools (ppt_table_*)

| Tool Name | Feature ID | Description |
|-----------|-----------|-------------|
| `ppt_table_create` | F-TABLE-001 | Create table |
| `ppt_table_set_cell` | F-TABLE-002 | Set cell text and formatting |
| `ppt_table_get_data` | F-TABLE-003 | Get table data as 2D array |
| `ppt_table_format_cell` | F-TABLE-004 | Format cell (fill, borders) |
| `ppt_table_modify_structure` | F-TABLE-005 | Add/delete rows and columns |
| `ppt_table_merge_cells` | F-TABLE-006 | Merge or split cells |

#### Formatting Tools (ppt_format_*)

| Tool Name | Feature ID | Description |
|-----------|-----------|-------------|
| `ppt_format_set_fill` | F-FILL-001 | Set shape fill |
| `ppt_format_set_line` | F-LINE-001 | Set shape line/border |
| `ppt_format_set_paragraph` | F-PARA-001 | Set paragraph formatting |
| `ppt_format_set_bullet` | F-PARA-002 | Set bullet/numbering |
| `ppt_format_set_background` | F-BG-001 | Set slide background |
| `ppt_format_set_shadow` | F-EFFECT-001 | Set shape shadow |
| `ppt_format_set_glow` | F-EFFECT-002 | Set shape glow |
| `ppt_format_set_reflection` | F-EFFECT-003 | Set shape reflection |
| `ppt_format_set_soft_edge` | F-EFFECT-004 | Set shape soft edge |

#### Export Tools (ppt_export_*)

| Tool Name | Feature ID | Description |
|-----------|-----------|-------------|
| `ppt_export_pdf` | F-EXPORT-001 | Export to PDF |
| `ppt_export_images` | F-EXPORT-002 | Export slides as images |

#### SlideShow Tools (ppt_slideshow_*)

| Tool Name | Feature ID | Description |
|-----------|-----------|-------------|
| `ppt_slideshow_start` | F-SS-001 | Start slide show |
| `ppt_slideshow_control` | F-SS-002 | Navigate slide show |
| `ppt_slideshow_get_status` | F-SS-003 | Get slide show status |

#### Notes Tools (ppt_notes_*)

| Tool Name | Feature ID | Description |
|-----------|-----------|-------------|
| `ppt_notes_set` | F-NOTES-001 | Set slide notes |
| `ppt_notes_get` | F-NOTES-002 | Get slide notes |

#### Chart Tools (ppt_chart_*)

| Tool Name | Feature ID | Description |
|-----------|-----------|-------------|
| `ppt_chart_create` | F-CHART-001 | Create chart |
| `ppt_chart_set_data` | F-CHART-002 | Set chart data |
| `ppt_chart_format` | F-CHART-003 | Format chart elements |

#### SmartArt Tools (ppt_smartart_*)

| Tool Name | Feature ID | Description |
|-----------|-----------|-------------|
| `ppt_smartart_create` | F-SMART-001 | Create SmartArt |
| `ppt_smartart_modify` | F-SMART-002 | Modify SmartArt nodes |

#### Animation Tools (ppt_animation_*)

| Tool Name | Feature ID | Description |
|-----------|-----------|-------------|
| `ppt_animation_set_transition` | F-ANIM-001 | Set slide transition |
| `ppt_animation_add_effect` | F-ANIM-002 | Add shape animation |

#### Media Tools (ppt_media_*)

| Tool Name | Feature ID | Description |
|-----------|-----------|-------------|
| `ppt_media_add_video` | F-MEDIA-001 | Insert video |
| `ppt_media_add_audio` | F-MEDIA-002 | Insert audio |
| `ppt_media_set_settings` | F-MEDIA-003 | Configure media playback |

#### Theme Tools (ppt_theme_*)

| Tool Name | Feature ID | Description |
|-----------|-----------|-------------|
| `ppt_theme_apply` | F-THEME-001 | Apply theme file |
| `ppt_theme_get_colors` | F-THEME-002 | Get theme color scheme |
| `ppt_theme_manage_masters` | F-THEME-003 | Manage masters and layouts |

#### Other Tools

| Tool Name | Feature ID | Description |
|-----------|-----------|-------------|
| `ppt_hyperlink_add` | F-LINK-001 | Add hyperlink |
| `ppt_hyperlink_list` | F-LINK-002 | List hyperlinks |
| `ppt_ole_add` | F-OLE-001 | Insert OLE object |
| `ppt_connector_add` | F-CONN-001 | Add connector shape |
| `ppt_group_shapes` | F-GROUP-001 | Group/ungroup shapes |
| `ppt_print` | F-PRINT-001 | Print presentation |
| `ppt_section_manage` | F-SECTION-001 | Manage sections |
| `ppt_property_set` | F-PROP-001 | Set document properties |
| `ppt_3d_set` | F-3D-001 | Set 3D effects |
| `ppt_headersfooters_set` | F-HF-001 | Set headers/footers |

---

## 6. Data Flow: Detailed Example

### 6.1 Example: `ppt_text_format_range` Tool Call

This example traces a request to format characters 5-10 of a shape's text as bold red.

```
Step 1: LLM sends MCP tool call
  {
    "tool": "ppt_text_format_range",
    "arguments": {
      "slide_index": 1,
      "shape_name": "Title 1",
      "start": 5,
      "length": 6,
      "bold": true,
      "color": "#FF0000"
    }
  }

Step 2: FastMCP routes to registered handler in server.py
  -> calls ppt_com.text.format_range(slide_index=1, shape_name="Title 1", ...)

Step 3: text.py validates inputs
  - Checks start > 0, length > 0
  - Parses color "#FF0000" via utils/color.py -> BGR integer 0x0000FF
  - Gets Application via ppt_com/app.py

Step 4: text.py acquires COM objects
  - app = PowerPointApp.get_app()
  - pres = app.ActivePresentation
  - slide = pres.Slides(1)
  - shape = slide.Shapes("Title 1")
  - Verifies shape.HasTextFrame

Step 5: text.py executes COM operations (via com_wrapper.py)
  - text_range = shape.TextFrame.TextRange
  - Validates start + length <= text_range.Length
  - char_range = text_range.Characters(5, 6)
  - char_range.Font.Bold = -1  (msoTrue)
  - char_range.Font.Color.RGB = 0x0000FF  (red in BGR)

Step 6: text.py builds response
  {
    "success": true,
    "formatted_text": "World!",
    "slide_index": 1,
    "shape_name": "Title 1",
    "range": {"start": 5, "length": 6},
    "applied": {"bold": true, "color": "#FF0000"}
  }

Step 7: FastMCP returns JSON response to LLM
```

### 6.2 Error Handling Flow

```
COM operation fails
        |
        v
pywintypes.com_error raised
        |
        v
com_wrapper.py catches it
  - Extracts HRESULT, strerror, excepinfo
  - Checks for retryable errors (RPC_E_CALL_REJECTED)
  - Checks for connection-lost errors (RPC_E_DISCONNECTED)
        |
        +---> Retryable? -> sleep -> retry (max 2 attempts)
        |
        +---> Connection lost? -> PowerPointApp.release()
        |                      -> raise ConnectionError
        |
        +---> Other error? -> raise COMOperationError(hresult, message)
                |
                v
Tool function catches COMOperationError
        |
        v
Returns structured MCP error response:
  {
    "error": true,
    "error_type": "com_error",
    "message": "Cannot access shape 'NonExistent' on slide 1",
    "details": {
      "hresult": "0x80004005",
      "source": "Microsoft PowerPoint"
    }
  }
```

---

## 7. Development Phases

### Phase 1: Core Infrastructure + MVP Tools

**Goal**: A working MCP server that can connect to PowerPoint, create content, and perform basic operations.

**Deliverables**:
1. `utils/com_wrapper.py` -- COM lifecycle and error handling
2. `utils/units.py` -- Unit conversion
3. `utils/color.py` -- Color conversion (RGB/hex to BGR)
4. `ppt_com/constants.py` -- Essential constants (layouts, shape types, placeholder types)
5. `ppt_com/app.py` -- Application connection (`ppt_app_connect`, `ppt_app_get_info`, `ppt_app_get_active_window`)
6. `ppt_com/presentation.py` -- Presentation CRUD (`ppt_pres_create`, `ppt_pres_open`, `ppt_pres_save`, `ppt_pres_close`, `ppt_pres_list`, `ppt_pres_get_info`)
7. `ppt_com/slides.py` -- Slide CRUD (`ppt_slide_add`, `ppt_slide_delete`, `ppt_slide_duplicate`, `ppt_slide_move`, `ppt_slide_list`, `ppt_slide_get_info`, `ppt_slide_goto`)
8. `ppt_com/shapes.py` -- Shape basics (`ppt_shape_add_shape`, `ppt_shape_add_textbox`, `ppt_shape_add_image`, `ppt_shape_add_line`, `ppt_shape_delete`, `ppt_shape_set_position`, `ppt_shape_list`, `ppt_shape_get_info`, `ppt_shape_duplicate`, `ppt_shape_set_zorder`)
9. `ppt_com/text.py` -- Text operations (`ppt_text_set`, `ppt_text_get`, `ppt_text_set_font`, `ppt_text_format_range`, `ppt_text_insert`, `ppt_text_replace`)
10. `ppt_com/placeholders.py` -- Placeholder operations (`ppt_placeholder_list`, `ppt_placeholder_set_text`, `ppt_placeholder_get_info`, `ppt_placeholder_list_layouts`)
11. `server.py` -- FastMCP server with all Phase 1 tools registered

**Estimated Effort**: Core sprint

---

### Phase 2: Formatting + Tables + Export + SlideShow

**Goal**: Rich formatting capabilities, table support, export features, and slide show control.

**Deliverables**:
1. `ppt_com/formatting.py` -- Fill, Line, Shadow, Glow, Reflection, SoftEdge (`ppt_format_set_fill`, `ppt_format_set_line`, `ppt_format_set_shadow`, `ppt_format_set_glow`, `ppt_format_set_reflection`, `ppt_format_set_soft_edge`)
2. `ppt_com/text.py` additions -- Paragraph format, bullet format (`ppt_format_set_paragraph`, `ppt_format_set_bullet`)
3. `ppt_com/tables.py` -- Full table support (`ppt_table_create`, `ppt_table_set_cell`, `ppt_table_get_data`, `ppt_table_format_cell`, `ppt_table_modify_structure`, `ppt_table_merge_cells`)
4. `ppt_com/slides.py` additions -- Slide background (`ppt_format_set_background`)
5. `ppt_com/export.py` -- PDF and image export (`ppt_export_pdf`, `ppt_export_images`)
6. `ppt_com/slideshow.py` -- SlideShow control (`ppt_slideshow_start`, `ppt_slideshow_control`, `ppt_slideshow_get_status`)
7. Notes support in `ppt_com/slides.py` -- (`ppt_notes_set`, `ppt_notes_get`)

**Estimated Effort**: Second sprint

---

### Phase 3: Charts + Media + Animation + Advanced

**Goal**: Complete feature coverage for advanced PowerPoint capabilities.

**Deliverables**:
1. `ppt_com/charts.py` -- Chart operations (`ppt_chart_create`, `ppt_chart_set_data`, `ppt_chart_format`)
2. `ppt_com/media.py` -- Media operations (`ppt_media_add_video`, `ppt_media_add_audio`, `ppt_media_set_settings`)
3. `ppt_com/animation.py` -- Animation and transitions (`ppt_animation_set_transition`, `ppt_animation_add_effect`)
4. `ppt_com/themes.py` -- Theme management (`ppt_theme_apply`, `ppt_theme_get_colors`, `ppt_theme_manage_masters`, `ppt_headersfooters_set`)
5. Advanced shape operations -- Connectors, groups, OLE, SmartArt (`ppt_connector_add`, `ppt_group_shapes`, `ppt_ole_add`, `ppt_smartart_create`, `ppt_smartart_modify`)
6. Miscellaneous -- Hyperlinks, sections, properties, print, 3D (`ppt_hyperlink_add`, `ppt_hyperlink_list`, `ppt_section_manage`, `ppt_property_set`, `ppt_print`, `ppt_3d_set`)
7. `ppt_com/constants.py` extensions -- Complete enum coverage

**Estimated Effort**: Third sprint

---

## Appendix A: Key COM Object Hierarchy Reference

```
Application
├── Presentations (collection)
│   └── Presentation
│       ├── Slides (collection)
│       │   └── Slide
│       │       ├── Shapes (collection)
│       │       │   ├── Shape
│       │       │   │   ├── TextFrame / TextFrame2
│       │       │   │   │   └── TextRange
│       │       │   │   │       ├── Font / ColorFormat
│       │       │   │   │       ├── ParagraphFormat / BulletFormat
│       │       │   │   │       └── Characters() / Words() / Paragraphs()
│       │       │   │   ├── Fill (FillFormat)
│       │       │   │   ├── Line (LineFormat)
│       │       │   │   ├── Shadow (ShadowFormat)
│       │       │   │   ├── Glow (GlowFormat)
│       │       │   │   ├── Reflection (ReflectionFormat)
│       │       │   │   ├── SoftEdge (SoftEdgeFormat)
│       │       │   │   ├── ThreeD (ThreeDFormat)
│       │       │   │   ├── PlaceholderFormat
│       │       │   │   ├── Table
│       │       │   │   │   ├── Rows / Columns
│       │       │   │   │   └── Cell -> Shape -> TextFrame
│       │       │   │   ├── Chart
│       │       │   │   │   ├── ChartData -> Workbook (Excel)
│       │       │   │   │   ├── SeriesCollection
│       │       │   │   │   └── Axes / Legend / ChartTitle
│       │       │   │   ├── SmartArt -> AllNodes / Nodes
│       │       │   │   ├── MediaFormat
│       │       │   │   ├── AnimationSettings
│       │       │   │   ├── ActionSettings -> Hyperlink
│       │       │   │   └── ConnectorFormat
│       │       │   └── Placeholders (collection of Shape)
│       │       ├── SlideShowTransition
│       │       ├── TimeLine -> MainSequence (Effects)
│       │       ├── NotesPage -> Shapes -> Placeholders(2) -> TextFrame
│       │       ├── Background -> Fill
│       │       ├── HeadersFooters
│       │       └── CustomLayout
│       ├── SlideMaster (Master)
│       │   ├── CustomLayouts (collection)
│       │   │   └── CustomLayout -> Shapes -> Placeholders
│       │   ├── Shapes / Placeholders
│       │   ├── TextStyles (Title/Body/Default)
│       │   ├── Theme -> ThemeColorScheme
│       │   ├── Background
│       │   └── HeadersFooters
│       ├── Designs (collection)
│       │   └── Design -> SlideMaster
│       ├── PageSetup
│       ├── SlideShowSettings
│       ├── PrintOptions
│       ├── SectionProperties
│       ├── BuiltInDocumentProperties
│       └── NotesMaster / HandoutMaster
├── Windows (collection)
│   └── DocumentWindow
│       ├── View (zoom, GotoSlide)
│       ├── Selection (Type, ShapeRange, TextRange, SlideRange)
│       └── Panes
├── SlideShowWindows (collection)
│   └── SlideShowWindow
│       └── SlideShowView (navigation, state, pointer)
└── SmartArtLayouts / SmartArtColors / SmartArtQuickStyles
```

## Appendix B: Unit Conversion Reference

| From | To Points | Formula |
|------|----------|---------|
| Inches | Points | `value * 72` |
| Centimeters | Points | `value * 28.3465` |
| EMU | Points | `value / 12700` |
| Points | Inches | `value / 72` |
| Points | Centimeters | `value / 28.3465` |
| Points | EMU | `value * 12700` |

**Standard slide sizes in points**:
- 16:9 (Widescreen): 960 x 540 pt (13.333" x 7.5")
- 4:3 (Standard): 720 x 540 pt (10" x 7.5")
- A4 Landscape: 842 x 595 pt (11.693" x 8.268")

## Appendix C: BGR Color Encoding

PowerPoint COM stores colors as BGR (Blue-Green-Red) integers:

```
BGR_value = R + (G * 256) + (B * 65536)
         = R + (G << 8) + (B << 16)
```

| Color | RGB | BGR Integer | Hex (as stored) |
|-------|-----|-------------|-----------------|
| Red | (255, 0, 0) | 255 | 0x0000FF |
| Green | (0, 255, 0) | 65280 | 0x00FF00 |
| Blue | (0, 0, 255) | 16711680 | 0xFF0000 |
| White | (255, 255, 255) | 16777215 | 0xFFFFFF |
| Black | (0, 0, 0) | 0 | 0x000000 |

The server's `utils/color.py` module must transparently convert user-provided `#RRGGBB` hex strings or `(R, G, B)` tuples to BGR integers before passing to COM.
