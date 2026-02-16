# ppt-com-mcp: Requirements Document

> **Project**: ppt-com-mcp
> **Version**: 1.0
> **Date**: 2026-02-16
> **Status**: Draft

---

## 1. Project Overview

### 1.1 Project Name

**ppt-com-mcp** -- PowerPoint COM MCP Server

### 1.2 Goal

Build the world's best PowerPoint MCP (Model Context Protocol) server that provides real-time COM control of a running Microsoft PowerPoint instance. Unlike static file-generation libraries (e.g., python-pptx), this server enables live, interactive manipulation of PowerPoint through an LLM (Claude, etc.) via the MCP protocol.

### 1.3 Key Differentiator

**Real-time control of a running PowerPoint application** -- not just file generation. The server connects to an active PowerPoint COM instance, allowing an LLM to:

- See and modify the current presentation in real time
- Control slide shows (start, stop, navigate)
- Read the current selection and active window state
- Make changes that are immediately visible in the PowerPoint UI
- Leverage the full PowerPoint COM object model (180+ shape types, all formatting options, transitions, animations, etc.)

### 1.4 Technology Stack

| Component | Technology |
|-----------|-----------|
| Language | Python 3.10+ |
| COM Bridge | `win32com.client` (pywin32) |
| MCP Framework | FastMCP (Python) |
| Platform | Windows only (COM dependency) |
| Target Application | Microsoft PowerPoint (Office 2016 / 2019 / 2021 / 365) |

---

## 2. Feature Requirements

### Priority 1 -- MVP (Must Have)

These features form the Minimum Viable Product. They cover the core loop of connecting to PowerPoint, creating content, and manipulating it.

---

#### F-APP-001: Connect to PowerPoint

| Field | Value |
|-------|-------|
| **Feature ID** | F-APP-001 |
| **Name** | Connect to running PowerPoint |
| **Description** | Connect to an already-running PowerPoint instance via `GetActiveObject`, or launch a new instance via `Dispatch`. Return application version and status info. |
| **COM Objects** | `Application` (`win32com.client.GetActiveObject`, `Dispatch`) |
| **Priority** | P1 |

#### F-APP-002: Get Application Info

| Field | Value |
|-------|-------|
| **Feature ID** | F-APP-002 |
| **Name** | Get application info |
| **Description** | Return the PowerPoint application state: version, name, window state, visible status, number of open presentations, and active presentation name. |
| **COM Objects** | `Application.Version`, `.Name`, `.Visible`, `.WindowState`, `.Presentations.Count`, `.ActivePresentation` |
| **Priority** | P1 |

#### F-APP-003: Get Active Window Info

| Field | Value |
|-------|-------|
| **Feature ID** | F-APP-003 |
| **Name** | Get active window and selection info |
| **Description** | Return the active document window caption, view type, current slide index, and current selection type (none/slides/shapes/text). |
| **COM Objects** | `Application.ActiveWindow`, `.Selection`, `.View`, `DocumentWindow.ViewType` |
| **Priority** | P1 |

---

#### F-PRES-001: Create New Presentation

| Field | Value |
|-------|-------|
| **Feature ID** | F-PRES-001 |
| **Name** | Create new presentation |
| **Description** | Create a new blank presentation. Optionally set the slide size (16:9, 4:3, custom width/height). |
| **COM Objects** | `Presentations.Add()`, `PageSetup.SlideWidth`, `.SlideHeight` |
| **Priority** | P1 |

#### F-PRES-002: Open Presentation

| Field | Value |
|-------|-------|
| **Feature ID** | F-PRES-002 |
| **Name** | Open existing presentation |
| **Description** | Open a presentation from a file path. Support read-only mode and opening without a window (background). |
| **COM Objects** | `Presentations.Open(FileName, ReadOnly, WithWindow)` |
| **Priority** | P1 |

#### F-PRES-003: Save Presentation

| Field | Value |
|-------|-------|
| **Feature ID** | F-PRES-003 |
| **Name** | Save presentation |
| **Description** | Save the active or specified presentation. Support Save (overwrite), SaveAs (with format and path), and SaveCopyAs. |
| **COM Objects** | `Presentation.Save()`, `.SaveAs(FileName, FileFormat)`, `.SaveCopyAs()` |
| **Priority** | P1 |

#### F-PRES-004: Close Presentation

| Field | Value |
|-------|-------|
| **Feature ID** | F-PRES-004 |
| **Name** | Close presentation |
| **Description** | Close the active or specified presentation. Support closing without saving (via `Saved = True` before close). |
| **COM Objects** | `Presentation.Close()`, `.Saved` |
| **Priority** | P1 |

#### F-PRES-005: List Presentations

| Field | Value |
|-------|-------|
| **Feature ID** | F-PRES-005 |
| **Name** | List open presentations |
| **Description** | Return a list of all currently open presentations with their name, path, slide count, and saved status. |
| **COM Objects** | `Application.Presentations`, `Presentation.Name`, `.FullName`, `.Slides.Count`, `.Saved` |
| **Priority** | P1 |

#### F-PRES-006: Get Presentation Info

| Field | Value |
|-------|-------|
| **Feature ID** | F-PRES-006 |
| **Name** | Get presentation details |
| **Description** | Return detailed information about a presentation: name, path, slide count, slide size, template name, and built-in document properties (title, author, subject, etc.). |
| **COM Objects** | `Presentation.Name`, `.FullName`, `.PageSetup`, `.BuiltInDocumentProperties`, `.TemplateName` |
| **Priority** | P1 |

---

#### F-SLIDE-001: Add Slide

| Field | Value |
|-------|-------|
| **Feature ID** | F-SLIDE-001 |
| **Name** | Add slide |
| **Description** | Add a new slide at a specified position. Support layout specification by PpSlideLayout integer or by CustomLayout name. |
| **COM Objects** | `Slides.Add(Index, Layout)`, `Slides.AddSlide(Index, pCustomLayout)`, `SlideMaster.CustomLayouts` |
| **Priority** | P1 |

#### F-SLIDE-002: Delete Slide

| Field | Value |
|-------|-------|
| **Feature ID** | F-SLIDE-002 |
| **Name** | Delete slide |
| **Description** | Delete a slide by index. |
| **COM Objects** | `Slide.Delete()` |
| **Priority** | P1 |

#### F-SLIDE-003: Duplicate Slide

| Field | Value |
|-------|-------|
| **Feature ID** | F-SLIDE-003 |
| **Name** | Duplicate slide |
| **Description** | Duplicate a slide. The copy is inserted immediately after the original. Optionally move it to a target position. |
| **COM Objects** | `Slide.Duplicate()`, `Slide.MoveTo(toPos)` |
| **Priority** | P1 |

#### F-SLIDE-004: Move Slide

| Field | Value |
|-------|-------|
| **Feature ID** | F-SLIDE-004 |
| **Name** | Move slide |
| **Description** | Move a slide from its current position to a new position in the slide collection. |
| **COM Objects** | `Slide.MoveTo(toPos)` |
| **Priority** | P1 |

#### F-SLIDE-005: List Slides

| Field | Value |
|-------|-------|
| **Feature ID** | F-SLIDE-005 |
| **Name** | List slides |
| **Description** | Return a list of all slides with index, slide number, name, layout name, and placeholder summary. |
| **COM Objects** | `Slides`, `Slide.SlideIndex`, `.SlideNumber`, `.Name`, `.CustomLayout.Name`, `.Shapes.Placeholders` |
| **Priority** | P1 |

#### F-SLIDE-006: Get Slide Info

| Field | Value |
|-------|-------|
| **Feature ID** | F-SLIDE-006 |
| **Name** | Get slide details |
| **Description** | Return detailed information about a slide: index, name, layout, shape count, placeholder list with types, and notes text. |
| **COM Objects** | `Slide.SlideIndex`, `.Name`, `.CustomLayout`, `.Shapes`, `.NotesPage` |
| **Priority** | P1 |

#### F-SLIDE-007: Go To Slide

| Field | Value |
|-------|-------|
| **Feature ID** | F-SLIDE-007 |
| **Name** | Navigate to slide |
| **Description** | Navigate the active window to display a specific slide by index. |
| **COM Objects** | `ActiveWindow.View.GotoSlide(Index)` |
| **Priority** | P1 |

---

#### F-SHAPE-001: Add Shape

| Field | Value |
|-------|-------|
| **Feature ID** | F-SHAPE-001 |
| **Name** | Add auto shape |
| **Description** | Add an auto shape (rectangle, oval, arrow, star, etc.) to a slide at a specified position and size. Accept MsoAutoShapeType integer or a friendly name. |
| **COM Objects** | `Shapes.AddShape(Type, Left, Top, Width, Height)` |
| **Priority** | P1 |

#### F-SHAPE-002: Add Text Box

| Field | Value |
|-------|-------|
| **Feature ID** | F-SHAPE-002 |
| **Name** | Add text box |
| **Description** | Add a text box with specified position, size, and optional initial text. |
| **COM Objects** | `Shapes.AddTextbox(Orientation, Left, Top, Width, Height)` |
| **Priority** | P1 |

#### F-SHAPE-003: Add Image

| Field | Value |
|-------|-------|
| **Feature ID** | F-SHAPE-003 |
| **Name** | Add image |
| **Description** | Insert an image from a file path onto a slide at a specified position and size. Support lock aspect ratio. |
| **COM Objects** | `Shapes.AddPicture(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)` |
| **Priority** | P1 |

#### F-SHAPE-004: Add Line

| Field | Value |
|-------|-------|
| **Feature ID** | F-SHAPE-004 |
| **Name** | Add line |
| **Description** | Add a line between two points on a slide. |
| **COM Objects** | `Shapes.AddLine(BeginX, BeginY, EndX, EndY)` |
| **Priority** | P1 |

#### F-SHAPE-005: Delete Shape

| Field | Value |
|-------|-------|
| **Feature ID** | F-SHAPE-005 |
| **Name** | Delete shape |
| **Description** | Delete a shape by name or index from a slide. |
| **COM Objects** | `Shape.Delete()` |
| **Priority** | P1 |

#### F-SHAPE-006: Set Shape Position/Size

| Field | Value |
|-------|-------|
| **Feature ID** | F-SHAPE-006 |
| **Name** | Set shape position and size |
| **Description** | Set or adjust a shape's Left, Top, Width, Height, and Rotation properties. Accept points, inches, or centimeters. |
| **COM Objects** | `Shape.Left`, `.Top`, `.Width`, `.Height`, `.Rotation` |
| **Priority** | P1 |

#### F-SHAPE-007: List Shapes

| Field | Value |
|-------|-------|
| **Feature ID** | F-SHAPE-007 |
| **Name** | List shapes on slide |
| **Description** | Return all shapes on a slide with name, type, position, size, and whether they have text. |
| **COM Objects** | `Slide.Shapes`, `Shape.Name`, `.Type`, `.Left`, `.Top`, `.Width`, `.Height`, `.HasTextFrame` |
| **Priority** | P1 |

#### F-SHAPE-008: Get Shape Info

| Field | Value |
|-------|-------|
| **Feature ID** | F-SHAPE-008 |
| **Name** | Get shape details |
| **Description** | Return detailed properties of a shape: type, position, size, rotation, visibility, z-order, text content (if any), fill info, line info. |
| **COM Objects** | `Shape.*` (multiple properties) |
| **Priority** | P1 |

#### F-SHAPE-009: Duplicate Shape

| Field | Value |
|-------|-------|
| **Feature ID** | F-SHAPE-009 |
| **Name** | Duplicate shape |
| **Description** | Duplicate a shape on the same slide. Optionally offset the duplicate position. |
| **COM Objects** | `Shape.Duplicate()` |
| **Priority** | P1 |

#### F-SHAPE-010: Set Shape Z-Order

| Field | Value |
|-------|-------|
| **Feature ID** | F-SHAPE-010 |
| **Name** | Set shape z-order |
| **Description** | Change a shape's stacking order: bring to front, send to back, bring forward, send backward. |
| **COM Objects** | `Shape.ZOrder(MsoZOrderCmd)` |
| **Priority** | P1 |

---

#### F-TEXT-001: Set Text

| Field | Value |
|-------|-------|
| **Feature ID** | F-TEXT-001 |
| **Name** | Set shape text |
| **Description** | Set or replace the full text content of a shape's text frame. Use `\r` for paragraph breaks. |
| **COM Objects** | `Shape.TextFrame.TextRange.Text` |
| **Priority** | P1 |

#### F-TEXT-002: Get Text

| Field | Value |
|-------|-------|
| **Feature ID** | F-TEXT-002 |
| **Name** | Get shape text |
| **Description** | Get the full text content of a shape's text frame, including paragraph count and run information. |
| **COM Objects** | `Shape.TextFrame.TextRange.Text`, `.Paragraphs()`, `.Runs()` |
| **Priority** | P1 |

#### F-TEXT-003: Set Font Format

| Field | Value |
|-------|-------|
| **Feature ID** | F-TEXT-003 |
| **Name** | Set font formatting |
| **Description** | Set font properties for an entire shape or a character range: name, size, bold, italic, underline, color (RGB or theme), shadow, subscript, superscript. |
| **COM Objects** | `TextRange.Font.Name`, `.Size`, `.Bold`, `.Italic`, `.Underline`, `.Color.RGB`, `.Color.ObjectThemeColor` |
| **Priority** | P1 |

#### F-TEXT-004: Set Partial Text Format

| Field | Value |
|-------|-------|
| **Feature ID** | F-TEXT-004 |
| **Name** | Format partial text (character range) |
| **Description** | Apply font/color formatting to a subset of text using `Characters(Start, Length)`. This is the key differentiator over static file generation. |
| **COM Objects** | `TextRange.Characters(Start, Length).Font.*` |
| **Priority** | P1 |

#### F-TEXT-005: Insert Text

| Field | Value |
|-------|-------|
| **Feature ID** | F-TEXT-005 |
| **Name** | Insert text before/after |
| **Description** | Insert text before or after existing text in a text frame. The inserted range can be independently formatted. |
| **COM Objects** | `TextRange.InsertBefore(NewText)`, `.InsertAfter(NewText)` |
| **Priority** | P1 |

#### F-TEXT-006: Find and Replace Text

| Field | Value |
|-------|-------|
| **Feature ID** | F-TEXT-006 |
| **Name** | Find and replace text |
| **Description** | Search for text within a shape and optionally replace it. Support case-sensitive and whole-word matching. |
| **COM Objects** | `TextRange.Find(FindWhat, After, MatchCase, WholeWords)`, `.Replace(FindWhat, ReplaceWhat)` |
| **Priority** | P1 |

---

#### F-PH-001: List Placeholders

| Field | Value |
|-------|-------|
| **Feature ID** | F-PH-001 |
| **Name** | List placeholders on slide |
| **Description** | Return all placeholders on a slide with their index, type (PpPlaceholderType), name, position, size, and contained type. |
| **COM Objects** | `Slide.Shapes.Placeholders`, `PlaceholderFormat.Type`, `.ContainedType` |
| **Priority** | P1 |

#### F-PH-002: Set Placeholder Text

| Field | Value |
|-------|-------|
| **Feature ID** | F-PH-002 |
| **Name** | Set placeholder text |
| **Description** | Set text in a placeholder by index or type. Support title, subtitle, body, and other text-capable placeholders. |
| **COM Objects** | `Placeholders(index).TextFrame.TextRange.Text` |
| **Priority** | P1 |

#### F-PH-003: Get Placeholder Info

| Field | Value |
|-------|-------|
| **Feature ID** | F-PH-003 |
| **Name** | Get placeholder details |
| **Description** | Return detailed information about a specific placeholder: type, contained type, text content, position, size, and formatting. |
| **COM Objects** | `PlaceholderFormat.*`, `Shape.TextFrame`, `.Left`, `.Top`, `.Width`, `.Height` |
| **Priority** | P1 |

#### F-PH-004: List Layouts

| Field | Value |
|-------|-------|
| **Feature ID** | F-PH-004 |
| **Name** | List available layouts |
| **Description** | Return all available CustomLayouts from the slide master, with name, index, and placeholder type summary. |
| **COM Objects** | `SlideMaster.CustomLayouts`, `CustomLayout.Name`, `.Shapes.Placeholders` |
| **Priority** | P1 |

---

### Priority 2 -- Important

These features extend the MVP with richer formatting, table support, export capabilities, and slide show control.

---

#### F-TABLE-001: Create Table

| Field | Value |
|-------|-------|
| **Feature ID** | F-TABLE-001 |
| **Name** | Create table |
| **Description** | Add a table with specified rows, columns, position, and size to a slide. |
| **COM Objects** | `Shapes.AddTable(NumRows, NumColumns, Left, Top, Width, Height)` |
| **Priority** | P2 |

#### F-TABLE-002: Set Cell Text

| Field | Value |
|-------|-------|
| **Feature ID** | F-TABLE-002 |
| **Name** | Set table cell text |
| **Description** | Set text content and basic formatting (font, color, alignment) of a specific table cell. |
| **COM Objects** | `Table.Cell(Row, Col).Shape.TextFrame.TextRange` |
| **Priority** | P2 |

#### F-TABLE-003: Get Table Data

| Field | Value |
|-------|-------|
| **Feature ID** | F-TABLE-003 |
| **Name** | Get table data |
| **Description** | Read all cell text from a table as a 2D array. Include row/column count. |
| **COM Objects** | `Table.Cell(Row, Col).Shape.TextFrame.TextRange.Text`, `Table.Rows.Count`, `Table.Columns.Count` |
| **Priority** | P2 |

#### F-TABLE-004: Format Table Cell

| Field | Value |
|-------|-------|
| **Feature ID** | F-TABLE-004 |
| **Name** | Format table cell |
| **Description** | Set cell background color, border styles, and text alignment. |
| **COM Objects** | `Cell.Shape.Fill`, `Cell.Borders()`, `Cell.Shape.TextFrame.TextRange.ParagraphFormat` |
| **Priority** | P2 |

#### F-TABLE-005: Table Row/Column Operations

| Field | Value |
|-------|-------|
| **Feature ID** | F-TABLE-005 |
| **Name** | Add/delete rows and columns |
| **Description** | Add or delete rows and columns. Set row height and column width. |
| **COM Objects** | `Table.Rows.Add()`, `Table.Columns.Add()`, `Row.Delete()`, `Column.Delete()`, `Row.Height`, `Column.Width` |
| **Priority** | P2 |

#### F-TABLE-006: Merge/Split Cells

| Field | Value |
|-------|-------|
| **Feature ID** | F-TABLE-006 |
| **Name** | Merge and split table cells |
| **Description** | Merge a range of cells or split a cell into multiple rows/columns. |
| **COM Objects** | `Cell.Merge(MergeTo)`, `Cell.Split(NumRows, NumColumns)` |
| **Priority** | P2 |

---

#### F-FILL-001: Set Shape Fill

| Field | Value |
|-------|-------|
| **Feature ID** | F-FILL-001 |
| **Name** | Set shape fill |
| **Description** | Set a shape's fill: solid color (RGB or theme), gradient (1-color, 2-color, preset), pattern, texture, or picture. Support transparency. |
| **COM Objects** | `Shape.Fill.Solid()`, `.ForeColor.RGB`, `.TwoColorGradient()`, `.PresetGradient()`, `.Patterned()`, `.UserPicture()`, `.Transparency` |
| **Priority** | P2 |

#### F-LINE-001: Set Shape Line

| Field | Value |
|-------|-------|
| **Feature ID** | F-LINE-001 |
| **Name** | Set shape line/border |
| **Description** | Set a shape's line/border: color, weight, dash style, line style (single/double/triple), arrow heads (style, length, width), visibility. |
| **COM Objects** | `Shape.Line.ForeColor.RGB`, `.Weight`, `.DashStyle`, `.Style`, `.BeginArrowheadStyle`, `.EndArrowheadStyle`, `.Visible` |
| **Priority** | P2 |

#### F-PARA-001: Set Paragraph Format

| Field | Value |
|-------|-------|
| **Feature ID** | F-PARA-001 |
| **Name** | Set paragraph formatting |
| **Description** | Set paragraph properties: alignment (left, center, right, justify, distribute), line spacing, space before/after, indent level, text direction. |
| **COM Objects** | `ParagraphFormat.Alignment`, `.SpaceWithin`, `.SpaceBefore`, `.SpaceAfter`, `.LineRuleWithin`, `TextRange.IndentLevel` |
| **Priority** | P2 |

#### F-PARA-002: Set Bullet Format

| Field | Value |
|-------|-------|
| **Feature ID** | F-PARA-002 |
| **Name** | Set bullet/numbering |
| **Description** | Configure bullet or numbered list formatting: type (none, symbol, numbered, picture), character, font, relative size, start value, numbering style. |
| **COM Objects** | `ParagraphFormat.Bullet.Type`, `.Character`, `.Font`, `.RelativeSize`, `.StartValue`, `.Style` |
| **Priority** | P2 |

---

#### F-BG-001: Set Slide Background

| Field | Value |
|-------|-------|
| **Feature ID** | F-BG-001 |
| **Name** | Set slide background |
| **Description** | Set a slide's background: solid color, gradient, pattern, texture, or image. Control `FollowMasterBackground` to override or inherit master. |
| **COM Objects** | `Slide.FollowMasterBackground`, `Slide.Background.Fill.*` |
| **Priority** | P2 |

#### F-EFFECT-001: Set Shape Shadow

| Field | Value |
|-------|-------|
| **Feature ID** | F-EFFECT-001 |
| **Name** | Set shape shadow |
| **Description** | Apply shadow effect to a shape: color, offset X/Y, blur, transparency, size, style (inner/outer). |
| **COM Objects** | `Shape.Shadow.Visible`, `.ForeColor`, `.OffsetX`, `.OffsetY`, `.Blur`, `.Transparency`, `.Style` |
| **Priority** | P2 |

#### F-EFFECT-002: Set Shape Glow

| Field | Value |
|-------|-------|
| **Feature ID** | F-EFFECT-002 |
| **Name** | Set shape glow |
| **Description** | Apply glow effect to a shape: color, radius, transparency. |
| **COM Objects** | `Shape.Glow.Color`, `.Radius`, `.Transparency` |
| **Priority** | P2 |

#### F-EFFECT-003: Set Shape Reflection

| Field | Value |
|-------|-------|
| **Feature ID** | F-EFFECT-003 |
| **Name** | Set shape reflection |
| **Description** | Apply reflection effect: preset type or custom blur/offset/size/transparency. |
| **COM Objects** | `Shape.Reflection.Type`, `.Blur`, `.Offset`, `.Size`, `.Transparency` |
| **Priority** | P2 |

#### F-EFFECT-004: Set Shape Soft Edge

| Field | Value |
|-------|-------|
| **Feature ID** | F-EFFECT-004 |
| **Name** | Set shape soft edge |
| **Description** | Apply soft edge (feather) effect: preset type or custom radius. |
| **COM Objects** | `Shape.SoftEdge.Type`, `.Radius` |
| **Priority** | P2 |

---

#### F-EXPORT-001: Export to PDF

| Field | Value |
|-------|-------|
| **Feature ID** | F-EXPORT-001 |
| **Name** | Export to PDF |
| **Description** | Export the presentation as PDF. Support quality (screen/print), slide range, and options (hidden slides, document properties). |
| **COM Objects** | `Presentation.ExportAsFixedFormat()`, `Presentation.SaveAs(FileFormat=ppSaveAsPDF)` |
| **Priority** | P2 |

#### F-EXPORT-002: Export Slides as Images

| Field | Value |
|-------|-------|
| **Feature ID** | F-EXPORT-002 |
| **Name** | Export slides as images |
| **Description** | Export individual slides or all slides as PNG/JPG/BMP/GIF/TIFF images. Support custom resolution (width/height in pixels). |
| **COM Objects** | `Slide.Export(Path, FilterName, ScaleWidth, ScaleHeight)`, `Presentation.SaveAs(FileFormat=ppSaveAsPNG)` |
| **Priority** | P2 |

---

#### F-SS-001: Start Slide Show

| Field | Value |
|-------|-------|
| **Feature ID** | F-SS-001 |
| **Name** | Start slide show |
| **Description** | Start a slide show with configurable settings: range (all, specific slides), show type (speaker, window, kiosk), loop, advance mode. |
| **COM Objects** | `SlideShowSettings.*`, `.Run()` |
| **Priority** | P2 |

#### F-SS-002: Control Slide Show

| Field | Value |
|-------|-------|
| **Feature ID** | F-SS-002 |
| **Name** | Control slide show navigation |
| **Description** | Navigate during a running slide show: next, previous, first, last, go to specific slide. Set state (running, paused, black screen, white screen). |
| **COM Objects** | `SlideShowView.Next()`, `.Previous()`, `.First()`, `.Last()`, `.GotoSlide()`, `.State`, `.Exit()` |
| **Priority** | P2 |

#### F-SS-003: Get Slide Show Status

| Field | Value |
|-------|-------|
| **Feature ID** | F-SS-003 |
| **Name** | Get slide show status |
| **Description** | Return current slide show state: running/paused/done, current slide position, elapsed time, pointer type. |
| **COM Objects** | `SlideShowView.State`, `.CurrentShowPosition`, `.Slide`, `.PresentationElapsedTime` |
| **Priority** | P2 |

---

#### F-NOTES-001: Set Slide Notes

| Field | Value |
|-------|-------|
| **Feature ID** | F-NOTES-001 |
| **Name** | Set slide notes |
| **Description** | Set or append text to a slide's speaker notes. |
| **COM Objects** | `Slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text` |
| **Priority** | P2 |

#### F-NOTES-002: Get Slide Notes

| Field | Value |
|-------|-------|
| **Feature ID** | F-NOTES-002 |
| **Name** | Get slide notes |
| **Description** | Read the speaker notes text from a slide. |
| **COM Objects** | `Slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text` |
| **Priority** | P2 |

---

### Priority 3 -- Nice to Have

Advanced features that expand the server's capabilities to cover the full PowerPoint feature set.

---

#### F-CHART-001: Create Chart

| Field | Value |
|-------|-------|
| **Feature ID** | F-CHART-001 |
| **Name** | Create chart |
| **Description** | Add a chart to a slide with specified type (column, bar, line, pie, scatter, etc.), position, and size. |
| **COM Objects** | `Shapes.AddChart2(Style, Type, Left, Top, Width, Height)` |
| **Priority** | P3 |

#### F-CHART-002: Set Chart Data

| Field | Value |
|-------|-------|
| **Feature ID** | F-CHART-002 |
| **Name** | Set chart data |
| **Description** | Populate a chart with data via the embedded Excel Workbook. Set categories, series names, and values. |
| **COM Objects** | `Chart.ChartData.Activate()`, `.Workbook.Worksheets(1)`, `Chart.SetSourceData()` |
| **Priority** | P3 |

#### F-CHART-003: Format Chart

| Field | Value |
|-------|-------|
| **Feature ID** | F-CHART-003 |
| **Name** | Format chart elements |
| **Description** | Set chart title, legend, axes (titles, scale, gridlines), data labels, series colors, and chart style. |
| **COM Objects** | `Chart.ChartTitle`, `.HasLegend`, `.Legend`, `.Axes()`, `.SeriesCollection()`, `.ChartStyle` |
| **Priority** | P3 |

---

#### F-SMART-001: Create SmartArt

| Field | Value |
|-------|-------|
| **Feature ID** | F-SMART-001 |
| **Name** | Create SmartArt |
| **Description** | Add a SmartArt graphic to a slide with a specified layout (list, process, cycle, hierarchy, etc.). |
| **COM Objects** | `Shapes.AddSmartArt(Layout, Left, Top, Width, Height)`, `Application.SmartArtLayouts` |
| **Priority** | P3 |

#### F-SMART-002: Modify SmartArt

| Field | Value |
|-------|-------|
| **Feature ID** | F-SMART-002 |
| **Name** | Modify SmartArt nodes |
| **Description** | Add, remove, and set text on SmartArt nodes. Change layout and color scheme. |
| **COM Objects** | `SmartArt.AllNodes`, `.Nodes`, `SmartArtNode.TextFrame2`, `.Delete()`, `SmartArt.Layout`, `.Color` |
| **Priority** | P3 |

---

#### F-ANIM-001: Set Transition

| Field | Value |
|-------|-------|
| **Feature ID** | F-ANIM-001 |
| **Name** | Set slide transition |
| **Description** | Set a slide's transition effect, speed/duration, advance settings (on click, after time), and sound. |
| **COM Objects** | `SlideShowTransition.EntryEffect`, `.Duration`, `.AdvanceOnClick`, `.AdvanceOnTime`, `.AdvanceTime` |
| **Priority** | P3 |

#### F-ANIM-002: Add Animation

| Field | Value |
|-------|-------|
| **Feature ID** | F-ANIM-002 |
| **Name** | Add shape animation |
| **Description** | Add animation effects to shapes using the Timeline API: appear, fade, fly in, etc. Set trigger (on click, with previous, after previous) and timing. |
| **COM Objects** | `Slide.TimeLine.MainSequence.AddEffect(Shape, effectId, trigger)` |
| **Priority** | P3 |

---

#### F-MEDIA-001: Add Video

| Field | Value |
|-------|-------|
| **Feature ID** | F-MEDIA-001 |
| **Name** | Insert video |
| **Description** | Add a video file to a slide. Support embedded and linked modes. Set position and size. |
| **COM Objects** | `Shapes.AddMediaObject2(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)` |
| **Priority** | P3 |

#### F-MEDIA-002: Add Audio

| Field | Value |
|-------|-------|
| **Feature ID** | F-MEDIA-002 |
| **Name** | Insert audio |
| **Description** | Add an audio file to a slide. Support embedded and linked modes. |
| **COM Objects** | `Shapes.AddMediaObject2(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)` |
| **Priority** | P3 |

#### F-MEDIA-003: Media Settings

| Field | Value |
|-------|-------|
| **Feature ID** | F-MEDIA-003 |
| **Name** | Configure media playback |
| **Description** | Set media volume, mute, trim (start/end), fade in/out, and play settings (auto-play, loop, hide when not playing). |
| **COM Objects** | `Shape.MediaFormat.Volume`, `.Muted`, `.StartPoint`, `.EndPoint`, `.FadeInDuration`, `.FadeOutDuration`, `AnimationSettings.PlaySettings.*` |
| **Priority** | P3 |

---

#### F-THEME-001: Apply Theme

| Field | Value |
|-------|-------|
| **Feature ID** | F-THEME-001 |
| **Name** | Apply theme |
| **Description** | Apply a theme (.thmx) or template (.potx) file to a presentation or slide master. |
| **COM Objects** | `Presentation.ApplyTheme()`, `.ApplyTemplate()`, `Master.ApplyTheme()` |
| **Priority** | P3 |

#### F-THEME-002: Get Theme Info

| Field | Value |
|-------|-------|
| **Feature ID** | F-THEME-002 |
| **Name** | Get theme color scheme |
| **Description** | Return the current theme color scheme (12 theme colors with their RGB values). |
| **COM Objects** | `Master.Theme.ThemeColorScheme` |
| **Priority** | P3 |

#### F-THEME-003: Master/Layout Management

| Field | Value |
|-------|-------|
| **Feature ID** | F-THEME-003 |
| **Name** | Manage slide masters and layouts |
| **Description** | List, add, delete, and modify slide masters (via Designs collection). List and manage custom layouts. |
| **COM Objects** | `Presentation.Designs`, `Design.SlideMaster`, `Master.CustomLayouts` |
| **Priority** | P3 |

---

#### F-LINK-001: Add Hyperlink

| Field | Value |
|-------|-------|
| **Feature ID** | F-LINK-001 |
| **Name** | Add hyperlink |
| **Description** | Add a hyperlink to a shape or text range: URL, file, email, or slide reference. Set action type (click, mouse-over). |
| **COM Objects** | `Shape.ActionSettings(ppMouseClick).Action`, `.Hyperlink.Address`, `.Hyperlink.SubAddress` |
| **Priority** | P3 |

#### F-LINK-002: Get Hyperlinks

| Field | Value |
|-------|-------|
| **Feature ID** | F-LINK-002 |
| **Name** | List hyperlinks |
| **Description** | Return all hyperlinks on a slide with their address, sub-address, and associated shape. |
| **COM Objects** | `Slide.Hyperlinks` |
| **Priority** | P3 |

---

#### F-OLE-001: Add OLE Object

| Field | Value |
|-------|-------|
| **Feature ID** | F-OLE-001 |
| **Name** | Insert OLE object |
| **Description** | Embed or link an OLE object (Excel worksheet, Word document, etc.) on a slide. |
| **COM Objects** | `Shapes.AddOLEObject(ClassName, FileName, Link, DisplayAsIcon, ...)` |
| **Priority** | P3 |

#### F-CONN-001: Add Connector

| Field | Value |
|-------|-------|
| **Feature ID** | F-CONN-001 |
| **Name** | Add connector shape |
| **Description** | Add a connector (straight, elbow, curve) between two shapes. |
| **COM Objects** | `Shapes.AddConnector(Type, ...)`, `ConnectorFormat.BeginConnect()`, `.EndConnect()` |
| **Priority** | P3 |

#### F-GROUP-001: Group/Ungroup Shapes

| Field | Value |
|-------|-------|
| **Feature ID** | F-GROUP-001 |
| **Name** | Group and ungroup shapes |
| **Description** | Group multiple shapes together or ungroup an existing group. Access group items. |
| **COM Objects** | `ShapeRange.Group()`, `Shape.Ungroup()`, `Shape.GroupItems` |
| **Priority** | P3 |

#### F-PRINT-001: Print Presentation

| Field | Value |
|-------|-------|
| **Feature ID** | F-PRINT-001 |
| **Name** | Print presentation |
| **Description** | Print the presentation with configurable options: range, copies, collation, output type, color mode, printer. |
| **COM Objects** | `Presentation.PrintOut()`, `Presentation.PrintOptions.*` |
| **Priority** | P3 |

#### F-SECTION-001: Manage Sections

| Field | Value |
|-------|-------|
| **Feature ID** | F-SECTION-001 |
| **Name** | Manage presentation sections |
| **Description** | Add, rename, move, and delete sections. List sections with their slide counts. |
| **COM Objects** | `Presentation.SectionProperties.*` |
| **Priority** | P3 |

#### F-PROP-001: Set Document Properties

| Field | Value |
|-------|-------|
| **Feature ID** | F-PROP-001 |
| **Name** | Set document properties |
| **Description** | Set built-in document properties: title, author, subject, keywords, comments, category, company. |
| **COM Objects** | `Presentation.BuiltInDocumentProperties("PropertyName").Value` |
| **Priority** | P3 |

#### F-3D-001: Set 3D Format

| Field | Value |
|-------|-------|
| **Feature ID** | F-3D-001 |
| **Name** | Set 3D effects |
| **Description** | Apply 3D effects to shapes: bevel (type, depth, inset), extrusion (depth, color), rotation (X/Y/Z), lighting, material. |
| **COM Objects** | `Shape.ThreeD.BevelTopType`, `.Depth`, `.RotationX`, `.RotationY`, `.PresetMaterial`, `.PresetLightingDirection` |
| **Priority** | P3 |

#### F-HF-001: Set Headers/Footers

| Field | Value |
|-------|-------|
| **Feature ID** | F-HF-001 |
| **Name** | Set headers and footers |
| **Description** | Configure slide header/footer/date/slide-number visibility and text at master or slide level. |
| **COM Objects** | `HeadersFooters.Footer.Visible`, `.Footer.Text`, `.DateAndTime.*`, `.SlideNumber.Visible` |
| **Priority** | P3 |

---

## 3. Non-Functional Requirements

### 3.1 Error Handling Strategy

| Requirement | Description |
|------------|-------------|
| **COM Error Wrapping** | All `pywintypes.com_error` exceptions must be caught and translated into human-readable MCP error responses. Extract `hresult`, `strerror`, and `excepinfo` fields. |
| **Graceful Degradation** | If a COM operation fails (e.g., PowerPoint crashed), the server should attempt to reconnect or return a clear "connection lost" error. |
| **Input Validation** | All tool inputs must be validated before COM calls. Invalid slide indices, negative dimensions, and out-of-range values should produce clear error messages. |
| **Idempotency** | Where possible, operations should be idempotent or at least safe to retry. |

### 3.2 COM Threading Model

| Requirement | Description |
|------------|-------------|
| **Apartment Threading** | COM operations must run in a Single-Threaded Apartment (STA). Use `pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)` for any non-main-thread COM access. |
| **Single-Thread Access** | All COM calls must be serialized on a single thread. MCP tool invocations from different requests must be queued. |
| **Lifecycle Management** | COM references must be explicitly released when no longer needed. Use `None` assignment and `gc.collect()` for cleanup. |
| **CoInitialize/CoUninitialize** | Ensure proper `CoInitialize` on thread start and `CoUninitialize` on thread end for any worker threads. |

### 3.3 Performance Considerations

| Requirement | Description |
|------------|-------------|
| **Lazy Connection** | Do not connect to PowerPoint until the first tool call. Cache the connection. |
| **Batch Operations** | Where the COM model allows, batch multiple changes before forcing a UI refresh. |
| **Timeout Protection** | Long-running operations (PDF export, video export) should have configurable timeouts. |
| **Minimal Round-Trips** | Gather multiple properties in a single tool call where practical to reduce MCP protocol overhead. |

### 3.4 Logging

| Requirement | Description |
|------------|-------------|
| **Structured Logging** | Use Python `logging` module with structured log messages. |
| **Log Levels** | DEBUG: COM call details. INFO: Tool invocations and results. WARNING: Recoverable issues. ERROR: Failed operations. |
| **COM Error Logging** | Log full COM error details (HRESULT, source, description) at ERROR level. |

### 3.5 Unit Conversion

The server must support transparent unit conversion for all position/size parameters:

| Unit | Abbreviation | Conversion to Points |
|------|-------------|---------------------|
| Points | `pt` | 1 pt = 1 pt (native) |
| Inches | `in` | 1 in = 72 pt |
| Centimeters | `cm` | 1 cm = 28.3465 pt |
| EMU (English Metric Unit) | `emu` | 1 pt = 12700 emu |

Tool parameters that accept position/size values should accept a numeric value with an optional unit suffix. Default unit is points.

---

## 4. Constraints and Risks

### 4.1 Platform Constraint: Windows Only

The server depends on the Windows COM infrastructure (`win32com`, `pythoncom`). It cannot run on macOS or Linux. Users must have a Windows machine with PowerPoint installed.

### 4.2 PowerPoint Must Be Installed

Microsoft PowerPoint (desktop version) must be installed on the machine. The web version of PowerPoint does not expose a COM interface.

### 4.3 Single COM Instance Limitation

PowerPoint can only run as a single instance (unlike Word or Excel). Multiple `Dispatch("PowerPoint.Application")` calls return the same instance. The MCP server must account for this by treating the application connection as a shared singleton.

### 4.4 COM Object Lifecycle Management

COM objects are reference-counted. Failure to release references can lead to:
- Memory leaks
- PowerPoint processes that refuse to close
- Stale references that cause `com_error` exceptions

Mitigation: The server must implement careful COM object lifecycle management with explicit `None` assignment and `gc.collect()` in cleanup routines.

### 4.5 PowerPoint UI Interactions

Some COM operations trigger UI dialogs (save confirmation, password prompts). The server should:
- Set `Presentation.Saved = True` before closing without saving
- Avoid operations that require user interaction in automated mode
- Document any operations that may trigger dialogs

### 4.6 File Path Handling

All file paths passed to COM methods must be absolute Windows paths (e.g., `C:\Users\...`). The server should validate and normalize paths before passing them to COM.

### 4.7 BGR Color Format

PowerPoint COM uses BGR (Blue-Green-Red) color encoding, not standard RGB. The server must provide a transparent conversion layer so users can specify colors in standard `#RRGGBB` hex format or `(R, G, B)` tuples.

### 4.8 Version Compatibility

Different Office versions may have slightly different COM APIs:
- `AddChart2` is Office 2013+
- `AddMediaObject2` replaces deprecated `AddMediaObject`
- SmartArt layout indices vary by version

The server should detect the PowerPoint version and handle API differences gracefully.

---

## Appendix A: Glossary

| Term | Definition |
|------|-----------|
| **MCP** | Model Context Protocol -- a standard for LLM tool interaction |
| **FastMCP** | A Python framework for building MCP servers |
| **COM** | Component Object Model -- Windows inter-process communication technology |
| **STA** | Single-Threaded Apartment -- COM threading model |
| **EMU** | English Metric Unit -- PowerPoint's internal unit (914400 EMU = 1 inch) |
| **BGR** | Blue-Green-Red -- PowerPoint's color encoding (reverse of standard RGB) |
| **PpSlideLayout** | PowerPoint enumeration for standard slide layouts |
| **MsoAutoShapeType** | Office enumeration for auto shape types (180+ shapes) |
| **PlaceholderFormat** | COM object describing a placeholder's type and contained content |
