# Tools

ppt-mcp has <!--tools-->157<!--/tools--> tools in <!--categories-->26<!--/categories--> categories. Each tool's own description, which your client shows, says what it takes. On macOS some tools work differently or refuse: [MACOS_PORT.md](MACOS_PORT.md) lists them.

| Category | Tools | Description |
|----------|------:|-------------|
| **App** | 5 | Connect to PowerPoint, app info, active window, window state, list presentations |
| **Presentation** | 8 | Create (with templates), open (a local file, or on Windows a SharePoint or OneDrive URL), save, close, info, activate target, list templates |
| **Slides** | 10 | Add, delete (bulk), duplicate (positional/multi), move (bulk), copy (cross-presentation), list, info, notes, navigation |
| **Shapes** | 10 | Add shapes/textboxes/pictures/lines (with z-order placement, including behind the text), list, info (position, fill, line, text frame), update (one shape, a set, or the whole slide, absolute or by offset), delete, z-order |
| **Text** | 10 | Set text (whole frame, one span keeping its formatting, or as runs), get text (with measurement and overflow), format one span or many in a batch, paragraph format, bullets, find/replace (optionally inside groups), textframe, extract all text as Markdown, typography check |
| **Placeholders** | 6 | List, get, set placeholder content |
| **Formatting** | 3 | Fill, line, shadow (on the shape or on its text) |
| **Tables** | 13 | Add tables, get/set cells, batch set data, merge/split cells, add/delete rows/columns, styles, layout, borders |
| **Export** | 4 | PDF, images, slide preview, clipboard copy |
| **Slideshow** | 6 | Start, stop, next, previous, go to slide, status |
| **Charts** | 7 | Add charts (on Windows, all but a few rarer types such as 3-D area, cone and stock without starting Excel), set/get data (on Windows, setting it swaps in a copy of the chart under a new shape id without starting Excel where it can, then reads the chart back, and data that did not land is an error while a label Excel read differently is a warning; a scatter or bubble chart takes its x values from the categories, and every bubble series needs its `sizes`; a bubble chart that would have to go through Excel, which drops them, is refused), format, format axis, series, change type |
| **Animation** | 6 | Transitions, add/list/update/remove/clear animations (entrance, exit, emphasis, motion path, interactive sequences) |
| **Themes** | 4 | Apply themes, get/set theme colors, headers/footers |
| **Groups** | 3 | Group, ungroup, get group items |
| **Connectors** | 2 | Add, format |
| **Hyperlinks** | 3 | Add, get, remove |
| **Sections** | 3 | Add, list, manage |
| **Properties** | 2 | Set/get presentation metadata |
| **Media** | 3 | Video, audio, media settings |
| **SmartArt** | 3 | Add, modify, list layouts |
| **Edit Operations** | 6 | Undo and redo (on the target presentation), copy shapes/formatting between slides |
| **Layout** | 8 | Align, distribute, slide size, background, flip, merge shapes, layout check |
| **Effects** | 3 | Glow and reflection (on the shape or on its text), soft edge |
| **Comments** | 3 | Add, list, delete |
| **Advanced** | 19 | Tags, fonts (list the deck's or this computer's, set defaults, bulk replace), crop, picture format, shape export, visibility, selection, view, animation copy, picture from URL, SVG icons, icon search, aspect ratio lock, batch apply, default shape style |
| **Freeform** | 7 | Build freeform paths, get/set node positions, insert/delete nodes, node editing type, segment type |
| | **157** | |
