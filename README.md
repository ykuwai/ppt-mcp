<p align="center">
  <img src="https://raw.githubusercontent.com/ykuwai/ppt-mcp/main/assets/ppt-mcp-logo-letter.png" alt="PowerPoint MCP" width="480">
</p>

<p align="center">
  <a href="README_ja.md">日本語版はこちら</a>
</p>

<p align="center">
  <a href="https://www.python.org/"><img src="https://img.shields.io/badge/Python-3.10%2B-blue.svg" alt="Python"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/License-MIT-green.svg" alt="License"></a>
  <img src="https://img.shields.io/badge/Platform-Windows-0078d4.svg" alt="Platform">
  <a href="https://pepy.tech/projects/ppt-mcp"><img src="https://static.pepy.tech/personalized-badge/ppt-mcp?period=total&units=ABBREVIATION&left_color=BLACK&right_color=GREEN&left_text=downloads" alt="Downloads"></a>
</p>

<p align="center">
  <strong>Real-time PowerPoint control through COM automation —<br>an MCP server with 147 tools for AI agents and developers.</strong>
</p>

---

An MCP (Model Context Protocol) server that gives AI agents full control over a live Microsoft PowerPoint instance via COM automation. Unlike file-based libraries like python-pptx, this server interacts directly with a running PowerPoint application.

## ✨ Key Features

- **Real-time control** — Directly manipulates a running PowerPoint instance; changes appear instantly on screen
- **147 tools across 26 categories** — Slides, shapes, text, tables, charts, animations, SmartArt, media, freeform paths, and more
- **Safe for AI agents** — `ppt_activate_presentation` locks all tools to a specific file, preventing accidental edits to the wrong presentation
- **[Google Material Symbols](https://fonts.google.com/icons) icons** — Search 2,500+ icons by keyword and insert as SVG with theme colors
- **Theme color awareness** — Use `accent1`, `accent2`, etc. instead of hardcoded RGB values

## 📋 Requirements

- Windows 11
- Microsoft PowerPoint
- [uv](https://docs.astral.sh/uv/getting-started/installation/)

## 🚀 Getting Started

Standard config — works in Claude Desktop, Cursor, `.mcp.json`, and most other MCP clients:

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "uvx",
      "args": ["ppt-mcp"]
    }
  }
}
```

### Claude Code

```bash
# User-scoped (available in all projects)
claude mcp add powerpoint uvx ppt-mcp

# Project-scoped (stored in .mcp.json, shared with your team)
claude mcp add --scope project powerpoint uvx ppt-mcp
```

### VS Code

```bash
code --add-mcp '{"name":"powerpoint","command":"uvx","args":["ppt-mcp"]}'
```

### Cursor

[![Install in Cursor](https://cursor.com/deeplink/mcp-install-dark.svg)](https://cursor.com/install-mcp?name=ppt-mcp&config=eyJ0eXBlIjoic3RkaW8iLCJjb21tYW5kIjoidXZ4IiwiYXJncyI6WyJwcHQtbWNwIl19)

Or add manually to `~/.cursor/mcp.json` using the standard config above.

### Claude Desktop

Edit `%APPDATA%\Claude\claude_desktop_config.json` using the standard config above.

### Codex

Edit `~/.codex/config.toml`:

```toml
[mcp_servers.ppt-mcp]
command = "uvx"
args = ["ppt-mcp"]
```

Or use the standard JSON config above in `.codex/config.json`.

### From source

```bash
git clone https://github.com/ykuwai/ppt-mcp.git
cd ppt-mcp
uv sync
```

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "uv",
      "args": [
        "--directory",
        "C:\\path\\to\\ppt-mcp",
        "run",
        "mcp",
        "run",
        "src/server.py"
      ]
    }
  }
}
```

## 🛠️ Tool Categories

| Category | Tools | Description |
|----------|------:|-------------|
| **App** | 5 | Connect to PowerPoint, app info, active window, window state, list presentations |
| **Presentation** | 8 | Create (with templates), open, save, close, info, activate target, list templates |
| **Slides** | 9 | Add, delete, duplicate, move, list, info, notes, navigation |
| **Shapes** | 10 | Add shapes/textboxes/pictures/lines, list, info, update, delete, z-order |
| **Text** | 8 | Set/get text, format text ranges, paragraph format, bullets, find/replace, textframe |
| **Placeholders** | 6 | List, get, set placeholder content |
| **Formatting** | 3 | Fill, line, shadow |
| **Tables** | 12 | Add tables, get/set cells, merge/split cells, add/delete rows/columns, styles, layout, borders |
| **Export** | 3 | PDF, images, slide preview |
| **Slideshow** | 6 | Start, stop, next, previous, go to slide, status |
| **Charts** | 6 | Add charts, set/get data, format, series, change type |
| **Animation** | 5 | Transitions, add/list/remove/clear animations |
| **Themes** | 3 | Apply themes, get theme colors, headers/footers |
| **Groups** | 3 | Group, ungroup, get group items |
| **Connectors** | 2 | Add, format |
| **Hyperlinks** | 3 | Add, get, remove |
| **Sections** | 3 | Add, list, manage |
| **Properties** | 2 | Set/get presentation metadata |
| **Media** | 3 | Video, audio, media settings |
| **SmartArt** | 3 | Add, modify, list layouts |
| **Edit Operations** | 6 | Undo, redo, copy shapes/formatting between slides |
| **Layout** | 7 | Align, distribute, slide size, background, flip, merge shapes |
| **Effects** | 3 | Glow, reflection, soft edge |
| **Comments** | 3 | Add, list, delete |
| **Advanced** | 18 | Tags, fonts (set defaults + bulk replace), crop, shape export, visibility, selection, view, animation copy, picture from URL, SVG icons, icon search, aspect ratio lock, batch apply, default shape style |
| **Freeform** | 7 | Build freeform paths, get/set node positions, insert/delete nodes, node editing type, segment type |
| | **147** | |

## 💡 Example Workflow

```python
# 1. Target a specific presentation (prevents editing the wrong file)
ppt_list_presentations()
ppt_activate_presentation(presentation_name="demo.pptx")

# 2. Create a slide from a personal template
ppt_list_templates()
ppt_create_presentation(template_path="C:\\...\\MyTemplate.potx")

# 3. Add a slide and set content
ppt_add_slide(layout_index=2)
ppt_set_text(slide_index=1, shape_name_or_index="Title 1", text="Hello World")

# 4. Set presentation-wide fonts (Latin + East Asian separately)
ppt_set_default_fonts(latin="Segoe UI", east_asian="Meiryo")

# 5. Insert a Google Material Symbols icon with theme color
ppt_add_svg_icon(slide_index=1, icon_name="rocket",
                 left=500, top=100, width=72, height=72,
                 color="accent1", style="rounded", filled=True)

# 6. Export to PDF
ppt_export_pdf(file_path="C:\\output\\presentation.pdf")
```

## 🔍 Features in Detail

### 🎯 Presentation Targeting

`ppt_activate_presentation` sets a session-level target so every subsequent tool call operates on that specific file — regardless of which window is active in PowerPoint. Switch targets anytime by calling it again.

```python
ppt_activate_presentation(presentation_name="report.pptx")
# All tools now operate on report.pptx
ppt_activate_presentation(presentation_name="demo.pptx")
# Switched — all tools now operate on demo.pptx
```

### 📁 Template Support

Auto-detects your personal PowerPoint templates folder (registry, OneDrive, or default paths). Use `ppt_list_templates` to discover available templates, then `ppt_create_presentation(template_path=...)` to create a new presentation from any template.

### 🎨 Google Material Symbols Icons

Search 2,500+ [Google Material Symbols](https://fonts.google.com/icons) icons with `ppt_search_icons(query="...")` and insert them as SVG with `ppt_add_svg_icon`:
- **3 styles**: outlined, rounded, sharp
- **Filled variants**: set `filled=True`
- **Theme colors**: `color="accent1"` uses the presentation's accent color
- **Auto-fit**: preserves aspect ratio within the specified area

### ⚡ Real-Time Navigation

Every write operation automatically navigates PowerPoint to the target slide. You see changes happening in real-time — no need to manually switch slides.

### ✍️ Text Formatting

- `\n` — Paragraph break (Enter). Each paragraph gets its own bullet/indent level.
- `\v` — Line break (Shift+Enter). Stays in the same paragraph, preserving formatting.
- Per-character formatting with `ppt_format_text_range`
- Auto-fit control: shrink text to fit, resize shape, or overflow

## ⚙️ Advanced Configuration

### Handling PowerPoint Modal Dialogs

When PowerPoint has a modal dialog open (e.g., SmartArt layout picker, Save dialog, Insert dialog), COM calls return `RPC_E_CALL_REJECTED`. The MCP server **automatically retries for up to 15 seconds** (5 retries × 3 s), so the server stays connected and responsive even when a dialog is blocking PowerPoint.

**Auto-dismiss (opt-in):** By default, the server waits for you to close the dialog manually. To have the server automatically send ESC on the first retry — dismissing the dialog without user interaction — set `PPT_AUTO_DISMISS_DIALOG=true`:

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "uvx",
      "args": ["ppt-mcp"],
      "env": {
        "PPT_AUTO_DISMISS_DIALOG": "true"
      }
    }
  }
}
```

ESC cancels without committing, so there are no destructive side effects. This is particularly useful in automated workflows where no human is present to close dialogs.

## 📄 License

MIT

## 🙏 Credits

- [FastMCP](https://github.com/jlowin/fastmcp) — Pythonic MCP server framework
- [pywin32](https://github.com/mhammond/pywin32) — Windows COM automation
- [Model Context Protocol](https://modelcontextprotocol.io/) — by Anthropic
