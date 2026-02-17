# PowerPoint Template API via COM Automation (pywin32)

Research document for implementing template support in the ppt-com-mcp server.

---

## 1. Getting the Personal Templates Folder Path via COM

### 1.1 There is NO `Application.Options.DefaultFilePath` in PowerPoint

Unlike Word and Excel, **PowerPoint does not expose** an `Application.Options.DefaultFilePath` property or equivalent to retrieve template folder paths programmatically.

- Word has `Application.Options.DefaultFilePath(wdUserTemplatesPath)` — PowerPoint has nothing similar.
- PowerPoint's object model simply does not include a property to query configured template paths.

### 1.2 Registry Keys

The personal templates path is stored in the Windows registry. The relevant keys:

```
HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\PowerPoint\Options
  Value: PersonalTemplates (REG_SZ)
```

And the shared user templates path (used by all Office apps):

```
HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\General
  Value: UserTemplates (REG_SZ)
```

**Important:** These keys may not exist if the user has never changed the default. When absent, Office uses the default path.

### 1.3 Default Paths (When Registry Keys Are Absent)

| Location | Path |
|---|---|
| **Custom Office Templates** (PowerPoint 2013+) | `%USERPROFILE%\Documents\Custom Office Templates\` |
| **User Templates** (legacy/roaming) | `%APPDATA%\Microsoft\Templates\` |
| **Document Themes** | `%APPDATA%\Microsoft\Templates\Document Themes\` |

For OneDrive-synced documents, the path may differ. For example:
```
C:\Users\kuwam\OneDrive\ドキュメント\Office のカスタム テンプレート
```

### 1.4 Python Code to Resolve the Path

```python
import os
import winreg

def get_personal_templates_path() -> str:
    """Get the PowerPoint personal templates folder path.

    Checks registry first, then falls back to the default location.
    """
    # Try reading from registry
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Office\16.0\PowerPoint\Options"
        )
        path, _ = winreg.QueryValueEx(key, "PersonalTemplates")
        winreg.CloseKey(key)
        if path and os.path.isdir(path):
            return path
    except (FileNotFoundError, OSError):
        pass

    # Fallback: default Custom Office Templates folder
    docs = os.path.join(os.environ["USERPROFILE"], "Documents")
    default = os.path.join(docs, "Custom Office Templates")
    if os.path.isdir(default):
        return default

    return default  # Return the path even if it doesn't exist yet
```

### 1.5 Getting Path via COM (Indirect)

There is no direct COM property, but you can read the registry from Python with `winreg` (as above) or read environment variables. Within a COM context:

```python
def _get_templates_path_impl() -> str:
    """Run on COM thread — but doesn't actually need COM."""
    # Environ("USERPROFILE") equivalent in Python
    return os.path.join(os.environ["USERPROFILE"], "Documents", "Custom Office Templates")
```

---

## 2. Listing Available Templates via COM

### 2.1 No Built-in COM API for Enumerating Templates

PowerPoint's COM API does **not** provide a `Templates` collection or any method to enumerate available template files. There is no `FileDialog` approach that lists templates either.

### 2.2 Solution: Scan the Folder for .potx/.potm Files

The only reliable approach is to scan the templates directory for template files:

```python
import os
import glob

def list_templates(templates_dir: str) -> list[dict]:
    """List all PowerPoint template files in the given directory.

    Searches for .potx (template) and .potm (macro-enabled template) files.
    """
    templates = []
    for ext in ("*.potx", "*.potm", "*.pot"):
        pattern = os.path.join(templates_dir, ext)
        for filepath in glob.glob(pattern):
            stat = os.stat(filepath)
            templates.append({
                "name": os.path.splitext(os.path.basename(filepath))[0],
                "file_name": os.path.basename(filepath),
                "file_path": filepath,
                "size_bytes": stat.st_size,
                "modified": stat.st_mtime,
            })
    return sorted(templates, key=lambda t: t["name"])
```

### 2.3 Also Consider Subdirectories

Some users organize templates in subdirectories. A recursive scan may be useful:

```python
for ext in ("**/*.potx", "**/*.potm"):
    pattern = os.path.join(templates_dir, ext)
    for filepath in glob.glob(pattern, recursive=True):
        ...
```

---

## 3. Creating a New Presentation FROM a Template via COM

There are three distinct approaches, each with different behavior:

### 3.1 Approach A: `Presentations.Open()` with `Untitled=True` (RECOMMENDED)

This is the **best approach** for creating a new presentation from a template. It replicates the behavior of double-clicking a .potx file in Explorer.

**VBA Syntax:**
```vb
Set pres = Application.Presentations.Open( _
    FileName:="C:\path\to\template.potx", _
    ReadOnly:=False, _
    Untitled:=True, _
    WithWindow:=True)
```

**pywin32 Code:**
```python
def _create_from_template_impl(template_path: str) -> dict:
    """Create a new presentation from a template file.

    Uses Presentations.Open with Untitled=msoTrue, which:
    - Opens the template file
    - Creates a new untitled presentation (like double-clicking the .potx)
    - Preserves ALL template content: slides, layouts, themes, custom XML
    - The original template file is NOT modified
    """
    app = ppt._get_app_impl()

    abs_path = os.path.abspath(template_path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"Template not found: {abs_path}")

    # msoFalse=0, msoTrue=-1
    pres = app.Presentations.Open(
        abs_path,    # FileName
        0,           # ReadOnly = msoFalse
        -1,          # Untitled = msoTrue  <-- KEY PARAMETER
        -1,          # WithWindow = msoTrue
    )

    return {
        "success": True,
        "name": pres.Name,
        "slides_count": pres.Slides.Count,
        "template_name": pres.TemplateName,
        "slide_width": pres.PageSetup.SlideWidth,
        "slide_height": pres.PageSetup.SlideHeight,
    }
```

**Key details about `Untitled=msoTrue`:**
- The `Untitled` parameter (3rd positional arg) is the critical setting.
- When `Untitled=msoTrue`, PowerPoint opens the file **without assigning the file name as the title**. This is documented as "equivalent to creating a copy of the file."
- The resulting presentation has no file path — it behaves like a new, unsaved presentation.
- **All template content is preserved**: slide masters, custom layouts, sample slides, themes, fonts, custom XML data, everything.
- The original .potx file is **not modified** and is not locked.

**Why use positional args:** Per project conventions, COM method calls should use positional args (keyword args are unreliable with late binding via pywin32).

### 3.2 Approach B: `Presentations.Add()` + `ApplyTemplate()`

**VBA Syntax:**
```vb
Set pres = Application.Presentations.Add()
pres.ApplyTemplate "C:\path\to\template.potx"
```

**pywin32 Code:**
```python
def _create_with_apply_template_impl(template_path: str) -> dict:
    app = ppt._get_app_impl()
    pres = app.Presentations.Add()
    pres.ApplyTemplate(os.path.abspath(template_path))
    return {
        "success": True,
        "name": pres.Name,
        "slides_count": pres.Slides.Count,
    }
```

**Behavior and limitations:**
- Creates a blank presentation first, then applies the template's **design only** (theme, colors, fonts, slide master, layouts).
- **Does NOT copy sample/content slides** from the template. You only get an empty presentation with the template's design.
- **Does NOT preserve custom XML data** stored in the template.
- The `ApplyTemplate` method essentially copies the slide master and its associated layouts from the template file.

**When to use:** This is appropriate when you want to change the design/theme of an **existing** presentation to match a template, without adding any content slides.

### 3.3 Approach C: `Presentations.Open()` without `Untitled` (Opens Template for Editing)

**VBA Syntax:**
```vb
Set pres = Application.Presentations.Open("C:\path\to\template.potx")
```

**pywin32 Code:**
```python
pres = app.Presentations.Open(abs_path, 0, 0, -1)
```

**Behavior:**
- Opens the .potx file **for editing** — the title bar shows the template name.
- Saving will **overwrite the original template file**.
- This is the mode for editing/modifying the template itself, not for creating new presentations from it.

### 3.4 Comparison Table

| Aspect | Open + Untitled=True | Add + ApplyTemplate | Open (default) |
|---|---|---|---|
| Sample slides from template | Preserved | NOT copied | Preserved |
| Theme/design | Applied | Applied | Applied |
| Custom layouts | Preserved | Copied (design only) | Preserved |
| Custom XML data | Preserved | Lost | Preserved |
| Result is untitled | Yes | Yes | No (has file path) |
| Modifies template file | No | No | Yes (if saved) |
| **Best for** | **New pres from template** | Change existing design | Edit template itself |

### 3.5 `Presentations.Add()` Has No Template Parameter

The `Presentations.Add()` method only accepts a single optional parameter:

```
Presentations.Add(WithWindow)
```

- `WithWindow` (MsoTriState): Whether to show the presentation window.
- **There is no template path parameter.** To use a template, you must use one of the approaches above.

### 3.6 Related Methods

#### `Presentation.ApplyTemplate2(FileName, Variant)`
Applies a design template with a specific theme variant. Requires both a file path and a variant ID string. The variant can be obtained from `Application.OpenThemeFile(path).ThemeVariants(n).Id`.

#### `Slide.ApplyTemplate(FileName)` / `Slide.ApplyTemplate2(FileName, Variant)`
Applies a template design to a **specific slide** rather than the entire presentation. This allows different slides to have different designs (multiple slide masters).

#### `Presentation.ApplyTheme(FileName)`
Applies a .thmx theme file (not a full template). Only changes colors, fonts, and effects — does not include slide layouts.

#### `Presentation.TemplateName` (Read-only property)
Returns the name of the first design/master associated with the presentation. Does not include the full path. Useful for checking which template is currently applied.

---

## 4. Best Practice Recommendation for MCP Tool Implementation

### 4.1 Recommended Approach

Use **Approach A: `Presentations.Open()` with `Untitled=msoTrue`** as the primary method. This perfectly replicates what happens when a user double-clicks a .potx file.

### 4.2 Proposed Tool: `ppt_create_presentation` Enhancement

Extend the existing `ppt_create_presentation` tool to accept an optional `template_path` parameter:

```python
class CreatePresentationInput(BaseModel):
    """Input for creating a new presentation."""
    model_config = ConfigDict(str_strip_whitespace=True)

    template_path: Optional[str] = Field(
        default=None,
        description=(
            "Absolute path to a PowerPoint template file (.potx, .potm, .pot). "
            "When provided, creates a new presentation based on this template, "
            "preserving all slides, layouts, themes, and custom data. "
            "Other size parameters are ignored when a template is used."
        ),
    )
    slide_width: Optional[float] = Field(default=None, ...)
    slide_height: Optional[float] = Field(default=None, ...)
    preset: Optional[str] = Field(default=None, ...)
```

**Implementation:**

```python
def _create_presentation_impl(
    template_path: Optional[str],
    slide_width: Optional[float],
    slide_height: Optional[float],
    preset: Optional[str],
) -> dict:
    app = ppt._get_app_impl()

    if template_path:
        # Create from template using Open + Untitled
        abs_path = os.path.abspath(template_path)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"Template not found: {abs_path}")

        pres = app.Presentations.Open(
            abs_path,   # FileName
            0,          # ReadOnly = msoFalse
            -1,         # Untitled = msoTrue
            -1,         # WithWindow = msoTrue
        )
    else:
        # Create blank presentation (existing behavior)
        pres = app.Presentations.Add()

        if preset:
            preset_key = preset.strip()
            if preset_key not in SLIDE_SIZE_PRESETS:
                raise ValueError(f"Unknown preset '{preset}'.")
            w, h = SLIDE_SIZE_PRESETS[preset_key]
            pres.PageSetup.SlideWidth = w
            pres.PageSetup.SlideHeight = h
        elif slide_width is not None and slide_height is not None:
            pres.PageSetup.SlideWidth = slide_width
            pres.PageSetup.SlideHeight = slide_height

    template_name = ""
    try:
        template_name = pres.TemplateName
    except Exception:
        pass

    return {
        "success": True,
        "name": pres.Name,
        "slides_count": pres.Slides.Count,
        "slide_width": pres.PageSetup.SlideWidth,
        "slide_height": pres.PageSetup.SlideHeight,
        "template_name": template_name,
    }
```

### 4.3 Proposed Tool: `ppt_list_templates`

A new read-only tool to discover available templates:

```python
class ListTemplatesInput(BaseModel):
    """Input for listing available templates."""
    model_config = ConfigDict(str_strip_whitespace=True)

    templates_dir: Optional[str] = Field(
        default=None,
        description=(
            "Directory to scan for template files. "
            "If omitted, scans the default personal templates folder."
        ),
    )
    recursive: bool = Field(
        default=False,
        description="If true, also search subdirectories.",
    )

def _list_templates_impl(templates_dir: Optional[str], recursive: bool) -> dict:
    import glob as glob_mod

    if templates_dir is None:
        templates_dir = _get_default_templates_dir()

    if not os.path.isdir(templates_dir):
        return {
            "templates": [],
            "templates_dir": templates_dir,
            "error": f"Directory not found: {templates_dir}",
        }

    templates = []
    pattern_prefix = "**/" if recursive else ""
    for ext in ("potx", "potm", "pot"):
        pattern = os.path.join(templates_dir, f"{pattern_prefix}*.{ext}")
        for filepath in glob_mod.glob(pattern, recursive=recursive):
            stat = os.stat(filepath)
            templates.append({
                "name": os.path.splitext(os.path.basename(filepath))[0],
                "file_name": os.path.basename(filepath),
                "file_path": os.path.abspath(filepath),
                "size_bytes": stat.st_size,
            })

    templates.sort(key=lambda t: t["name"])
    return {
        "templates_dir": templates_dir,
        "count": len(templates),
        "templates": templates,
    }

def _get_default_templates_dir() -> str:
    """Resolve the default personal templates folder."""
    import winreg

    # Try registry first
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Office\16.0\PowerPoint\Options"
        )
        path, _ = winreg.QueryValueEx(key, "PersonalTemplates")
        winreg.CloseKey(key)
        if path and os.path.isdir(path):
            return path
    except (FileNotFoundError, OSError):
        pass

    # Fallback to default
    return os.path.join(os.environ["USERPROFILE"], "Documents", "Custom Office Templates")
```

### 4.4 Proposed Tool: `ppt_apply_template`

For applying a template design to an existing presentation:

```python
class ApplyTemplateInput(BaseModel):
    """Input for applying a template to a presentation."""
    model_config = ConfigDict(str_strip_whitespace=True)

    template_path: str = Field(
        ...,
        description="Absolute path to the template file (.potx, .potm, .pot, .pptx, .thmx).",
    )
    slide_number: Optional[int] = Field(
        default=None,
        description=(
            "1-based slide number to apply the template to. "
            "If omitted, applies to the entire presentation."
        ),
    )
    presentation_index: Optional[int] = Field(
        default=None,
        description="1-based presentation index. If omitted, uses active presentation.",
    )

def _apply_template_impl(
    template_path: str,
    slide_number: Optional[int],
    presentation_index: Optional[int],
) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(app, presentation_index)
    abs_path = os.path.abspath(template_path)

    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"Template not found: {abs_path}")

    if slide_number is not None:
        slide = pres.Slides(slide_number)
        slide.ApplyTemplate(abs_path)
    else:
        pres.ApplyTemplate(abs_path)

    return {
        "success": True,
        "template_name": pres.TemplateName,
        "applied_to": f"slide {slide_number}" if slide_number else "entire presentation",
    }
```

### 4.5 COM Thread Safety Notes

All the above implementations follow the project's existing pattern:

- The `_impl` functions run on the COM STA thread via `ppt.execute()`.
- Use positional arguments for COM method calls (not keyword arguments).
- Use integer constants directly (`-1` for msoTrue, `0` for msoFalse).
- File paths must be absolute Windows paths.

### 4.6 Testing Checklist

Since there is no automated test suite, manually test against a live PowerPoint instance:

1. **Create from template** — Verify slides, layouts, themes, and fonts are all preserved.
2. **Template with sample slides** — Verify all sample slides appear in the new presentation.
3. **Save after creation** — Verify SaveAs works on the untitled presentation.
4. **Apply template to existing** — Verify design changes without losing existing slide content.
5. **List templates** — Verify correct enumeration of .potx files in the templates folder.
6. **Non-existent template** — Verify clear error message.
7. **Japanese characters in path** — Verify `C:\Users\kuwam\OneDrive\ドキュメント\Office のカスタム テンプレート` works correctly.

---

## 5. References

- [Presentations.Open method (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentations.open)
- [Presentations.Add method (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentations.add)
- [Presentation.ApplyTemplate method (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.applytemplate)
- [Presentation.ApplyTemplate2 method (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.applytemplate2)
- [Slide.ApplyTemplate method (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Slide.ApplyTemplate)
- [Presentation.TemplateName property (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.templatename)
- [CustomLayouts object (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.customlayouts)
- [Create new presentation from template with VBA - Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/65a19455-f237-4971-9bd0-61b172a16785/create-new-presentation-from-template-with-vba)
- [Office templates registry locations - Experts Exchange](https://www.experts-exchange.com/questions/28738548/Office-templates-location-s-in-relation-to-windows-registry.html)
- [Custom theme locations - Indezine](https://www.indezine.com/products/powerpoint/learn/themes/custom-theme-locations.html)
- [Default file location in VBA - microsoft.public.powerpoint](https://microsoft.public.powerpoint.narkive.com/ECBLFxHU/default-file-location-in-vba)
- [Personal templates path policy - admx.help](https://admx.help/?Category=Office2016&Policy=ppt16.Office.Microsoft.Policies.Windows::L_PersonalTemplatesPath)
