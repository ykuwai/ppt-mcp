"""Export tools for PowerPoint COM automation.

Export presentations to PDF or images (PNG/JPG).
"""

import json
import logging
import os
import tempfile
from typing import Optional

import pythoncom
from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from ppt_com.constants import (
    ppFixedFormatTypePDF,
    ppSaveAsPDF,
    ppSaveAsPNG,
    ppSaveAsJPG,
)

logger = logging.getLogger(__name__)

IMAGE_FORMAT_MAP = {
    "png": ppSaveAsPNG,
    "jpg": ppSaveAsJPG,
}

IMAGE_FILTER_MAP = {
    "png": "PNG",
    "jpg": "JPG",
}


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class ExportPDFInput(BaseModel):
    """Input for exporting a presentation to PDF."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_path: str = Field(
        ...,
        description="Full path for the output PDF file.",
    )
    slide_range_start: Optional[int] = Field(
        default=None,
        description="1-based starting slide index for partial export.",
    )
    slide_range_end: Optional[int] = Field(
        default=None,
        description="1-based ending slide index for partial export.",
    )


class ExportImagesInput(BaseModel):
    """Input for exporting slides as images."""
    model_config = ConfigDict(str_strip_whitespace=True)

    output_dir: str = Field(
        ...,
        description="Directory to save exported images.",
    )
    format: str = Field(
        default="png",
        description="Image format: 'png' or 'jpg'.",
    )
    slide_index: Optional[int] = Field(
        default=None,
        description="1-based slide index to export a single slide. If omitted, exports all slides.",
    )
    width: Optional[int] = Field(
        default=None,
        description="Image width in pixels (for single slide export).",
    )
    height: Optional[int] = Field(
        default=None,
        description="Image height in pixels (for single slide export).",
    )


# ---------------------------------------------------------------------------
# Implementation functions (run on COM thread via ppt.execute)
# ---------------------------------------------------------------------------
def _export_pdf_all_slides(pres, abs_path: str) -> None:
    """Export all slides to PDF using InvokeTypes to bypass pywin32 bug.

    ExportAsFixedFormat has a known pywin32 issue where the ExternalExporter
    parameter (VT_VARIANT|VT_BYREF) causes a "Python instance cannot be
    converted to COM object" error. We work around this by calling InvokeTypes
    directly with corrected parameter type flags.
    """
    pres._oleobj_.InvokeTypes(
        2096, 0, 1,                     # dispid, lcid, DISPATCH_METHOD
        (24, 32),                        # return: void
        (
            (8, 1), (3, 1),              # Path (BSTR), FixedFormatType (LONG)
            (3, 49), (3, 49), (3, 49),   # Intent, FrameSlides, HandoutOrder
            (3, 49), (3, 49),            # OutputType, PrintHiddenSlides
            (9, 49), (3, 49),            # PrintRange (IDispatch), RangeType
            (8, 49),                     # SlideShowName
            (11, 49), (11, 49),          # IncludeDocProperties, KeepIRMSettings
            (11, 49), (11, 49),          # DocStructureTags, BitmapMissingFonts
            (11, 49),                    # UseISO19005_1
            (12, 49),                    # ExternalExporter (fixed: VT_VARIANT optional)
        ),
        abs_path, ppFixedFormatTypePDF,
        1, 0, 1, 1, 0,                  # Intent=screen, defaults
        None, 1,                         # PrintRange=None, RangeType=ppPrintAll
        '',
        False, True, True, True, False,
        pythoncom.Empty,                 # ExternalExporter
    )


def _export_pdf_slide_range(app, pres, abs_path: str, start: int, end: int) -> None:
    """Export a slide range to PDF by creating a temporary copy.

    pywin32 cannot marshal the PrintRange COM object through InvokeTypes,
    so we work around this by saving a copy, deleting unwanted slides,
    and exporting the trimmed copy as PDF.
    """
    tmp_file = os.path.join(tempfile.gettempdir(), 'ppt_export_temp.pptx')
    try:
        pres.SaveCopyAs(tmp_file)
        tmp_pres = app.Presentations.Open(tmp_file, WithWindow=False)
        try:
            total = tmp_pres.Slides.Count
            # Delete slides after the range (high to low to avoid reindexing)
            for i in range(total, end, -1):
                tmp_pres.Slides(i).Delete()
            # Delete slides before the range
            for i in range(start - 1, 0, -1):
                tmp_pres.Slides(i).Delete()
            tmp_pres.SaveAs(abs_path, ppSaveAsPDF)
        finally:
            tmp_pres.Close()
    finally:
        if os.path.exists(tmp_file):
            os.remove(tmp_file)


def _export_pdf_impl(
    file_path: str,
    slide_range_start: Optional[int],
    slide_range_end: Optional[int],
) -> dict:
    app = ppt._get_app_impl()
    if app.Presentations.Count == 0:
        raise RuntimeError(
            "No presentation is open. "
            "Use ppt_create_presentation or ppt_open_presentation first."
        )
    pres = app.ActivePresentation

    # COM requires absolute Windows-style paths
    abs_path = os.path.abspath(file_path)

    # Ensure output directory exists
    out_dir = os.path.dirname(abs_path)
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    if slide_range_start is not None and slide_range_end is not None:
        total = pres.Slides.Count
        if slide_range_start < 1 or slide_range_start > total:
            raise ValueError(
                f"slide_range_start {slide_range_start} out of range (1-{total})"
            )
        if slide_range_end < slide_range_start or slide_range_end > total:
            raise ValueError(
                f"slide_range_end {slide_range_end} out of range "
                f"({slide_range_start}-{total})"
            )
        _export_pdf_slide_range(app, pres, abs_path, slide_range_start, slide_range_end)
    else:
        _export_pdf_all_slides(pres, abs_path)

    return {
        "success": True,
        "file_path": abs_path,
        "slide_range_start": slide_range_start,
        "slide_range_end": slide_range_end,
        "total_slides": pres.Slides.Count,
    }


def _export_images_impl(
    output_dir: str,
    format: str,
    slide_index: Optional[int],
    width: Optional[int],
    height: Optional[int],
) -> dict:
    app = ppt._get_app_impl()
    if app.Presentations.Count == 0:
        raise RuntimeError(
            "No presentation is open. "
            "Use ppt_create_presentation or ppt_open_presentation first."
        )
    pres = app.ActivePresentation

    fmt_key = format.lower().strip()
    if fmt_key not in IMAGE_FORMAT_MAP:
        raise ValueError(
            f"Unknown image format '{format}'. Supported: {list(IMAGE_FORMAT_MAP.keys())}"
        )

    filter_name = IMAGE_FILTER_MAP[fmt_key]

    if slide_index is not None:
        # Export single slide
        if slide_index < 1 or slide_index > pres.Slides.Count:
            raise ValueError(
                f"Slide index {slide_index} out of range (1-{pres.Slides.Count})"
            )

        # COM requires absolute Windows-style paths
        abs_dir = os.path.abspath(output_dir)
        if not os.path.exists(abs_dir):
            os.makedirs(abs_dir, exist_ok=True)

        file_name = f"Slide{slide_index}.{fmt_key}"
        abs_file_path = os.path.join(abs_dir, file_name)

        # Slide.Export positional args: FileName, FilterName, ScaleWidth, ScaleHeight
        slide = pres.Slides(slide_index)
        if width is not None and height is not None:
            slide.Export(abs_file_path, filter_name, width, height)
        elif width is not None:
            slide.Export(abs_file_path, filter_name, width)
        else:
            slide.Export(abs_file_path, filter_name)

        return {
            "success": True,
            "output_dir": abs_dir,
            "format": fmt_key,
            "slide_index": slide_index,
            "files": [abs_file_path],
        }
    else:
        # Export all slides - SaveAs creates a folder with individual images
        abs_dir = os.path.abspath(output_dir)
        pres.SaveAs(abs_dir, IMAGE_FORMAT_MAP[fmt_key])

        # Collect exported files
        exported_files = []
        if os.path.isdir(abs_dir):
            for f in sorted(os.listdir(abs_dir)):
                if f.lower().endswith(f".{fmt_key}"):
                    exported_files.append(os.path.join(abs_dir, f))

        return {
            "success": True,
            "output_dir": abs_dir,
            "format": fmt_key,
            "total_slides": pres.Slides.Count,
            "files_count": len(exported_files),
            "files": exported_files,
        }


# ---------------------------------------------------------------------------
# MCP tool functions (return JSON strings)
# ---------------------------------------------------------------------------
def export_pdf(params: ExportPDFInput) -> str:
    """Export the active presentation to PDF."""
    try:
        result = ppt.execute(
            _export_pdf_impl,
            params.file_path,
            params.slide_range_start,
            params.slide_range_end,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def export_images(params: ExportImagesInput) -> str:
    """Export slides as images (PNG or JPG)."""
    try:
        result = ppt.execute(
            _export_images_impl,
            params.output_dir,
            params.format,
            params.slide_index,
            params.width,
            params.height,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register all export tools with the MCP server."""

    @mcp.tool(
        name="ppt_export_pdf",
        annotations={
            "title": "Export to PDF",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        },
    )
    async def tool_export_pdf(params: ExportPDFInput) -> str:
        """Export the active presentation to a PDF file.

        Optionally export a specific range of slides by providing
        slide_range_start and slide_range_end.
        """
        return export_pdf(params)

    @mcp.tool(
        name="ppt_export_images",
        annotations={
            "title": "Export as Images",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        },
    )
    async def tool_export_images(params: ExportImagesInput) -> str:
        """Export slides as images (PNG or JPG).

        Export all slides to a directory, or a single slide by index.
        For single slide export, optionally specify width and height in pixels.
        For all slides, PowerPoint creates a folder of individual images.
        """
        return export_images(params)
