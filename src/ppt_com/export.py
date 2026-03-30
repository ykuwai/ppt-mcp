"""Export tools for PowerPoint COM automation.

Export presentations to PDF, images (PNG/JPG), or copy slides to clipboard.
"""

import atexit
import ctypes
import ctypes.wintypes
import json
import logging
import os
import shutil
import struct
import tempfile
from typing import List, Optional

import pythoncom
from pydantic import BaseModel, Field, ConfigDict, model_validator

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
    file_name: Optional[str] = Field(
        default=None,
        description="Custom filename for single-slide export (e.g. 'cover.png'). If omitted, defaults to 'Slide{N}.{format}'. Requires slide_index.",
    )

    @model_validator(mode="after")
    def file_name_requires_slide_index(self):
        if self.file_name is not None and self.slide_index is None:
            raise ValueError("file_name requires slide_index to be set")
        return self


class CopyToClipboardInput(BaseModel):
    """Input for copying slides as images to the clipboard."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_indices: Optional[List[int]] = Field(
        default=None,
        description=(
            "1-based slide indices to copy. "
            "If omitted, copies the currently viewed slide."
        ),
    )
    width: Optional[int] = Field(
        default=None,
        description="Image width in pixels. If omitted, uses default resolution.",
    )
    height: Optional[int] = Field(
        default=None,
        description="Image height in pixels. If omitted, uses default resolution.",
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
    pres = ppt._get_pres_impl()

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
    file_name: Optional[str],
) -> dict:
    app = ppt._get_app_impl()
    if app.Presentations.Count == 0:
        raise RuntimeError(
            "No presentation is open. "
            "Use ppt_create_presentation or ppt_open_presentation first."
        )
    pres = ppt._get_pres_impl()

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

        if file_name:
            # Ensure correct extension (strip wrong extension to avoid double ext)
            base, ext = os.path.splitext(file_name)
            if ext.lower() != f".{fmt_key}":
                file_name = f"{base}.{fmt_key}"
        else:
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
# Clipboard helpers (Windows only)
# ---------------------------------------------------------------------------
CF_DIB = 8
CF_HDROP = 15
GHND = 0x0042  # GMEM_MOVEABLE | GMEM_ZEROINIT

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
ole32 = ctypes.windll.ole32

OpenClipboard = user32.OpenClipboard
CloseClipboard = user32.CloseClipboard
EmptyClipboard = user32.EmptyClipboard
SetClipboardData = user32.SetClipboardData
SetClipboardData.argtypes = [ctypes.wintypes.UINT, ctypes.wintypes.HANDLE]
SetClipboardData.restype = ctypes.wintypes.HANDLE
GlobalAlloc = kernel32.GlobalAlloc
GlobalAlloc.argtypes = [ctypes.wintypes.UINT, ctypes.c_size_t]
GlobalAlloc.restype = ctypes.wintypes.HGLOBAL
GlobalLock = kernel32.GlobalLock
GlobalLock.argtypes = [ctypes.wintypes.HGLOBAL]
GlobalLock.restype = ctypes.c_void_p
GlobalUnlock = kernel32.GlobalUnlock
GlobalUnlock.argtypes = [ctypes.wintypes.HGLOBAL]
GlobalFree = kernel32.GlobalFree
GlobalFree.argtypes = [ctypes.wintypes.HGLOBAL]


def _png_to_dib(png_path: str) -> bytes:
    """Convert a PNG file to DIB (Device Independent Bitmap) bytes.

    Reads the PNG via COM-free WIC (Windows Imaging Component) is overkill;
    instead, we simply load with the ``PIL``-free BMP approach: export as BMP
    from PowerPoint is not available, so we use a minimal PNG→BMP decoder
    via ctypes GDI+.
    """
    # Use GDI+ to load PNG and convert to DIB
    from ctypes import byref, c_int, c_uint, c_void_p, POINTER

    gdiplus = ctypes.windll.gdiplus

    # GDI+ startup
    class GdiplusStartupInput(ctypes.Structure):
        _fields_ = [
            ("GdiplusVersion", c_uint),
            ("DebugEventCallback", c_void_p),
            ("SuppressBackgroundThread", c_int),
            ("SuppressExternalCodecs", c_int),
        ]

    token = ctypes.c_ulong()
    startup_input = GdiplusStartupInput(1, None, 0, 0)
    gdiplus.GdiplusStartup(byref(token), byref(startup_input), None)

    try:
        # Load image from file
        bitmap = c_void_p()
        status = gdiplus.GdipCreateBitmapFromFile(png_path, byref(bitmap))
        if status != 0:
            raise RuntimeError(f"GDI+ failed to load image: status {status}")

        try:
            # Get HBITMAP
            hbitmap = ctypes.wintypes.HBITMAP()
            # Background color: ARGB white
            status = gdiplus.GdipCreateHBITMAPFromBitmap(
                bitmap, byref(hbitmap), 0xFFFFFFFF
            )
            if status != 0:
                raise RuntimeError(
                    f"GdipCreateHBITMAPFromBitmap failed: status {status}"
                )

            # Convert HBITMAP to DIB bytes via GetDIBits
            gdi32 = ctypes.windll.gdi32

            class BITMAPINFOHEADER(ctypes.Structure):
                _fields_ = [
                    ("biSize", ctypes.wintypes.DWORD),
                    ("biWidth", ctypes.wintypes.LONG),
                    ("biHeight", ctypes.wintypes.LONG),
                    ("biPlanes", ctypes.wintypes.WORD),
                    ("biBitCount", ctypes.wintypes.WORD),
                    ("biCompression", ctypes.wintypes.DWORD),
                    ("biSizeImage", ctypes.wintypes.DWORD),
                    ("biXPelsPerMeter", ctypes.wintypes.LONG),
                    ("biYPelsPerMeter", ctypes.wintypes.LONG),
                    ("biClrUsed", ctypes.wintypes.DWORD),
                    ("biClrImportant", ctypes.wintypes.DWORD),
                ]

            class BITMAP(ctypes.Structure):
                _fields_ = [
                    ("bmType", ctypes.wintypes.LONG),
                    ("bmWidth", ctypes.wintypes.LONG),
                    ("bmHeight", ctypes.wintypes.LONG),
                    ("bmWidthBytes", ctypes.wintypes.LONG),
                    ("bmPlanes", ctypes.wintypes.WORD),
                    ("bmBitsPixel", ctypes.wintypes.WORD),
                    ("bmBits", c_void_p),
                ]

            try:
                bm = BITMAP()
                gdi32.GetObjectW(hbitmap, ctypes.sizeof(BITMAP), byref(bm))

                bih = BITMAPINFOHEADER()
                bih.biSize = ctypes.sizeof(BITMAPINFOHEADER)
                bih.biWidth = bm.bmWidth
                bih.biHeight = bm.bmHeight  # positive = bottom-up
                bih.biPlanes = 1
                bih.biBitCount = 32
                bih.biCompression = 0  # BI_RGB

                row_size = ((bm.bmWidth * 32 + 31) // 32) * 4
                bih.biSizeImage = row_size * bm.bmHeight

                # Allocate pixel buffer
                pixel_buf = (ctypes.c_byte * bih.biSizeImage)()

                hdc = user32.GetDC(0)
                gdi32.GetDIBits(
                    hdc, hbitmap, 0, bm.bmHeight,
                    pixel_buf, byref(bih), 0  # DIB_RGB_COLORS
                )
                user32.ReleaseDC(0, hdc)

                # DIB = BITMAPINFOHEADER + pixel data
                return bytes(bih) + bytes(pixel_buf)
            finally:
                gdi32.DeleteObject(hbitmap)

        finally:
            gdiplus.GdipDisposeImage(bitmap)
    finally:
        gdiplus.GdiplusShutdown(token)


def _set_clipboard_dib(dib_data: bytes) -> None:
    """Place DIB data on the clipboard (single image)."""
    hglob = GlobalAlloc(GHND, len(dib_data))
    if not hglob:
        raise RuntimeError("GlobalAlloc failed")
    ptr = GlobalLock(hglob)
    ctypes.memmove(ptr, dib_data, len(dib_data))
    GlobalUnlock(hglob)

    if not OpenClipboard(0):
        GlobalFree(hglob)
        raise RuntimeError("Cannot open clipboard")
    try:
        EmptyClipboard()
        if not SetClipboardData(CF_DIB, hglob):
            GlobalFree(hglob)
            raise RuntimeError("SetClipboardData failed")
        # After successful SetClipboardData, system owns the memory
    finally:
        CloseClipboard()


def _set_clipboard_hdrop(file_paths: list[str]) -> None:
    """Place file paths on the clipboard as HDROP (file drop list).

    This allows pasting multiple images into Word, PowerPoint, etc.
    """
    # DROPFILES structure:
    #   DWORD pFiles (offset to file list)
    #   POINT pt (unused, 0,0)
    #   BOOL  fNC (FALSE)
    #   BOOL  fWide (TRUE for Unicode)
    # Followed by double-null-terminated wide-char file list
    offset = 20  # sizeof(DROPFILES)

    # Build the file list: each path null-terminated, extra null at end
    file_list = ""
    for p in file_paths:
        file_list += p + "\0"
    file_list += "\0"  # double-null terminator

    file_list_bytes = file_list.encode("utf-16-le")
    total_size = offset + len(file_list_bytes)

    # Pack DROPFILES header
    header = struct.pack("IiiII", offset, 0, 0, 0, 1)  # fWide=1

    hglob = GlobalAlloc(GHND, total_size)
    if not hglob:
        raise RuntimeError("GlobalAlloc failed")
    ptr = GlobalLock(hglob)
    ctypes.memmove(ptr, header, len(header))
    ctypes.memmove(ptr + offset, file_list_bytes, len(file_list_bytes))
    GlobalUnlock(hglob)

    if not OpenClipboard(0):
        GlobalFree(hglob)
        raise RuntimeError("Cannot open clipboard")
    try:
        EmptyClipboard()
        if not SetClipboardData(CF_HDROP, hglob):
            GlobalFree(hglob)
            raise RuntimeError("SetClipboardData failed")
    finally:
        CloseClipboard()


def _copy_to_clipboard_impl(
    slide_indices: Optional[List[int]],
    width: Optional[int],
    height: Optional[int],
) -> dict:
    """Export slide(s) as PNG to temp files, then copy to clipboard."""
    from utils.navigation import goto_slide

    app = ppt._get_app_impl()
    if app.Presentations.Count == 0:
        raise RuntimeError(
            "No presentation is open. "
            "Use ppt_create_presentation or ppt_open_presentation first."
        )
    pres = ppt._get_pres_impl()
    total_slides = pres.Slides.Count

    # Resolve slide indices
    if slide_indices is None or len(slide_indices) == 0:
        # Default: currently viewed slide
        try:
            current = app.ActiveWindow.View.Slide.SlideIndex
        except Exception:
            current = 1
        slide_indices = [current]

    # Validate indices
    for idx in slide_indices:
        if idx < 1 or idx > total_slides:
            raise ValueError(
                f"Slide index {idx} out of range (1-{total_slides})"
            )

    # Navigate to the last slide being copied so the user can see it
    goto_slide(app, slide_indices[-1])

    # Export to temp directory
    tmp_dir = tempfile.mkdtemp(prefix="ppt_clipboard_")
    exported_files = []
    try:
        for idx in slide_indices:
            fname = os.path.join(tmp_dir, f"Slide{idx}.png")
            slide = pres.Slides(idx)
            if width is not None and height is not None:
                slide.Export(fname, "PNG", width, height)
            elif width is not None:
                slide.Export(fname, "PNG", width)
            else:
                slide.Export(fname, "PNG")
            exported_files.append(fname)

        # Copy to clipboard
        if len(exported_files) == 1:
            # Single image: use CF_DIB for direct paste as image
            dib_data = _png_to_dib(exported_files[0])
            _set_clipboard_dib(dib_data)
        else:
            # Multiple images: use CF_HDROP for file drop
            _set_clipboard_hdrop(exported_files)

    except Exception:
        # Clean up temp files on error
        shutil.rmtree(tmp_dir, ignore_errors=True)
        raise

    # Single image: temp files can be cleaned up immediately
    # HDROP: temp files must persist until paste; register atexit cleanup
    if len(exported_files) == 1:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    else:
        atexit.register(shutil.rmtree, tmp_dir, True)

    return {
        "success": True,
        "slide_indices": slide_indices,
        "count": len(slide_indices),
        "method": "CF_DIB" if len(slide_indices) == 1 else "CF_HDROP",
        "temp_dir": tmp_dir if len(exported_files) > 1 else None,
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
            params.file_name,
        )
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def copy_to_clipboard(params: CopyToClipboardInput) -> str:
    """Copy slides as PNG images to the clipboard."""
    try:
        result = ppt.execute(
            _copy_to_clipboard_impl,
            params.slide_indices,
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

    @mcp.tool(
        name="ppt_copy_to_clipboard",
        annotations={
            "title": "Copy Slides to Clipboard",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_copy_to_clipboard(params: CopyToClipboardInput) -> str:
        """Copy slides as PNG images to the Windows clipboard.

        By default, copies the currently viewed slide.
        Specify slide_indices to copy specific slides.
        Single slide is placed as a bitmap (paste directly as image).
        Multiple slides are placed as file drop (paste inserts all images).
        """
        return copy_to_clipboard(params)
