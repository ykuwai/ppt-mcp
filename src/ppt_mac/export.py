"""Export tools, on Apple Events.

Mirrors ``ppt_com/export.py``. This is the module where macOS differs most, for
two reasons that are worth stating plainly rather than discovering at runtime.

**There is no export command.** PowerPoint's Apple Event dictionary has no
counterpart to ``Slide.Export``, and the one thing that looks like it,
``save <pres> in <dir> as save as PNG``, reports success in 0.15 seconds and
writes nothing at all. Verified repeatedly, including from a deck with a real
file path saved inside PowerPoint's own container. So slide images are produced
by exporting the deck to PDF, which does work, and rendering the pages with
Quartz, which ships with macOS.

**PowerPoint is sandboxed.** Exporting to a directory it has not written to
before blocks for tens of seconds and then kills the application with -609.
Reproduced three times during the port. Its own container is always writable,
so every export is staged there and moved out afterwards by this process, which
is not sandboxed.
"""

import logging
import os
import shutil
import tempfile
from typing import List, Optional

from backend.mac_ae import EXPORT_STAGING_DIR, count, mactypes, ppt
from backend.mac_enums import PpSaveAsFileType, to_keyword
from ppt_com.constants import ppSaveAsPDF

logger = logging.getLogger(__name__)

# Quartz renders the staged PDF. It arrives with pyobjc, which is a macOS only
# dependency, and it is imported lazily so that merely importing this module
# does not pay for it.
_QUARTZ_HINT = (
    "Slide images need Quartz, which comes from the pyobjc-framework-Quartz "
    "package. Reinstall ppt-mcp so the macOS dependencies are present."
)


def _staging_path(suffix: str) -> str:
    """A path inside PowerPoint's container, which it is always allowed to write.

    Anywhere else risks the sandbox stall that kills the application, so no
    export is ever pointed straight at the caller's directory.
    """
    os.makedirs(EXPORT_STAGING_DIR, exist_ok=True)
    handle, path = tempfile.mkstemp(
        prefix="ppt_mcp_export_", suffix=suffix, dir=EXPORT_STAGING_DIR
    )
    os.close(handle)
    # PowerPoint wants to create the file itself, so leave only the name behind.
    os.remove(path)
    return path


def _save_pdf(pres, destination: str) -> None:
    """Export the whole deck to PDF, and confirm that it happened.

    PowerPoint answers without error for saves that write nothing, so the file
    on disk is the only evidence worth acting on.
    """
    pres.save(
        in_=mactypes.File(destination),
        as_=to_keyword(PpSaveAsFileType, ppSaveAsPDF, "save format"),
    )
    if not os.path.exists(destination) or os.path.getsize(destination) == 0:
        raise RuntimeError(
            "PowerPoint reported success but wrote no PDF. This usually means "
            "the presentation has never been saved to a file, or that it was "
            "asked to write somewhere its sandbox does not allow."
        )


def _render_pdf_pages(pdf_path, pages, width=None, height=None,
                      uti="public.png", suffix=".png"):
    """Render chosen PDF pages to image files with Quartz.

    ``pages`` is 1 based, matching slide numbering everywhere else. When no
    size is given the page is rendered at twice its natural size, which for a
    standard 960 by 540 point deck is 1920 by 1080.

    ``uti`` and ``suffix`` have to agree. They were once fixed at PNG while the
    caller was free to ask for jpg, so a file named .jpg held PNG bytes and
    anything decoding by the format it asked for choked on it.
    """
    try:
        import Quartz
        from CoreFoundation import (
            CFURLCreateFromFileSystemRepresentation,
            kCFAllocatorDefault,
        )
    except ImportError as exc:  # pragma: no cover - depends on the install
        raise RuntimeError(_QUARTZ_HINT) from exc

    def _url(path):
        encoded = path.encode()
        return CFURLCreateFromFileSystemRepresentation(
            kCFAllocatorDefault, encoded, len(encoded), False
        )

    document = Quartz.CGPDFDocumentCreateWithURL(_url(pdf_path))
    if document is None:
        raise RuntimeError("The exported PDF could not be read back.")

    total = Quartz.CGPDFDocumentGetNumberOfPages(document)
    rendered = []
    for number in pages:
        if number < 1 or number > total:
            raise ValueError(f"Slide {number} is not in the exported PDF (1-{total})")
        page = Quartz.CGPDFDocumentGetPage(document, number)
        if page is None:
            # Answering None rather than raising is how Quartz declines a page
            # it has, and drawing None draws nothing at all.
            raise RuntimeError(
                f"Slide {number} is missing from the deck's PDF export, so no "
                "image could be rendered for it."
            )
        box = Quartz.CGPDFPageGetBoxRect(page, Quartz.kCGPDFMediaBox)

        if width and height:
            out_w, out_h = int(width), int(height)
        elif width:
            out_w = int(width)
            out_h = int(round(out_w * box.size.height / box.size.width))
        else:
            out_w = int(box.size.width * 2)
            out_h = int(box.size.height * 2)

        context = Quartz.CGBitmapContextCreate(
            None, out_w, out_h, 8, 0,
            Quartz.CGColorSpaceCreateDeviceRGB(),
            Quartz.kCGImageAlphaPremultipliedFirst | Quartz.kCGBitmapByteOrder32Host,
        )
        # A slide is opaque. Without this the areas the PDF leaves untouched
        # come out transparent, which reads as black in most viewers.
        Quartz.CGContextSetRGBFillColor(context, 1, 1, 1, 1)
        Quartz.CGContextFillRect(context, Quartz.CGRectMake(0, 0, out_w, out_h))
        Quartz.CGContextScaleCTM(
            context, out_w / box.size.width, out_h / box.size.height
        )
        Quartz.CGContextDrawPDFPage(context, page)
        image = Quartz.CGBitmapContextCreateImage(context)

        out_path = _staging_path(suffix)
        destination = Quartz.CGImageDestinationCreateWithURL(
            _url(out_path), uti, 1, None
        )
        Quartz.CGImageDestinationAddImage(destination, image, None)
        if not Quartz.CGImageDestinationFinalize(destination):
            raise RuntimeError(f"Quartz could not write the image for slide {number}")
        rendered.append((number, out_path, out_w, out_h))
    return rendered


def _export_pdf_impl(
    file_path: str,
    slide_range_start: Optional[int],
    slide_range_end: Optional[int],
) -> dict:
    app = ppt._get_app_impl()
    if count(app.presentations) == 0:
        raise RuntimeError(
            "No presentation is open. "
            "Use ppt_create_presentation or ppt_open_presentation first."
        )
    pres = ppt._get_pres_impl()
    total = count(pres.slides)

    abs_path = os.path.abspath(file_path)
    out_dir = os.path.dirname(abs_path)
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    if slide_range_start is not None and slide_range_end is not None:
        if slide_range_start < 1 or slide_range_start > total:
            raise ValueError(
                f"slide_range_start {slide_range_start} out of range (1-{total})"
            )
        if slide_range_end < slide_range_start or slide_range_end > total:
            raise ValueError(
                f"slide_range_end {slide_range_end} out of range "
                f"({slide_range_start}-{total})"
            )

    staged = _staging_path(".pdf")
    try:
        _save_pdf(pres, staged)
        if slide_range_start is not None and slide_range_end is not None:
            # PowerPoint for Mac cannot export a range, so the full deck is
            # exported and the wanted pages are copied into a new document.
            # Quartz is already here for slide images, and this keeps the tool
            # behaving the same on both platforms.
            _write_pdf_range(staged, abs_path, slide_range_start, slide_range_end)
        else:
            shutil.move(staged, abs_path)
    finally:
        if os.path.exists(staged):
            os.remove(staged)

    # The size as well as the file, the same pair `_save_pdf` checks. A PDF
    # context that wrote no page still leaves a file behind, so existence on
    # its own is not evidence.
    if not os.path.exists(abs_path) or os.path.getsize(abs_path) == 0:
        raise RuntimeError(
            "PowerPoint reported success but wrote no PDF. Check the output "
            "directory, and that the presentation has been saved to a file."
        )

    return {
        "success": True,
        "file_path": abs_path,
        "slide_range_start": slide_range_start,
        "slide_range_end": slide_range_end,
        "total_slides": total,
    }


def _write_pdf_range(source: str, destination: str, start: int, end: int) -> None:
    """Copy a page range out of one PDF into another, and count what landed.

    ``CGPDFDocumentGetPage`` answers None for a page it will not give, and
    drawing None draws nothing, so a range that quietly lost pages used to come
    back as a file that exists and is short. Every page is checked as it is
    taken, and the finished document is opened again and counted.
    """
    try:
        import Quartz
        from CoreFoundation import (
            CFURLCreateFromFileSystemRepresentation,
            kCFAllocatorDefault,
        )
    except ImportError as exc:  # pragma: no cover - depends on the install
        raise RuntimeError(_QUARTZ_HINT) from exc

    def _url(path):
        encoded = path.encode()
        return CFURLCreateFromFileSystemRepresentation(
            kCFAllocatorDefault, encoded, len(encoded), False
        )

    document = Quartz.CGPDFDocumentCreateWithURL(_url(source))
    if document is None:
        raise RuntimeError("The exported PDF could not be read back.")

    context = Quartz.CGPDFContextCreateWithURL(_url(destination), None, None)
    try:
        for number in range(start, end + 1):
            page = Quartz.CGPDFDocumentGetPage(document, number)
            if page is None:
                raise RuntimeError(
                    f"Slide {number} is missing from the deck's PDF export, so "
                    "the range could not be written."
                )
            box = Quartz.CGPDFPageGetBoxRect(page, Quartz.kCGPDFMediaBox)
            Quartz.CGContextBeginPage(context, box)
            Quartz.CGContextDrawPDFPage(context, page)
            Quartz.CGContextEndPage(context)
    finally:
        Quartz.CGPDFContextClose(context)

    wanted = end - start + 1
    written = Quartz.CGPDFDocumentCreateWithURL(_url(destination))
    if written is None:
        raise RuntimeError(
            "The slide range was written and the file cannot be read back as a "
            "PDF, so nothing usable came out of it."
        )
    pages = Quartz.CGPDFDocumentGetNumberOfPages(written)
    if pages != wanted:
        raise RuntimeError(
            f"The slide range should hold {wanted} page(s) and the file holds "
            f"{pages}. Quartz reported no error for the difference."
        )


def _export_images_impl(
    output_dir: str,
    format: str,
    slide_index: Optional[int],
    slide_indices: Optional[List[int]],
    from_index: Optional[int],
    to_index: Optional[int],
    width: Optional[int],
    height: Optional[int],
    file_name: Optional[str],
) -> dict:
    app = ppt._get_app_impl()
    if count(app.presentations) == 0:
        raise RuntimeError(
            "No presentation is open. "
            "Use ppt_create_presentation or ppt_open_presentation first."
        )
    pres = ppt._get_pres_impl()
    total_slides = count(pres.slides)

    fmt_key = format.lower().strip()
    # Quartz writes PNG and JPEG. The other formats the Windows side offers
    # exist only as PowerPoint export filters, which macOS does not have.
    uti = {"png": "public.png", "jpg": "public.jpeg", "jpeg": "public.jpeg"}.get(fmt_key)
    if uti is None:
        raise ValueError(
            f"Slide images on macOS are rendered with Quartz, which writes png "
            f"and jpg. '{format}' is not available here."
        )

    abs_dir = os.path.abspath(output_dir)
    os.makedirs(abs_dir, exist_ok=True)

    if slide_index is not None:
        targets = [slide_index]
    elif slide_indices is not None:
        targets = list(slide_indices)
    elif from_index is not None or to_index is not None:
        targets = list(range(from_index or 1, (to_index or total_slides) + 1))
    else:
        targets = list(range(1, total_slides + 1))

    for number in targets:
        if number < 1 or number > total_slides:
            raise ValueError(f"Slide index {number} out of range (1-{total_slides})")

    staged_pdf = _staging_path(".pdf")
    exported = []
    try:
        _save_pdf(pres, staged_pdf)
        rendered = _render_pdf_pages(
            staged_pdf, targets, width, height, uti, "." + fmt_key
        )
        for (number, temp_image, out_w, out_h) in rendered:
            if file_name and len(targets) == 1:
                name = file_name
            elif file_name:
                stem, ext = os.path.splitext(file_name)
                name = f"{stem}_{number}{ext or '.' + fmt_key}"
            else:
                name = f"Slide{number}.{fmt_key}"
            final = os.path.join(abs_dir, name)
            shutil.move(temp_image, final)
            # The size as well as the file, the same pair `_save_pdf` checks.
            # An empty file is what a move onto a full disk leaves behind, and
            # it reads as a written image to anything checking existence alone.
            if not os.path.exists(final) or os.path.getsize(final) == 0:
                raise RuntimeError(
                    f"Slide {number} was not written to {final}, or arrived "
                    "there empty."
                )
            exported.append({
                "slide_index": number,
                "file_path": final,
                "width": out_w,
                "height": out_h,
            })
    finally:
        if os.path.exists(staged_pdf):
            os.remove(staged_pdf)

    return {
        "success": True,
        "output_dir": abs_dir,
        "format": fmt_key,
        "count": len(exported),
        "files": exported,
        # Worth saying once, because the route is not the one Windows takes and
        # it explains why the images are vector sharp rather than screen sized.
        "note": (
            "Rendered from the deck's PDF export with Quartz, because "
            "PowerPoint for Mac has no slide export command."
        ),
    }


def _copy_to_clipboard_impl(*args, **kwargs) -> dict:
    """Refuse, because this is built on the Win32 clipboard API.

    The Windows implementation talks to user32 and kernel32 directly to put a
    DIB or an HDROP on the clipboard. macOS has an equivalent in NSPasteboard,
    but nothing of the existing code carries over, so this is a rewrite rather
    than a translation and it is not part of this port.
    """
    return {
        "error": "ppt_copy_to_clipboard is not available on macOS",
        "reason": (
            "The Windows implementation writes DIB and HDROP formats through "
            "the Win32 clipboard API, which has no counterpart here. macOS "
            "would need an NSPasteboard implementation written from scratch."
        ),
        "platform": "macOS",
        "alternatives": ["ppt_export_images", "ppt_export_pdf"],
    }


def render_slide_png(slide_index: int, width: Optional[int] = None) -> bytes:
    """Return one slide as PNG bytes, for ppt_get_slide_preview.

    Lives here rather than in server.py so the whole PDF and Quartz route sits
    in one file.
    """
    pres = ppt._get_pres_impl()
    total = count(pres.slides)
    if slide_index < 1 or slide_index > total:
        raise ValueError(f"Slide index {slide_index} out of range (1-{total})")

    staged_pdf = _staging_path(".pdf")
    png_path = None
    try:
        _save_pdf(pres, staged_pdf)
        (_, png_path, _, _) = _render_pdf_pages(staged_pdf, [slide_index], width)[0]
        with open(png_path, "rb") as handle:
            return handle.read()
    finally:
        for path in (staged_pdf, png_path):
            if path and os.path.exists(path):
                os.remove(path)
