"""ppt_get_slide_preview says the image's scale beside it (#194).

The image of a 960 pt slide is 1920 px across, and nothing said so: a caller
measuring a chart in it took pixels for points and twice concluded the chart
was ignoring its width. The result now carries, after the image, the image's
size in pixels, the slide's in points and the pixels per point.
"""

import json
import struct
import sys
import zlib
from unittest.mock import patch

import anyio
import pytest

sys.path.insert(0, "src")


def png(width, height):
    """A real, tiny PNG header and body of the given size."""
    def chunk(kind, body):
        return (struct.pack(">I", len(body)) + kind + body
                + struct.pack(">I", zlib.crc32(kind + body)))
    rows = b"".join(b"\0" + b"\0\0\0" * width for _ in range(height))
    return (b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b""))


def test_the_size_is_read_from_the_png_header():
    import server

    assert server.png_size(png(48, 27)) == (48, 27)
    assert server.png_size(b"not a png") is None
    assert server.png_size(b"") is None


def test_the_scale_converts_points_to_pixels():
    import server

    said = server.preview_scale(3, png(1920, 1080), (960.0, 540.0))
    assert said == {"slide_index": 3, "image_width_px": 1920, "image_height_px": 1080,
                    "slide_width_pt": 960.0, "slide_height_pt": 540.0,
                    "pixels_per_point": 2.0}


def test_a_slide_size_that_could_not_be_read_leaves_out_the_scale():
    import server

    said = server.preview_scale(1, png(1920, 1080), None)
    assert "pixels_per_point" not in said and said["image_width_px"] == 1920


@pytest.mark.parametrize("points, pixels", [
    ((960.0, 540.0), (1920, 1080)),   # 16:9
    ((720.0, 540.0), (1440, 1080)),   # 4:3
    ((612.0, 792.0), (1224, 1584)),   # letter, portrait
])
def test_windows_asks_for_two_pixels_per_point(points, pixels):
    """Slide.Export with no size follows the display's DPI, 1280 px across
    for a 960 pt slide at 96 DPI; the preview asks for the Mac's 2 px/pt."""
    import server

    calls = []

    class Slide:
        def Export(self, *args):
            calls.append(args)

    class PageSetup:
        SlideWidth, SlideHeight = points

    server.export_preview_png(Slide(), PageSetup(), "out.png")
    assert calls == [("out.png", "PNG") + pixels]


def test_the_tool_answers_the_image_and_then_its_scale():
    import server
    from backend import ppt

    image = png(64, 36)
    with patch.object(ppt, "execute", lambda fn, *args, **kw: (image, (32.0, 18.0))):
        result = anyio.run(server.mcp.call_tool, "ppt_get_slide_preview",
                           {"params": {"slide_index": 2}})
    blocks = server._content_of(result)
    assert [b.type for b in blocks] == ["image", "text"]
    # 1.x calls it mimeType, 2.x mime_type.
    mime = getattr(blocks[0], "mime_type", None) or getattr(blocks[0], "mimeType", None)
    assert mime == "image/png"
    said = json.loads(blocks[1].text)
    assert said["pixels_per_point"] == 2.0
    assert (said["image_width_px"], said["slide_width_pt"]) == (64, 32.0)


@pytest.mark.skipif(sys.platform != "darwin", reason="drives the macOS render path")
@pytest.mark.parametrize("size, expected", [
    ({"width_points": 960.0, "height_points": 540.0}, 0.0667),
    (RuntimeError("no deck"), None),
])
def test_the_slide_size_is_read_beside_the_render(monkeypatch, size, expected):
    import server
    from backend import ppt
    from ppt_com import layout
    from ppt_mac import export
    from utils import navigation

    def slide_size():
        if isinstance(size, Exception):
            raise size
        return size

    monkeypatch.setattr(export, "render_slide_png", lambda index: png(64, 36))
    monkeypatch.setattr(navigation, "goto_slide", lambda app, index: None)
    monkeypatch.setattr(layout, "_get_slide_size_impl", slide_size)
    monkeypatch.setattr(ppt, "_get_app_impl", lambda: None)
    monkeypatch.setattr(ppt, "execute", lambda fn, *args, **kw: fn(*args))
    result = anyio.run(server.mcp.call_tool, "ppt_get_slide_preview",
                       {"params": {"slide_index": 1}})
    said = json.loads(server._content_of(result)[1].text)
    assert said.get("pixels_per_point") == expected
    assert said["image_width_px"] == 64


@pytest.mark.parametrize("raised, said", [
    (PermissionError("macOS is asking on screen"), "macOS is asking on screen"),
    (TimeoutError(), "TimeoutError"),
])
def test_a_preview_that_fails_answers_why(raised, said):
    """Raised, the reason reached the caller only as far as its client showed
    the SDK's "Error executing tool" text, and once not at all."""
    import server
    from backend import ppt

    def fail(fn, *args, **kw):
        raise raised

    with patch.object(ppt, "execute", fail):
        result = anyio.run(server.mcp.call_tool, "ppt_get_slide_preview",
                           {"params": {"slide_index": 4}})
    answer = json.loads(server._text_of(server._content_of(result)))
    assert answer["error"] == f"Failed to preview slide 4: {said}"
