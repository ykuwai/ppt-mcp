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


def test_the_tool_answers_the_image_and_then_its_scale():
    import server
    from backend import ppt

    image = png(64, 36)
    with patch.object(ppt, "execute", lambda fn, *args, **kw: (image, (32.0, 18.0))):
        result = anyio.run(server.mcp.call_tool, "ppt_get_slide_preview",
                           {"params": {"slide_index": 2}})
    blocks = server._content_of(result)
    assert [b.type for b in blocks] == ["image", "text"]
    assert blocks[0].mimeType == "image/png"
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
