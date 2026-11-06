"""Trimming media with ppt_set_media_settings on Windows.

start_point went to MediaFormat.SetDisplayPictureFromPoint, which MediaFormat
does not have, so any call with a start_point failed, after the volume and
mute had already been written. end_point was never used. The trim is
MediaFormat.StartPoint and EndPoint, and PowerPoint refuses either unless the
start is before the end and the end is within the clip.

Nothing here reaches PowerPoint. The fake MediaFormat refuses what PowerPoint
refused when tried live (PowerPoint 16.0 on Windows 11, a 3 s clip).
"""

import sys

import pytest
from pydantic import ValidationError

sys.path.insert(0, "src")

from ppt_com.media import SetMediaSettingsInput  # noqa: E402

# On macOS the ppt_com modules carry the Mac implementations.
windows_only = pytest.mark.skipif(
    sys.platform != "win32", reason="drives the COM implementation"
)

SHAPE = {"slide_index": 1, "shape_name_or_index": "Clip"}


class TestArguments:
    def test_a_start_before_the_end_is_taken(self):
        inp = SetMediaSettingsInput(**SHAPE, start_point=500, end_point=2000)
        assert (inp.start_point, inp.end_point) == (500, 2000)

    def test_a_start_at_or_after_the_end_is_refused(self):
        with pytest.raises(ValidationError, match="must be before end_point"):
            SetMediaSettingsInput(**SHAPE, start_point=2000, end_point=2000)
        with pytest.raises(ValidationError, match="must be before end_point"):
            SetMediaSettingsInput(**SHAPE, start_point=2500, end_point=2000)

    def test_either_alone_is_taken(self):
        assert SetMediaSettingsInput(**SHAPE, start_point=500).end_point is None
        assert SetMediaSettingsInput(**SHAPE, end_point=500).start_point is None


# ---------------------------------------------------------------------------
# Windows
# ---------------------------------------------------------------------------
class _MediaFormat:
    """Refuses a trim the way PowerPoint does ("Invalid request. Illegal value.")."""

    def __init__(self, length=3000, start=0, end=None):
        object.__setattr__(self, "Length", length)
        object.__setattr__(self, "StartPoint", start)
        object.__setattr__(self, "EndPoint", length if end is None else end)
        object.__setattr__(self, "Volume", 0.8)
        object.__setattr__(self, "Muted", False)
        object.__setattr__(self, "FadeInDuration", 0)
        object.__setattr__(self, "FadeOutDuration", 0)

    def __setattr__(self, name, value):
        if name == "StartPoint" and not 0 <= value < self.EndPoint:
            raise RuntimeError("MediaFormat.StartPoint : Invalid request.  Illegal value.")
        if name == "EndPoint" and not self.StartPoint < value <= self.Length:
            raise RuntimeError("MediaFormat.EndPoint : Invalid request.  Illegal value.")
        object.__setattr__(self, name, value)


class _Shape:
    def __init__(self, media):
        self.Name = "Clip"
        self.MediaFormat = media


class _Slide:
    def __init__(self, shape):
        self.shape = shape


class _Deck:
    def __init__(self, shape):
        self.slide = _Slide(shape)

    def Slides(self, index):  # noqa: N802 - the COM name
        return self.slide


@pytest.fixture
def media(monkeypatch):
    import ppt_com.media as mod

    fmt = _MediaFormat()
    deck = _Deck(_Shape(fmt))
    monkeypatch.setattr(mod.ppt, "_get_app_impl", lambda *a, **kw: object())
    monkeypatch.setattr(mod.ppt, "_get_pres_impl", lambda *a, **kw: deck)
    monkeypatch.setattr(mod, "_get_shape", lambda slide, name: slide.shape)
    return fmt


def _set(**kw):
    from ppt_com.media import _set_media_settings_impl

    args = dict(volume=None, muted=None, start_point=None, end_point=None,
                fade_in=None, fade_out=None, loop=None, hide_while_not_playing=None)
    args.update(kw)
    return _set_media_settings_impl(1, "Clip", **args)


@windows_only
class TestOnWindows:
    def test_start_and_end_trim_the_clip(self, media):
        result = _set(start_point=500, end_point=2000)
        assert (media.StartPoint, media.EndPoint) == (500, 2000)
        assert result["start_point"] == 500
        assert result["end_point"] == 2000
        assert result["length"] == 3000

    def test_the_start_alone_keeps_the_end(self, media):
        _set(start_point=500)
        assert (media.StartPoint, media.EndPoint) == (500, 3000)

    def test_the_end_alone_keeps_the_start(self, media):
        _set(end_point=1200)
        assert (media.StartPoint, media.EndPoint) == (0, 1200)

    def test_a_trim_past_the_current_one_sets_the_end_first(self, media):
        object.__setattr__(media, "StartPoint", 500)
        object.__setattr__(media, "EndPoint", 1000)
        _set(start_point=1500, end_point=2500)
        assert (media.StartPoint, media.EndPoint) == (1500, 2500)

    def test_a_trim_before_the_current_one_sets_the_start_first(self, media):
        object.__setattr__(media, "StartPoint", 1500)
        object.__setattr__(media, "EndPoint", 2500)
        _set(start_point=100, end_point=200)
        assert (media.StartPoint, media.EndPoint) == (100, 200)

    def test_an_end_past_the_clip_changes_nothing(self, media):
        with pytest.raises(ValueError, match="3000 ms long"):
            _set(volume=0.3, end_point=5000)
        assert media.Volume == 0.8
        assert (media.StartPoint, media.EndPoint) == (0, 3000)

    def test_a_start_past_the_current_end_changes_nothing(self, media):
        object.__setattr__(media, "EndPoint", 1000)
        with pytest.raises(ValueError, match="start must be before the end"):
            _set(muted=True, start_point=2000)
        assert media.Muted is False
        assert media.StartPoint == 0

    def test_without_a_trim_the_answer_is_as_before(self, media):
        assert _set(volume=0.5) == {"success": True, "shape_name": "Clip"}
        assert media.Volume == 0.5
