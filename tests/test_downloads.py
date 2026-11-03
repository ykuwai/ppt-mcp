"""Downloads have a time limit and say where they could not get to.

ppt_search_icons, ppt_add_svg_icon and ppt_add_picture_from_url opened
their URL with no timeout, so on a network that drops requests to a host
rather than refusing them, the call hung until the operating system gave
up. The icon set and the icon list were fixed addresses, so a network that
blocks jsDelivr or Google Fonts had no way to get icons at all. Now every
download has a time limit, an unreachable host is named in the error with
the setting that moves it, and PPT_ICON_BASE_URL and PPT_ICON_METADATA_URL
point the icons at a mirror or a local copy.

No COM, no PowerPoint and no network required.
"""

from __future__ import annotations

import json
import os
import socket
import subprocess
import sys
import urllib.error
from pathlib import Path

import pytest

_src_dir = str(Path(__file__).resolve().parents[1] / "src")
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)

from utils import download  # noqa: E402


def _refusing(exc):
    def urlopen(request, timeout=None):
        raise exc
    return urlopen


class TestATimeLimit:
    def test_every_download_is_given_one(self, monkeypatch):
        seen = []
        monkeypatch.setattr(download.urllib.request, "urlopen",
                            lambda request, timeout=None: seen.append(timeout))
        download.open_url("https://example.test/a.svg", "download the icon")
        assert seen == [download.DOWNLOAD_TIMEOUT]

    def test_a_caller_can_ask_for_a_shorter_one(self, monkeypatch):
        seen = []
        monkeypatch.setattr(download.urllib.request, "urlopen",
                            lambda request, timeout=None: seen.append(timeout))
        download.open_url("https://example.test/a.svg", "check an icon", timeout=5)
        assert seen == [5]

    def test_a_head_request_stays_a_head_request(self, monkeypatch):
        seen = []
        monkeypatch.setattr(download.urllib.request, "urlopen",
                            lambda request, timeout=None: seen.append(request.get_method()))
        download.open_url("https://example.test/a.svg", "check an icon", method="HEAD")
        assert seen == ["HEAD"]


class TestWhenTheHostCannotBeReached:
    def test_a_timeout_names_the_host_and_the_setting(self, monkeypatch):
        monkeypatch.setattr(download.urllib.request, "urlopen",
                            _refusing(socket.timeout("timed out")))
        with pytest.raises(ConnectionError) as caught:
            download.open_url("https://cdn.jsdelivr.net/npm/x/a.svg", "download the icon",
                              "PPT_ICON_BASE_URL")
        said = str(caught.value)
        assert "Could not reach cdn.jsdelivr.net to download the icon" in said
        assert f"no answer within {download.DOWNLOAD_TIMEOUT:g} s" in said
        assert "set PPT_ICON_BASE_URL to a mirror or a local copy" in said

    def test_a_name_that_does_not_resolve_says_so(self, monkeypatch):
        monkeypatch.setattr(download.urllib.request, "urlopen", _refusing(
            urllib.error.URLError(socket.gaierror(8, "nodename nor servname provided"))))
        with pytest.raises(ConnectionError, match="Could not reach fonts.google.com"):
            download.open_url("https://fonts.google.com/metadata/icons",
                              "download the icon list", "PPT_ICON_METADATA_URL")

    def test_a_picture_has_no_setting_to_offer(self, monkeypatch):
        monkeypatch.setattr(download.urllib.request, "urlopen",
                            _refusing(ConnectionResetError("reset")))
        with pytest.raises(ConnectionError) as caught:
            download.open_url("https://example.test/a.png", "download the picture")
        assert "Could not reach example.test to download the picture (reset)." == str(
            caught.value)

    def test_a_missing_file_is_still_a_404(self, monkeypatch):
        missing = urllib.error.HTTPError("https://example.test/a.svg", 404, "Not Found",
                                         None, None)
        monkeypatch.setattr(download.urllib.request, "urlopen", _refusing(missing))
        with pytest.raises(urllib.error.HTTPError) as caught:
            download.open_url("https://example.test/a.svg", "download the icon")
        assert caught.value.code == 404


class TestALocalCopy:
    def test_a_file_is_read(self, tmp_path):
        icon = tmp_path / "outlined" / "home.svg"
        icon.parent.mkdir()
        icon.write_text("<svg/>")
        with download.open_url(icon.as_uri(), "download the icon") as resp:
            assert resp.read() == b"<svg/>"

    def test_a_file_that_is_not_there_is_a_404(self, tmp_path):
        with pytest.raises(urllib.error.HTTPError) as caught:
            download.open_url((tmp_path / "nope.svg").as_uri(), "download the icon")
        assert caught.value.code == 404

    def test_the_icon_check_answers_from_it(self, tmp_path, monkeypatch):
        from ppt_com import advanced_ops

        (tmp_path / "outlined").mkdir()
        (tmp_path / "outlined" / "home.svg").write_text("<svg/>")
        monkeypatch.setattr(advanced_ops, "ICON_PACKAGE_BASE", tmp_path.as_uri())
        monkeypatch.setattr(advanced_ops, "_icon_on_cdn_cache", {})
        assert advanced_ops._icon_is_on_cdn("home") is True
        assert advanced_ops._icon_is_on_cdn("auto_awesome") is False

    def test_the_icon_list_is_read_from_a_saved_copy(self, tmp_path, monkeypatch):
        from ppt_com import advanced_ops

        saved = tmp_path / "icons.json"
        # Saved as Google Fonts serves it, guard line and all.
        saved.write_text(")]}'\n" + json.dumps({"icons": [{"name": "home"}]}))
        monkeypatch.setattr(advanced_ops, "ICON_METADATA_URL", saved.as_uri())
        monkeypatch.setattr(advanced_ops, "_icon_cache", None)
        assert advanced_ops._fetch_icon_metadata() == [{"name": "home"}]

    def test_and_from_a_copy_without_the_guard_line(self, tmp_path, monkeypatch, caplog):
        from ppt_com import advanced_ops

        saved = tmp_path / "icons.json"
        saved.write_text(json.dumps({"icons": [{"name": "home"}]}, indent=2))
        monkeypatch.setattr(advanced_ops, "ICON_METADATA_URL", saved.as_uri())
        monkeypatch.setattr(advanced_ops, "_icon_cache", None)
        with caplog.at_level("INFO", logger=advanced_ops.logger.name):
            assert advanced_ops._fetch_icon_metadata() == [{"name": "home"}]
        # The log names where the list came from, not always Google Fonts.
        assert f"Read 1 icons from {saved.as_uri()}" in caplog.text


class TestTheSettings:
    def _read(self, **env):
        """The three values as a fresh interpreter reads them."""
        code = (
            "import sys; sys.path.insert(0, 'src');"
            "from ppt_com import constants; from utils import download;"
            "print(constants.ICON_PACKAGE_BASE); print(constants.ICON_METADATA_URL);"
            "print(download.DOWNLOAD_TIMEOUT)"
        )
        clean = {k: v for k, v in os.environ.items() if not k.startswith("PPT_")}
        out = subprocess.run([sys.executable, "-c", code], env={**clean, **env},
                             cwd=Path(__file__).resolve().parents[1],
                             capture_output=True, text=True, check=True)
        return out.stdout.split()

    def test_the_defaults_are_unchanged(self):
        assert self._read() == [
            "https://cdn.jsdelivr.net/npm/@material-symbols/svg-400@0.31.3",
            "https://fonts.google.com/metadata/icons",
            "15.0",
        ]

    def test_each_can_be_set(self):
        assert self._read(
            PPT_ICON_BASE_URL="https://unpkg.com/@material-symbols/svg-400@0.31.3/",
            PPT_ICON_METADATA_URL="file:///opt/icons.json",
            PPT_DOWNLOAD_TIMEOUT="4",
        ) == [
            "https://unpkg.com/@material-symbols/svg-400@0.31.3",
            "file:///opt/icons.json",
            "4.0",
        ]


    @pytest.mark.parametrize("bad", ["abc", "", "0", "-3", "nan", "inf"])
    def test_a_bad_timeout_falls_back_to_the_default(self, bad):
        assert self._read(PPT_DOWNLOAD_TIMEOUT=bad)[-1] == "15.0"


class TestTheTimeInTheMessage:
    def test_it_is_the_one_the_caller_asked_for(self, monkeypatch):
        monkeypatch.setattr(download.urllib.request, "urlopen",
                            _refusing(socket.timeout("timed out")))
        with pytest.raises(ConnectionError, match="no answer within 5 s"):
            download.open_url("https://example.test/a.svg", "check an icon", timeout=5)

    def test_an_incomplete_read_is_a_connection_error(self, monkeypatch):
        import http.client

        monkeypatch.setattr(download.urllib.request, "urlopen",
                            _refusing(http.client.IncompleteRead(b"")))
        with pytest.raises(ConnectionError, match="Could not reach example.test"):
            download.open_url("https://example.test/a.svg", "download the icon")


class TestAnIconNameIsAName:
    def _model(self, **kwargs):
        from ppt_com.advanced_ops import AddSvgIconInput

        return AddSvgIconInput(slide_index=1, **kwargs)

    @pytest.mark.parametrize("name", ["../x", "a/b", "bolt.svg", "Bolt", "a b", ""])
    def test_a_path_is_refused(self, name):
        from pydantic import ValidationError

        with pytest.raises(ValidationError):
            self._model(icon_name=name)

    def test_real_names_pass(self):
        assert self._model(icon_name="auto_awesome").icon_name == "auto_awesome"
        assert self._model(icon_name="123", style="sharp").style == "sharp"

    def test_the_style_is_one_of_the_three(self):
        from pydantic import ValidationError

        with pytest.raises(ValidationError):
            self._model(icon_name="bolt", style="../outlined")


class TestTheToolSaysSo:
    def test_a_search_on_a_network_that_blocks_google_fonts(self, monkeypatch):
        from ppt_com import advanced_ops

        monkeypatch.setattr(advanced_ops, "_icon_cache", None)
        monkeypatch.setattr(download.urllib.request, "urlopen",
                            _refusing(socket.timeout("timed out")))
        said = json.loads(advanced_ops.search_icons(
            advanced_ops.SearchIconsInput(query="home")))["error"]
        assert said.startswith("Failed to search icons: Could not reach fonts.google.com")
        assert "PPT_ICON_METADATA_URL" in said
