"""Platform backend selection.

ppt-mcp drives a live PowerPoint. On Windows that means COM; on macOS it means
Apple Events. The two speak the same object model under different names, so the
split is kept here and nowhere else: every module says ``from backend import ppt``
and gets the wrapper for the platform it is running on.

The wrappers present the same lifecycle surface (``start``, ``stop``,
``execute``, ``connect``, ``get_app``, ``ensure_presentation``, and the
``_*_impl`` helpers that run on the worker thread), so calling code does not
branch. What genuinely differs is how an operation walks the object model, and
that lives in the per-platform ``_impl`` functions of each tool module.

See MACOS_PORT.md for the measurements behind this design.
"""

import sys

IS_WINDOWS = sys.platform == "win32"
IS_MACOS = sys.platform == "darwin"

PLATFORM_NAME = "Windows" if IS_WINDOWS else ("macOS" if IS_MACOS else sys.platform)

if IS_WINDOWS:
    from utils.com_wrapper import (  # noqa: F401
        AUTO_DISMISS_DIALOG,
        handle_com_error,
        ppt,
    )
elif IS_MACOS:
    from backend.mac_ae import (  # noqa: F401
        AUTO_DISMISS_DIALOG,
        handle_com_error,
        ppt,
    )
else:  # pragma: no cover - unsupported platform
    raise RuntimeError(
        f"ppt-mcp drives a live PowerPoint and supports Windows and macOS. "
        f"This is {sys.platform!r}, where PowerPoint does not run."
    )

from backend.unsupported import unsupported  # noqa: E402,F401

__all__ = [
    "ppt",
    "handle_com_error",
    "AUTO_DISMISS_DIALOG",
    "unsupported",
    "IS_WINDOWS",
    "IS_MACOS",
    "PLATFORM_NAME",
]
