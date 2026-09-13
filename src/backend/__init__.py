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


def use_mac_impls(namespace, module) -> None:
    """Swap a tool module's COM implementations for their Apple Event ones.

    Everything a tool module does apart from walking the object model is
    already platform neutral: the pydantic models, the validation, the warnings,
    the JSON it returns. Only the ``_*_impl`` functions differ, and they have
    the same names and signatures on both sides, so the port is a swap rather
    than a fork.

    Called at the bottom of a ported module::

        if IS_MACOS:
            from ppt_mac import shapes as _mac
            use_mac_impls(globals(), _mac)

    The public functions look their implementation up in module globals at call
    time, so replacing the name is enough. A module only defines the ones it
    has ported; anything it leaves out keeps the COM version, which then fails
    loudly rather than silently doing the wrong thing.
    """
    for name in dir(module):
        if name.startswith("_") and name.endswith("_impl"):
            namespace[name] = getattr(module, name)

__all__ = [
    "ppt",
    "use_mac_impls",
    "handle_com_error",
    "AUTO_DISMISS_DIALOG",
    "unsupported",
    "IS_WINDOWS",
    "IS_MACOS",
    "PLATFORM_NAME",
]
