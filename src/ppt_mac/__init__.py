"""Apple Event implementations of the tool internals, for macOS.

Each module here mirrors one under ``ppt_com`` and defines only the ``_*_impl``
functions, the part that actually walks PowerPoint's object model. Everything
else about a tool, its pydantic model, its validation, the JSON it returns, is
platform neutral and stays in ``ppt_com``.

The swap happens at the bottom of the ``ppt_com`` module, through
``backend.use_mac_impls``. A function not defined here keeps its COM version,
which fails loudly on macOS rather than quietly doing the wrong thing.
"""
