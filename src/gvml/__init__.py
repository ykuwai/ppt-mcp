"""GVML clipboard packages, read and written without PowerPoint.

PowerPoint for Mac's Apple Event dictionary has no chart, no freeform builder
and no way to group shapes, but its clipboard does. A shape copied with
``copy shape`` lands on the pasteboard as ``com.microsoft.Art--GVML-ClipFormat``,
a small OPC zip holding one DrawingML ``lockedCanvas``, and ``paste object``
takes the same package back. Everything on the canvas is honoured, charts and
custom geometry and groups included, so a package this module writes is a
route to every one of them, and a package it reads is how a chart's numbers or
a freeform's points come back out.

This package is the pure half of that route. It imports nothing but the
standard library, knows nothing about PowerPoint or appscript, and runs on
Windows, which is where the CI is. ``ppt_mac/gvml_paste.py`` is the other half,
the one that puts a package on the pasteboard and drives the paste.

Two facts from the measurements behind ``docs/gvml-design.md`` shape every
module here. PowerPoint says nothing about a package it cannot use: ``paste
object`` returns no error and the slide does not change, whether the bytes are
not a zip, the XML is broken, the canvas is empty or the elements are in the
wrong order. And the element order that trips it is real: a ``graphicFrame``
written as ``nvGraphicFramePr, xfrm, graphic`` is dropped, and only
``nvGraphicFramePr, graphic, xfrm`` lands. The XML here is written from string
templates whose order was copied from packages PowerPoint itself produced, and
``tests/test_gvml.py`` holds those packages as fixtures so a change to a
template that moves an element is caught on the CI rather than by a silent
paste.

Units: OOXML measures in EMU, and this project in points. ``canvas.emu`` and
``canvas.pt`` convert; 12700 EMU is one point.
"""

from gvml.package import (  # noqa: F401
    GVML_UTI,
    Package,
    PackageError,
    build,
    read,
    validate,
)

__all__ = ["GVML_UTI", "Package", "PackageError", "build", "read", "validate"]
