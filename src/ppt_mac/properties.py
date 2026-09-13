"""Document property tools, on Apple Events.

Mirrors ``ppt_com/properties.py``. Same function names, same signatures, same
returned shapes; what differs is the walk through PowerPoint's object model.

Three things about document properties on this side are worth knowing before
reading on.

**The collection is a mixed one, so it is counted and then indexed.**
``custom document property`` inherits from ``document property``, which makes
``document properties`` exactly the collection MACOS_PORT section 5.1 warns
about, where PowerPoint answers with references addressed by subclass and the
ones addressed as the other subclass do not resolve. So ``positional`` builds
every reference here and each name is read back through the reference this
module built, rather than asking the collection for all its names at once.

**Names are matched without case.** Windows asks
``BuiltInDocumentProperties("Last Author")`` and gets an answer. PowerPoint for
Mac spells some of the same properties differently, ``Last author`` where
Windows writes ``Last Author``, so the lookup folds case. The keys this module
answers with are still the Windows spellings, because a deck read on a Mac has
to read the same as one read on Windows.

**A property the deck does not carry cannot be created.** Windows writes to a
name whether or not the document has it yet. The only route to a new element
here is the Standard Suite ``make``, which has never been tried against
``document property``, so a name that is not there is reported in ``warnings``
rather than invented.
"""

import logging

from appscript.reference import CommandError

from backend.mac_ae import is_missing, positional, ppt
from backend.unsupported import refusal as _refusal

logger = logging.getLogger(__name__)


# The reason both tools give when the deck answers with no properties at all.
# Written once because it is the same reason each time, and it is a refusal
# rather than a page of nulls because a deck always carries built in
# properties, so an empty answer means the collection could not be read and not
# that the fields are blank.
_UNREADABLE = (
    "PowerPoint answered with no document properties at all. Every deck "
    "carries the built in ones, so an empty collection means it could not be "
    "read rather than that the fields are empty, and reporting each of them as "
    "null would read as an answer."
)


def _property_index(pres) -> dict:
    """Map each document property's lowercased name to its own reference.

    Built by counting and then indexing, never by asking the collection for
    every name in one event. See the module docstring for why a mixed
    collection cannot be trusted to answer that way.
    """
    index = {}
    for ref in positional(pres.document_properties):
        try:
            name = ref.name()
        except CommandError:
            # One property that will not answer is not a reason to lose the
            # rest, and the caller sees the gap as a null for that field.
            logger.warning("A document property would not give its name", exc_info=True)
            continue
        if not isinstance(name, str) or not name.strip():
            continue
        index[name.strip().lower()] = ref
    return index


# ---------------------------------------------------------------------------
# Apple Event implementation functions
# ---------------------------------------------------------------------------
def _set_properties_impl(title, author, subject, keywords, comments, category, company):
    ppt._get_app_impl()
    pres = ppt._get_pres_impl()

    index = _property_index(pres)
    if not index:
        return _refusal("ppt_set_properties", _UNREADABLE)

    # The same map the Windows implementation builds, in the same order, so
    # both platforms report the same names in the same sequence.
    field_map = {
        "Title": title,
        "Author": author,
        "Subject": subject,
        "Keywords": keywords,
        "Comments": comments,
        "Category": category,
        "Company": company,
    }

    set_names = []
    warnings = []
    for prop_name, value in field_map.items():
        if value is None:
            continue
        ref = index.get(prop_name.lower())
        if ref is None:
            warnings.append(
                f"This deck has no document property called {prop_name}, and "
                "PowerPoint for Mac offers no way to add one, so it was left "
                "alone."
            )
            continue

        ref.value.set(value)

        # Nothing is trusted because it did not raise. A property that reports
        # a clean write and keeps its old text is the silent no-op of
        # MACOS_PORT section 5, and it is counted as unset rather than set.
        try:
            written = ref.value()
        except CommandError:
            written = None
        if not isinstance(written, str) or written != value:
            warnings.append(
                f"{prop_name} still reads {written!r} after the write, so "
                "PowerPoint accepted the value and did not keep it."
            )
            continue
        set_names.append(prop_name)

    result = {
        "success": True,
        "properties_set": len(set_names),
        "set_names": set_names,
    }
    if warnings:
        result["warnings"] = warnings
    return result


def _get_properties_impl():
    # Imported lazily. ppt_com/properties.py imports this module at the bottom
    # of its own file, so importing it back at module scope would let an
    # "import ppt_mac.properties first" ordering run that swap block against a
    # module that has defined nothing yet, and the swap would silently not
    # happen. By call time both modules are fully loaded.
    from ppt_com.properties import READABLE_PROPERTIES

    ppt._get_app_impl()
    pres = ppt._get_pres_impl()

    index = _property_index(pres)
    if not index:
        return _refusal("ppt_get_properties", _UNREADABLE)

    result = {}
    for prop_name in READABLE_PROPERTIES:
        ref = index.get(prop_name.lower())
        if ref is None:
            result[prop_name] = None
            continue
        try:
            value = ref.value()
        except CommandError:
            result[prop_name] = None
            continue

        if is_missing(value) or value == "":
            # An empty string is what PowerPoint answers for a property nobody
            # has filled in, and the Windows tool documents that as null, so
            # both platforms say null rather than one of them saying "".
            result[prop_name] = None
        elif hasattr(value, "strftime"):
            # `value` is declared as text here, so a date usually arrives
            # already formatted. appscript still turns a real date descriptor
            # into a datetime, and both platforms have to answer the same
            # string, so that case is formatted the way Windows formats it.
            result[prop_name] = value.strftime("%Y-%m-%d %H:%M:%S")
        elif isinstance(value, (str, int, float, bool)):
            result[prop_name] = value
        else:
            result[prop_name] = str(value)

    return {
        "success": True,
        "properties": result,
    }
