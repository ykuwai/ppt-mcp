"""Batch formatting, on Apple Events.

Mirrors ``ppt_com/batch_apply.py``. Same function name, same signature, same
returned shape.

Almost none of this module needed porting. ``_dispatch_op`` reaches the
formatting, effect and text implementations through their modules rather than
by name, and those modules swap their own implementations at import time, so a
batch run on macOS is already calling the Apple Event versions. What is left is
the walk this file does for itself, resolving the slide and checking each shape
exists, which is COM in the original.

The other half is a bug rather than a port. The Windows loop records a step as
``"status": "success"`` whenever the call did not raise, and on macOS a tool
that cannot be honoured returns a refusal dict instead of raising. Recorded the
Windows way, a batch of seven refusals reads as seven successes, which is the
exact failure MACOS_PORT section 7.3 is written against. So the return value is
inspected here and a refusal is recorded as an error carrying its reason.

Refusing is not the only thing an operation can say. Several apply most of what
they were asked and name the rest in ``warnings``, or in ``partial`` and
``unsupported`` for the paragraph format tools. Those travel onto the step
rather than being flattened into a bare success, for the same reason.
"""

import logging

from backend.mac_ae import ppt, slide_at as _slide
from ppt_mac.shapes import _get_shape

logger = logging.getLogger(__name__)


def _describe_failure(result):
    """Return the message to record when an operation did not really succeed.

    ``None`` means the operation was honoured. Anything else is the text to put
    in the step's ``error``. Only a dict carrying an ``error`` key counts as a
    refusal; a successful operation returns a dict too, so truthiness is not
    the test.
    """
    if not isinstance(result, dict):
        return None
    if "error" not in result:
        return None
    reason = result.get("reason")
    return f"{result['error']}. {reason}" if reason else str(result["error"])


# What an operation can say about itself besides working or not. A tool that
# applied nine of its ten properties reports the tenth here, and a batch that
# threw those away would be the same lie this module exists to stop, one step
# quieter. `partial` and `unsupported` come from the paragraph format tools and
# `warnings` from most of the rest.
_CARRIED_KEYS = ("warnings", "partial", "unsupported", "note")


def _carried_over(result):
    """The parts of an operation's answer that survive into the batch's."""
    if not isinstance(result, dict):
        return {}
    return {key: result[key] for key in _CARRIED_KEYS if result.get(key)}


def _batch_apply_impl(slide_index, shapes, operations):
    """Apply several formatting operations to several shapes."""
    # Lazy import. ppt_com/batch_apply.py imports this module at the bottom of
    # its own file, so importing it back at module scope would let an
    # "import ppt_mac.batch_apply first" ordering run that swap block against a
    # module that has defined nothing yet. By call time both are fully loaded.
    from ppt_com.batch_apply import _dispatch_op

    ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = _slide(pres, slide_index)

    results = []
    for shape_id in shapes:
        # Verify the shape is there before spending events on its operations.
        try:
            _get_shape(slide, shape_id)
        except Exception as e:  # noqa: BLE001 - reported per shape, not raised
            results.append({
                "shape": str(shape_id),
                "error": str(e),
                "operations": [],
            })
            continue

        shape_results = []
        for op in operations:
            try:
                outcome = _dispatch_op(
                    slide_index, shape_id, op["tool"], op.get("params", {})
                )
            except Exception as e:  # noqa: BLE001 - recorded, not raised
                shape_results.append({
                    "tool": op["tool"],
                    "status": "error",
                    "error": str(e),
                })
                continue

            failure = _describe_failure(outcome)
            if failure is None:
                step = {"tool": op["tool"], "status": "success"}
                step.update(_carried_over(outcome))
                shape_results.append(step)
            else:
                shape_results.append({
                    "tool": op["tool"],
                    "status": "error",
                    "error": failure,
                })

        results.append({"shape": str(shape_id), "operations": shape_results})

    return {"results": results}
