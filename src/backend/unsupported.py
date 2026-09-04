"""Structured refusal for tools a platform genuinely cannot perform.

Some PowerPoint features have no counterpart on the other platform. Charts,
SmartArt and freeform path building have no words at all in PowerPoint for
Mac's Apple Event dictionary, so no amount of translation reaches them.

Two tempting answers are both wrong. Silently doing nothing leaves the model
believing the deck now contains something it does not, and it keeps building on
that belief. Hiding the tool leaves the model unable to see that the capability
exists at all, so it hunts for it instead of choosing something else.

So the tool stays listed and says plainly what happened, why, and what to reach
for instead. That is enough for the model to recover in one turn. Hiding buys
nothing anyway: an unlisted tool that is called still comes back as
``Unknown tool: ppt_add_chart``, which is the same wire shape with a worse
message.

One thing to know about this payload. Returned from a tool function it arrives
with ``isError: false``, a successful call whose result happens to carry an
``error`` key. That is deliberate here only because it is what every one of the
156 tools already does for every failure, and a platform refusal that behaved
differently from an ordinary error would be the odd one out. Moving the server
to ``isError: true`` is worth doing, but it is a decision about all 156 tools
rather than about this one, so it is tracked separately rather than smuggled in
with the port.
"""

import json
from typing import List, Optional

from backend import PLATFORM_NAME


def unsupported(
    tool_name: str,
    reason: str,
    alternatives: Optional[List[str]] = None,
) -> str:
    """Return the JSON body a tool sends when this platform cannot do it.

    Args:
        tool_name: The MCP tool name, e.g. ``ppt_add_chart``.
        reason: Why it cannot be done here, in one sentence, concrete enough
            that nobody re-derives it. Name the mechanism, not just the fact.
        alternatives: Tool names worth trying instead, best first. Omit rather
            than pad; a wrong suggestion costs more than no suggestion.

    Returns:
        A JSON string, the same shape every other tool returns on failure.
    """
    payload = {
        "error": f"{tool_name} is not available on {PLATFORM_NAME}",
        "reason": reason,
        "platform": PLATFORM_NAME,
    }
    if alternatives:
        payload["alternatives"] = alternatives
    return json.dumps(payload)
