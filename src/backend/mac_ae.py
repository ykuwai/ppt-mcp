"""Apple Event connection lifecycle for PowerPoint on macOS.

The counterpart of ``utils/com_wrapper.py``. PowerPoint for Mac exposes the
same object model as the Windows VBA one, with spaces in the names, over Apple
Events. This module owns the connection to it and absorbs every quirk of that
channel so the tool modules never have to know about them.

Two things are worth knowing before reading further.

**The terminology argument is mandatory.** ``app(id=..., terms='sdef')`` is what
makes any of this work. Without it appscript asks PowerPoint for terminology the
old AETE way, gets nothing, silently falls back to fourteen built-in words, and
every attribute access raises ``AttributeError``. It is undocumented and it is
almost certainly why the received wisdom says Python cannot drive Office on a
Mac.

**Nothing is trusted because it did not raise.** PowerPoint has several
operations that report success and do nothing. Callers verify; see
``MACOS_PORT.md`` section 5 for the list.

Everything runs on one worker thread, as on Windows. Apple Events do not need
an STA apartment, but PowerPoint wedges under concurrent access, and keeping the
same shape means the tool modules see one lifecycle surface on both platforms.
"""

import logging
import os
import shutil
import subprocess
import threading
import time
from concurrent.futures import Future
from concurrent.futures import TimeoutError as FutureTimeout
from queue import Queue
from typing import Any, Callable, Optional

from appscript import app, its, k, mactypes  # noqa: F401  (re-exported for tools)
from appscript.reference import CommandError, Reference

# Only for the cancellation contract. `utils.com_wrapper` imports nothing from
# this package, so this direction is safe, and the module is importable on any
# platform because of the guard at its top (#185).
from utils.com_wrapper import pending_com_futures

logger = logging.getLogger(__name__)

BUNDLE_ID = "com.microsoft.Powerpoint"

# Apple Event errors worth naming. Everything else is passed through as is.
AE_NO_SUCH_OBJECT = -1728      # the reference does not resolve right now
AE_NOT_HANDLED = -1708         # PowerPoint does not implement this verb
AE_TIMED_OUT = -1712           # PowerPoint did not answer in time
AE_CONNECTION_INVALID = -609   # PowerPoint died mid call
AE_APP_NOT_RUNNING = -600      # PowerPoint is not there at all
AE_NOT_AUTHORISED = -1743      # the user declined the Automation prompt
AE_PARAMETER = -50             # the command was refused for what it was given

# Errors that mean the call never landed, so retrying is safe. This is the
# Apple Event analogue of _BUSY_HRESULTS on the Windows side.
_RETRYABLE = frozenset({AE_TIMED_OUT, AE_CONNECTION_INVALID, AE_APP_NOT_RUNNING})
_RETRY_MAX = 2         # total attempts = _RETRY_MAX + 1
_RETRY_INTERVAL = 2    # seconds between retries

# Per call ceiling. PowerPoint wedges often enough that appscript's own 60
# second default turns a stuck app into a stuck server; twenty seconds is long
# enough for a real save and short enough to surface as an error the model can
# act on.
DEFAULT_TIMEOUT = int(os.getenv("PPT_AE_TIMEOUT", "20"))

# What one call is allowed, once it starts. The worker retries a call that
# never landed, so this has to cover every attempt and the pauses between them.
_CALL_BUDGET = DEFAULT_TIMEOUT * (_RETRY_MAX + 1) + _RETRY_INTERVAL * _RETRY_MAX + 5

# How long a call will queue behind the ones already in front of it before it
# gives up. Only one thing at a time reaches PowerPoint, so a caller that fires
# several tools at once puts the rest in line here.
#
# It has to be more than one whole call, not less. A single call in front that
# times out and is retried spends the full budget, so anything shorter would
# take back every call behind a slow one that was going to recover, which is
# the same wrong answer as before wearing better wording.
_QUEUE_WAIT = _CALL_BUDGET * 2

# The Windows wrapper's ESC-the-dialog escape hatch has no Apple Event
# equivalent, so the flag exists only to keep the two module surfaces identical.
AUTO_DISMISS_DIALOG: bool = False

# PowerPoint is sandboxed and can only write where it has a grant. Exporting to
# a directory it has not written to before blocks for tens of seconds and then
# kills the application. Its own container is always writable, and an
# unsandboxed Python process can read files back out of it freely, so exports
# stage here and get moved afterwards.
EXPORT_STAGING_DIR = os.path.expanduser(
    "~/Library/Containers/com.microsoft.Powerpoint/Data/Documents"
)


class AppleEventError(RuntimeError):
    """An Apple Event failure carrying its OSError number."""

    def __init__(self, message: str, errornumber: Optional[int] = None):
        super().__init__(message)
        self.errornumber = errornumber


def error_number(exc: BaseException) -> Optional[int]:
    """Return the Apple Event error number of an exception, if it has one."""
    return getattr(exc, "errornumber", None)


def is_missing(value: Any) -> bool:
    """True when PowerPoint answered with ``missing value``.

    It uses that in place of an empty result far more often than the dictionary
    suggests, including for collection counts and for theme colour schemes.

    Compared by equality, not identity. ``k.missing_value`` builds a fresh
    Keyword on every access, so ``is`` silently never matches.
    """
    return value is None or value == k.missing_value


def elements(ref: Reference) -> list:
    """Return the elements of a collection reference as a list.

    PowerPoint raises -1728 for an empty collection instead of returning an
    empty list, so an empty result and a broken reference look identical from
    the outside. Both are reported as empty here, which is what every caller
    wants.
    """
    try:
        value = ref.get()
    except CommandError as exc:
        if error_number(exc) in (AE_NO_SUCH_OBJECT, AE_NOT_HANDLED):
            return []
        raise
    if isinstance(value, list):
        return value
    if is_missing(value):
        return []
    return [value]


def count(ref: Reference) -> int:
    """Return how many elements a collection has.

    Counts what ``get`` returns, because asking a collection reference to count
    itself answers -1708. Asking its container instead does work, which is what
    ``count_of`` is for, and it is the safer of the two.
    """
    return len(elements(ref))


def positional(collection) -> list:
    """Return one positional reference per element of a collection.

    Asking PowerPoint for a collection hands back references addressed by
    subclass, so a slide holding a text box and an autoshape answers with
    ``text_boxes[1]`` and ``shapes[2]``, and the second does not resolve.
    Counting and then indexing avoids it at the cost of one Apple Event.
    """
    total = count(collection)
    return [collection[i] for i in range(1, total + 1)]


def shapes_of(container) -> list:
    """Return one positional reference per shape, in z order.

    Not ``elements(container.shapes)``. Asking PowerPoint for the shapes of a
    slide hands back references addressed by subclass, so a text box comes back
    as ``text_boxes[1]`` while the autoshape beside it comes back as
    ``shapes[2]``, and the second of those does not resolve. Every use of it
    fails with -1728 the moment a slide holds more than one kind of shape.

    Counting them and addressing each one as ``shapes[i]`` avoids the whole
    problem, and costs one extra Apple Event.
    """
    return positional(container.shapes)


def stage_into_container(path: str) -> str:
    """Copy a file into PowerPoint's container and return the copy's path.

    Every path this project hands PowerPoint goes through here first. Two
    things make that necessary rather than tidy. PowerPoint is sandboxed, so a
    location it has no grant for either stalls until the call times out or
    makes macOS ask the user to grant access, and a user who inserts ten
    pictures from ten folders is asked ten times. Its own container is the one
    place it never has to ask about.

    The caller is responsible for removing the copy once PowerPoint has read
    it, where the thing being read is embedded rather than linked.
    """
    os.makedirs(EXPORT_STAGING_DIR, exist_ok=True)
    staged = os.path.join(EXPORT_STAGING_DIR, os.path.basename(path))
    if os.path.abspath(staged) != os.path.abspath(path):
        shutil.copy2(path, staged)
    return staged


def count_of(container, each) -> int:
    """Ask a container how many of a class it holds, without materialising any.

    The difference between this and asking the collection itself is not
    cosmetic, it decides whether PowerPoint survives.
    ``sequence.effects.count()`` asks the effects collection how many elements
    it has, and kills PowerPoint with -609 exactly as ``.get()`` on it does.
    ``sequence.count(each=k.effect)`` asks the sequence how many effects it
    holds, answers in one round trip, and is safe. Both were run against the
    same slide.

    So where a collection is known to be dangerous, this is the cheap way to
    count it and ``probe_count`` is the fallback for anything that will not
    answer this either.
    """
    return container.count(each=each)


# How far probe_count will walk before deciding a collection is unreasonably
# large. Nothing it is used for has thousands of members, and the point is to
# fail rather than hang if PowerPoint starts answering strangely.
_PROBE_CEILING = 500


def probe_count(collection) -> int:
    """Count a collection by walking it, without ever materialising it.

    For collections where asking for the whole thing is not safe. A slide's
    animation effects are the case that forced this: `main_sequence.effects.get()`
    kills PowerPoint with -609 and takes the open deck with it, while
    `main_sequence.effects[1]` answers perfectly. So this asks for element one,
    then two, and stops when one is not there.

    Slower than `count_of`, one Apple Event per element, so reach for that
    first and keep this for collections that will not answer it.
    """
    total = 0
    while total < _PROBE_CEILING:
        try:
            collection[total + 1].get()
        except CommandError:
            break
        total += 1
    return total


def probe_elements(collection) -> list:
    """Positional references for a collection that cannot be materialised."""
    return [collection[i] for i in range(1, probe_count(collection) + 1)]


def raw(ref: Reference, code: bytes) -> Reference:
    """Reach a property by its raw four character code.

    A handful of property names collide with enumerator names in PowerPoint's
    dictionary, and appscript resolves the collision toward the enumerator, so
    the property becomes unreachable by name. ``rotation`` (``ShRt``) and
    ``dash style`` (``LFds``) are the two that matter.

        raw(shape, b'ShRt').set(33)
        raw(shape.line_format, b'LFds').get()
    """
    return Reference(ref.AS_appdata, ref.AS_aemreference.property(code))


def osascript(script: str, timeout: float = DEFAULT_TIMEOUT) -> str:
    """Run one line of AppleScript and return its output.

    A deliberate fallback, not a general escape hatch. A few Standard Suite
    verbs (``open``, ``duplicate``) are absent from PowerPoint's dictionary and
    misbehave through appscript while working perfectly from AppleScript, so
    those go through here. Everything else stays on the fast in-process path.
    """
    try:
        proc = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired as exc:
        raise AppleEventError(
            f"PowerPoint did not respond within {timeout}s", AE_TIMED_OUT
        ) from exc
    if proc.returncode != 0:
        raise AppleEventError((proc.stderr or "").strip() or "AppleScript failed")
    return (proc.stdout or "").strip()


class _Job:
    """One piece of work waiting for the worker thread.

    It carries an extra handshake the Windows wrapper does not need. Work is
    queued from whichever thread the tool call arrived on, and a caller that
    gives up waiting used to leave its job in the queue, where the worker ran it
    minutes later against a deck that had moved on. That is where a caller got
    back an error with no message at all and then found the picture on the slide
    anyway, and where retrying it put a second copy there. So the caller and the
    worker agree on exactly one of two outcomes before any of it runs.
    """

    __slots__ = (
        "func", "args", "kwargs", "idempotent",
        "future", "started", "settled", "_lock", "_dropped",
    )

    def __init__(self, func: Callable, args: tuple, kwargs: dict,
                 idempotent: bool = False):
        self.func = func
        self.args = args
        self.kwargs = kwargs
        self.idempotent = idempotent
        self.future: Future = Future()
        self.started = threading.Event()
        # Set by whichever of claim and drop gets there first, so the caller
        # waiting in the queue is woken by either outcome. Waiting on `started`
        # alone meant a job taken back was correctly discarded and its caller
        # still sat there for the whole queue budget, holding a thread out of
        # the pool that everything else shares.
        self.settled = threading.Event()
        self._lock = threading.Lock()
        self._dropped = False

    def claim(self) -> bool:
        """Worker side. True when the job is still wanted."""
        with self._lock:
            if self._dropped:
                return False
            self.started.set()
            self.settled.set()
            return True

    @property
    def dropped(self) -> bool:
        with self._lock:
            return self._dropped

    def drop(self) -> bool:
        """Caller side. True when the job was taken back before it began."""
        with self._lock:
            if self.started.is_set():
                return False
            self._dropped = True
        self.settled.set()
        return True

    def cancel(self) -> bool:
        """`drop` under the name `utils.com_wrapper.QueuedCalls` calls.

        That class holds whatever a request has queued and calls `.cancel()`
        on each of it when the caller goes away. On Windows those are COM
        futures; here they are jobs, and taking one back before the worker
        claims it is the same promise. A job already running is not recalled,
        which is the honest outcome for an Apple Event in flight.
        """
        return self.drop()


class PowerPointAppleEventWrapper:
    """Manages the connection to PowerPoint over Apple Events.

    Mirrors ``PowerPointCOMWrapper`` so the tool modules see one lifecycle
    surface on both platforms. All work is routed through a single worker
    thread, which serialises access; PowerPoint is not reliable when several
    Apple Events are in flight at once.
    """

    def __init__(self):
        self._app = None
        self._thread: Optional[threading.Thread] = None
        self._queue: Queue = Queue()
        self._running = False
        # Session level target, held by full name so a second file with the
        # same basename cannot steal it. This matters more here than on
        # Windows: `active presentation` raises -1728 whenever PowerPoint's
        # start gallery is the frontmost window, which users hit every day.
        self._target_pres_full_name: Optional[str] = None

    # -- lifecycle ---------------------------------------------------------

    def start(self) -> None:
        """Start the worker thread. Does not touch PowerPoint."""
        if self._running:
            return
        self._running = True
        self._thread = threading.Thread(
            target=self._worker, daemon=True, name="AppleEvent-Worker"
        )
        self._thread.start()
        logger.info("Apple Event worker thread started")

    def stop(self) -> None:
        """Stop the worker thread. Leaves PowerPoint running."""
        if not self._running:
            return
        self._running = False
        self._queue.put(None)
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=5.0)
        self._app = None
        logger.info("Apple Event worker thread stopped")

    def _worker(self) -> None:
        while self._running:
            item = self._queue.get()
            if item is None:
                break
            if not item.claim():
                # The caller stopped waiting and already said so. Running this
                # now would change the deck long after the tool call that asked
                # for it reported failure.
                logger.warning(
                    "Skipping %s: the call that queued it gave up waiting",
                    getattr(item.func, "__name__", item.func),
                )
                continue
            func, args, kwargs, future = item.func, item.args, item.kwargs, item.future
            for attempt in range(_RETRY_MAX + 1):
                try:
                    future.set_result(func(*args, **kwargs))
                    break
                except CommandError as exc:
                    number = error_number(exc)
                    if number in _RETRYABLE and not item.idempotent:
                        # One impl is many Apple Events. PowerPoint may have
                        # applied the first few before it stopped answering, so
                        # running the whole thing again adds a second shape
                        # rather than recovering the first. Windows reached the
                        # same conclusion in #200 and refuses here too; only
                        # callers that say they are safe to repeat are retried.
                        future.set_exception(AppleEventError(
                            "PowerPoint stopped answering part-way through "
                            f"(Apple Event error {number}). Some of the request "
                            "may already have been applied, so it was not sent "
                            "again. Look at the slide before retrying."
                        ))
                        break
                    if number in _RETRYABLE and attempt < _RETRY_MAX:
                        logger.warning(
                            "PowerPoint did not answer (error %s). "
                            "Retrying in %ds... (%d/%d)",
                            number, _RETRY_INTERVAL, attempt + 1, _RETRY_MAX,
                        )
                        # A dead application leaves a stale reference behind.
                        if number in (AE_CONNECTION_INVALID, AE_APP_NOT_RUNNING):
                            self._app = None
                        time.sleep(_RETRY_INTERVAL)
                        continue
                    future.set_exception(self._translate(exc))
                    break
                except Exception as exc:  # noqa: BLE001 - reported to the caller
                    future.set_exception(exc)
                    break

    @staticmethod
    def _translate(exc: CommandError) -> Exception:
        """Turn an Apple Event failure into something worth reading.

        Only the cases where the raw message would send someone in the wrong
        direction are rewritten. Everything else is passed through.
        """
        number = error_number(exc)
        if number == AE_NOT_AUTHORISED:
            # Written for the model to relay, not for a developer to debug.
            # The user has to fix this themselves, so the useful thing is to
            # route them there in one step. Note whose permission it is: macOS
            # attributes automation consent to the responsible parent process,
            # so it is the terminal or editor that launched the server, not
            # Python, that appears in the list.
            return AppleEventError(
                "macOS refused permission to control PowerPoint.\n"
                "Open System Settings > Privacy & Security > Automation and "
                "allow the application that launched this server (the terminal "
                "or editor, not Python itself) to control Microsoft PowerPoint.\n"
                "To open that pane directly, run:\n"
                "  open \"x-apple.systempreferences:com.apple.preference.security"
                "?Privacy_Automation\"\n"
                "The prompt cannot be answered over ssh or with no one at the "
                "screen.",
                number,
            )
        if number == AE_CONNECTION_INVALID:
            return AppleEventError(
                "PowerPoint stopped responding and the connection was lost. It "
                "may have quit. Reopen the presentation and try again.",
                number,
            )
        if number == AE_TIMED_OUT:
            return AppleEventError(
                f"PowerPoint did not answer within {DEFAULT_TIMEOUT}s. It is "
                "usually waiting on a dialog, or being asked to write somewhere "
                "its sandbox does not allow.",
                number,
            )
        return AppleEventError(str(exc), number)

    def execute(self, func: Callable, *args: Any,
                idempotent: bool = False, **kwargs: Any) -> Any:
        """Run ``func`` on the worker thread and return its result.

        The single entry point for every operation, matching the Windows
        wrapper. Blocks until the work finishes or the worker gives up.

        Args:
            func: The work to run.
            *args: Positional arguments for func.
            idempotent: Keyword-only, and not forwarded to func. Pass True
                only when running func twice is the same as running it once,
                which is what allows the worker to retry it whole. Connecting
                is the case that qualifies; anything that edits a deck is not.
            **kwargs: Keyword arguments for func.
        """
        job = _Job(func, args, kwargs, idempotent)
        self._queue.put(job)

        # Register with whatever is watching this request, so that a caller who
        # goes away takes its queued work with it. `utils.offload` puts a
        # `QueuedCalls` here for the duration of a tool call and cancels it on
        # the way out. Without this macOS got half of #198 and #199: the event
        # loop stayed free, but a cancelled request's queue still ran, minutes
        # later, against a deck that had moved on. Nobody is watching for
        # internal callers, and then this does nothing.
        watcher = pending_com_futures.get()
        if watcher is not None:
            watcher.add(job)

        # Two waits, not one. The first is for the queue, and it is the caller's
        # to abandon; the second is for PowerPoint, and it starts only once the
        # work does. Timing both together used to charge a call for the time it
        # spent in line, which is how a handful of parallel tool calls made the
        # ones at the back fail while their work went ahead regardless.
        if not job.settled.wait(timeout=_QUEUE_WAIT):
            if job.drop():
                raise AppleEventError(
                    f"PowerPoint was still busy with an earlier request after "
                    f"{_QUEUE_WAIT}s, so this one was taken back rather than "
                    "left to run later on its own. Nothing was changed. Only "
                    "one request at a time reaches PowerPoint on macOS, so "
                    "call these tools one after another rather than several in "
                    "the same turn."
                )
            # It started while that was being decided, so wait for it properly.
        elif job.dropped:
            # Cancelled from outside while it sat in the queue. The work is
            # discarded either way; returning now is what frees this thread,
            # which is the whole reason `settled` exists.
            raise AppleEventError(
                "The request that queued this call was cancelled before "
                "PowerPoint reached it, so it was discarded. Nothing was "
                "changed."
            )

        try:
            return job.future.result(timeout=_CALL_BUDGET)
        except FutureTimeout:
            # `str()` on this one is the empty string, so letting it out reaches
            # the caller as "Failed to add picture: " with nothing after it.
            raise AppleEventError(
                f"PowerPoint did not finish this within {_CALL_BUDGET}s and the "
                "request was abandoned. It may still be working, so check the "
                "deck before asking for the same thing again.",
                AE_TIMED_OUT,
            ) from None

    # -- connection --------------------------------------------------------

    def connect(self, visible: Optional[bool] = None, allow_launch: bool = True) -> Any:
        """Connect to PowerPoint, launching it when allowed.

        Idempotent, and says so. Connecting twice is connecting once, and it
        edits nothing, so it is one of the few things the worker may safely run
        again after a retryable failure. `ppt_com/app.py` passes the same flag
        for `ppt_connect`; these two were left without it and the difference
        was an oversight rather than a distinction.
        """
        return self.execute(self._connect_impl, visible, allow_launch,
                            idempotent=True)

    def _connect_impl(
        self, visible: Optional[bool] = None, allow_launch: bool = True
    ) -> Any:
        """Internal: connect on the worker thread.

        ``visible`` has no counterpart here. PowerPoint for Mac has no headless
        mode and no ``Application.Visible``, so the flag is accepted for API
        symmetry and only used to decide whether to bring the app forward.
        """
        if self._app is not None:
            try:
                self._app.version()
                return self._app
            except CommandError:
                logger.warning("Stale Apple Event reference, reconnecting...")
                self._app = None

        candidate = app(id=BUNDLE_ID, terms="sdef")
        if not candidate.isrunning():
            if not allow_launch:
                raise ConnectionError(
                    "PowerPoint is not running. Call ppt_connect, "
                    "ppt_create_presentation, or ppt_open_presentation first."
                )
            candidate.activate()
            # Launching is not instant, and the first event against a starting
            # application fails rather than waiting.
            deadline = time.time() + 20
            while time.time() < deadline:
                try:
                    candidate.version()
                    break
                except CommandError:
                    time.sleep(0.3)
            else:
                raise ConnectionError(
                    "PowerPoint did not finish starting. Is it installed?"
                )
        elif visible:
            candidate.activate()

        self._app = candidate
        logger.info("Connected to PowerPoint over Apple Events")
        return self._app

    def get_app(self, allow_launch: bool = False) -> Any:
        """Get the application reference, reconnecting if needed.

        Idempotent for the same reason as `connect`: it reaches for the
        application and reconnects when the reference is stale, and touches no
        deck on the way.
        """
        return self.execute(self._get_app_impl, allow_launch, idempotent=True)

    def _get_app_impl(self, allow_launch: bool = False) -> Any:
        """Internal: get the application on the worker thread.

        Attach only by default. Most tools operate on an already open
        presentation and should fail fast rather than silently starting
        PowerPoint, which is the same contract as the Windows side.
        """
        if self._app is None:
            return self._connect_impl(allow_launch=allow_launch)
        try:
            self._app.version()
            return self._app
        except CommandError:
            logger.warning("Apple Event connection lost, reconnecting...")
            self._app = None
            return self._connect_impl(allow_launch=allow_launch)

    # -- presentations -----------------------------------------------------

    def _presentations(self, app_ref) -> list:
        return elements(app_ref.presentations)

    def _get_pres_impl(self) -> Any:
        """Internal: get the target presentation on the worker thread.

        Returns the session target when one is set and its file is still open,
        without bringing PowerPoint forward. Falls back to the active
        presentation, and then to the first open one, because ``active
        presentation`` raises whenever PowerPoint's start gallery is frontmost.
        """
        app_ref = self._get_app_impl()
        presentations = self._presentations(app_ref)

        if self._target_pres_full_name:
            for pres in presentations:
                try:
                    if pres.full_name() == self._target_pres_full_name:
                        # Deliberately not activated. This runs on the way into
                        # every tool call, and activating a window brings
                        # PowerPoint to the front of whatever the user is
                        # actually doing. Showing the slide being edited is
                        # what was wanted, and `goto_slide` does that by
                        # driving this deck's own window without raising it.
                        return pres
                except CommandError:
                    continue
            logger.warning(
                "Target presentation '%s' is no longer open; "
                "falling back to the active presentation",
                self._target_pres_full_name,
            )
            self._target_pres_full_name = None

        try:
            active = app_ref.active_presentation
            active.name()
            return active
        except CommandError:
            if presentations:
                return presentations[0]
            raise RuntimeError(
                "No presentation is open in PowerPoint. "
                "Use ppt_create_presentation or ppt_open_presentation first."
            ) from None

    def _set_target_pres_impl(self, name_or_index) -> dict:
        """Internal: set the session target presentation on the worker thread."""
        app_ref = self._get_app_impl()
        presentations = self._presentations(app_ref)
        if not presentations:
            raise RuntimeError("No presentation is open in PowerPoint.")

        if isinstance(name_or_index, int):
            if name_or_index < 1 or name_or_index > len(presentations):
                raise ValueError(
                    f"Presentation index {name_or_index} out of range "
                    f"(1-{len(presentations)})"
                )
            pres = presentations[name_or_index - 1]
        else:
            wanted = name_or_index.lower()
            matches = [
                p for p in presentations
                if p.name().lower() == wanted or p.full_name().lower() == wanted
            ]
            if not matches:
                open_names = [p.name() for p in presentations]
                raise ValueError(
                    f"Presentation '{name_or_index}' not found. "
                    f"Open presentations: {open_names}"
                )
            if len(matches) > 1:
                raise ValueError(
                    f"Multiple presentations match '{name_or_index}': "
                    f"{[p.name() for p in matches]}. Use a more specific name."
                )
            pres = matches[0]

        # Reported, not silently targeted. This used to activate
        # `document_windows[1]` and log whatever came back, so a deck that had
        # outlived its window became the session target anyway and every tool
        # after it failed with a bare -1728 instead. `target_window` raises
        # with the explanation, and it raises before the target is set, so a
        # refused call leaves the session pointing where it already was.
        window = target_window(pres)
        try:
            window.activate()
        except CommandError as exc:
            # A window that exists and will not come forward is a nuisance, not
            # a reason to refuse; the deck is still editable either way.
            logger.warning("Could not activate presentation window: %s", exc)

        full_name = pres.full_name()
        self._target_pres_full_name = full_name
        index = None
        for i, p in enumerate(presentations, start=1):
            try:
                if p.full_name() == full_name:
                    index = i
                    break
            except CommandError:
                continue
        return {
            "success": True,
            "name": pres.name(),
            "full_name": full_name,
            "index": index,
        }

    def ensure_presentation(self) -> Any:
        """Ensure at least one presentation is open, and return the target."""
        return self.execute(self._ensure_presentation_impl)

    def _ensure_presentation_impl(self) -> Any:
        """Internal: ensure a presentation on the worker thread."""
        app_ref = self._get_app_impl()
        if not self._presentations(app_ref):
            raise RuntimeError(
                "No presentation is open in PowerPoint. "
                "Use ppt_create_presentation or ppt_open_presentation first."
            )
        return self._get_pres_impl()

    # -- verbs PowerPoint's dictionary leaves out ---------------------------

    def open_presentation(self, path: str) -> None:
        """Open a file, through AppleScript rather than appscript.

        ``PP.open(...)`` returns None and opens nothing, because PowerPoint's
        dictionary declares no Standard Suite commands and appscript's default
        ``odoc`` event goes unanswered. The AppleScript form works, so it is
        what runs here.
        """
        escaped = path.replace("\\", "\\\\").replace('"', '\\"')
        osascript(
            'tell application "Microsoft PowerPoint" to open POSIX file "%s"' % escaped
        )


def handle_com_error(exc: BaseException) -> dict:
    """Parse a failure into the structured dict the error responses expect.

    Named for its Windows counterpart so ``ppt_com.app`` needs no branch. The
    keys match, with the Apple Event error number standing in for the HRESULT.
    """
    number = error_number(exc)
    return {
        "hresult": number,
        "message": str(exc) or "Unknown Apple Event error",
        "source": "Microsoft PowerPoint",
        "description": getattr(exc, "errormessage", None),
    }


# Global singleton, matching the Windows module.
ppt = PowerPointAppleEventWrapper()


# ---------------------------------------------------------------------------
# Walking the object model
#
# The helpers every tool module needs before it can touch anything, the slide,
# the shape on it, the presentation it belongs to, and the number Windows would
# have reported for a macOS enumerator. They live here rather than in one of
# the tool modules because a tool module that imported another one would meet
# the import cycle ``ppt_com`` already sets up, and lose.
# ---------------------------------------------------------------------------


def target_window(pres):
    """The window a deck is edited through, or a refusal naming what is wrong.

    A document can outlive its window. PowerPoint's own logs show it closing a
    window and never freeing the document behind it, and the deck then sits in
    `presentations` answering questions, holding all its slides, and invisible
    to the person at the machine. Every tool here that shows a slide reaches
    for `document_windows[1]`, which in that state raises -1728 and says
    nothing about why.

    The state is recoverable: closing the deck and opening the file again
    brings a window back, and a saved file is untouched by any of it.
    """
    if count_of(pres, k.document_window) == 0:
        raise AppleEventError(
            "This presentation is still open inside PowerPoint but has no "
            "window, so there is no editor to drive and nothing on screen. "
            "That happens when a window is closed and the document behind it "
            "is not, which PowerPoint does on its own. Its slides are intact "
            "and a saved file is untouched. Close it with "
            "ppt_close_presentation(save_changes=false) and open it again "
            "with ppt_open_presentation."
        )
    return pres.document_windows[1]


def slide_at(pres, slide_index: int):
    """Return a slide reference, checking the index first.

    An out of range element reference does not fail where it is built, it fails
    somewhere later with -1728 and no mention of the index, so the range is
    checked here where the number is still in hand.
    """
    total = count(pres.slides)
    if total == 0:
        # `elements` turns -1728 into an empty list, so "this deck has no
        # slides" and "this reference no longer reaches a deck" arrive here
        # looking identical. A caller once read "The presentation has 0 slides"
        # while PowerPoint was in the act of dying underneath it, and went
        # looking for the missing slides rather than the missing application
        # (#191). Asking the deck its own name separates the two.
        try:
            pres.name()
        except CommandError as exc:
            raise AppleEventError(
                "The presentation this call was working on can no longer be "
                f"reached (Apple Event error {error_number(exc)}). It is not "
                "an empty deck; the reference itself is dead, which happens "
                "when the file is renamed by a save, closed, or PowerPoint "
                "restarts underneath the session. Call "
                "ppt_list_presentations to see what is open now."
            ) from exc
    if slide_index < 1 or slide_index > total:
        raise ValueError(
            f"Slide index {slide_index} is out of range. "
            f"The presentation has {total} slides (1-based)."
        )
    return pres.slides[slide_index]


def shape_by_name_or_index(slide, name_or_index):
    """Find a shape on a slide by name or 1-based index.

    Built out of ``shapes_of`` rather than ``slide.shapes``, because asking
    PowerPoint for a slide's shapes hands back references addressed by
    subclass and the second of those does not resolve.
    """
    shapes = shapes_of(slide)
    if isinstance(name_or_index, int):
        if name_or_index < 1 or name_or_index > len(shapes):
            raise ValueError(
                f"Shape index {name_or_index} out of range (1-{len(shapes)})"
            )
        return shapes[name_or_index - 1]
    for shape in shapes:
        if shape.name() == name_or_index:
            return shape
    raise ValueError(f"Shape '{name_or_index}' not found on slide")


def windows_constant(table, keyword, default=None):
    """Turn a macOS enumerator back into the Windows constant it stands for.

    ``to_keyword`` goes one way, and reading a property needs the other. The
    generated tables are keyed by the Windows constant, and they are small
    enough that a scan costs less than keeping a second index in step with
    them.
    """
    for value, word in table.items():
        if word == keyword:
            return value
    return default


def full_names(app) -> list:
    """Every open presentation's full name, in one Apple Event."""
    return [str(name) for name in elements(app.presentations.full_name)]


def resolve_presentation(
    app,
    presentation_index: Optional[int] = None,
    presentation_name: Optional[str] = None,
):
    """Return a presentation by index, by name, or the session target.

    The counterpart of the helper of the same name on the Windows side, raising
    the same errors with the same wording.
    """
    if presentation_index is not None and presentation_name is not None:
        raise ValueError(
            "Specify either presentation_index or presentation_name, not both"
        )

    presentations = elements(app.presentations)

    if presentation_index is not None:
        total = len(presentations)
        if presentation_index < 1 or presentation_index > total:
            raise ValueError(
                f"Presentation index {presentation_index} out of range (1-{total})"
            )
        return presentations[presentation_index - 1]

    if presentation_name is not None:
        if not presentations:
            raise RuntimeError(
                "No presentation is open. "
                "Use ppt_create_presentation or ppt_open_presentation first."
            )
        matches = []
        available = []
        for index, pres in enumerate(presentations, start=1):
            name = pres.name()
            available.append(f"  [{index}] {name}")
            if name == presentation_name:
                matches.append((index, pres))
        if len(matches) == 1:
            return matches[0][1]
        if len(matches) > 1:
            match_list = ", ".join(
                f"[{index}] {presentation_name}" for index, _ in matches
            )
            raise ValueError(
                f"Multiple presentations match name '{presentation_name}': "
                f"{match_list}. Use presentation_index to disambiguate."
            )
        raise ValueError(
            f"No presentation named '{presentation_name}'. "
            f"Available presentations:\n" + "\n".join(available)
        )

    if not presentations:
        raise RuntimeError(
            "No presentation is open. "
            "Use ppt_create_presentation or ppt_open_presentation first."
        )
    return ppt._get_pres_impl()
