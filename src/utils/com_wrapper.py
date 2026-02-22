"""COM connection lifecycle management for PowerPoint automation.

Handles CoInitialize, Dispatch, GetActiveObject, error recovery, and cleanup.
PowerPoint supports only a single running instance, so this provides
singleton-like access to the Application COM object.
"""

import gc
import logging
import threading
import time
from concurrent.futures import Future
from queue import Queue
from typing import Any, Callable, Optional

import pythoncom
import pywintypes
import win32com.client

logger = logging.getLogger(__name__)

# HRESULTs that indicate PowerPoint is temporarily busy (e.g. modal dialog open).
# RPC_E_CALL_REJECTED (0x80010001): server rejected the call outright.
# RPC_E_SERVERCALL_RETRYLATER (0x8001010A): server explicitly says retry later.
# Both mean the call was never started, so retrying is always safe.
_BUSY_HRESULTS = frozenset({-2147418111, -2147417846})
_RETRY_MAX = 5       # maximum number of retries
_RETRY_INTERVAL = 3  # seconds between retries


def _try_dismiss_ppt_dialog() -> None:
    """Send ESC to the PowerPoint window to dismiss any open modal dialog.

    Called once on the first RPC_E_CALL_REJECTED so the next retry can
    succeed without waiting for the user to notice.  ESC is safe: it cancels
    without committing, so no destructive side-effects occur.

    Implementation notes:
    - Uses win32gui (part of pywin32) to find the PowerPoint main window by
      class name "PPTFrameClass", then SetForegroundWindow + WScript.Shell
      SendKeys to deliver the keystroke reliably.
    - All errors are swallowed — this is best-effort only.
    """
    try:
        import win32gui  # part of pywin32, already a project dependency
        hwnd = win32gui.FindWindow("PPTFrameClass", None)
        if not hwnd:
            logger.debug("_try_dismiss_ppt_dialog: PPTFrameClass window not found")
            return
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.15)  # brief pause for focus to settle before SendKeys
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys("{ESCAPE}")
        logger.info("Sent ESC to PowerPoint to dismiss open dialog")
    except Exception as exc:
        logger.debug("_try_dismiss_ppt_dialog failed (ignored): %s", exc)


class PowerPointCOMWrapper:
    """Manages the lifecycle of a PowerPoint COM Application object.

    All COM operations are routed through a dedicated STA thread to ensure
    thread safety. The MCP server (which runs async) calls methods on this
    wrapper, which internally queues operations to the COM thread.
    """

    def __init__(self):
        self._app = None
        self._com_thread: Optional[threading.Thread] = None
        self._queue: Queue = Queue()
        self._running = False
        self._target_pres_full_name: Optional[str] = None  # session-level target (FullName for uniqueness)

    def start(self) -> None:
        """Start the COM worker thread."""
        if self._running:
            return
        self._running = True
        self._com_thread = threading.Thread(
            target=self._com_worker, daemon=True, name="COM-Worker"
        )
        self._com_thread.start()
        logger.info("COM worker thread started")

    def stop(self) -> None:
        """Stop the COM worker thread and clean up."""
        if not self._running:
            return
        self._running = False
        # Send a sentinel to unblock the worker
        self._queue.put(None)
        if self._com_thread and self._com_thread.is_alive():
            self._com_thread.join(timeout=5.0)
        logger.info("COM worker thread stopped")

    def _com_worker(self) -> None:
        """Worker thread that processes COM operations in an STA apartment."""
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        try:
            while self._running:
                item = self._queue.get()
                if item is None:
                    break
                func, args, kwargs, future = item
                for attempt in range(_RETRY_MAX):
                    try:
                        result = func(*args, **kwargs)
                        future.set_result(result)
                        break
                    except pywintypes.com_error as e:
                        if e.hresult in _BUSY_HRESULTS and attempt < _RETRY_MAX - 1:
                            logger.warning(
                                "PowerPoint is busy (modal dialog open?). "
                                "Retrying in %ds... (%d/%d)",
                                _RETRY_INTERVAL, attempt + 1, _RETRY_MAX - 1,
                            )
                            if attempt == 0:
                                # On the very first failure, try to dismiss the
                                # blocking dialog automatically via ESC so the
                                # next retry likely succeeds immediately.
                                _try_dismiss_ppt_dialog()
                            time.sleep(_RETRY_INTERVAL)
                        else:
                            future.set_exception(e)
                            break
                    except Exception as e:
                        future.set_exception(e)
                        break
        finally:
            self._cleanup_com()
            pythoncom.CoUninitialize()

    def execute(self, func: Callable, *args: Any, **kwargs: Any) -> Any:
        """Execute a function on the COM thread and return its result.

        This is the main entry point for all COM operations from async code.
        It queues the operation and blocks until completion.

        Args:
            func: The function to execute on the COM thread
            *args, **kwargs: Arguments to pass to the function

        Returns:
            The return value of func

        Raises:
            Any exception raised by func
        """
        future: Future = Future()
        self._queue.put((func, args, kwargs, future))
        return future.result(timeout=30.0)

    def connect(self, visible: Optional[bool] = None) -> Any:
        """Connect to PowerPoint (runs on COM thread).

        Args:
            visible: If True, make PowerPoint visible. If False, headless mode.
                    If None, don't change visibility (keep current state).

        Returns:
            PowerPoint.Application COM object
        """
        return self.execute(self._connect_impl, visible)

    def _connect_impl(self, visible: Optional[bool] = None) -> Any:
        """Internal: connect to PowerPoint on the COM thread."""
        if self._app is not None:
            try:
                _ = self._app.Name
                if visible is not None:
                    self._app.Visible = visible
                return self._app
            except (pywintypes.com_error, AttributeError):
                logger.warning("Stale COM reference, reconnecting...")
                self._app = None

        # Try existing instance first
        try:
            self._app = win32com.client.GetActiveObject("PowerPoint.Application")
            logger.info("Connected to existing PowerPoint instance")
        except pywintypes.com_error:
            try:
                self._app = win32com.client.Dispatch("PowerPoint.Application")
                logger.info("Created new PowerPoint instance via Dispatch")
            except pywintypes.com_error as e:
                raise ConnectionError(
                    f"Failed to connect to PowerPoint. Is it installed? Error: {e.strerror}"
                ) from e

        if visible is not None:
            self._app.Visible = visible
        elif not self._app.Visible:
            # Default: make visible if launching new
            self._app.Visible = True

        return self._app

    def get_app(self) -> Any:
        """Get the Application object, reconnecting if needed."""
        return self.execute(self._get_app_impl)

    def _get_app_impl(self) -> Any:
        """Internal: get app on COM thread."""
        if self._app is None:
            return self._connect_impl()
        try:
            _ = self._app.Name
            return self._app
        except (pywintypes.com_error, AttributeError):
            logger.warning("COM connection lost, reconnecting...")
            self._app = None
            return self._connect_impl()

    def _get_pres_impl(self) -> Any:
        """Internal: get target presentation on COM thread.

        Returns the session-level target presentation if one has been set via
        _set_target_pres_impl and the file is still open.  Also activates its
        window so subsequent goto_slide / ActiveWindow calls use the right window.
        Falls back to ActivePresentation when no target is set or when the
        target was closed.
        """
        app = self._get_app_impl()
        if self._target_pres_full_name:
            for i in range(1, app.Presentations.Count + 1):
                try:
                    p = app.Presentations(i)
                    if p.FullName == self._target_pres_full_name:
                        # Ensure this presentation's window is active so
                        # goto_slide / app.ActiveWindow operate on the right deck.
                        try:
                            p.Windows(1).Activate()
                        except Exception:
                            pass
                        return p
                except Exception:
                    pass
            # Target was closed since last activation — clear and fall back
            logger.warning(
                "Target presentation '%s' is no longer open; "
                "falling back to ActivePresentation",
                self._target_pres_full_name,
            )
            self._target_pres_full_name = None
        return app.ActivePresentation

    def _set_target_pres_impl(self, name_or_index) -> dict:
        """Internal: set session-level target presentation on COM thread."""
        app = self._get_app_impl()
        if app.Presentations.Count == 0:
            raise RuntimeError("No presentation is open in PowerPoint.")

        pres = None
        if isinstance(name_or_index, int):
            if name_or_index < 1 or name_or_index > app.Presentations.Count:
                raise ValueError(
                    f"Presentation index {name_or_index} out of range "
                    f"(1-{app.Presentations.Count})"
                )
            pres = app.Presentations(name_or_index)
        else:
            name_lower = name_or_index.lower()
            matches = []
            for i in range(1, app.Presentations.Count + 1):
                p = app.Presentations(i)
                if p.Name.lower() == name_lower or p.FullName.lower() == name_lower:
                    matches.append(p)
            if len(matches) == 0:
                open_names = [
                    app.Presentations(i).Name
                    for i in range(1, app.Presentations.Count + 1)
                ]
                raise ValueError(
                    f"Presentation '{name_or_index}' not found. "
                    f"Open presentations: {open_names}"
                )
            if len(matches) > 1:
                raise ValueError(
                    f"Multiple presentations match '{name_or_index}': "
                    f"{[p.Name for p in matches]}. Use a more specific name."
                )
            pres = matches[0]

        # Bring the presentation's window to the front
        try:
            pres.Windows(1).Activate()
        except Exception as e:
            logger.warning("Could not activate presentation window: %s", e)

        # Store FullName (includes path) to uniquely identify the presentation
        # even if another file with the same basename is later opened.
        self._target_pres_full_name = pres.FullName
        index = None
        for i in range(1, app.Presentations.Count + 1):
            if app.Presentations(i).FullName == pres.FullName:
                index = i
                break
        return {
            "success": True,
            "name": pres.Name,
            "full_name": pres.FullName,
            "index": index,
        }

    def ensure_presentation(self) -> Any:
        """Ensure at least one presentation is open, return the active one."""
        return self.execute(self._ensure_presentation_impl)

    def _ensure_presentation_impl(self) -> Any:
        """Internal: ensure presentation on COM thread."""
        app = self._get_app_impl()
        if app.Presentations.Count == 0:
            raise RuntimeError(
                "No presentation is open in PowerPoint. "
                "Use ppt_create_presentation or ppt_open_presentation first."
            )
        return app.ActivePresentation

    def _cleanup_com(self) -> None:
        """Release COM references."""
        if self._app is not None:
            try:
                # Don't quit PowerPoint - the user may be using it
                self._app = None
            except Exception:
                pass
        gc.collect()


def handle_com_error(e: pywintypes.com_error) -> dict:
    """Parse a COM error into a structured dict for error responses."""
    result = {
        "hresult": e.hresult,
        "message": str(e.strerror) if e.strerror else "Unknown COM error",
        "source": None,
        "description": None,
    }
    if e.excepinfo:
        result["source"] = e.excepinfo[1] if len(e.excepinfo) > 1 else None
        result["description"] = e.excepinfo[2] if len(e.excepinfo) > 2 else None
    return result


# Global singleton instance
ppt = PowerPointCOMWrapper()
