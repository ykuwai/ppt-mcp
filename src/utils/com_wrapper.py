"""COM connection lifecycle management for PowerPoint automation.

Handles CoInitialize, Dispatch, GetActiveObject, error recovery, and cleanup.
PowerPoint supports only a single running instance, so this provides
singleton-like access to the Application COM object.
"""

import gc
import logging
import threading
from concurrent.futures import Future
from queue import Queue
from typing import Any, Callable, Optional

import pythoncom
import pywintypes
import win32com.client

logger = logging.getLogger(__name__)


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
                try:
                    result = func(*args, **kwargs)
                    future.set_result(result)
                except Exception as e:
                    future.set_exception(e)
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
