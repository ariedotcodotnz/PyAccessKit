# pyright: basic
"""Make Ctrl+C effective while Python is blocked inside a long COM call.

Python only raises ``KeyboardInterrupt`` when the main thread runs bytecode again, which does not happen
while it waits for Access. A console control handler (called by Windows on its own thread) lets the
watchdogs terminate the *owned* Access process whose call is in flight; the COM call then returns and the
interrupt surfaces immediately. The handler returns ``False`` so Python's default handling still runs.
"""

from __future__ import annotations

import logging
import threading
from collections.abc import Callable

import win32api
import win32con

__all__ = ["register_interrupt_callback", "unregister_interrupt_callback"]

logger = logging.getLogger("pyaccesskit.process")

_lock = threading.Lock()
_callbacks: set[Callable[[], None]] = set()
_installed = False


def _handler(ctrl_type: int) -> bool:
    if ctrl_type in (win32con.CTRL_C_EVENT, win32con.CTRL_BREAK_EVENT):
        with _lock:
            callbacks = list(_callbacks)
        for callback in callbacks:
            try:
                callback()
            except Exception:
                logger.exception("interrupt callback failed")
    return False


def register_interrupt_callback(callback: Callable[[], None]) -> None:
    """Call ``callback`` (from a Windows-owned thread) when the user presses Ctrl+C or Ctrl+Break."""
    global _installed  # noqa: PLW0603 - one process-wide console handler
    with _lock:
        _callbacks.add(callback)
        if not _installed:
            try:
                win32api.SetConsoleCtrlHandler(_handler, True)
                _installed = True
            except Exception as exc:
                logger.debug("console control handler unavailable: %s", exc)


def unregister_interrupt_callback(callback: Callable[[], None]) -> None:
    """Remove a callback registered with :func:`register_interrupt_callback`."""
    with _lock:
        _callbacks.discard(callback)
