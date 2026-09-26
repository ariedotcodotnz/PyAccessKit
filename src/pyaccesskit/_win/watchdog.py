# pyright: basic
"""The dialog watchdog: detects, reports and dismisses modal dialogs of *one owned* Access process.

A modal dialog inside a hidden Access blocks the current COM call forever. pywin32 releases the GIL while
it waits for the call, so this thread keeps running (spike S8). While a call is in flight it:

1. looks for visible top-level windows belonging to **our PID only** (a hidden Access has no other visible
   windows; in visible mode only standard dialog windows count);
2. records title, text and buttons, then dismisses the dialog: ``WM_CLOSE`` first, then a click on
   Cancel/No/OK (Yes/No boxes ignore ``WM_CLOSE``), and finally terminates our Access if nothing works;
3. terminates our Access when a call exceeds ``call_timeout``.

It implements :class:`~pyaccesskit._com.gateway.CallMonitor`.
"""

from __future__ import annotations

import contextlib
import logging
import threading
import time
from collections.abc import Callable, Iterator

import win32con
import win32gui
import win32process

from pyaccesskit._win.console import register_interrupt_callback, unregister_interrupt_callback
from pyaccesskit.enums import DialogPolicy
from pyaccesskit.errors import DialogInfo

__all__ = ["DialogWatchdog"]

logger = logging.getLogger("pyaccesskit.process")

BM_CLICK = 0x00F5
DIALOG_CLASS = "#32770"
BUTTON_PREFERENCE = ("cancel", "no", "ok", "close", "abort", "ignore", "end")
INTERRUPT_GRACE = 0.5


def _plain(text: str) -> str:
    return text.replace("&", "").strip().rstrip(".").lower()


class DialogWatchdog:
    """Watches one Access process for modal dialogs and runaway calls.

    Args:
        pid: PID of the owned Access process.
        policy: What to do about dialogs (``OFF`` disables dialog handling; timeouts still apply).
        call_timeout: Seconds a single operation may run before Access is terminated (``None`` = never).
        visible: Whether the Access window is visible (then only standard dialogs are considered).
        terminate: Callback that terminates the owned Access process (given a reason).
        poll_interval: Seconds between window scans while a call is in flight.
    """

    def __init__(
        self,
        pid: int,
        *,
        policy: DialogPolicy,
        call_timeout: float | None,
        visible: bool,
        terminate: Callable[[str], None],
        poll_interval: float = 0.2,
    ) -> None:
        self.pid = pid
        self.policy = policy
        self._call_timeout = call_timeout
        self._visible = visible
        self._terminate = terminate
        self._poll = poll_interval
        self._lock = threading.Lock()
        self._idle = threading.Condition(self._lock)
        self._busy = 0
        self._inflight_since: float | None = None
        self._description = ""
        self._events: list[DialogInfo] = []
        self._handled: set[int] = set()
        self._termination: tuple[str, str] | None = None
        self._stop = threading.Event()
        self._thread = threading.Thread(
            target=self._run, name=f"pyaccesskit-watchdog-{pid}", daemon=True
        )

    # ----------------------------------------------------------------------------------- lifecycle
    def start(self) -> None:
        """Start the watchdog thread and hook Ctrl+C."""
        self._thread.start()
        register_interrupt_callback(self._on_interrupt)

    def stop(self) -> None:
        """Stop the watchdog thread."""
        unregister_interrupt_callback(self._on_interrupt)
        self._stop.set()
        if self._thread.is_alive() and self._thread is not threading.current_thread():
            self._thread.join(2.0)

    # --------------------------------------------------------------------------------- CallMonitor
    def begin(self, description: str) -> None:
        with self._lock:
            self._inflight_since = time.monotonic()
            self._description = description
            self._events = []

    def end(self) -> tuple[DialogInfo, ...]:
        with self._idle:
            # A dismissed dialog unblocks the COM call immediately; wait until its handling is recorded.
            self._idle.wait_for(lambda: self._busy == 0, timeout=10.0)
            self._inflight_since = None
            events = tuple(self._events)
            self._events = []
            return events

    @property
    def termination(self) -> tuple[str, str] | None:
        """``(kind, reason)`` if the watchdog terminated Access; kind is timeout/dialog/interrupt."""
        return self._termination

    @contextlib.contextmanager
    def watching(self, description: str) -> Iterator[None]:
        """Monitor an operation whose dialogs are handled but not reported (e.g. shutdown)."""
        self.begin(description)
        try:
            yield
        finally:
            self.end()

    # ------------------------------------------------------------------------------------- thread
    def _run(self) -> None:
        while not self._stop.wait(self._poll):
            with self._lock:
                since = self._inflight_since
                description = self._description
            if since is None:
                continue
            try:
                if self.policy is not DialogPolicy.OFF:
                    for hwnd in self._dialog_windows():
                        self._handle(hwnd)
                if self._call_timeout is not None and time.monotonic() - since > self._call_timeout:
                    self._kill(
                        "timeout", f"'{description}' did not finish within {self._call_timeout:g}s"
                    )
                    with self._lock:
                        self._inflight_since = None
            except Exception:
                logger.exception("dialog watchdog error")

    def _dialog_windows(self) -> list[int]:
        found: list[int] = []

        def collect(hwnd: int, _extra: object) -> bool:
            with contextlib.suppress(Exception):
                if win32process.GetWindowThreadProcessId(hwnd)[1] != self.pid:
                    return True
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                if self._visible and win32gui.GetClassName(hwnd) != DIALOG_CLASS:
                    return True
                found.append(hwnd)
            return True

        with contextlib.suppress(Exception):
            win32gui.EnumWindows(collect, None)
        return [hwnd for hwnd in found if hwnd not in self._handled]

    @staticmethod
    def _children(hwnd: int) -> list[tuple[int, str, str]]:
        children: list[tuple[int, str, str]] = []

        def collect(child: int, _extra: object) -> bool:
            with contextlib.suppress(Exception):
                children.append(
                    (child, win32gui.GetClassName(child) or "", win32gui.GetWindowText(child) or "")
                )
            return True

        with contextlib.suppress(Exception):
            win32gui.EnumChildWindows(hwnd, collect, None)
        return children

    @staticmethod
    def _gone(hwnd: int, wait: float) -> bool:
        deadline = time.monotonic() + wait
        while time.monotonic() < deadline:
            if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd):
                return True
            time.sleep(0.05)
        return not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd)

    def _handle(self, hwnd: int) -> None:
        self._handled.add(hwnd)
        title = win32gui.GetWindowText(hwnd)
        children = self._children(hwnd)
        buttons = [(h, text) for h, cls, text in children if cls == "Button" and text]
        texts = [text for _h, cls, text in children if cls != "Button" and text.strip()]
        with contextlib.suppress(Exception):
            message = win32gui.GetDlgItemText(hwnd, 0xFFFF)
            if message and message not in texts:
                texts.insert(0, message)
        logger.warning("Access (PID %s) showed a dialog: %r %r", self.pid, title, texts)

        info = DialogInfo(
            title=title,
            text=" ".join(" ".join(texts).split()),
            buttons=tuple(text.replace("&", "") for _h, text in buttons),
            action="dismissing",
        )
        # Record the dialog *before* dismissing it: dismissal unblocks the COM call, whose end() must see it.
        with self._idle:
            self._busy += 1
            events = self._events
            events.append(info)
            position = len(events) - 1
        action = "dismissal failed"
        try:
            action = self._dismiss(hwnd, buttons)
        finally:
            with self._idle:
                if position < len(events) and events[position] is info:
                    events[position] = DialogInfo(info.title, info.text, info.buttons, action)
                self._busy -= 1
                self._idle.notify_all()

    def _dismiss(self, hwnd: int, buttons: list[tuple[int, str]]) -> str:
        with contextlib.suppress(Exception):
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        if self._gone(hwnd, 1.0):
            return "closed"
        by_text = {_plain(text): (button, text) for button, text in buttons}
        for wanted in BUTTON_PREFERENCE:
            if wanted in by_text:
                button, text = by_text[wanted]
                with contextlib.suppress(Exception):
                    win32gui.PostMessage(button, BM_CLICK, 0, 0)
                if self._gone(hwnd, 1.5):
                    return f"clicked {text.replace('&', '')!r}"
        self._kill("dialog", "a modal dialog could not be dismissed")
        return "terminated Access"

    def _kill(self, kind: str, reason: str) -> None:
        if self._termination is None:
            self._termination = (kind, reason)
        logger.error("terminating owned Access process %s: %s", self.pid, reason)
        self._terminate(reason)

    def _on_interrupt(self) -> None:
        with self._lock:
            since = self._inflight_since
        if since is not None and time.monotonic() - since > INTERRUPT_GRACE:
            self._kill("interrupt", "interrupted by the user (Ctrl+C)")
