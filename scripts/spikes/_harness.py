"""Shared helpers for the Phase-0 spikes.

Safety rules (mirrors the library design):
* Access is only ever created with CoCreateInstanceEx(CLSCTX_LOCAL_SERVER) — never Dispatch/GetObject,
  which would attach to a user's running instance.
* Only processes launched here (tracked by an open process handle) are ever terminated.
"""

from __future__ import annotations

import contextlib
import ctypes
import shutil
import time
from collections.abc import Iterator
from ctypes import wintypes
from pathlib import Path
from typing import Any

import pythoncom
import pywintypes
import win32api
import win32event
import win32process
from win32com.client import dynamic

ROOT = Path(__file__).resolve().parents[2]
SCRATCH = ROOT / "scratch" / "spikes"

SYNCHRONIZE = 0x0010_0000
PROCESS_TERMINATE = 0x0001
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
PROCESS_SET_QUOTA = 0x0100

acQuitSaveNone = 2


def query_image(handle: Any) -> str:
    """Full image path of a process via QueryFullProcessImageNameW (works with limited rights)."""
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    fn = kernel32.QueryFullProcessImageNameW
    fn.argtypes = [wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
    fn.restype = wintypes.BOOL
    size = wintypes.DWORD(1024)
    buf = ctypes.create_unicode_buffer(size.value)
    if not fn(int(handle), 0, buf, ctypes.byref(size)):
        raise ctypes.WinError(ctypes.get_last_error())
    return buf.value


class OwnedAccess:
    """A hidden Access instance that this script launched and therefore owns."""

    def __init__(self, progid: str = "Access.Application") -> None:
        started = time.perf_counter()
        clsid = pywintypes.IID(progid)
        unk = pythoncom.CoCreateInstanceEx(
            clsid, None, pythoncom.CLSCTX_LOCAL_SERVER, None, (pythoncom.IID_IDispatch,)
        )[0]
        self.app: Any = dynamic.Dispatch(unk)
        self.launch_seconds = time.perf_counter() - started
        hwnd = self.app.hWndAccessApp()
        _tid, self.pid = win32process.GetWindowThreadProcessId(hwnd)
        rights = SYNCHRONIZE | PROCESS_TERMINATE | PROCESS_QUERY_LIMITED_INFORMATION | PROCESS_SET_QUOTA
        self.handle = win32api.OpenProcess(rights, False, self.pid)
        self.image = query_image(self.handle)
        if not self.image.lower().endswith("msaccess.exe"):
            raise RuntimeError(f"hWndAccessApp resolved to unexpected image {self.image!r}")

    def alive(self) -> bool:
        return win32event.WaitForSingleObject(self.handle, 0) == win32event.WAIT_TIMEOUT

    def quit(self, timeout: float = 30.0) -> float:
        """Quit gracefully; terminate *our* process if it does not exit. Returns seconds to exit.

        A dialog watcher stays active for the whole quit so a prompt can't block the exit.
        """
        started = time.perf_counter()
        with DialogWatcher(self.pid):
            if self.app is not None:
                with contextlib.suppress(pywintypes.com_error):
                    self.app.Quit(acQuitSaveNone)
                self.app = None
            rc = win32event.WaitForSingleObject(self.handle, int(timeout * 1000))
        if rc == win32event.WAIT_TIMEOUT:
            print(f"  !! Access PID {self.pid} did not exit within {timeout}s — terminating our own process")
            win32api.TerminateProcess(self.handle, 1)
            win32event.WaitForSingleObject(self.handle, 10_000)
        return time.perf_counter() - started


@contextlib.contextmanager
def owned_access(progid: str = "Access.Application") -> Iterator[OwnedAccess]:
    acc = OwnedAccess(progid)
    try:
        yield acc
    finally:
        acc.quit()


class DialogWatcher:
    """Background thread that records (and optionally closes) visible windows of *one* PID.

    Our Access instance is hidden, so any visible top-level window it owns is a dialog.
    """

    def __init__(self, pid: int, *, close: bool = True, interval: float = 0.2) -> None:
        import threading

        self.pid = pid
        self.close = close
        self.interval = interval
        self.events: list[dict[str, Any]] = []
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._run, name=f"dialog-watch-{pid}", daemon=True)

    def __enter__(self) -> DialogWatcher:
        self._thread.start()
        return self

    def __exit__(self, *exc: object) -> None:
        self._stop.set()
        self._thread.join(5)

    def _texts(self, hwnd: int) -> list[tuple[str, str]]:
        import win32gui

        found: list[tuple[str, str]] = []

        def child(h: int, _: object) -> bool:
            text = win32gui.GetWindowText(h)
            if text:
                found.append((win32gui.GetClassName(h), text))
            return True

        with contextlib.suppress(Exception):
            win32gui.EnumChildWindows(hwnd, child, None)
        return found

    def _run(self) -> None:
        import win32con
        import win32gui

        seen: set[int] = set()
        while not self._stop.wait(self.interval):
            windows: list[int] = []

            def collect(hwnd: int, _: object) -> bool:
                try:
                    if win32process.GetWindowThreadProcessId(hwnd)[1] == self.pid and win32gui.IsWindowVisible(hwnd):
                        windows.append(hwnd)
                except Exception:
                    pass
                return True

            with contextlib.suppress(Exception):
                win32gui.EnumWindows(collect, None)
            for hwnd in windows:
                if hwnd in seen:
                    continue
                seen.add(hwnd)
                event = {
                    "time": time.time(),
                    "class": win32gui.GetClassName(hwnd),
                    "title": win32gui.GetWindowText(hwnd),
                    "children": self._texts(hwnd),
                }
                self.events.append(event)
                print(f"  [watcher] visible window: {event}")
                if self.close:
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)


def fresh_dir(name: str) -> Path:
    path = SCRATCH / name
    if path.exists():
        shutil.rmtree(path, ignore_errors=True)
    path.mkdir(parents=True, exist_ok=True)
    return path


def describe_com_error(exc: pywintypes.com_error) -> str:
    hresult, message, excepinfo, argerr = exc.args
    parts = [f"hresult=0x{hresult & 0xFFFFFFFF:08X} ({message})"]
    if excepinfo:
        wcode, source, description, _help, _ctx, scode = excepinfo
        parts.append(f"source={source!r} desc={description!r} wcode={wcode}")
        if scode is not None:
            parts.append(f"scode=0x{scode & 0xFFFFFFFF:08X} (low16={scode & 0xFFFF})")
    if argerr is not None:
        parts.append(f"argerr={argerr}")
    return " | ".join(parts)


def dao_errors(dbengine: Any) -> list[tuple[int, str, str]]:
    out: list[tuple[int, str, str]] = []
    try:
        errs = dbengine.Errors
        for i in range(errs.Count):
            e = errs.Item(i)
            out.append((int(e.Number), str(e.Description), str(e.Source)))
    except pywintypes.com_error as exc:  # pragma: no cover - diagnostic only
        out.append((-1, f"<could not read DBEngine.Errors: {exc}>", ""))
    return out
