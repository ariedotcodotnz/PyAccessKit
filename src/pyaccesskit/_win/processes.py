# pyright: basic
"""Win32 process helpers: identity (PID + creation time + image), liveness, termination.

A process is identified by PID **and** creation time: PIDs are recycled by Windows, so a PID alone could
point at an unrelated process later. While PyAccessKit holds an open handle, the PID cannot be reused.
"""

from __future__ import annotations

import ctypes
import os
from ctypes import wintypes
from dataclasses import dataclass
from pathlib import PureWindowsPath
from typing import Any

import pywintypes
import win32api
import win32event
import win32process

__all__ = [
    "ProcessIdentity",
    "access_process_ids",
    "creation_time",
    "identity_of",
    "is_alive",
    "open_process",
    "query_image",
    "wait_for_exit",
]

SYNCHRONIZE = 0x0010_0000
PROCESS_TERMINATE = 0x0001
PROCESS_SET_QUOTA = 0x0100
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
OWNED_PROCESS_RIGHTS = (
    SYNCHRONIZE | PROCESS_TERMINATE | PROCESS_SET_QUOTA | PROCESS_QUERY_LIMITED_INFORMATION
)

_kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
_QueryFullProcessImageNameW = _kernel32.QueryFullProcessImageNameW
_QueryFullProcessImageNameW.argtypes = [
    wintypes.HANDLE,
    wintypes.DWORD,
    wintypes.LPWSTR,
    ctypes.POINTER(wintypes.DWORD),
]
_QueryFullProcessImageNameW.restype = wintypes.BOOL


@dataclass(frozen=True)
class ProcessIdentity:
    """A process pinned down by PID, creation time (seconds since the epoch) and image path."""

    pid: int
    creation_time: float
    image: str

    @property
    def image_name(self) -> str:
        """The executable file name, e.g. ``MSACCESS.EXE``."""
        return PureWindowsPath(self.image).name


def open_process(pid: int, rights: int = OWNED_PROCESS_RIGHTS) -> Any:
    """Open a handle to ``pid`` (raises ``pywintypes.error`` if it no longer exists or access is denied)."""
    return win32api.OpenProcess(rights, False, pid)


def query_image(handle: Any) -> str:
    """Full path of the process image (works with ``PROCESS_QUERY_LIMITED_INFORMATION``)."""
    size = wintypes.DWORD(1024)
    buffer = ctypes.create_unicode_buffer(size.value)
    if not _QueryFullProcessImageNameW(int(handle), 0, buffer, ctypes.byref(size)):
        raise ctypes.WinError(ctypes.get_last_error())
    return buffer.value


def creation_time(handle: Any) -> float:
    """Process creation time as a POSIX timestamp."""
    created = win32process.GetProcessTimes(handle)["CreationTime"]
    return float(created.timestamp())


def identity_of(handle: Any, pid: int) -> ProcessIdentity:
    """Build the identity of an open process."""
    return ProcessIdentity(pid, creation_time(handle), query_image(handle))


def is_alive(handle: Any) -> bool:
    """Whether the process behind ``handle`` is still running."""
    return win32event.WaitForSingleObject(handle, 0) == win32event.WAIT_TIMEOUT


def wait_for_exit(handle: Any, timeout: float) -> bool:
    """Wait up to ``timeout`` seconds for the process to exit; returns whether it exited."""
    milliseconds = max(0, int(timeout * 1000))
    return win32event.WaitForSingleObject(handle, milliseconds) == win32event.WAIT_OBJECT_0


def current_identity() -> ProcessIdentity:
    """Identity of the running Python process."""
    pid = os.getpid()
    handle = open_process(pid, PROCESS_QUERY_LIMITED_INFORMATION)
    try:
        return identity_of(handle, pid)
    finally:
        win32api.CloseHandle(handle)


def identity_matches(expected: ProcessIdentity) -> bool:
    """Whether a live process still has exactly this identity (same PID *and* creation time)."""
    try:
        # SYNCHRONIZE is required by the liveness check (WaitForSingleObject), not just query rights.
        handle = open_process(expected.pid, SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION)
    except pywintypes.error:
        return False
    try:
        if not is_alive(handle):
            return False
        return abs(creation_time(handle) - expected.creation_time) < 0.001
    except (pywintypes.error, OSError):
        return False
    finally:
        win32api.CloseHandle(handle)


def access_process_ids() -> set[int]:
    """PIDs of every running ``MSACCESS.EXE`` visible to this user (used for diagnostics and tests only)."""
    pids: set[int] = set()
    for pid in win32process.EnumProcesses():
        if pid == 0:
            continue
        try:
            handle = open_process(pid, PROCESS_QUERY_LIMITED_INFORMATION)
        except pywintypes.error:
            continue
        try:
            if PureWindowsPath(query_image(handle)).name.lower() == "msaccess.exe":
                pids.add(pid)
        except OSError:
            pass
        finally:
            win32api.CloseHandle(handle)
    return pids


def terminate(handle: Any, exit_code: int = 1) -> None:
    """Terminate the process behind an **owned** handle."""
    win32api.TerminateProcess(handle, exit_code)


def close_handle(handle: Any) -> None:
    """Close a Win32 handle, ignoring errors."""
    try:
        win32api.CloseHandle(handle)
    except pywintypes.error:
        pass
