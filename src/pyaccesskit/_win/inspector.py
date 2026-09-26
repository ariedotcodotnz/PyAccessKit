# pyright: basic
"""Win32 implementation of the ledger's :class:`~pyaccesskit._ledger.ProcessInspector`."""

from __future__ import annotations

from pathlib import PureWindowsPath

import pywintypes

from pyaccesskit._win.processes import (
    PROCESS_QUERY_LIMITED_INFORMATION,
    PROCESS_TERMINATE,
    SYNCHRONIZE,
    ProcessIdentity,
    close_handle,
    identity_matches,
    is_alive,
    open_process,
    query_image,
    terminate,
    wait_for_exit,
)
from pyaccesskit._win.processes import creation_time as process_creation_time

__all__ = ["Win32Inspector"]


class Win32Inspector:
    """Checks and terminates processes by *identity* (PID + creation time + image name)."""

    def matches(self, pid: int, creation_time: float) -> bool:
        return identity_matches(ProcessIdentity(pid, creation_time, ""))

    def terminate(self, pid: int, creation_time: float, image_name: str) -> bool:
        try:
            handle = open_process(
                pid, PROCESS_TERMINATE | SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION
            )
        except pywintypes.error:
            return False
        try:
            if not is_alive(handle):
                return True
            same_process = abs(process_creation_time(handle) - creation_time) < 0.001
            same_image = PureWindowsPath(query_image(handle)).name.lower() == image_name.lower()
            if not (same_process and same_image):
                return False
            terminate(handle)
            return wait_for_exit(handle, 10)
        except (pywintypes.error, OSError):
            return False
        finally:
            close_handle(handle)
