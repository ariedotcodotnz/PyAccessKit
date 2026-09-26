# pyright: basic
"""A Windows job object that kills its processes when PyAccessKit's Python process dies.

Once our Access process is assigned to the job, closing the job handle — which Windows does automatically
when Python exits for *any* reason (crash, ``os._exit``, Task Manager) — terminates it. Verified in spike S1
for DCOM-launched Access. Assignment is best effort; the ownership ledger covers the rest.
"""

from __future__ import annotations

import logging
from typing import Any

import pywintypes
import win32job

from pyaccesskit._win.processes import close_handle

__all__ = ["KillOnCloseJob"]

logger = logging.getLogger("pyaccesskit.process")


class KillOnCloseJob:
    """A job object with ``JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE``."""

    def __init__(self) -> None:
        self._handle: Any = win32job.CreateJobObject(None, "")
        self._set_kill_on_close(True)
        self.assigned = False

    def _set_kill_on_close(self, enabled: bool) -> None:
        info = win32job.QueryInformationJobObject(
            self._handle, win32job.JobObjectExtendedLimitInformation
        )
        flags = info["BasicLimitInformation"]["LimitFlags"]
        if enabled:
            flags |= win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        else:
            flags &= ~win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        info["BasicLimitInformation"]["LimitFlags"] = flags
        win32job.SetInformationJobObject(
            self._handle, win32job.JobObjectExtendedLimitInformation, info
        )

    def assign(self, process_handle: Any) -> bool:
        """Put the process in the job; returns ``False`` (and logs) if Windows refuses."""
        try:
            win32job.AssignProcessToJobObject(self._handle, process_handle)
        except pywintypes.error as exc:
            logger.warning("could not assign the Access process to a kill-on-close job: %s", exc)
            return False
        self.assigned = True
        return True

    def release(self) -> None:
        """Stop killing the processes when the job closes (used when handing Access over to the user)."""
        if self._handle is not None:
            self._set_kill_on_close(False)

    def close(self) -> None:
        """Close the job handle (kills remaining processes unless :meth:`release` was called)."""
        if self._handle is not None:
            close_handle(self._handle)
            self._handle = None
