# pyright: basic
"""An ``MSACCESS.EXE`` process that PyAccessKit started and therefore owns.

Launch (ADR 0001, spike S1):

1. ``CoCreateInstanceEx(CLSCTX_LOCAL_SERVER)`` — always a *new* process, never a user's running Access;
2. ``hWndAccessApp()`` → PID → an open process handle; identity = PID + creation time + image;
3. assignment to a kill-on-close job object (Access dies with Python, whatever kills Python);
4. an ownership-ledger entry (for ``pyaccesskit cleanup`` should everything else fail);
5. the dialog watchdog; then ``AutomationSecurity``, ``Visible`` and ``SetWarnings``.

Shutdown closes open forms/reports without saving, closes the database, calls
``Quit(acQuitSaveNone)``, waits for the process to exit and — only for *this* process, identified by the
handle we hold — terminates it if it lingers. Every step runs even if an earlier one failed.
"""

from __future__ import annotations

import contextlib
import gc
import logging
import threading
import time
from dataclasses import dataclass
from typing import Any

import pywintypes
import win32process

from pyaccesskit import _ledger
from pyaccesskit._com import constants as c
from pyaccesskit._com.dispatch import create_local_server
from pyaccesskit._com.errors import OpContext, translate
from pyaccesskit._com.gateway import Com, call, get, put, retry_busy
from pyaccesskit._version import __version__
from pyaccesskit._win.job import KillOnCloseJob
from pyaccesskit._win.processes import (
    ProcessIdentity,
    close_handle,
    current_identity,
    identity_of,
    is_alive,
    open_process,
    terminate,
    wait_for_exit,
)
from pyaccesskit._win.watchdog import DialogWatchdog
from pyaccesskit.enums import DialogPolicy, MacroSecurity
from pyaccesskit.errors import AccessNotInstalledError, DaoError

__all__ = ["AccessLaunchOptions", "AccessProcess", "read_dao_errors"]

logger = logging.getLogger("pyaccesskit.process")

_MACRO_SECURITY = {
    MacroSecurity.DISABLE: c.MsoAutomationSecurity.msoAutomationSecurityForceDisable,
    MacroSecurity.USE_UI: c.MsoAutomationSecurity.msoAutomationSecurityByUI,
    MacroSecurity.ENABLE: c.MsoAutomationSecurity.msoAutomationSecurityLow,
}
_NO_CURRENT_DATABASE = 2467


@dataclass(frozen=True)
class AccessLaunchOptions:
    """How to start and supervise Access."""

    progid: str = "Access.Application"
    visible: bool = False
    macro_security: MacroSecurity = MacroSecurity.DISABLE
    dialog_policy: DialogPolicy = DialogPolicy.FAIL
    call_timeout: float | None = 600.0
    quit_timeout: float = 30.0
    kill_on_parent_exit: bool = True


def read_dao_errors(dbengine: Any) -> list[DaoError]:
    """Snapshot ``DBEngine.Errors`` (used to enrich error messages)."""
    errors: list[DaoError] = []
    collection = dbengine.Errors
    for index in range(collection.Count):
        item = get(collection, "Item", index)
        errors.append(DaoError(int(item.Number), str(item.Description), str(item.Source)))
    return errors


class AccessProcess:
    """A hidden Access instance owned by one PyAccessKit session."""

    def __init__(self, options: AccessLaunchOptions) -> None:
        self.options = options
        self.app: Any = None
        self.identity: ProcessIdentity | None = None
        self.watchdog: DialogWatchdog | None = None
        self.is_runtime = False
        self._handle: Any = None
        self._job: KillOnCloseJob | None = None
        self._ledger_path: Any = None
        self._lock = threading.Lock()
        self._closed = False
        self.com = Com(dao_errors=self._dao_errors)

    # ------------------------------------------------------------------------------------- launch
    @classmethod
    def launch(cls, options: AccessLaunchOptions, *, database: str | None = None) -> AccessProcess:
        """Start and configure a new, owned Access process.

        Raises:
            AccessNotInstalledError: If ``Access.Application`` cannot be started.
        """
        process = cls(options)
        try:
            process.app = create_local_server(options.progid)
        except pywintypes.com_error as exc:
            error = translate(exc, OpContext("start Microsoft Access"))
            raise AccessNotInstalledError(
                f"Microsoft Access ({options.progid}) could not be started: {error.details.summary() if error.details else exc}",
                diagnosis=(
                    "Install Microsoft Access (Microsoft 365, 2016 or later), or use engine='dao' with the "
                    "Microsoft 365 Access Runtime and a Python of the same bitness. Run 'pyaccesskit doctor'."
                ),
                details=error.details,
            ) from exc
        try:
            process._take_ownership(database)
            process._configure()
        except BaseException:
            process.shutdown()
            raise
        return process

    def _take_ownership(self, database: str | None) -> None:
        try:
            hwnd = retry_busy(call, self.app, "hWndAccessApp")
            _thread, pid = win32process.GetWindowThreadProcessId(int(hwnd))
            self._handle = open_process(pid)
            identity = identity_of(self._handle, pid)
        except (pywintypes.error, pywintypes.com_error, OSError) as exc:
            logger.warning(
                "could not identify the Access process (%s); it will be shut down with Quit() only",
                exc,
            )
            return
        if identity.image_name.lower() != "msaccess.exe":
            logger.warning(
                "hWndAccessApp resolved to %s, not MSACCESS.EXE; not taking ownership",
                identity.image,
            )
            close_handle(self._handle)
            self._handle = None
            return
        self.identity = identity
        logger.info("started Microsoft Access (PID %s)", identity.pid)

        self.watchdog = DialogWatchdog(
            identity.pid,
            policy=self.options.dialog_policy,
            call_timeout=self.options.call_timeout,
            visible=self.options.visible,
            terminate=self.terminate,
        )
        self.watchdog.start()
        self.com = Com(monitor=self.watchdog, dao_errors=self._dao_errors)

        if self.options.kill_on_parent_exit:
            job = KillOnCloseJob()
            if job.assign(self._handle):
                self._job = job
            else:
                job.close()
        try:
            owner = current_identity()
            entry = _ledger.OwnedProcess(
                pid=identity.pid,
                creation_time=identity.creation_time,
                image=identity.image,
                owner_pid=owner.pid,
                owner_creation_time=owner.creation_time,
                database=database,
                started_at=time.time(),
                version=__version__,
            )
            self._ledger_path = _ledger.record(entry)
        except (OSError, pywintypes.error) as exc:
            logger.warning("could not write the ownership ledger entry: %s", exc)

    def _configure(self) -> None:
        with self.com.op("configure Microsoft Access"):
            put(self.app, "AutomationSecurity", int(_MACRO_SECURITY[self.options.macro_security]))
            if self.options.visible:
                put(self.app, "Visible", True)
            with contextlib.suppress(pywintypes.com_error):
                self.is_runtime = bool(
                    call(self.app, "SysCmd", int(c.AcSysCmdAction.acSysCmdRuntime))
                )

    def database_opened(self) -> None:
        """Session settings that need an open current database (``SetWarnings`` fails without one)."""
        with contextlib.suppress(pywintypes.com_error):
            call(get(self.app, "DoCmd"), "SetWarnings", False)

    def _dao_errors(self) -> list[DaoError]:
        return read_dao_errors(get(self.app, "DBEngine")) if self.app is not None else []

    # --------------------------------------------------------------------------------------- state
    @property
    def pid(self) -> int | None:
        """PID of the owned process (``None`` if ownership could not be established)."""
        return self.identity.pid if self.identity else None

    @property
    def alive(self) -> bool:
        """Whether the process is still running (``True`` if unknown)."""
        return is_alive(self._handle) if self._handle is not None else self.app is not None

    def terminate(self, reason: str) -> None:
        """Terminate the owned process immediately (safe from any thread)."""
        with self._lock:
            handle = self._handle
        if handle is None:
            logger.error(
                "cannot terminate Access (%s): its process could not be identified", reason
            )
            return
        with contextlib.suppress(pywintypes.error):
            if is_alive(handle):
                logger.warning("terminating owned Access process %s: %s", self.pid, reason)
                terminate(handle)

    # ------------------------------------------------------------------------------------ shutdown
    def close_open_objects(self) -> None:
        """Close every open form and report without saving (avoids 'Save As' prompts)."""
        if self.app is None:
            return
        for collection, object_type in (
            ("Forms", c.AcObjectType.acForm),
            ("Reports", c.AcObjectType.acReport),
        ):
            try:
                opened = get(self.app, collection)
                names = [
                    get(get(opened, "Item", index), "Name") for index in range(get(opened, "Count"))
                ]
            except pywintypes.com_error:
                continue
            for name in names:
                with contextlib.suppress(pywintypes.com_error):
                    call(
                        get(self.app, "DoCmd"),
                        "Close",
                        int(object_type),
                        name,
                        int(c.AcCloseSave.acSaveNo),
                    )

    def close_database(self) -> None:
        """Close the current database, if any (ignoring 'no database open')."""
        if self.app is None:
            return
        try:
            call(self.app, "CloseCurrentDatabase")
        except pywintypes.com_error as exc:
            excepinfo = exc.args[2] if len(exc.args) > 2 else None
            scode = excepinfo[5] if excepinfo else None
            if scode is None or (scode & 0xFFFF) != _NO_CURRENT_DATABASE:
                raise

    def shutdown(self, timeout: float | None = None) -> list[BaseException]:
        """Quit Access and make sure the owned process is gone. Returns the errors of failed steps."""
        errors: list[BaseException] = []
        with self._lock:
            if self._closed:
                return errors
            self._closed = True
        wait = self.options.quit_timeout if timeout is None else timeout
        # Release unreachable COM proxies while Access is still alive: releasing them after it exits makes
        # COM raise (and handle) RPC_E_DISCONNECTED, which faulthandler reports as a "fatal exception".
        gc.collect()
        watch = (
            self.watchdog.watching("shut down Microsoft Access")
            if self.watchdog
            else contextlib.nullcontext()
        )
        # If the process is already gone (e.g. terminated by the watchdog), COM steps can only fail.
        dead = self._handle is not None and not is_alive(self._handle)
        try:
            with watch:
                for step in (
                    () if dead else (self.close_open_objects, self.close_database, self._quit)
                ):
                    try:
                        step()
                    except Exception as exc:
                        errors.append(exc)
                self.app = None
                self.com = Com()
                if self._handle is not None and not wait_for_exit(self._handle, wait):
                    logger.warning(
                        "Access (PID %s) did not exit %.0fs after Quit; terminating it",
                        self.pid,
                        wait,
                    )
                    self.terminate("did not exit after Quit")
                    wait_for_exit(self._handle, 10)
        finally:
            self._release()
        return errors

    def _quit(self) -> None:
        if self.app is None:
            return
        with contextlib.suppress(pywintypes.com_error):
            call(self.app, "Quit", int(c.AcQuitOption.acQuitSaveNone))

    def _release(self) -> None:
        if self.watchdog is not None:
            self.watchdog.stop()
        if self._job is not None:
            self._job.close()
            self._job = None
        exited = self._handle is None or not is_alive(self._handle)
        if self._ledger_path is not None and exited:
            _ledger.remove(self._ledger_path)
            self._ledger_path = None
        if self._handle is not None:
            close_handle(self._handle)
            self._handle = None

    def detach(self) -> None:
        """Hand the process over to the user: it keeps running after this Python process exits."""
        if self.app is not None:
            put(self.app, "Visible", True)
            put(self.app, "UserControl", True)
        if self._job is not None:
            self._job.release()
        self.app = None
        with self._lock:
            self._closed = True
        if self.watchdog is not None:
            self.watchdog.stop()
        if self._ledger_path is not None:
            _ledger.remove(self._ledger_path)
            self._ledger_path = None
        if self._job is not None:
            self._job.close()
            self._job = None
        if self._handle is not None:
            close_handle(self._handle)
            self._handle = None
