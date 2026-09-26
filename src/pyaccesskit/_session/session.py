"""The Session: sole owner of an engine, its state machine and every cleanup guarantee (COM-free).

* **Engine selection** — ``auto`` prefers in-process DAO, else Access; the first design feature lazily
  switches an in-process DAO session to Access (name-based handles survive; raw proxies are revoked).
* **Thread affinity** — COM objects belong to the apartment that created them; a session refuses calls
  from other threads (:class:`~pyaccesskit.errors.WrongThreadError`).
* **Atomic creation** — ``create()`` builds a sibling temp file and moves it into place only when the
  session closes without error; otherwise the temp file is deleted and the target is untouched.
* **Cleanup** — failures while closing never mask the exception that caused the close (they are attached
  as notes); on a normal close they are raised together as :class:`~pyaccesskit.errors.CleanupError`.
"""

from __future__ import annotations

import contextlib
import logging
import threading
import time
from collections.abc import Callable
from pathlib import Path
from typing import Any, Literal

from pyaccesskit._backends.protocols import DesignBackend, SchemaBackend
from pyaccesskit._com.raw import ProxyRegistry
from pyaccesskit._session.protocols import EngineHandle, EnginePlan, EnvironmentProbe, RawKind
from pyaccesskit.enums import Engine, Transport
from pyaccesskit.errors import (
    AccessNotInstalledError,
    CapabilityError,
    CleanupError,
    DaoNotAvailableError,
    DatabaseLockedError,
    EngineUnavailableError,
    ReadOnlyError,
    SessionClosedError,
    WrongThreadError,
)
from pyaccesskit.options import SessionOptions

__all__ = ["DefaultProbe", "EngineFactory", "Session"]

logger = logging.getLogger("pyaccesskit.session")

EngineFactory = Callable[[EnginePlan], EngineHandle]

_ACCESS_ADVICE = (
    "Install Microsoft Access (Microsoft 365, 2016 or later) for full functionality, or the Microsoft 365 Access "
    "Runtime together with a Python of the same bitness for DAO-only work. Run 'pyaccesskit doctor' for details."
)


class DefaultProbe:
    """Probes the real machine (imports the COM layer lazily)."""

    def inproc_dao(self) -> tuple[bool, str]:
        from pyaccesskit._engines.probe import inproc_dao

        result = inproc_dao()
        return result.available, result.reason

    def access(self, progid: str) -> bool:
        from pyaccesskit._engines.probe import (
            access_registered,
        )

        return access_registered(progid)


def default_factory(plan: EnginePlan) -> EngineHandle:
    """Open a real engine (imports the COM layer lazily)."""
    from pyaccesskit._engines import open_engine

    engine: EngineHandle = open_engine(plan)
    return engine


class Session:
    """Owns one engine for one database file."""

    def __init__(
        self,
        *,
        target: Path,
        working: Path,
        create: bool,
        readonly: bool,
        exclusive: bool,
        password: str | None,
        engine: Engine,
        options: SessionOptions,
        factory: EngineFactory | None = None,
        probe: EnvironmentProbe | None = None,
    ) -> None:
        self.target = target
        self.working = working
        self.created = create
        self.readonly = readonly
        self.exclusive = exclusive
        self.password = password
        self.engine_choice = engine
        self.options = options
        self.thread_id = threading.get_ident()
        self.state: Literal["opening", "open", "closing", "closed"] = "opening"
        self._factory = factory or default_factory
        self._probe = probe or DefaultProbe()
        self._raw = ProxyRegistry()
        self._kind = self._select()
        self._engine: EngineHandle = self._factory(
            self._plan(self._kind, create=create, design=False)
        )
        self.state = "open"
        logger.info("opened %s via %s", target, self._engine.transport.value)

    # ------------------------------------------------------------------------------ engine choice
    def _plan(self, kind: Literal["dao", "access"], *, create: bool, design: bool) -> EnginePlan:
        return EnginePlan(
            kind=kind,
            path=self.working,
            create=create,
            readonly=self.readonly,
            exclusive=self.exclusive,
            password=self.password,
            design=design,
            options=self.options,
        )

    def _select(self) -> Literal["dao", "access"]:
        progid = self.options.access_progid
        if self.engine_choice is Engine.DAO:
            available, reason = self._probe.inproc_dao()
            if not available:
                raise DaoNotAvailableError("in-process DAO is not available", diagnosis=reason)
            return "dao"
        if self.engine_choice is Engine.ACCESS:
            if not self._probe.access(progid):
                raise AccessNotInstalledError(
                    "Microsoft Access is not installed", diagnosis=_ACCESS_ADVICE
                )
            return "access"
        available, reason = self._probe.inproc_dao()
        if available:
            return "dao"
        if self._probe.access(progid):
            return "access"
        raise EngineUnavailableError(
            "neither in-process DAO nor Microsoft Access is available",
            diagnosis=f"In-process DAO: {reason}\nMicrosoft Access: not installed.\n{_ACCESS_ADVICE}",
        )

    # ------------------------------------------------------------------------------------ checks
    def check(self) -> None:
        """Raise if the session is closed or used from the wrong thread."""
        if self.state != "open":
            raise SessionClosedError(f"the session for {self.target} is closed")
        if threading.get_ident() != self.thread_id:
            raise WrongThreadError(
                "a PyAccessKit session can only be used from the thread that opened it (COM objects are "
                "bound to their apartment); open a separate session per thread or process"
            )

    def check_writable(self, action: str) -> None:
        """Raise :class:`ReadOnlyError` for read-only sessions."""
        self.check()
        if self.readonly:
            raise ReadOnlyError(
                f"cannot {action}: {self.target.name} was opened with readonly=True"
            )

    @property
    def transport(self) -> Transport:
        """How the database is currently reached."""
        return self._engine.transport

    @property
    def engine_handle(self) -> EngineHandle:
        """The current engine (for diagnostics and tests)."""
        return self._engine

    # ---------------------------------------------------------------------------------- backends
    def schema(self) -> SchemaBackend:
        """The schema backend of the current engine."""
        self.check()
        return self._engine.schema()

    def design(self) -> DesignBackend:
        """The design backend, switching to Microsoft Access first if needed."""
        self.check()
        self._ensure_access()
        return self._engine.design()

    def _ensure_access(self) -> None:
        if self._kind == "access":
            return
        if self.engine_choice is Engine.DAO:
            raise CapabilityError(
                "forms, reports, modules and text import/export need Microsoft Access; this session was "
                "opened with engine='dao'"
            )
        if self.readonly:
            raise CapabilityError("read-only sessions cannot use design features")
        if not self._probe.access(self.options.access_progid):
            raise AccessNotInstalledError(
                "this feature needs Microsoft Access, which is not installed",
                diagnosis=_ACCESS_ADVICE,
            )
        logger.info("switching %s from in-process DAO to Microsoft Access", self.target)
        self._raw.revoke_all()
        self._raw = ProxyRegistry()
        errors = self._engine.close()
        if errors:
            raise CleanupError(
                "could not release in-process DAO before switching to Access", errors=tuple(errors)
            )
        self._engine = self._factory(self._plan("access", create=False, design=True))
        self._kind = "access"

    def raw(self, which: RawKind) -> Any:
        """A revocable proxy for a raw COM object (escape hatch)."""
        self.check()
        if which == "access":
            self._ensure_access()
        return self._raw.wrap(self._engine.raw(which), f"db.raw.{which}")

    # ------------------------------------------------------------------------------------- close
    @property
    def atomic(self) -> bool:
        """Whether the database is being built in a temp file."""
        return self.working != self.target

    def close(self, error: BaseException | None = None) -> None:
        """Close the session. With ``error`` set, a newly created database is discarded."""
        if self.state in ("closing", "closed"):
            return
        self.state = "closing"
        cleanup: list[BaseException] = []
        interrupt: BaseException | None = None
        try:
            self._raw.revoke_all()
            cleanup.extend(self._engine.close())
        except (KeyboardInterrupt, SystemExit) as exc:
            interrupt = exc
            with contextlib.suppress(Exception):
                self._engine.terminate()
        except Exception as exc:
            cleanup.append(exc)
        finally:
            try:
                if self.atomic:
                    if error is None and interrupt is None:
                        self._commit()
                    else:
                        self._discard()
            except Exception as exc:
                cleanup.append(exc)
            self.state = "closed"
            logger.info("closed %s", self.target)
        if interrupt is not None:
            raise interrupt
        if cleanup:
            if error is not None:
                for problem in cleanup:
                    error.add_note(
                        f"PyAccessKit cleanup also failed: {type(problem).__name__}: {problem}"
                    )
            else:
                raise CleanupError(
                    f"closing {self.target.name} did not complete cleanly", errors=tuple(cleanup)
                )

    def terminate(self) -> None:
        """Emergency stop (safe from any thread): terminate an owned Access process, discard temp files."""
        if self.state == "closed":
            return
        self.state = "closed"
        with contextlib.suppress(Exception):
            self._engine.terminate()
        if self.atomic:
            with contextlib.suppress(Exception):
                self._discard()

    @staticmethod
    def _retry(action: Callable[[], object], attempts: int = 40, delay: float = 0.25) -> None:
        for attempt in range(attempts):
            try:
                action()
                return
            except PermissionError:
                if attempt == attempts - 1:
                    raise
                time.sleep(delay)

    def _commit(self) -> None:
        try:
            self._retry(lambda: self.working.replace(self.target))
        except OSError as exc:
            raise DatabaseLockedError(
                f"could not move the new database into place at {self.target} ({exc}); it was left at {self.working}",
                path=self.target,
            ) from exc
        self._remove_lock_file(self.working)

    def _discard(self) -> None:
        self._retry(lambda: self.working.unlink(missing_ok=True))
        self._remove_lock_file(self.working)

    @staticmethod
    def _remove_lock_file(path: Path) -> None:
        lock_suffix = ".ldb" if path.suffix.lower() in (".mdb", ".mde") else ".laccdb"
        with contextlib.suppress(OSError):
            path.with_suffix(lock_suffix).unlink(missing_ok=True)
