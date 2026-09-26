"""The :class:`AccessDatabase` facade — the entry point of PyAccessKit."""

from __future__ import annotations

import os
import threading
import uuid
import warnings
import weakref
from collections.abc import Mapping
from pathlib import Path
from types import TracebackType
from typing import Any, Self

from pyaccesskit._backends.protocols import PropertyTarget
from pyaccesskit._session.session import Session
from pyaccesskit.enums import Engine, Transport
from pyaccesskit.errors import (
    DatabaseExistsError,
    DatabaseLockedError,
    DatabaseNotFoundError,
    SpecError,
)
from pyaccesskit.forms.collection import FormCollection
from pyaccesskit.modules import ModuleCollection
from pyaccesskit.objects import AccessObjects
from pyaccesskit.options import SessionOptions
from pyaccesskit.properties import PropertyBag
from pyaccesskit.queries import QueryCollection
from pyaccesskit.relationships import RelationshipCollection
from pyaccesskit.tables import TableCollection

__all__ = ["AccessDatabase", "RawAccess"]

MAX_PATH_LENGTH = 255
_EXTENSIONS = (".accdb", ".mdb", ".accde", ".mde")


def _resolve(path: str | os.PathLike[str]) -> Path:
    resolved = Path(path).expanduser().resolve()
    if resolved.suffix.lower() not in _EXTENSIONS:
        raise SpecError(
            f"{resolved.name!r} is not an Access database file name (expected .accdb or .mdb)"
        )
    if len(str(resolved)) > MAX_PATH_LENGTH:
        raise SpecError(
            f"database paths are limited to {MAX_PATH_LENGTH} characters by DAO: {resolved}"
        )
    return resolved


def _abandon(session: Session) -> None:
    """Finalizer for sessions that were never closed (garbage collection or interpreter exit)."""
    if session.state != "open":
        return
    warnings.warn(
        f"PyAccessKit session for {session.target} was not closed; use 'with AccessDatabase...' or call "
        "close(). Newly created databases are discarded.",
        ResourceWarning,
        stacklevel=2,
    )
    if threading.get_ident() == session.thread_id:
        try:
            session.close(error=RuntimeError("session abandoned"))
        except BaseException:
            session.terminate()
    else:
        session.terminate()


class RawAccess:
    """Escape hatch to the underlying COM objects.

    Every object returned is a revocable proxy: it stops working (and releases its COM reference) when the
    session closes or switches engines. Save your own design changes; PyAccessKit quits Access without
    saving objects it did not manage.
    """

    def __init__(self, session: Session) -> None:
        self._session = session

    @property
    def dao(self) -> Any:
        """The DAO ``Database`` object."""
        return self._session.raw("dao")

    @property
    def dbengine(self) -> Any:
        """The DAO ``DBEngine`` object."""
        return self._session.raw("dbengine")

    @property
    def access(self) -> Any:
        """The ``Access.Application`` object (switches the session to a design session if needed)."""
        return self._session.raw("access")


class AccessDatabase:
    """An open Microsoft Access database.

    Create or open one with :meth:`create` / :meth:`open`, ideally as a context manager so that everything
    (including the Access process) is cleaned up even if an exception occurs::

        with AccessDatabase.create("crm.accdb") as db:
            db.tables.create("Customers", columns=[...])

    Attributes:
        tables: The tables (:class:`~pyaccesskit.tables.TableCollection`).
        relationships: The relationships.
        queries: The saved queries.
        forms: The forms (building forms needs Microsoft Access).
        modules: The VBA modules (needs Microsoft Access).
        objects: Text import/export and management of forms, reports, macros and modules.
        properties: The database's DAO properties (``AppTitle``, ``StartUpForm``...).
        raw: Escape hatch to the underlying COM objects.
    """

    def __init__(self, session: Session) -> None:
        self._session = session
        self.tables = TableCollection(session)
        self.relationships = RelationshipCollection(session)
        self.queries = QueryCollection(session)
        self.forms = FormCollection(session)
        self.modules = ModuleCollection(session)
        self.objects = AccessObjects(session)
        self.properties = PropertyBag(session, PropertyTarget.database())
        self.raw = RawAccess(session)
        self._finalizer = weakref.finalize(self, _abandon, session)

    # ----------------------------------------------------------------------------- construction
    @classmethod
    def create(
        cls,
        path: str | os.PathLike[str],
        *,
        overwrite: bool = False,
        atomic: bool = True,
        engine: Engine | str = Engine.AUTO,
        options: SessionOptions | None = None,
    ) -> AccessDatabase:
        """Create a new ``.accdb`` database.

        Args:
            path: Where to create the database.
            overwrite: Replace an existing file at ``path``.
            atomic: Build in a temporary sibling file and move it into place only when the session closes
                without an error; on error the target is left untouched.
            engine: ``"auto"`` (default), ``"dao"`` or ``"access"`` — see :class:`~pyaccesskit.enums.Engine`.
            options: Advanced session options.

        Raises:
            DatabaseExistsError: If ``path`` exists and ``overwrite`` is false.
        """
        target = _resolve(path)
        if target.suffix.lower() != ".accdb":
            raise SpecError("PyAccessKit creates .accdb databases (Access 2007 and later format)")
        if not target.parent.is_dir():
            raise DatabaseNotFoundError(f"folder {target.parent} does not exist", path=target)
        if target.exists() and not overwrite:
            raise DatabaseExistsError(
                f"{target} already exists (pass overwrite=True to replace it)", path=target
            )
        if atomic:
            working = target.with_name(f".{target.stem}.pak-{uuid.uuid4().hex[:8]}{target.suffix}")
        else:
            working = target
            if target.exists():
                try:
                    target.unlink()
                except PermissionError as exc:
                    raise DatabaseLockedError(
                        f"cannot overwrite {target}: it is in use", path=target
                    ) from exc
        session = Session(
            target=target,
            working=working,
            create=True,
            readonly=False,
            exclusive=True,
            password=None,
            engine=Engine(engine),
            options=options or SessionOptions(),
            overwrite=overwrite,
        )
        return cls(session)

    @classmethod
    def open(
        cls,
        path: str | os.PathLike[str],
        *,
        readonly: bool = False,
        exclusive: bool | None = None,
        password: str | None = None,
        engine: Engine | str = Engine.AUTO,
        options: SessionOptions | None = None,
    ) -> AccessDatabase:
        """Open an existing database.

        Args:
            path: The database file.
            readonly: Open read-only (also allows opening a database that is open in Access elsewhere).
            exclusive: Open exclusively (default: ``True`` unless ``readonly``). Design changes need it.
            password: Database password, if any.
            engine: ``"auto"`` (default), ``"dao"`` or ``"access"``.
            options: Advanced session options.

        Raises:
            DatabaseNotFoundError: If the file does not exist.
            DatabaseLockedError: If it is in use elsewhere.
        """
        target = _resolve(path)
        if not target.is_file():
            raise DatabaseNotFoundError(f"{target} does not exist", path=target)
        session = Session(
            target=target,
            working=target,
            create=False,
            readonly=readonly,
            exclusive=(not readonly) if exclusive is None else exclusive,
            password=password,
            engine=Engine(engine),
            options=options or SessionOptions(),
        )
        return cls(session)

    # ----------------------------------------------------------------------------------- facts
    @property
    def path(self) -> Path:
        """The database file (for atomic creation: where it will be when the session closes)."""
        return self._session.target

    @property
    def transport(self) -> Transport:
        """How the database is currently reached (in-process DAO, Access-hosted DAO, design session)."""
        return self._session.transport

    @property
    def format_version(self) -> str:
        """The database engine format (DAO ``Database.Version``): ``"12.0"`` for ``.accdb``, ``"4.0"`` for Jet 4."""
        return self._session.schema().database_info().version

    @property
    def readonly(self) -> bool:
        """Whether the database was opened read-only."""
        return self._session.readonly

    @property
    def is_open(self) -> bool:
        """Whether the session is still open."""
        return self._session.state == "open"

    @property
    def access_pid(self) -> int | None:
        """PID of the owned ``MSACCESS.EXE`` (``None`` when no Access process is used)."""
        return getattr(self._session.engine_handle, "pid", None)

    # ------------------------------------------------------------------------------------- data
    def execute(self, sql: str, params: Mapping[str, Any] | None = None) -> int:
        """Run an action statement (INSERT/UPDATE/DELETE/DDL); returns the number of affected rows.

        ``params`` binds parameters by name (``[Name]`` references in the SQL); values are never
        interpolated into the SQL text.
        """
        self._session.check_writable("run SQL")
        return self._session.schema().execute(sql, params)

    def fetch_all(
        self, sql: str, params: Mapping[str, Any] | None = None, *, limit: int | None = None
    ) -> list[dict[str, Any]]:
        """Run a SELECT and return its rows as dictionaries."""
        return self._session.schema().fetch(sql, params, limit=limit).as_dicts()

    # ------------------------------------------------------------------------------------ close
    def close(self) -> None:
        """Close the session and release everything (idempotent; only from the thread that opened it)."""
        self._session.check_thread()
        self._finalizer.detach()
        self._session.close()

    def __enter__(self) -> Self:
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: TracebackType | None,
    ) -> None:
        self._session.check_thread()
        self._finalizer.detach()
        self._session.close(error=exc)

    def __repr__(self) -> str:
        state = self.transport.value if self.is_open else "closed"
        return f"<AccessDatabase {str(self.path)!r} ({state})>"
