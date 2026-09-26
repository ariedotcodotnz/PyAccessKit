# pyright: basic
"""Engine: Microsoft Access automation (works with any Python/Office bitness combination).

Schema work uses **Access-hosted DAO** (``app.DBEngine.OpenDatabase``): the database is *not* opened in the
Access UI, so AutoExec macros and startup forms never run. The first design feature (forms, modules, text
import/export, Decimal DDL) upgrades the engine to a **design session** (``OpenCurrentDatabase``).
New databases are created with ``NewCurrentDatabase`` and start in design mode (native Access defaults).
"""

from __future__ import annotations

import contextlib
import os
from collections.abc import Callable
from pathlib import Path
from typing import Any, Literal

import pywintypes

from pyaccesskit._backends.access.design import AccessDesignBackend
from pyaccesskit._backends.dao.schema import DaoSchemaBackend
from pyaccesskit._backends.protocols import DesignBackend, SchemaBackend
from pyaccesskit._com import constants as c
from pyaccesskit._com.errors import OpContext, translate
from pyaccesskit._com.gateway import call, get
from pyaccesskit._win.access_process import AccessLaunchOptions, AccessProcess
from pyaccesskit.enums import Transport
from pyaccesskit.errors import AccessRuntimeOnlyError, CapabilityError, DatabaseLockedError

__all__ = ["AccessEngine"]


def _same_path(left: str, right: Path) -> bool:
    return os.path.normcase(str(Path(left).resolve())) == os.path.normcase(str(right.resolve()))


class AccessEngine:
    """Owns one Access process and the database opened in it."""

    def __init__(
        self,
        path: Path,
        *,
        create: bool,
        readonly: bool,
        exclusive: bool,
        password: str | None,
        design: bool,
        options: AccessLaunchOptions,
        on_created: Callable[[], None] | None = None,
    ) -> None:
        self._path = path
        self._readonly = readonly
        self._exclusive = exclusive
        self._password = password
        self._mode: Literal["hosted", "design"] = "hosted"
        self._db: Any = None
        self._design: AccessDesignBackend | None = None
        self._on_created = on_created
        self._process = AccessProcess.launch(options, database=str(path))
        try:
            if create:
                self._new_current()
            elif design and not readonly:
                self._open_current()
            else:
                self._open_hosted()
        except BaseException:
            self._process.shutdown()
            raise
        self._schema = DaoSchemaBackend(
            com=self._process.com,
            database=lambda: self._db,
            path=path,
            transport=lambda: self.transport,
            run_ddl=self._run_ddl,
        )

    @property
    def _app(self) -> Any:
        return self._process.app

    @property
    def transport(self) -> Transport:
        return Transport.ACCESS_DESIGN if self._mode == "design" else Transport.ACCESS_HOSTED

    @property
    def supports_design(self) -> bool:
        return not self._readonly and not self._process.is_runtime

    @property
    def pid(self) -> int | None:
        return self._process.pid

    # ------------------------------------------------------------------------------ open / create
    def _connect(self) -> str:
        return f";PWD={self._password}" if self._password else ""

    def _open_hosted(self) -> None:
        with self._process.com.op(f"open database {self._path}", path=self._path):
            dbengine = get(self._app, "DBEngine")
            self._db = call(
                dbengine,
                "OpenDatabase",
                str(self._path),
                self._exclusive,
                self._readonly,
                self._connect(),
            )
        self._mode = "hosted"

    def _new_current(self) -> None:
        with self._process.com.op(f"create database {self._path}", path=self._path):
            call(
                self._app,
                "NewCurrentDatabase",
                str(self._path),
                int(c.AcNewDatabaseFormat.acNewDatabaseFormatAccess2007),
            )
        self._verify_current()  # confirms Access created *this* file and it is now its current database
        if self._on_created is not None:
            self._on_created()

    def _open_current(self) -> None:
        with self._process.com.op(f"open database {self._path} in Access", path=self._path):
            call(
                self._app,
                "OpenCurrentDatabase",
                str(self._path),
                self._exclusive,
                self._password or "",
            )
        self._verify_current()

    def _verify_current(self) -> None:
        """``OpenCurrentDatabase`` can fail *silently* (ADR 0001, S9b): check, and explain failures."""
        with self._process.com.op(f"open database {self._path} in Access", path=self._path):
            full_name = str(get(get(self._app, "CurrentProject"), "FullName") or "")
            db = call(self._app, "CurrentDb") if full_name else None
        if full_name and db is not None and _same_path(full_name, self._path):
            self._db = db
            self._mode = "design"
            self._process.database_opened()
            return
        # Ask DAO why Access refused: it reports locked / unrecognized format / missing precisely.
        try:
            probe = call(
                get(self._app, "DBEngine"),
                "OpenDatabase",
                str(self._path),
                False,
                True,
                self._connect(),
            )
        except pywintypes.com_error as exc:
            raise translate(
                exc, OpContext(f"open database {self._path} in Access", path=self._path)
            ) from exc
        with contextlib.suppress(pywintypes.com_error):
            call(probe, "Close")
        raise DatabaseLockedError(
            f"Microsoft Access could not open {self._path} as its current database (it may be open "
            "exclusively in another Access instance)",
            path=self._path,
        )

    def upgrade(self) -> None:
        """Switch from Access-hosted DAO to a design session (``OpenCurrentDatabase``)."""
        if self._mode == "design":
            return
        if self._readonly:
            raise CapabilityError(
                "read-only sessions cannot use design features (forms, modules, DDL)"
            )
        if self._process.is_runtime:
            raise AccessRuntimeOnlyError(
                "only the Access Runtime is installed; design features need full Microsoft Access",
                diagnosis="Install Microsoft Access (not the Runtime) to build forms, reports and modules.",
            )
        with self._process.com.op(f"release database {self._path}", path=self._path):
            call(self._db, "Close")
        self._db = None
        self._open_current()

    def _run_ddl(self, sql: str) -> None:
        self.upgrade()
        with self._process.com.op("run DDL", sql=sql, path=self._path):
            connection = get(get(self._app, "CurrentProject"), "Connection")
            call(connection, "Execute", sql)

    # ------------------------------------------------------------------------------------ engine
    def schema(self) -> SchemaBackend:
        return self._schema

    def design(self) -> DesignBackend:
        self.upgrade()
        if self._design is None:
            self._design = AccessDesignBackend(
                com=lambda: self._process.com, app=lambda: self._process.app
            )
        return self._design

    def raw(self, which: str) -> Any:
        if which == "dao":
            return self._db
        if which == "dbengine":
            return get(self._app, "DBEngine")
        self.upgrade()
        return self._app

    def close(self) -> list[BaseException]:
        errors: list[BaseException] = []
        if self._design is not None:
            self._design.cleanup()
            self._design = None
        if self._db is not None and self._mode == "hosted":
            try:
                call(self._db, "Close")
            except pywintypes.com_error as exc:
                errors.append(exc)
        self._db = None
        errors.extend(self._process.shutdown())
        return errors

    def terminate(self) -> None:
        self._process.terminate("the PyAccessKit session was abandoned")
