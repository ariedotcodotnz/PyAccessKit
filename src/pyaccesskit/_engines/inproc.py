# pyright: basic
"""Engine: DAO loaded into the Python process (fastest; needs Python bitness = Office/ACE bitness)."""

from __future__ import annotations

import contextlib
from pathlib import Path
from typing import Any

import pywintypes

from pyaccesskit._backends.dao.profile import apply_native_defaults
from pyaccesskit._backends.dao.schema import DaoSchemaBackend
from pyaccesskit._backends.protocols import DesignBackend, SchemaBackend
from pyaccesskit._com import constants as c
from pyaccesskit._com.dispatch import create_inproc
from pyaccesskit._com.gateway import Com
from pyaccesskit._engines.probe import DAO_PROGID, inproc_dao
from pyaccesskit._win.access_process import read_dao_errors
from pyaccesskit.enums import Transport
from pyaccesskit.errors import CapabilityError, DaoNotAvailableError

__all__ = ["InProcDaoEngine"]

_ACE_PROVIDERS = ("Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0")


class InProcDaoEngine:
    """Owns an in-process ``DBEngine`` and one open ``Database``."""

    transport = Transport.DAO_INPROC
    supports_design = False

    def __init__(
        self,
        path: Path,
        *,
        create: bool,
        readonly: bool,
        exclusive: bool,
        password: str | None,
        native_defaults: bool = True,
    ) -> None:
        self._path = path
        self._readonly = readonly
        self._exclusive = exclusive
        self._password = password
        try:
            self._dbengine: Any = create_inproc(DAO_PROGID)
        except pywintypes.com_error as exc:
            raise DaoNotAvailableError(
                "in-process DAO is not available in this Python process",
                diagnosis=inproc_dao().reason or str(exc),
            ) from exc
        self._com = Com(dao_errors=lambda: read_dao_errors(self._dbengine))
        self._db: Any = None
        verb = "create" if create else "open"
        with self._com.op(f"{verb} database {path}", path=path):
            if create:
                locale = c.dbLangGeneral + (f";pwd={password}" if password else "")
                self._db = self._dbengine.CreateDatabase(
                    str(path), locale, int(c.DatabaseTypeEnum.dbVersion120)
                )
                if native_defaults:
                    apply_native_defaults(self._db)
            else:
                self._db = self._open()
        self._schema = DaoSchemaBackend(
            com=self._com,
            database=lambda: self._db,
            path=path,
            transport=lambda: Transport.DAO_INPROC,
            run_ddl=self._run_ddl,
        )

    def _open(self) -> Any:
        connect = f";PWD={self._password}" if self._password else ""
        return self._dbengine.OpenDatabase(
            str(self._path), self._exclusive, self._readonly, connect
        )

    def _run_ddl(self, sql: str) -> None:
        """Run ANSI-92 DDL (e.g. DECIMAL(p,s)) with ADO; DAO must release its exclusive handle meanwhile."""
        with self._com.op("run DDL", sql=sql, path=self._path):
            self._db.Close()
            self._db = None
            try:
                last_error: Exception | None = None
                for provider in _ACE_PROVIDERS:
                    connection = create_inproc("ADODB.Connection")
                    extra = (
                        f"Jet OLEDB:Database Password={self._password};" if self._password else ""
                    )
                    try:
                        connection.Open(f"Provider={provider};Data Source={self._path};{extra}")
                    except pywintypes.com_error as exc:
                        last_error = exc
                        continue
                    try:
                        connection.Execute(sql)
                    finally:
                        with contextlib.suppress(pywintypes.com_error):
                            connection.Close()
                    return
                assert last_error is not None
                raise last_error
            finally:
                self._db = self._open()

    # ------------------------------------------------------------------------------------ engine
    def schema(self) -> SchemaBackend:
        return self._schema

    def design(self) -> DesignBackend:
        raise CapabilityError(
            "forms, reports, modules and text import/export need Microsoft Access; open the database with "
            "engine='access' (or engine='auto') instead of engine='dao'"
        )

    def raw(self, which: str) -> Any:
        if which == "dao":
            return self._db
        if which == "dbengine":
            return self._dbengine
        raise CapabilityError("this session uses in-process DAO; there is no Access.Application")

    def close(self) -> list[BaseException]:
        errors: list[BaseException] = []
        if self._db is not None:
            try:
                self._db.Close()
            except pywintypes.com_error as exc:
                errors.append(exc)
            self._db = None
        self._dbengine = None
        return errors

    def terminate(self) -> None:
        """Nothing to terminate: in-process DAO has no separate process."""
