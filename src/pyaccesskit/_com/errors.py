# pyright: basic
"""Translate ``pywintypes.com_error`` into PyAccessKit exceptions.

Access and DAO report errors as ``DISP_E_EXCEPTION`` whose EXCEPINFO carries ``scode = 0x800A0000 | number``
(FACILITY_CONTROL) — verified for every number below in spike S9. DAO errors have a ``DAO.<Collection>``
source; Access application errors have no source but a filled description.
"""

from __future__ import annotations

import traceback
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from pyaccesskit._com import constants as c
from pyaccesskit.enums import ObjectKind
from pyaccesskit.errors import (
    AccessProcessDiedError,
    ComError,
    DaoError,
    DatabaseExistsError,
    DatabaseLockedError,
    DatabaseNotFoundError,
    EngineUnavailableError,
    ErrorDetails,
    IntegrityViolationError,
    InvalidPasswordError,
    MissingParameterError,
    ObjectExistsError,
    ObjectInUseError,
    ObjectNotFoundError,
    PyAccessKitError,
    ReadOnlyError,
    RelationshipError,
    SchemaError,
    SqlSyntaxError,
    UnrecognizedFormatError,
)

__all__ = ["OpContext", "details_from", "translate"]

FACILITY_CONTROL = 0x0A

NOT_FOUND = frozenset({3011, 3078, 3265, 2102, 2103, 2450, 2451, 2465, 7874})
EXISTS = frozenset({3010, 3012, 3191, 3284, 3367, 3380})
DB_EXISTS = frozenset({3204, 7865})
DB_NOT_FOUND = frozenset({3024, 3044})
LOCKED = frozenset({3006, 3045, 3050, 3356, 3734, 7866})
IN_USE = frozenset({3008, 3009, 3211, 3262})
PASSWORD = frozenset({3031})
FORMAT = frozenset({3049, 3343})
SQL_SYNTAX = frozenset(
    {3075, 3122, 3129, 3131, 3134, 3141, 3144, 3145, 3292, 3293, 3306, 3319, 3346}
)
MISSING_PARAMETER = frozenset({3061})
RELATIONSHIP = frozenset({3366, 3368, 3609})
INTEGRITY = frozenset({3022, 3200, 3201})
SCHEMA = frozenset({3125, 3259, 3280, 3283, 3303, 3375})
READ_ONLY = frozenset({3027})
PROCESS_DIED = frozenset(
    {
        c.RPC_E_DISCONNECTED,
        c.RPC_S_SERVER_UNAVAILABLE,
        c.RPC_S_CALL_FAILED,
        c.RPC_E_SERVERFAULT,
        c.CO_E_OBJNOTCONNECTED,
    }
)
UNAVAILABLE = frozenset(
    {c.REGDB_E_CLASSNOTREG, c.CO_E_SERVER_EXEC_FAILURE, -2147221005}
)  # + CO_E_CLASSSTRING


@dataclass(frozen=True)
class OpContext:
    """What PyAccessKit was doing when a COM call failed (used to build precise messages)."""

    action: str
    kind: ObjectKind | None = None
    name: str | None = None
    path: Path | None = None
    sql: str | None = None


def _signed(value: int | None) -> int | None:
    if value is None:
        return None
    value &= 0xFFFFFFFF
    return value - 0x1_0000_0000 if value & 0x8000_0000 else value


def details_from(exc: Any, dao_errors: Sequence[DaoError] = ()) -> ErrorDetails:
    """Extract :class:`ErrorDetails` from a ``pywintypes.com_error`` (or any exception with its ``args``)."""
    args = tuple(getattr(exc, "args", ()))
    hresult = _signed(args[0]) if args and isinstance(args[0], int) else None
    message = args[1] if len(args) > 1 and isinstance(args[1], str) else None
    excepinfo = args[2] if len(args) > 2 and isinstance(args[2], tuple) else None
    source = description = None
    scode = None
    if excepinfo and len(excepinfo) >= 6:
        source = excepinfo[1] or None
        description = excepinfo[2] or None
        scode = _signed(excepinfo[5]) if isinstance(excepinfo[5], int) else None
    number = None
    if scode is not None and ((scode >> 16) & 0x1FFF) == FACILITY_CONTROL:
        number = scode & 0xFFFF
    if description is None and scode is None:
        description = message
    matching: tuple[DaoError, ...] = ()
    if dao_errors and number is not None and dao_errors[-1].number == number:
        matching = tuple(
            dao_errors
        )  # DBEngine.Errors keeps stale entries; only trust a matching one
    return ErrorDetails(
        hresult=hresult,
        scode=scode,
        number=number,
        source=source,
        description=description.strip() if description else None,
        dao_errors=matching,
    )


def _message(ctx: OpContext, details: ErrorDetails) -> str:
    summary = details.summary() or "unknown COM error"
    return f"cannot {ctx.action}: {summary}"


def translate(
    exc: BaseException, ctx: OpContext, dao_errors: Sequence[DaoError] = ()
) -> PyAccessKitError:
    """Map a COM error to the most specific :mod:`pyaccesskit.errors` exception.

    The COM-holding locals of the original traceback are cleared (the traceback itself is kept for
    debugging), so an exception kept alive by the caller cannot keep an Access process alive.
    """
    if exc.__traceback__ is not None:
        traceback.clear_frames(exc.__traceback__)
    details = details_from(exc, dao_errors)
    number = details.number
    message = _message(ctx, details)
    common = {"operation": ctx.action, "details": details}

    if number is not None:
        if number in NOT_FOUND:
            return ObjectNotFoundError(message, kind=ctx.kind, name=ctx.name, **common)
        if number in EXISTS:
            return ObjectExistsError(message, kind=ctx.kind, name=ctx.name, **common)
        if number in DB_EXISTS:
            return DatabaseExistsError(message, path=ctx.path, **common)
        if number in DB_NOT_FOUND:
            return DatabaseNotFoundError(message, path=ctx.path, **common)
        if number in LOCKED:
            return DatabaseLockedError(message, path=ctx.path, **common)
        if number in IN_USE:
            return ObjectInUseError(message, kind=ctx.kind, name=ctx.name, **common)
        if number in PASSWORD:
            return InvalidPasswordError(message, path=ctx.path, **common)
        if number in FORMAT:
            return UnrecognizedFormatError(message, path=ctx.path, **common)
        if number in SQL_SYNTAX:
            return SqlSyntaxError(message, sql=ctx.sql, **common)
        if number in MISSING_PARAMETER:
            return MissingParameterError(message, sql=ctx.sql, **common)
        if number in RELATIONSHIP:
            return RelationshipError(message, **common)
        if number in INTEGRITY:
            return IntegrityViolationError(message, **common)
        if number in SCHEMA:
            return SchemaError(message, **common)
        if number in READ_ONLY:
            return ReadOnlyError(message, **common)

    for code in (details.scode, details.hresult):
        if code in PROCESS_DIED:
            return AccessProcessDiedError(
                f"cannot {ctx.action}: the Microsoft Access process stopped responding or exited",
                **common,
            )
        if code in UNAVAILABLE:
            return EngineUnavailableError(message, **common)
    return ComError(message, **common)
