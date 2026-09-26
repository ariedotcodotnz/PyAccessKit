"""Error translation (runs without pywin32: com_error is only duck-typed through ``args``)."""

from __future__ import annotations

from pathlib import Path

import pytest

from pyaccesskit._com import constants as c
from pyaccesskit._com.errors import OpContext, details_from, translate
from pyaccesskit.enums import ObjectKind
from pyaccesskit.errors import (
    AccessProcessDiedError,
    ComError,
    DaoError,
    DatabaseExistsError,
    DatabaseLockedError,
    DatabaseNotFoundError,
    EngineUnavailableError,
    IntegrityViolationError,
    InvalidPasswordError,
    MissingParameterError,
    ObjectExistsError,
    ObjectNotFoundError,
    RelationshipError,
    SqlSyntaxError,
    UnrecognizedFormatError,
)


class FakeComError(Exception):
    """Mimics ``pywintypes.com_error``: args = (hresult, message, excepinfo, argerr)."""


def dao_error(
    number: int, description: str = "boom", source: str = "DAO.TableDefs"
) -> FakeComError:
    scode = -0x7FF60000 + number  # 0x800A0000 | number, as a signed 32-bit int
    return FakeComError(
        c.DISP_E_EXCEPTION, "Exception occurred.", (0, source, description, None, 0, scode), None
    )


def hresult_error(hresult: int, message: str = "failure") -> FakeComError:
    return FakeComError(hresult, message, None, None)


def test_details_extract_the_access_error_number() -> None:
    details = details_from(dao_error(3010, "Table 'T' already exists."))
    assert details.number == 3010
    assert details.source == "DAO.TableDefs"
    assert details.summary() == "DAO error 3010: Table 'T' already exists."
    access = details_from(dao_error(2102, "The form name 'x' is misspelled", source=""))
    assert access.summary().startswith("Access error 2102")


def test_stale_dao_errors_are_ignored() -> None:
    stale = [DaoError(3061, "Too few parameters", "DAO.Database")]
    assert details_from(dao_error(2102), stale).dao_errors == ()
    matching = [DaoError(3265, "Item not found", "DAO.TableDefs")]
    assert details_from(dao_error(3265), matching).dao_errors == tuple(matching)


@pytest.mark.parametrize(
    ("number", "expected"),
    [
        (3265, ObjectNotFoundError),
        (2102, ObjectNotFoundError),
        (7874, ObjectNotFoundError),
        (3010, ObjectExistsError),
        (3012, ObjectExistsError),
        (3204, DatabaseExistsError),
        (7865, DatabaseExistsError),
        (3024, DatabaseNotFoundError),
        (3045, DatabaseLockedError),
        (7866, DatabaseLockedError),
        (3031, InvalidPasswordError),
        (3343, UnrecognizedFormatError),
        (3075, SqlSyntaxError),
        (3129, SqlSyntaxError),
        (3141, SqlSyntaxError),
        (3061, MissingParameterError),
        (3368, RelationshipError),
        (3609, RelationshipError),
        (3201, IntegrityViolationError),
        (9999, ComError),
    ],
)
def test_error_numbers_map_to_exceptions(number: int, expected: type[Exception]) -> None:
    ctx = OpContext(
        "do something", kind=ObjectKind.TABLE, name="T", path=Path("x.accdb"), sql="SELECT 1"
    )
    error = translate(dao_error(number), ctx)
    assert isinstance(error, expected)
    assert error.operation == "do something"
    assert error.details is not None and error.details.number == number
    assert str(error).startswith("cannot do something")


def test_context_is_attached() -> None:
    ctx = OpContext("read table 'T'", kind=ObjectKind.TABLE, name="T")
    error = translate(dao_error(3265), ctx)
    assert isinstance(error, ObjectNotFoundError)
    assert (error.kind, error.name) == (ObjectKind.TABLE, "T")
    sql_error = translate(dao_error(3075), OpContext("run", sql="SELECT FROM"))
    assert isinstance(sql_error, SqlSyntaxError) and sql_error.sql == "SELECT FROM"


@pytest.mark.parametrize(
    "hresult", [c.RPC_E_DISCONNECTED, c.RPC_S_SERVER_UNAVAILABLE, c.RPC_E_SERVERFAULT]
)
def test_dead_server_errors(hresult: int) -> None:
    assert isinstance(translate(hresult_error(hresult), OpContext("x")), AccessProcessDiedError)


def test_unavailable_engine() -> None:
    assert isinstance(
        translate(hresult_error(c.REGDB_E_CLASSNOTREG), OpContext("x")), EngineUnavailableError
    )


def test_translate_clears_traceback_frames() -> None:
    def fail() -> None:
        big_local = object()  # noqa: F841 - stands in for a COM proxy held by a frame
        raise dao_error(3265)

    try:
        fail()
    except FakeComError as exc:
        translate(exc, OpContext("x"))
        tb = exc.__traceback__
        assert tb is not None
        while tb.tb_next is not None:
            tb = tb.tb_next
        assert "big_local" not in tb.tb_frame.f_locals
