"""Pure SQL helpers of the DAO backend (the module needs pywin32 to import)."""

from __future__ import annotations

import pytest

pytest.importorskip("pywintypes")

from pyaccesskit._backends.dao.schema import _declare_binary_parameters


def test_binary_parameters_are_declared() -> None:
    sql = "INSERT INTO T (Blob) VALUES ([b])"
    assert _declare_binary_parameters(sql, {"b": b"x"}) == f"PARAMETERS [b] LongBinary;\n{sql}"
    assert _declare_binary_parameters(sql, {"b": "text"}) == sql
    assert _declare_binary_parameters(sql, None) == sql


def test_existing_parameters_clause_is_extended_not_duplicated() -> None:
    sql = "PARAMETERS [n] Long; INSERT INTO T (N, Blob) VALUES ([n], [b])"
    assert _declare_binary_parameters(sql, {"n": 1, "b": b"x"}) == (
        "PARAMETERS [b] LongBinary, [n] Long; INSERT INTO T (N, Blob) VALUES ([n], [b])"
    )
    declared = "PARAMETERS [b] LongBinary; INSERT INTO T (Blob) VALUES ([b])"
    assert _declare_binary_parameters(declared, {"b": b"x"}) == declared
