"""Real backends for the contract suite (registered as integration-marked parameters)."""

from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import TYPE_CHECKING

import pytest

if TYPE_CHECKING:
    from tests.contract.conftest import Backends


def inproc_dao_available() -> tuple[bool, str]:
    try:
        from pyaccesskit._engines.probe import inproc_dao
    except ImportError as exc:  # pywin32 missing (non-Windows)
        return False, str(exc)
    result = inproc_dao()
    return result.available, result.reason


@contextmanager
def _open(tmp_path: Path, engine: str) -> Iterator[Backends]:
    from pyaccesskit import AccessDatabase
    from tests.contract.conftest import Backends

    db = AccessDatabase.create(tmp_path / f"contract-{engine}.accdb", engine=engine, atomic=False)
    session = db._session  # the contract suite drives the backends directly
    try:
        design = session.design() if engine == "access" else None
        yield Backends(engine, session.schema(), design)  # type: ignore[arg-type]
    finally:
        db.close()


@contextmanager
def _access(tmp_path: Path) -> Iterator[Backends]:
    with _open(tmp_path, "access") as backends:
        yield backends


@contextmanager
def _dao(tmp_path: Path) -> Iterator[Backends]:
    available, reason = inproc_dao_available()
    if not available:
        pytest.skip(f"in-process DAO unavailable: {reason}")
    with _open(tmp_path, "dao") as backends:
        yield backends


REAL_FACTORIES = {"access": _access, "dao": _dao}
REAL_PARAMS = [
    pytest.param("access", id="access", marks=[pytest.mark.integration, pytest.mark.access]),
    pytest.param("dao", id="dao", marks=[pytest.mark.integration, pytest.mark.dao]),
]
