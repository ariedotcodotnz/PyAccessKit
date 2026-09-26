"""Session state machine, engine selection and cleanup guarantees (fake engines; no Access)."""

from __future__ import annotations

import threading
from pathlib import Path

import pytest

from pyaccesskit import (
    AccessNotInstalledError,
    CapabilityError,
    CleanupError,
    Column,
    DaoNotAvailableError,
    Engine,
    EngineUnavailableError,
    ReadOnlyError,
    SessionClosedError,
    Transport,
    WrongThreadError,
)
from tests.fakes import FakeFactory, FakeProbe, fake_database


# ---------------------------------------------------------------------------------- engine choice
@pytest.mark.parametrize(
    ("engine", "probe", "expected"),
    [
        (Engine.AUTO, FakeProbe(inproc=True), "dao"),
        (Engine.AUTO, FakeProbe(inproc=False), "access"),
        (Engine.DAO, FakeProbe(inproc=True), "dao"),
        (Engine.ACCESS, FakeProbe(inproc=True), "access"),
    ],
)
def test_engine_selection(tmp_path: Path, engine: Engine, probe: FakeProbe, expected: str) -> None:
    db, factory = fake_database(tmp_path, engine=engine, probe=probe)
    assert factory.engines[0].plan.kind == expected
    db.close()


@pytest.mark.parametrize(
    ("engine", "probe", "error"),
    [
        (Engine.AUTO, FakeProbe(inproc=False, access_installed=False), EngineUnavailableError),
        (Engine.DAO, FakeProbe(inproc=False), DaoNotAvailableError),
        (Engine.ACCESS, FakeProbe(access_installed=False), AccessNotInstalledError),
    ],
)
def test_engine_unavailable(
    tmp_path: Path, engine: Engine, probe: FakeProbe, error: type[Exception]
) -> None:
    with pytest.raises(error) as info:
        fake_database(tmp_path, engine=engine, probe=probe)
    assert "doctor" in str(info.value) or "DAO" in str(info.value)


def test_auto_session_switches_to_access_for_design(tmp_path: Path) -> None:
    db, factory = fake_database(tmp_path)
    assert db.transport is Transport.DAO_INPROC
    raw = db.raw.dao
    db.tables.create("T", columns=[Column.text("A")])
    db.modules.create("modA", "Sub X()\nEnd Sub")
    assert db.transport is Transport.ACCESS_DESIGN
    first, second = factory.engines
    assert first.closed and first.plan.kind == "dao"
    assert (second.plan.kind, second.plan.create, second.plan.design) == ("access", False, True)
    assert raw.revoked, "raw proxies from the DAO engine are revoked by the switch"
    with pytest.raises(SessionClosedError):
        raw.Name()
    assert db.tables.names() == ["T"], "name-based handles keep working after the switch"
    db.close()


def test_dao_only_session_refuses_design(tmp_path: Path) -> None:
    db, _ = fake_database(tmp_path, engine=Engine.DAO)
    with pytest.raises(CapabilityError, match="engine='dao'"):
        db.modules.create("modA", "Sub X()\nEnd Sub")
    db.close()


def test_readonly_session(tmp_path: Path) -> None:
    db, _ = fake_database(tmp_path, create=False, readonly=True)
    with pytest.raises(ReadOnlyError):
        db.tables.create("T", columns=[Column.text("A")])
    with pytest.raises(ReadOnlyError):
        db.modules.create("modA", "Sub X()\nEnd Sub")
    with pytest.raises(CapabilityError, match="read-only"):
        db.objects.export_text("module", "modA")  # a design *read* still cannot switch engines
    db.close()


# ------------------------------------------------------------------------------------ atomicity
def test_atomic_commit_on_success(tmp_path: Path) -> None:
    db, factory = fake_database(tmp_path)
    working = factory.engines[0].plan.path
    assert working.exists() and not db.path.exists()
    db.close()
    assert db.path.exists() and not working.exists()


def test_atomic_discard_on_error(tmp_path: Path) -> None:
    db, factory = fake_database(tmp_path)
    working = factory.engines[0].plan.path
    with pytest.raises(RuntimeError), db:
        raise RuntimeError("boom")
    assert not working.exists() and not db.path.exists()


def test_close_is_idempotent_and_final(tmp_path: Path) -> None:
    db, factory = fake_database(tmp_path)
    db.close()
    db.close()
    assert factory.engines[0].closed
    with pytest.raises(SessionClosedError):
        db.tables.names()


# --------------------------------------------------------------------------------- cleanup errors
def test_cleanup_errors_are_raised_on_normal_close(tmp_path: Path) -> None:
    factory = FakeFactory(close_errors=[RuntimeError("quit failed")])
    db, _ = fake_database(tmp_path, factory=factory)
    with pytest.raises(CleanupError, match="quit failed"):
        db.close()
    assert db.path.exists(), "the database is still committed"


def test_cleanup_errors_never_mask_the_original_exception(tmp_path: Path) -> None:
    factory = FakeFactory(close_errors=[RuntimeError("quit failed")])
    db, _ = fake_database(tmp_path, factory=factory)
    with pytest.raises(ValueError, match="original") as info, db:
        raise ValueError("original")
    assert any("quit failed" in note for note in getattr(info.value, "__notes__", []))


def test_interrupt_during_close_terminates_and_propagates(tmp_path: Path) -> None:
    factory = FakeFactory(close_interrupt=True)
    db, _ = fake_database(tmp_path, factory=factory)
    with pytest.raises(KeyboardInterrupt):
        db.close()
    assert factory.engines[0].terminated
    assert not db.path.exists(), "an interrupted build is discarded"


# ----------------------------------------------------------------------------------------- threads
def test_sessions_are_thread_affine(tmp_path: Path) -> None:
    db, _ = fake_database(tmp_path)
    caught: list[BaseException] = []

    def worker() -> None:
        try:
            db.tables.names()
        except BaseException as exc:
            caught.append(exc)

    thread = threading.Thread(target=worker)
    thread.start()
    thread.join()
    assert len(caught) == 1 and isinstance(caught[0], WrongThreadError)
    db.close()
