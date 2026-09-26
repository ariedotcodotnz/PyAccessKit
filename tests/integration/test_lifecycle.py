"""Tier 5: Access process lifecycle — no orphaned MSACCESS.EXE, whatever happens.

These tests start real (hidden) Access instances. They never touch Access processes they did not start;
the only process they terminate directly is a decoy instance the test itself launched.
"""

from __future__ import annotations

import gc
import os
import subprocess
import sys
import textwrap
import threading
import time
import warnings
from pathlib import Path

import pytest

from pyaccesskit import (
    AccessDatabase,
    AccessDialogError,
    AccessTimeoutError,
    Column,
    DatabaseExistsError,
    DatabaseLockedError,
    DatabaseNotFoundError,
    DialogPolicy,
    ReadOnlyError,
    SessionClosedError,
    SessionOptions,
    UnrecognizedFormatError,
    WrongThreadError,
)

pytestmark = [pytest.mark.integration, pytest.mark.access, pytest.mark.lifecycle]


def _alive(pid: int) -> bool:
    from pyaccesskit._win.processes import access_process_ids

    return pid in access_process_ids()


def _wait_gone(pid: int, timeout: float = 20.0) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if not _alive(pid):
            return True
        time.sleep(0.2)
    return not _alive(pid)


def _customers(db: AccessDatabase) -> None:
    db.tables.create(
        "Customers",
        columns=[Column.autonumber("CustomerID", primary_key=True), Column.text("CustomerName")],
    )


# --------------------------------------------------------------------------------------- basics
def test_create_is_atomic_on_success(tmp_path: Path) -> None:
    target = tmp_path / "ok.accdb"
    with AccessDatabase.create(target, engine="access") as db:
        pid = db.access_pid
        assert pid is not None and _alive(pid)
        assert not target.exists(), "the database appears only when the session closes"
        _customers(db)
    assert target.exists()
    assert _wait_gone(pid)
    assert [p.name for p in tmp_path.iterdir()] == ["ok.accdb"], "no temp or lock files left behind"


def test_exception_discards_new_database_and_quits_access(tmp_path: Path) -> None:
    target = tmp_path / "failed.accdb"
    pid: int | None = None
    with (  # noqa: PT012 - the failure must happen inside the session block
        pytest.raises(RuntimeError, match="boom"),
        AccessDatabase.create(target, engine="access") as db,
    ):
        pid = db.access_pid
        _customers(db)
        raise RuntimeError("boom")
    assert pid is not None and _wait_gone(pid)
    assert list(tmp_path.iterdir()) == [], "a failed build leaves nothing behind"


def test_existing_target_is_untouched_when_rebuild_fails(tmp_path: Path) -> None:
    target = tmp_path / "keep.accdb"
    with AccessDatabase.create(target, engine="access") as db:
        _customers(db)
    before = target.read_bytes()
    with (
        pytest.raises(RuntimeError),
        AccessDatabase.create(target, overwrite=True, engine="access"),
    ):
        raise RuntimeError("rebuild failed")
    assert target.read_bytes() == before


def test_close_is_idempotent_and_blocks_further_use(tmp_path: Path) -> None:
    db = AccessDatabase.create(tmp_path / "c.accdb", engine="access")
    pid = db.access_pid
    db.close()
    db.close()
    assert not db.is_open
    with pytest.raises(SessionClosedError):
        db.tables.names()
    assert pid is not None and _wait_gone(pid)


def test_abandoned_session_is_cleaned_up_by_gc(tmp_path: Path) -> None:
    db = AccessDatabase.create(tmp_path / "abandoned.accdb", engine="access")
    pid = db.access_pid
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        del db
        gc.collect()
    assert any(issubclass(w.category, ResourceWarning) for w in caught)
    assert pid is not None and _wait_gone(pid)
    assert list(tmp_path.iterdir()) == []


def test_session_is_thread_affine(tmp_path: Path) -> None:
    with AccessDatabase.create(tmp_path / "t.accdb", engine="access") as db:
        errors: list[BaseException] = []

        def worker() -> None:
            try:
                db.tables.names()
            except BaseException as exc:
                errors.append(exc)

        thread = threading.Thread(target=worker)
        thread.start()
        thread.join()
        assert len(errors) == 1 and isinstance(errors[0], WrongThreadError)


# ------------------------------------------------------------------------------------ crash safety
def _run_child(script: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, "-c", textwrap.dedent(script)],
        capture_output=True,
        text=True,
        timeout=120,
        env={**os.environ},
        check=False,
    )


def test_hard_python_crash_kills_owned_access(tmp_path: Path) -> None:
    """The kill-on-close job object terminates Access when Python dies without any cleanup."""
    result = _run_child(
        f"""
        import os, sys
        from pyaccesskit import AccessDatabase
        db = AccessDatabase.create(r"{tmp_path / "crash.accdb"}", engine="access")
        print(db.access_pid, flush=True)
        os._exit(3)
        """
    )
    assert result.returncode == 3, result.stderr
    pid = int(result.stdout.strip())
    assert _wait_gone(pid), "Access outlived the crashed Python process"


def test_orphans_are_reaped_via_the_ledger(tmp_path: Path) -> None:
    """Without the job object (kill_on_parent_exit=False), the ledger lets reap_orphans() clean up."""
    from pyaccesskit import reap_orphans

    result = _run_child(
        f"""
        import os
        from pyaccesskit import AccessDatabase, SessionOptions
        db = AccessDatabase.create(r"{tmp_path / "orphan.accdb"}", engine="access",
                                   options=SessionOptions(kill_on_parent_exit=False))
        print(db.access_pid, flush=True)
        os._exit(0)
        """
    )
    pid = int(result.stdout.strip())
    try:
        assert _alive(pid), "without the job object the orphan survives its owner"
        dry = [r for r in reap_orphans(dry_run=True) if r.entry.pid == pid]
        assert [r.action for r in dry] == ["would-terminate"]
        done = [r for r in reap_orphans() if r.entry.pid == pid]
        assert [r.action for r in done] == ["terminated"]
        assert _wait_gone(pid)
    finally:
        if _alive(pid):  # pragma: no cover - only if the assertions above failed
            reap_orphans()


def _msaccess_path() -> str | None:
    from pyaccesskit._engines.probe import access_facts

    return access_facts().executable


def test_never_attaches_to_a_running_user_instance(tmp_path: Path) -> None:
    """A user's Access (here a decoy started by the test) is neither used nor closed by PyAccessKit."""
    exe = _msaccess_path()
    if exe is None or not Path(exe).exists():
        pytest.skip("MSACCESS.EXE not found in App Paths")
    decoy = subprocess.Popen([exe, "/nostartup"])
    try:
        time.sleep(4)
        with AccessDatabase.create(tmp_path / "mine.accdb", engine="access") as db:
            assert db.access_pid != decoy.pid
        assert decoy.poll() is None, "the user's Access instance must survive"
    finally:
        decoy.kill()  # the decoy belongs to this test
        decoy.wait(10)


# ------------------------------------------------------------------------------ dialogs & timeouts
def _eval_in_access(db: AccessDatabase, expression: str) -> object:
    """Evaluate an Access expression inside a monitored operation (as PyAccessKit's own calls are)."""
    from pyaccesskit._com.gateway import call

    engine = db._session.engine_handle
    process = engine._process  # type: ignore[attr-defined]
    with process.com.op(f"evaluate {expression}"):
        return call(process.app, "Eval", expression)


def test_modal_dialog_is_detected_and_dismissed(tmp_path: Path) -> None:
    with AccessDatabase.create(tmp_path / "d.accdb", engine="access") as db:
        with pytest.raises(AccessDialogError) as info:
            _eval_in_access(db, 'MsgBox("Continue?", 4, "PyAccessKit test")')
        dialog = info.value.dialogs[0]
        assert dialog.title == "PyAccessKit test"
        assert "No" in dialog.buttons
        assert dialog.action.startswith(("closed", "clicked"))
        _customers(db)  # Access is still usable afterwards


def test_dialog_warn_policy(tmp_path: Path) -> None:
    options = SessionOptions(dialog_policy=DialogPolicy.WARN)
    with (
        AccessDatabase.create(tmp_path / "w.accdb", engine="access", options=options) as db,
        pytest.warns(Warning, match="PyAccessKit warn test"),
    ):
        assert _eval_in_access(db, 'MsgBox("Hi", 0, "PyAccessKit warn test")') == 1


def test_call_timeout_terminates_owned_access(tmp_path: Path) -> None:
    options = SessionOptions(dialog_policy=DialogPolicy.OFF, call_timeout=3)
    db = AccessDatabase.create(tmp_path / "timeout.accdb", engine="access", options=options)
    pid = db.access_pid
    try:
        with pytest.raises(AccessTimeoutError, match="did not finish within 3s"):
            _eval_in_access(db, 'MsgBox("blocks forever", 0, "PyAccessKit timeout test")')
    finally:
        db.close()  # the process is already gone: closing must be quiet
    assert pid is not None and _wait_gone(pid)


# --------------------------------------------------------------------------------- file errors
def test_file_errors(tmp_path: Path) -> None:
    with pytest.raises(DatabaseNotFoundError):
        AccessDatabase.open(tmp_path / "missing.accdb", engine="access")
    bogus = tmp_path / "bogus.accdb"
    bogus.write_text("not a database", encoding="ascii")
    with pytest.raises(UnrecognizedFormatError):
        AccessDatabase.open(bogus, engine="access")
    target = tmp_path / "exists.accdb"
    with AccessDatabase.create(target, engine="access") as db:
        _customers(db)
    with pytest.raises(DatabaseExistsError):
        AccessDatabase.create(target, engine="access")


def test_exclusive_lock_is_reported(tmp_path: Path) -> None:
    target = tmp_path / "locked.accdb"
    with AccessDatabase.create(target, engine="access") as db:
        _customers(db)
    with AccessDatabase.open(target, engine="access") as holder:
        holder.tables.names()
        with pytest.raises(DatabaseLockedError):
            AccessDatabase.open(target, engine="access")


def test_readonly_session(tmp_path: Path) -> None:
    target = tmp_path / "ro.accdb"
    with AccessDatabase.create(target, engine="access") as db:
        _customers(db)
    with AccessDatabase.open(target, readonly=True, engine="access") as db:
        assert db.tables.names() == ["Customers"]
        with pytest.raises(ReadOnlyError):
            db.tables.create("Other", columns=[Column.text("A")])


def test_startup_code_never_runs_by_default(tmp_path: Path) -> None:
    """An AutoExec macro calling VBA must not run when PyAccessKit opens the database (S2)."""
    target = tmp_path / "startup.accdb"
    marker = tmp_path / "marker.txt"
    with AccessDatabase.create(target, engine="access") as db:
        db.modules.create(
            "modMarker",
            "Public Function WriteMarker() As Boolean\n"
            f'    Open "{marker}" For Output As #1\n    Print #1, "ran"\n    Close #1\n'
            "    WriteMarker = True\nEnd Function\n",
        )
        db.objects.import_text(
            "macro",
            "AutoExec",
            'Version =196611\nColumnsShown =0\nBegin\n    Action ="RunCode"\n    Argument ="WriteMarker()"\nEnd\n',
        )
    with AccessDatabase.open(target, engine="access") as db:
        assert "modMarker" in db.modules.names()
        db.forms.names()
        _ = db.modules["modMarker"].code  # forces a design session (OpenCurrentDatabase)
    assert not marker.exists(), "AutoExec ran although macro_security defaults to DISABLE"


def test_ledger_is_empty_after_clean_sessions(tmp_path: Path) -> None:
    from pyaccesskit.maintenance import owned_processes

    with AccessDatabase.create(tmp_path / "l.accdb", engine="access") as db:
        pid = db.access_pid
        assert any(entry.pid == pid for entry in owned_processes())
    assert not any(entry.pid == pid for entry in owned_processes())
