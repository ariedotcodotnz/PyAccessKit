"""The command-line interface, with the environment and the database engine faked."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest
from typer.testing import CliRunner

from pyaccesskit import AccessDatabase, Column, diagnostics
from pyaccesskit.cli import app as cli_app
from pyaccesskit.cli import cleanup as cleanup_cmd
from pyaccesskit.cli import doctor as doctor_cmd
from pyaccesskit.cli import inspection
from pyaccesskit.cli._output import redact
from pyaccesskit.errors import DatabaseLockedError, EngineUnavailableError
from tests.fakes import FakeFactory, fake_database

runner = CliRunner()


def invoke(*args: str) -> Any:
    return runner.invoke(cli_app.app, list(args))


def report(
    *, usable: bool = True, probes: tuple[diagnostics.ProbeReport, ...] = ()
) -> diagnostics.Diagnosis:
    return diagnostics.Diagnosis(
        pyaccesskit_version="0.0.test",
        python_version="3.12.0",
        python_bits=64,
        python_executable="python.exe",
        operating_system="Windows-11",
        pywin32_version="311",
        access=diagnostics.AccessInstallation(
            progid="Access.Application",
            registered_version="Access.Application.16",
            executable=r"C:\Office\MSACCESS.EXE",
            version="16.0.1.2",
            bits=32,
            click_to_run=True,
            products=("O365ProPlusRetail",),
        ),
        engines=(
            diagnostics.EngineCheck("dao", False, "bitness mismatch"),
            diagnostics.EngineCheck("access", usable, "Access 16.0.1.2, 32-bit"),
        ),
        ace_oledb=(),
        auto_engine="access" if usable else None,
        owned_processes=(),
        probes=probes,
        problems=() if usable else ("nothing works",),
        notes=("a note",),
    )


def test_version() -> None:
    result = invoke("--version")
    assert result.exit_code == 0
    assert result.output.startswith("pyaccesskit ")


def test_usage_error_exit_code() -> None:
    assert invoke("no-such-command").exit_code == 2


@pytest.mark.parametrize(("usable", "code"), [(True, 0), (False, 3)])
def test_doctor_exit_codes(monkeypatch: pytest.MonkeyPatch, usable: bool, code: int) -> None:
    monkeypatch.setattr(
        doctor_cmd.diagnostics, "diagnose", lambda probe=False: report(usable=usable)
    )
    result = invoke("doctor")
    assert result.exit_code == code
    assert "Microsoft Access" in result.output
    assert ("ready" if usable else "not usable") in result.output


def test_doctor_json(monkeypatch: pytest.MonkeyPatch) -> None:
    failed = diagnostics.ProbeReport("access", False, 1.5, "boom")
    monkeypatch.setattr(
        doctor_cmd.diagnostics, "diagnose", lambda probe=False: report(probes=(failed,))
    )
    result = invoke("doctor", "--json", "--probe")
    data = json.loads(result.output)
    assert result.exit_code == 3, "a failed probe makes the environment unusable"
    assert data["usable"] is False
    assert data["probes"][0]["detail"] == "boom"
    assert data["access"]["bits"] == 32


@pytest.fixture
def sample(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    """A fake database with a little of everything, served to ``inspect`` through AccessDatabase.open."""
    factory = FakeFactory()
    db, _ = fake_database(tmp_path, create=False, factory=factory, engine="access")  # type: ignore[arg-type]
    with db:
        db.tables.create(
            "Customers",
            columns=[
                Column.autonumber("CustomerID", primary_key=True),
                Column.text("CustomerName", length=80),
            ],
        )
        db.tables.create("Orders", columns=[Column.number("CustomerID", required=True)])
        db.relationships.create("Customers.CustomerID", "Orders.CustomerID", cascade_delete=True)
        db.queries.create("qryCustomers", "SELECT *\nFROM Customers;")
        db.modules.create("modTools", "Public Sub Hello()\nEnd Sub\n")

    def open_fake(path: Any, **kwargs: Any) -> AccessDatabase:
        assert kwargs["readonly"] is True, "inspect must never open a database for writing"
        reopened, _ = fake_database(tmp_path, create=False, readonly=True, factory=factory)
        return reopened

    monkeypatch.setattr(inspection.AccessDatabase, "open", open_fake)
    return db.path


def test_inspect_text(sample: Path) -> None:
    result = invoke("inspect", str(sample))
    assert result.exit_code == 0, result.output
    for expected in (
        "Table Customers",
        "Short Text(80)",
        "PrimaryKey",
        "CustomersOrders",
        "qryCustomers",
        "modTools",
    ):
        assert expected in result.output


def test_inspect_json(sample: Path) -> None:
    result = invoke("inspect", str(sample), "--json")
    assert result.exit_code == 0, result.output
    data = json.loads(result.output)
    assert [t["name"] for t in data["tables"]] == ["Customers", "Orders"]
    assert data["tables"][0]["spec"]["columns"][1]["length"] == 80
    assert data["relationships"][0]["cascade_delete"] is True
    assert data["queries"][0]["kind"] == "select"
    assert data["modules"] == ["modTools"]
    assert data["warnings"] == []


@pytest.mark.parametrize(
    ("error", "code"),
    [
        (DatabaseLockedError("in use", path=Path("x.accdb")), 1),
        (EngineUnavailableError("no engine", diagnosis="why"), 3),
    ],
)
def test_inspect_errors(monkeypatch: pytest.MonkeyPatch, error: Exception, code: int) -> None:
    def fail(*_args: Any, **_kwargs: Any) -> AccessDatabase:
        raise error

    monkeypatch.setattr(inspection.AccessDatabase, "open", fail)
    result = invoke("inspect", "x.accdb")
    assert result.exit_code == code
    assert "error:" in result.output


def test_cleanup_with_empty_ledger() -> None:
    result = invoke("cleanup", "--dry-run")
    assert result.exit_code == 0
    assert "Nothing to do" in result.output


def test_cleanup_reports_failures(monkeypatch: pytest.MonkeyPatch) -> None:
    from pyaccesskit._ledger import OwnedProcess, ReapResult

    entry = OwnedProcess(1234, 1.0, "MSACCESS.EXE", 99, 2.0, None, 0.0, "test")
    monkeypatch.setattr(
        cleanup_cmd, "reap_orphans", lambda dry_run=False: [ReapResult(entry, "failed")]
    )
    result = invoke("cleanup", "--json")
    assert result.exit_code == 1
    assert json.loads(result.output) == [
        {"pid": 1234, "owner_pid": 99, "database": None, "action": "failed"}
    ]


def test_redact() -> None:
    assert redact("ODBC;DSN=x;UID=me;PWD=secret;APP=y") == "ODBC;DSN=x;UID=me;PWD=***;APP=y"
    assert redact(";Password=hunter2") == ";Password=***"
    assert redact(None) is None
