"""Regression tests for issues found in code review (each test names the behaviour it protects)."""

from __future__ import annotations

import threading
from pathlib import Path
from typing import Any

import pytest

from pyaccesskit import (
    CleanupError,
    Column,
    DatabaseExistsError,
    FormSpec,
    FormView,
    IndexSpec,
    ObjectExistsError,
    PassThroughOptions,
    QuerySpec,
    SessionClosedError,
    SpecError,
    SqlSyntaxError,
    TableSpec,
    Vba,
    WrongThreadError,
)
from pyaccesskit._backends.fake import FakeBackend
from pyaccesskit._com.raw import ProxyRegistry
from pyaccesskit._ops import schema as ops
from pyaccesskit.cli._output import redact
from pyaccesskit.forms import layout_form
from pyaccesskit.forms.controls import TextBoxSpec
from pyaccesskit.forms.spec import label_name_for
from pyaccesskit.forms.vba import EventBinding, build_module
from tests.fakes import FakeComObject, FakeFactory, fake_database

pytestmark = pytest.mark.filterwarnings("ignore::pyaccesskit.errors.AccessNameWarning")


# ------------------------------------------------------------------------------- session & facade
def test_form_build_after_switching_from_inproc_dao(tmp_path: Path) -> None:
    db, factory = fake_database(tmp_path)  # auto -> in-process DAO first
    with db:
        db.tables.create("Customers", columns=[Column.text("CustomerName")])
        with db.forms.create("frmCustomers", record_source="Customers") as form:
            form.textbox("CustomerName")
        assert db.forms.names() == ["frmCustomers"]
    assert [engine.plan.kind for engine in factory.engines] == ["dao", "access"]


def test_atomic_create_never_replaces_a_database_created_meanwhile(tmp_path: Path) -> None:
    first, _ = fake_database(tmp_path)
    second, _ = fake_database(tmp_path)
    first.close()
    first.path.write_bytes(b"first")
    with pytest.raises(CleanupError) as info:
        second.close()
    assert isinstance(info.value.errors[0], DatabaseExistsError)
    assert first.path.read_bytes() == b"first"


def test_atomic_create_with_overwrite_replaces(tmp_path: Path) -> None:
    (tmp_path / "fake.accdb").write_bytes(b"old")
    db, _ = fake_database(tmp_path, overwrite=True)
    db.close()
    assert (tmp_path / "fake.accdb").read_bytes() == b"fake accdb"


def test_engine_failure_during_create_leaves_no_working_file(tmp_path: Path) -> None:
    with pytest.raises(RuntimeError, match="configuring"):
        fake_database(tmp_path, factory=FakeFactory(fail_after_create=True))
    assert list(tmp_path.iterdir()) == []


def test_close_from_another_thread_is_rejected(tmp_path: Path) -> None:
    db, factory = fake_database(tmp_path)
    caught: list[BaseException] = []

    def worker() -> None:
        try:
            db.close()
        except BaseException as exc:
            caught.append(exc)

    thread = threading.Thread(target=worker)
    thread.start()
    thread.join()
    assert len(caught) == 1
    assert isinstance(caught[0], WrongThreadError)
    assert db.is_open
    assert not factory.engines[0].closed
    db.close()  # the owning thread can still close normally
    assert db.path.exists()


def test_query_text_import_does_not_overwrite_by_default(tmp_path: Path) -> None:
    db, _ = fake_database(tmp_path, engine="access")  # type: ignore[arg-type]
    with db:
        db.queries.create("qryOne", "SELECT 1")
        with pytest.raises(ObjectExistsError):
            db.objects.import_text("query", "qryone", 'dbMemo "SQL" ="SELECT 2"\n')
        db.objects.import_text("query", "qryOne", 'dbMemo "SQL" ="SELECT 2"\n', replace=True)


# ------------------------------------------------------------------------------------ schema ops
def _pass_through(sql: str) -> QuerySpec:
    return QuerySpec(name="qPT", sql=sql, pass_through=PassThroughOptions(connect="ODBC;DSN=x"))


def test_pass_through_replacement_compares_exactly() -> None:
    backend = FakeBackend(Path("m.accdb"))
    ops.create_or_replace_query(backend, _pass_through("SELECT $$hello$$"))
    assert ops.create_or_replace_query(backend, _pass_through("SELECT $$HELLO$$"), replace=True)
    assert "$$HELLO$$" in backend.read_query("qPT").sql


def test_failed_query_type_change_keeps_the_original(monkeypatch: pytest.MonkeyPatch) -> None:
    backend = FakeBackend(Path("m.accdb"))
    ops.create_or_replace_query(backend, _pass_through("SELECT 1"))

    def reject(spec: QuerySpec) -> None:
        raise SqlSyntaxError("bad SQL", sql=spec.sql)

    monkeypatch.setattr(backend, "create_query", reject)
    with pytest.raises(SqlSyntaxError):
        ops.create_or_replace_query(backend, QuerySpec(name="qPT", sql="SELEC oops"), replace=True)
    assert backend.read_query("qPT").pass_through is not None
    assert [q.name for q in backend.list_queries()] == ["qPT"]


def test_query_type_change_succeeds() -> None:
    backend = FakeBackend(Path("m.accdb"))
    ops.create_or_replace_query(backend, _pass_through("SELECT 1"))
    ops.create_or_replace_query(backend, QuerySpec(name="qPT", sql="SELECT 2"), replace=True)
    assert backend.read_query("qPT").pass_through is None
    assert [q.name for q in backend.list_queries()] == ["qPT"]


def _table_with_code_index() -> FakeBackend:
    backend = FakeBackend(Path("m.accdb"))
    ops.create_table(
        backend,
        TableSpec(name="T", columns=[Column.text("A")], indexes=[IndexSpec.on("Code", "A")]),
    )
    return backend


def test_add_column_checks_generated_index_names_first() -> None:
    backend = _table_with_code_index()
    with pytest.raises(ObjectExistsError, match="Code"):
        ops.add_column(backend, "T", Column.text("Code", unique=True))
    assert [c.name for c in backend.read_table("T").columns] == ["A"]


def test_add_column_is_rolled_back_when_its_index_fails(monkeypatch: pytest.MonkeyPatch) -> None:
    backend = _table_with_code_index()

    def fail(table: str, index: Any) -> None:
        raise RuntimeError("index creation failed")

    monkeypatch.setattr(backend, "create_index", fail)
    with pytest.raises(RuntimeError, match="index creation failed"):
        ops.add_column(backend, "T", Column.text("Other", unique=True))
    assert [c.name for c in backend.read_table("T").columns] == ["A"]


# ---------------------------------------------------------------------------------------- CLI
def test_redact_brace_quoted_passwords() -> None:
    assert redact("ODBC;DSN=x;PWD={first;second};APP=y") == "ODBC;DSN=x;PWD=***;APP=y"
    assert redact("ODBC;PWD={a}}b;c};X=1") == "ODBC;PWD=***;X=1"
    assert redact("ODBC;DSN={my;dsn};UID=u") == "ODBC;DSN={my;dsn};UID=u"


# ------------------------------------------------------------------------------------ raw proxies
class _Collection:
    _oleobj_ = object()

    def __iter__(self) -> Any:
        return iter([FakeComObject("a"), FakeComObject("b")])


def test_suspended_raw_iterators_are_revoked() -> None:
    registry = ProxyRegistry()
    iterator = iter(registry.wrap(_Collection(), "collection"))
    first = next(iterator)
    registry.revoke_all()
    with pytest.raises(SessionClosedError):
        next(iterator)
    assert iterator.revoked
    assert first.revoked


# ------------------------------------------------------------------------------------------ forms
def test_module_declarations_precede_event_procedures() -> None:
    text = build_module(
        [EventBinding("Form", "OnLoad", "Load", Vba("counter = 1"))],
        "Option Explicit\nPrivate counter As Long\n\n' helper\nPrivate Sub Helper()\nEnd Sub\n",
    )
    assert text.index("Private counter") < text.index("Private Sub Form_Load")
    assert text.index("Private Sub Form_Load") < text.index("' helper") < text.index("Sub Helper")
    assert text.count("Option Explicit") == 1


def test_headerless_tabular_form_is_rejected_before_building() -> None:
    spec = FormSpec(
        name="frm",
        default_view=FormView.CONTINUOUS,
        header=False,
        controls=(TextBoxSpec(field="A"),),
    )
    with pytest.raises(SpecError, match="header=False"):
        layout_form(spec)
    unlabelled = spec.model_copy(update={"controls": (TextBoxSpec(field="A", label=False),)})
    assert not layout_form(unlabelled).has_header


def test_label_names_stay_within_access_limits() -> None:
    long_a, long_b = "F" * 59 + "a", "F" * 59 + "b"
    names = {label_name_for(long_a), label_name_for(long_b)}
    assert len(names) == 2
    assert all(len(name) <= 64 and name.endswith("_Label") for name in names)
    assert label_name_for("Short") == "Short_Label"
    layout_form(FormSpec(name="frm", controls=(TextBoxSpec(field=long_a),)))
