"""The public API end-to-end over the in-memory engine (no Access required)."""

from __future__ import annotations

from pathlib import Path

import pytest

from pyaccesskit import (
    Column,
    ControlKind,
    ModuleKind,
    ObjectExistsError,
    ObjectNotFoundError,
    PropertyType,
    RelationshipSpec,
    SchemaError,
    TableSpec,
)
from pyaccesskit.errors import AccessNameWarning
from tests.fakes import fake_database

pytestmark = pytest.mark.filterwarnings("ignore::pyaccesskit.errors.AccessNameWarning")


def test_headline_example(tmp_path: Path) -> None:
    db, _ = fake_database(tmp_path, engine="access")  # type: ignore[arg-type]
    with db:
        customers = db.tables.create(
            "Customers",
            columns=[
                Column.autonumber("CustomerID", primary_key=True),
                Column.text("Name", length=200, required=True),
                Column.text("Email", length=255),
            ],
        )
        db.queries.create("qryCustomers", "SELECT *\nFROM Customers\nORDER BY Name;")
        with db.forms.create("frmCustomers", record_source="Customers") as form:
            form.textbox("Name", label="Customer name")
        assert customers.name == "Customers"
        assert db.tables.names() == ["Customers"]
        assert "customers" in db.tables
        assert [f.name for f in customers.fields] == ["CustomerID", "Name", "Email"]
        assert customers.fields["email"].size == 255
        assert customers.primary_key is not None
        assert db.queries["qrycustomers"].kind.value == "select"
        assert db.forms.names() == ["frmCustomers"]
        assert {c.kind for c in db.forms["frmCustomers"].controls()} == {
            ControlKind.TEXTBOX,
            ControlKind.LABEL,
        }


def test_reserved_word_warning_points_at_user_code(tmp_path: Path) -> None:
    db, _ = fake_database(tmp_path)
    with db, pytest.warns(AccessNameWarning, match="reserved word") as record:
        db.tables.create("People", columns=[Column.text("Name")])
    assert Path(record[0].filename).name == "test_api_fake.py"


def test_table_handles(tmp_path: Path) -> None:
    db, _ = fake_database(tmp_path)
    with db:
        table = db.tables.create(TableSpec(name="T", columns=[Column.text("A"), Column.text("B")]))
        table.add_column(Column.number("C", unique=True))
        table.rename_column("B", "B2")
        table.fields["B2"].rename("B3")
        table.description = "desc"
        assert table.description == "desc"
        table.properties.set("PakTag", 5, PropertyType.LONG)
        assert table.properties["PakTag"] == 5 and "paktag" in table.properties
        assert table.properties.get("missing", "dflt") == "dflt"
        table.drop_index("C")
        table.drop_column("C")
        table.rename("T2")
        assert db.tables.names() == ["T2"]
        assert [i.name for i in db.tables["T2"].indexes] == []
        with pytest.raises(ObjectNotFoundError):
            db.tables["nope"]
        assert db.tables.get("nope") is None
        db.tables.drop("T2")
        assert len(db.tables) == 0


def test_relationships_and_queries(tmp_path: Path) -> None:
    db, _ = fake_database(tmp_path)
    with db:
        db.tables.create("A", columns=[Column.autonumber("ID", primary_key=True)])
        db.tables.create("B", columns=[Column.number("AID")])
        rel = db.relationships.create("A.ID", "B.AID", cascade_delete=True)
        assert rel.name == "AB" and rel.to_spec().cascade_delete
        with pytest.raises(ObjectExistsError):
            db.relationships.create(RelationshipSpec.between("A.ID", "B.AID"))
        with pytest.raises(SchemaError):
            db.tables.drop("A")
        db.tables.drop("A", drop_relationships=True)
        assert db.relationships.names() == []

        query = db.queries.create("q1", "SELECT 1")
        query.rename("q2")
        assert db.queries.names() == ["q2"]
        assert db.queries.create("q2", "SELECT 2", replace=True).sql.startswith("SELECT 2")
        db.queries.drop("q2")
        assert len(db.queries) == 0


def test_modules_and_objects(tmp_path: Path) -> None:
    db, _ = fake_database(tmp_path, engine="access")  # type: ignore[arg-type]
    with db:
        module = db.modules.create("clsThing", "Public Title As String\n", kind=ModuleKind.CLASS)
        assert module.kind is ModuleKind.CLASS
        assert "Public Title As String" in module.code
        module.code = "Public Other As Long\n"
        assert "Public Other As Long" in db.modules["clsthing"].code
        text = db.objects.export_text("module", "clsThing")
        db.objects.import_text("module", "clsCopy", text)
        assert db.modules["clsCopy"].kind is ModuleKind.CLASS
        saved = db.objects.save_text("module", "clsCopy", tmp_path / "clsCopy.cls")
        assert saved.read_text(encoding="utf-8").startswith("Attribute VB_GlobalNameSpace")
        db.objects.load_text("module", "clsThird", saved)
        assert sorted(db.objects.names("module")) == ["clsCopy", "clsThing", "clsThird"]
        db.modules.drop("clsThird")
        db.objects.rename("module", "clsCopy", "clsRenamed")
        db.objects.delete("module", "clsRenamed")
        assert db.modules.names() == ["clsThing"]


def test_repr_and_facts(tmp_path: Path) -> None:
    db, _ = fake_database(tmp_path)
    assert "fake.accdb" in repr(db) and db.is_open and not db.readonly
    assert db.access_pid is None
    db.close()
    assert "closed" in repr(db)
