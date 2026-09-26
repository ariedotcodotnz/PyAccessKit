"""Tier 4: what PyAccessKit writes is exactly what it reads back — after closing and reopening the file."""

from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pytest

from pyaccesskit import (
    AccessDatabase,
    Column,
    Expr,
    IndexSpec,
    MissingParameterError,
    NumberSize,
    ObjectExistsError,
    SpecError,
    SqlSyntaxError,
    TableSpec,
    Transport,
)
from pyaccesskit.schema import sql_equivalent

pytestmark = [
    pytest.mark.integration,
    pytest.mark.access,
    pytest.mark.filterwarnings("ignore::pyaccesskit.errors.AccessNameWarning"),
]


def everything() -> TableSpec:
    """One column of every creatable type, with every option set to a non-default value somewhere."""
    return TableSpec(
        name="Everything",
        description="All column types",
        validation_rule="[Qty] >= 0 Or [Qty] Is Null",
        validation_text="Qty cannot be negative",
        columns=[
            Column.autonumber("ID"),
            Column.text(
                "Code",
                length=12,
                required=True,
                allow_zero_length=True,
                unicode_compression=False,
                input_mask=">LLL-000",
                caption="Product code",
                description="The SKU",
                format=">",
                default="AAA-000",
                validation_rule='Like "???-###"',
                validation_text="Use the AAA-000 format",
            ),
            Column.long_text("Notes", rich_text=True),
            Column.long_text("History", append_only=True),
            Column.number("Tiny", size=NumberSize.BYTE, default=1),
            Column.number("Qty", size=NumberSize.INTEGER, default=0, decimal_places=0),
            Column.number("Big", size=NumberSize.LONG_INTEGER),
            Column.number("Ratio", size=NumberSize.SINGLE, default=0.5, decimal_places=2),
            Column.number("Precise", size=NumberSize.DOUBLE, format="Percent"),
            Column.number("Ref", size=NumberSize.REPLICATION_ID),
            Column.decimal("Rate", precision=12, scale=4, default=Decimal("1.2500")),
            Column.currency("Price", default=Decimal("9.99"), decimal_places=2),
            Column.date_time("Created", default=Expr("Now()"), format="Short Date"),
            Column.date_time("Launch", default=date(2026, 1, 31)),
            Column.date_time("Stamp", default=datetime(2026, 1, 31, 13, 45, 0)),
            Column.yes_no("Active", default=False),
            Column.hyperlink("Website"),
            Column.ole_object("Blob"),
        ],
        indexes=[
            IndexSpec.primary_key("ID"),
            IndexSpec.on("ix_Code", "Code", unique=True),
            IndexSpec.on("ix_Recent", ("Created", "desc"), "Code", ignore_nulls=True),
        ],
    )


def test_every_column_type_round_trips(tmp_path: Path) -> None:
    path = tmp_path / "types.accdb"
    spec = everything()
    with AccessDatabase.create(path, engine="access") as db:
        db.tables.create(spec)
        assert db.tables["Everything"].to_spec() == spec.normalized()
    with AccessDatabase.open(path, readonly=True, engine="access") as db:
        assert db.transport is Transport.ACCESS_HOSTED
        assert db.tables["everything"].to_spec() == spec.normalized()


def test_autonumber_variants_round_trip(tmp_path: Path) -> None:
    path = tmp_path / "auto.accdb"
    specs = [
        TableSpec(name="Inc", columns=[Column.autonumber("ID", primary_key=True)]),
        TableSpec(
            name="Guid", columns=[Column.autonumber("ID", replication_id=True, primary_key=True)]
        ),
    ]
    with AccessDatabase.create(path, engine="access") as db:
        for spec in specs:
            db.tables.create(spec)
    with AccessDatabase.open(path, engine="access") as db:
        for spec in specs:
            assert db.tables[spec.name].to_spec() == spec.normalized()
        db.execute("INSERT INTO Inc (ID) VALUES (1)")
        assert db.fetch_all("SELECT ID FROM Inc") == [{"ID": 1}]


def test_composite_keys_and_relationships_round_trip(tmp_path: Path) -> None:
    path = tmp_path / "rel.accdb"
    with AccessDatabase.create(path, engine="access") as db:
        db.tables.create(
            "Orders",
            columns=[Column.number("OrderID"), Column.number("Rev")],
            primary_key=("OrderID", "Rev"),
        )
        db.tables.create(
            "Lines",
            columns=[
                Column.autonumber("LineID", primary_key=True),
                Column.number("OrderID"),
                Column.number("Rev"),
            ],
        )
        db.relationships.create(
            ("Orders", ["OrderID", "Rev"]), ("Lines", ["OrderID", "Rev"]), cascade_update=True
        )
        expected = db.relationships.specs()
    with AccessDatabase.open(path, engine="access") as db:
        assert db.relationships.specs() == expected
        assert expected[0].primary_columns == ("OrderID", "Rev")
        assert db.tables["Orders"].primary_key is not None
        # the hidden relationship index counts towards the 32-index budget but is not reported as an index
        assert [i.name for i in db.tables["Lines"].indexes] == ["PrimaryKey"]


def test_schema_changes_survive_reopen(tmp_path: Path) -> None:
    path = tmp_path / "alter.accdb"
    with AccessDatabase.create(path, engine="access") as db:
        table = db.tables.create(
            "People", columns=[Column.autonumber("ID", primary_key=True), Column.text("First")]
        )
        table.add_column(Column.text("Last", indexed=True))
        table.add_column(Column.decimal("Score", precision=5, scale=1))
        table.rename_column("First", "GivenName")
        table.description = "Everyone"
        table.fields["Last"].properties["Caption"] = "Surname"
        table.rename("Persons")
    with AccessDatabase.open(path, engine="access") as db:
        spec = db.tables["Persons"].to_spec()
        assert spec.column_names == ("ID", "GivenName", "Last", "Score")
        assert spec.description == "Everyone"
        assert spec.column("Last").caption == "Surname"
        assert spec.column("Score").model_dump()["precision"] == 5
        assert any(i.name == "Last" for i in spec.indexes)


def test_queries_and_data(tmp_path: Path) -> None:
    path = tmp_path / "data.accdb"
    with AccessDatabase.create(path, engine="access") as db:
        db.tables.create(
            "Items",
            columns=[
                Column.autonumber("ID", primary_key=True),
                Column.text("Name"),
                Column.currency("Price"),
                Column.date_time("Added"),
                Column.yes_no("Active"),
            ],
        )
        insert = "INSERT INTO Items (Name, Price, Added, Active) VALUES ([n], [p], [d], [a])"
        assert (
            db.execute(insert, {"n": "Pen", "p": Decimal("1.50"), "d": date(2026, 1, 2), "a": True})
            == 1
        )
        assert (
            db.execute(
                insert,
                {"n": "Pad", "p": Decimal("3.25"), "d": datetime(2026, 1, 3, 9, 30), "a": False},
            )
            == 1
        )
        rows = db.fetch_all("SELECT Name, Price, Added, Active FROM Items ORDER BY ID")
        assert rows == [
            {"Name": "Pen", "Price": Decimal("1.5"), "Added": datetime(2026, 1, 2), "Active": True},
            {
                "Name": "Pad",
                "Price": Decimal("3.25"),
                "Added": datetime(2026, 1, 3, 9, 30),
                "Active": False,
            },
        ]
        assert db.fetch_all("SELECT * FROM Items", limit=1)[0]["Name"] == "Pen"

        query = db.queries.create(
            "qryActive",
            "PARAMETERS [pMin] Currency; SELECT Name FROM Items WHERE Active AND Price >= [pMin]",
        )
        assert query.fetch({"pMin": 1}) == [{"Name": "Pen"}]
        assert [p.name for p in query.parameters] == ["pMin"]
        with pytest.raises(MissingParameterError):
            query.fetch()
        with pytest.raises(SpecError, match="unknown parameter"):
            query.fetch({"pMin": 1, "nope": 2})
        with pytest.raises(SqlSyntaxError):
            db.execute("UPDATE Items SET WHERE")

        update = db.queries.create(
            "qryDeactivate", "UPDATE Items SET Active = False WHERE Name = [who]"
        )
        assert update.execute({"who": "Pen"}) == 1
        assert db.tables["Items"].record_count() == 2
        query.sql = "SELECT Name FROM Items ORDER BY Name"
        stored = query.sql
    with AccessDatabase.open(path, readonly=True, engine="access") as db:
        assert sql_equivalent(db.queries["qryActive"].sql, stored)
        assert db.queries["qryDeactivate"].kind.value == "update"


def test_dao_created_database_gets_native_defaults(tmp_path: Path) -> None:
    """In-process DAO needs a bitness-matched Python; the Access engine checks the profile's effect instead."""
    path = tmp_path / "native.accdb"
    with AccessDatabase.create(path, engine="access") as db:
        assert db.properties["UseMDIMode"] == 0
        assert db.properties["ShowDocumentTabs"] is True


def test_table_rule_on_a_decimal_column_and_required_ole(tmp_path: Path) -> None:
    """The table rule is installed only after ADO has added the Decimal column it refers to."""
    path = tmp_path / "rules.accdb"
    spec = TableSpec(
        name="Priced",
        columns=[
            Column.autonumber("ID", primary_key=True),
            Column.decimal("Amount", precision=10, scale=2),
            Column.ole_object("Payload", required=True),
        ],
        validation_rule="[Amount] >= 0",
        validation_text="Amount cannot be negative",
    )
    with AccessDatabase.create(
        path
    ) as db:  # auto: covers in-process DAO when this Python can load it
        db.tables.create(spec)
    with AccessDatabase.open(path, readonly=True) as db:
        assert db.tables["Priced"].to_spec() == spec.normalized()


def test_add_column_with_clashing_index_name_changes_nothing(tmp_path: Path) -> None:
    with AccessDatabase.create(tmp_path / "clash.accdb") as db:
        table = db.tables.create(
            "T", columns=[Column.text("A")], indexes=[IndexSpec.on("Code", "A")]
        )
        with pytest.raises(ObjectExistsError):
            table.add_column(Column.text("Code", unique=True))
        assert table.to_spec().column_names == ("A",)
        table.add_column(Column.text("Code"))  # retrying without the shorthand works


def test_binary_parameters_round_trip_exactly(tmp_path: Path) -> None:
    """Implicit query parameters are text in DAO; bytes bound to them used to be corrupted (ADR 0002)."""
    data = bytes(range(256)) * 3 + b"end"
    with AccessDatabase.create(tmp_path / "blob.accdb") as db:
        db.tables.create(
            "Files",
            columns=[Column.autonumber("FileID", primary_key=True), Column.ole_object("Payload")],
        )
        db.execute("INSERT INTO Files (Payload) VALUES ([payload])", {"payload": data})
        assert db.fetch_all("SELECT Payload FROM Files") == [{"Payload": data}]
        db.queries.create("qryAddFile", "INSERT INTO Files (Payload) VALUES ([payload])")
        with pytest.raises(SpecError, match="LongBinary"):
            db.queries["qryAddFile"].execute({"payload": data})


def test_fresh_database_lists_no_design_objects(tmp_path: Path) -> None:
    """DAO-created files lack the Forms/Reports/... containers until Access has opened them."""
    with AccessDatabase.create(tmp_path / "fresh.accdb") as db:
        db.tables.create("T", columns=[Column.text("A")])
    with AccessDatabase.open(tmp_path / "fresh.accdb", readonly=True) as db:
        for kind in ("form", "report", "macro", "module"):
            assert db.objects.names(kind) == []
