"""Behaviour every SchemaBackend must share (run against the fake and, with --integration, real Access)."""

from __future__ import annotations

import warnings
from datetime import date
from decimal import Decimal

import pytest

from pyaccesskit._backends.protocols import PropertyTarget
from pyaccesskit._ops import schema as ops
from pyaccesskit.enums import JoinType, NumberSize, ObjectKind, QueryKind
from pyaccesskit.errors import (
    AccessNameWarning,
    ObjectExistsError,
    ObjectNotFoundError,
    RelationshipError,
    SchemaError,
)
from pyaccesskit.schema import (
    Column,
    Expr,
    IndexSpec,
    QuerySpec,
    RelationshipSpec,
    TableSpec,
    sql_equivalent,
)

from .conftest import Backends

pytestmark = pytest.mark.filterwarnings("ignore::pyaccesskit.errors.AccessNameWarning")


def customers() -> TableSpec:
    return TableSpec(
        name="Customers",
        description="People we sell to",
        columns=[
            Column.autonumber("CustomerID", primary_key=True),
            Column.text("CustomerName", length=200, required=True, caption="Customer name"),
            Column.text("Email", length=255, unique=True),
            Column.text("City", indexed=True),
            Column.currency(
                "CreditLimit",
                default=Decimal("1000"),
                validation_rule=">=0",
                validation_text="Not negative",
            ),
            Column.date_time("CreatedAt", default=Expr("Now()"), format="Short Date"),
            Column.yes_no("IsActive", default=True),
            Column.number("Score", size=NumberSize.INTEGER, default=0, description="0-100"),
            Column.number("Ratio", size=NumberSize.DOUBLE),
            Column.long_text("Notes"),
        ],
    )


def orders() -> TableSpec:
    return TableSpec(
        name="Orders",
        columns=[
            Column.autonumber("OrderID", primary_key=True),
            Column.number("CustomerID", required=True),
            Column.date_time("OrderDate", default=date(2026, 1, 1)),
            Column.text("Status", length=20, default="New"),
        ],
    )


# ------------------------------------------------------------------------------------------ tables
def test_create_and_read_back_round_trip(backends: Backends) -> None:
    created = ops.create_table(backends.schema, customers())
    assert created == customers().normalized()
    assert [t.name for t in backends.schema.list_tables() if not t.is_system] == ["Customers"]
    assert backends.schema.read_table("customers") == customers().normalized()


def test_names_are_case_insensitive_and_shared_with_queries(backends: Backends) -> None:
    ops.create_table(backends.schema, customers())
    with pytest.raises(ObjectExistsError, match="already exists"):
        ops.create_table(backends.schema, TableSpec(name="CUSTOMERS", columns=[Column.text("A")]))
    with pytest.raises(ObjectExistsError, match="table"):
        ops.create_or_replace_query(backends.schema, QuerySpec(name="customers", sql="SELECT 1;"))
    ops.create_or_replace_query(
        backends.schema, QuerySpec(name="qryOne", sql="SELECT * FROM Customers;")
    )
    with pytest.raises(ObjectExistsError, match="share one namespace"):
        ops.create_table(backends.schema, TableSpec(name="QRYONE", columns=[Column.text("A")]))


def test_missing_objects(backends: Backends) -> None:
    with pytest.raises(ObjectNotFoundError) as info:
        ops.drop_table(backends.schema, "Nope")
    assert info.value.kind is ObjectKind.TABLE


def test_add_drop_rename_columns(backends: Backends) -> None:
    ops.create_table(backends.schema, orders())
    ops.add_column(backends.schema, "Orders", Column.text("Reference", length=40, unique=True))
    spec = backends.schema.read_table("Orders")
    assert spec.column_names[-1] == "Reference"
    assert any(i.name == "Reference" and i.unique for i in spec.indexes)

    with pytest.raises(ObjectExistsError):
        ops.add_column(backends.schema, "Orders", Column.text("status"))
    with pytest.raises(SchemaError, match="index"):
        ops.drop_column(backends.schema, "Orders", "Reference")

    ops.drop_index(backends.schema, "Orders", "Reference")
    ops.rename_column(backends.schema, "Orders", "Reference", "ExternalRef")
    ops.drop_column(backends.schema, "Orders", "Status")
    assert backends.schema.read_table("Orders").column_names == (
        "OrderID",
        "CustomerID",
        "OrderDate",
        "ExternalRef",
    )


def test_rename_column_keeps_indexes(backends: Backends) -> None:
    ops.create_table(backends.schema, customers())
    ops.rename_column(backends.schema, "Customers", "Email", "EmailAddress")
    spec = backends.schema.read_table("Customers")
    email_index = next(i for i in spec.indexes if i.unique and not i.primary)
    assert email_index.field_names == ("EmailAddress",)


def test_indexes(backends: Backends) -> None:
    ops.create_table(backends.schema, orders())
    ops.create_index(
        backends.schema,
        "Orders",
        IndexSpec.on("ix_CustomerDate", "CustomerID", ("OrderDate", "desc")),
    )
    spec = backends.schema.read_table("Orders")
    index = next(i for i in spec.indexes if i.name == "ix_CustomerDate")
    assert [(f.name, f.descending) for f in index.fields] == [
        ("CustomerID", False),
        ("OrderDate", True),
    ]
    with pytest.raises(ObjectExistsError):
        ops.create_index(backends.schema, "Orders", IndexSpec.on("IX_CUSTOMERDATE", "Status"))
    with pytest.raises(SchemaError, match="primary key"):
        ops.create_index(backends.schema, "Orders", IndexSpec.primary_key("Status"))
    with pytest.raises(ObjectNotFoundError):
        ops.create_index(backends.schema, "Orders", IndexSpec.on("ix_bad", "Missing"))


def test_rename_and_drop_table(backends: Backends) -> None:
    ops.create_table(backends.schema, orders())
    ops.rename_table(backends.schema, "Orders", "SalesOrders")
    assert [t.name for t in backends.schema.list_tables() if not t.is_system] == ["SalesOrders"]
    ops.drop_table(backends.schema, "salesorders")
    assert [t.name for t in backends.schema.list_tables() if not t.is_system] == []


def test_lint_warnings_on_create(backends: Backends) -> None:
    with pytest.warns(AccessNameWarning, match="reserved word"):
        ops.create_table(backends.schema, TableSpec(name="People", columns=[Column.text("Name")]))
    with warnings.catch_warnings():
        warnings.simplefilter("error", AccessNameWarning)
        ops.create_table(
            backends.schema, TableSpec(name="Places", columns=[Column.text("PlaceName")]), lint=True
        )


# ----------------------------------------------------------------------------------- relationships
def test_relationships(backends: Backends) -> None:
    ops.create_table(backends.schema, customers())
    ops.create_table(backends.schema, orders())
    created = ops.create_relationship(
        backends.schema,
        RelationshipSpec.between(
            "customers.customerid", "orders.customerid", cascade_delete=True, join=JoinType.LEFT
        ),
    )
    assert created.name == "CustomersOrders"
    assert created.primary_table == "Customers" and created.foreign_columns == ("CustomerID",)
    assert backends.schema.list_relationships() == [created]

    with pytest.raises(ObjectExistsError, match="name="):
        ops.create_relationship(
            backends.schema, RelationshipSpec.between("Customers.CustomerID", "Orders.CustomerID")
        )
    with pytest.raises(SchemaError, match="relationship"):
        ops.drop_table(backends.schema, "Customers")
    with pytest.raises(SchemaError, match="relationship"):
        ops.drop_column(backends.schema, "Orders", "CustomerID")

    ops.drop_relationship(backends.schema, "customersorders")
    assert backends.schema.list_relationships() == []


def test_relationship_validation(backends: Backends) -> None:
    ops.create_table(backends.schema, customers())
    ops.create_table(backends.schema, orders())
    with pytest.raises(RelationshipError, match="storage type"):
        ops.create_relationship(
            backends.schema,
            RelationshipSpec.between("Customers.CustomerID", "Orders.Status", name="bad1"),
        )
    with pytest.raises(RelationshipError, match="unique index"):
        ops.create_relationship(
            backends.schema,
            RelationshipSpec.between("Customers.City", "Orders.Status", name="bad2"),
        )
    with pytest.raises(ObjectNotFoundError):
        ops.create_relationship(
            backends.schema, RelationshipSpec.between("Customers.Nope", "Orders.CustomerID")
        )


def test_drop_table_with_relationships_option(backends: Backends) -> None:
    ops.create_table(backends.schema, customers())
    ops.create_table(backends.schema, orders())
    ops.create_relationship(
        backends.schema, RelationshipSpec.between("Customers.CustomerID", "Orders.CustomerID")
    )
    ops.drop_table(backends.schema, "Orders", drop_relationships=True)
    assert backends.schema.list_relationships() == []


# ----------------------------------------------------------------------------------------- queries
def test_queries(backends: Backends) -> None:
    ops.create_table(backends.schema, customers())
    assert ops.create_or_replace_query(
        backends.schema, QuerySpec(name="qryActive", sql="SELECT * FROM Customers WHERE IsActive")
    )
    stored = backends.schema.read_query("qryactive")
    assert stored.name == "qryActive"
    assert sql_equivalent(stored.sql, "SELECT * FROM Customers WHERE IsActive")
    infos = {q.name: q for q in backends.schema.list_queries()}
    assert infos["qryActive"].kind is QueryKind.SELECT

    assert not ops.create_or_replace_query(
        backends.schema,
        QuerySpec(name="qryActive", sql="select *\nfrom customers where isactive;"),
        replace=True,
    )
    assert ops.create_or_replace_query(
        backends.schema,
        QuerySpec(name="qryActive", sql="SELECT CustomerName FROM Customers"),
        replace=True,
    )
    with pytest.raises(ObjectExistsError):
        ops.create_or_replace_query(backends.schema, QuerySpec(name="qryActive", sql="SELECT 1"))

    ops.rename_query(backends.schema, "qryActive", "qryNames")
    ops.drop_query(backends.schema, "qrynames")
    assert backends.schema.list_queries() == []


def test_query_parameters(backends: Backends) -> None:
    ops.create_table(backends.schema, customers())
    ops.create_or_replace_query(
        backends.schema,
        QuerySpec(
            name="qryParams",
            sql="PARAMETERS [pMin] Long, [pName] Text ( 255 ); SELECT * FROM Customers WHERE Score > [pMin] AND CustomerName <> [pName];",
        ),
    )
    params = backends.schema.query_parameters("qryParams")
    assert [p.name for p in params] == ["pMin", "pName"]


def test_action_query_kinds(backends: Backends) -> None:
    ops.create_table(backends.schema, customers())
    ops.create_or_replace_query(
        backends.schema,
        QuerySpec(name="qryDeactivate", sql="UPDATE Customers SET IsActive = False"),
    )
    ops.create_or_replace_query(
        backends.schema,
        QuerySpec(name="qryPurge", sql="DELETE FROM Customers WHERE IsActive = False"),
    )
    kinds = {q.name: q.kind for q in backends.schema.list_queries()}
    assert kinds == {"qryDeactivate": QueryKind.UPDATE, "qryPurge": QueryKind.DELETE}


# -------------------------------------------------------------------------------------- properties
def test_custom_and_mapped_properties(backends: Backends) -> None:
    ops.create_table(backends.schema, customers())
    table = PropertyTarget.table("Customers")
    assert backends.schema.get_property(table, "Description") == "People we sell to"
    backends.schema.set_property(table, "Description", "Changed")
    assert backends.schema.read_table("Customers").description == "Changed"

    field = PropertyTarget.column("Customers", "Email")
    backends.schema.set_property(field, "Caption", "E-mail")
    assert backends.schema.read_table("Customers").column("Email").caption == "E-mail"

    backends.schema.set_property(table, "PakCustomTag", "hello")
    assert backends.schema.get_property(table, "pakcustomtag") == "hello"
    backends.schema.delete_property(table, "PakCustomTag")
    with pytest.raises(ObjectNotFoundError):
        backends.schema.get_property(table, "PakCustomTag")
