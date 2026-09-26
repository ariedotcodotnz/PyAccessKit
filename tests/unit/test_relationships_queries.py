from __future__ import annotations

import pytest
from pydantic import ValidationError

from pyaccesskit.enums import JoinType, QueryKind
from pyaccesskit.errors import SpecError
from pyaccesskit.schema import (
    PassThroughOptions,
    QuerySpec,
    RelationshipSpec,
    detect_query_kind,
    normalize_sql,
    sql_equivalent,
)
from pyaccesskit.schema.queries import DAO_QUERY_KINDS


# ------------------------------------------------------------------------------------ relationships
def test_between_parses_table_column_references() -> None:
    rel = RelationshipSpec.between("Customers.CustomerID", "Orders.CustomerID", cascade_delete=True)
    assert (rel.primary_table, rel.primary_columns) == ("Customers", ("CustomerID",))
    assert (rel.foreign_table, rel.foreign_columns) == ("Orders", ("CustomerID",))
    assert rel.effective_name == "CustomersOrders"
    assert rel.normalized().name == "CustomersOrders"
    assert rel.cascade_delete


def test_composite_references() -> None:
    rel = RelationshipSpec.between(
        ("Orders", ["OrderID", "Rev"]), ("Lines", ("OrderID", "Rev")), name="OrdersLines"
    )
    assert rel.primary_columns == ("OrderID", "Rev")
    assert rel.normalized() is rel


@pytest.mark.parametrize("ref", ["NoDot", "a.b.c", ".x", "x.", 5])
def test_bad_references(ref: object) -> None:
    with pytest.raises(SpecError):
        RelationshipSpec.between(ref, "Orders.CustomerID")  # type: ignore[arg-type]


def test_relationship_validation() -> None:
    with pytest.raises(ValidationError, match="pair columns one to one"):
        RelationshipSpec(
            primary_table="A", primary_columns=("x", "y"), foreign_table="B", foreign_columns=("x",)
        )
    with pytest.raises(ValidationError, match="require enforce_integrity"):
        RelationshipSpec(
            primary_table="A",
            primary_columns=("x",),
            foreign_table="B",
            foreign_columns=("x",),
            enforce_integrity=False,
            cascade_delete=True,
        )


def test_relationship_json_round_trip() -> None:
    rel = RelationshipSpec.between("A.x", "B.y", join=JoinType.LEFT, one_to_one=True)
    assert RelationshipSpec.model_validate_json(rel.model_dump_json()) == rel


# ------------------------------------------------------------------------------------------ queries
def test_query_spec_normalizes_line_endings() -> None:
    spec = QuerySpec(name="q", sql="SELECT *\nFROM T\rWHERE 1=1")
    assert spec.normalized().sql == "SELECT *\r\nFROM T\r\nWHERE 1=1"


def test_query_spec_validation() -> None:
    with pytest.raises(ValidationError):
        QuerySpec(name="q", sql="   ")
    with pytest.raises(ValidationError, match="ODBC;"):
        PassThroughOptions(connect="DSN=x")


@pytest.mark.parametrize(
    ("left", "right"),
    [
        (
            "SELECT * FROM Customers ORDER BY Name;",
            "SELECT *\r\nFROM Customers\r\nORDER BY Name;\r\n",
        ),
        (
            "select id from customers where score>1",
            "SELECT id\r\nFROM customers\r\nWHERE score>1;\r\n",
        ),
        (
            "DELETE FROM Customers WHERE Score < 0;",
            "DELETE *\r\nFROM Customers\r\nWHERE Score < 0;\r\n",
        ),
        (
            "INSERT INTO Customers (CustomerName) SELECT CustomerName FROM Customers WHERE ID = 0;",
            "INSERT INTO Customers ( CustomerName )\r\nSELECT CustomerName\r\nFROM Customers\r\nWHERE ID = 0;\r\n",
        ),
        (
            "SELECT [Order Details].Qty FROM [Order Details]",
            "SELECT [ORDER DETAILS].qty FROM [order details];",
        ),
    ],
)
def test_sql_equivalent_sees_through_access_rewrites(left: str, right: str) -> None:
    assert sql_equivalent(left, right)


def test_sql_equivalent_keeps_literals_significant() -> None:
    assert not sql_equivalent("SELECT 'abc'", "SELECT 'ABC'")
    assert not sql_equivalent("SELECT * FROM A", "SELECT * FROM B")
    assert normalize_sql("SELECT  #2026-01-01#  ;") == "SELECT #2026-01-01#"


@pytest.mark.parametrize(
    ("sql", "kind"),
    [
        ("SELECT * FROM T", QueryKind.SELECT),
        ("SELECT * INTO T2 FROM T", QueryKind.MAKE_TABLE),
        ("SELECT A FROM T UNION SELECT A FROM U", QueryKind.UNION),
        ("SELECT (SELECT 1 UNION SELECT 2) AS X FROM T", QueryKind.SELECT),
        ("UPDATE T SET A = 1", QueryKind.UPDATE),
        ("DELETE FROM T", QueryKind.DELETE),
        ("INSERT INTO T (A) VALUES (1)", QueryKind.APPEND),
        ("TRANSFORM Sum(A) SELECT B FROM T GROUP BY B PIVOT C", QueryKind.CROSSTAB),
        ("CREATE TABLE Z (A LONG)", QueryKind.DDL),
        ("PARAMETERS [p] Long; SELECT * FROM T WHERE A > [p]", QueryKind.SELECT),
        ("PARAMETERS [p] Long; UPDATE T SET A = [p]", QueryKind.UPDATE),
        ("EXEC something", QueryKind.UNKNOWN),
        ("", QueryKind.UNKNOWN),
    ],
)
def test_detect_query_kind(sql: str, kind: QueryKind) -> None:
    assert detect_query_kind(sql) is kind


def test_pass_through_kind() -> None:
    spec = QuerySpec(
        name="pt", sql="SELECT 1", pass_through=PassThroughOptions(connect="ODBC;DSN=x")
    )
    assert spec.kind is QueryKind.PASS_THROUGH


def test_dao_query_type_table_matches_typelib_values() -> None:
    assert DAO_QUERY_KINDS[0] is QueryKind.SELECT
    assert DAO_QUERY_KINDS[112] is QueryKind.PASS_THROUGH
    assert DAO_QUERY_KINDS[128] is QueryKind.UNION
    assert len(DAO_QUERY_KINDS) == 13
