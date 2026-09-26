from __future__ import annotations

import pytest
from pydantic import ValidationError

from pyaccesskit.errors import SpecError
from pyaccesskit.schema import Column, IndexField, IndexSpec, TableSpec


def customers() -> TableSpec:
    return TableSpec(
        name="Customers",
        columns=[
            Column.autonumber("CustomerID", primary_key=True),
            Column.text("CustomerName", length=200, required=True),
            Column.text("Email", unique=True),
            Column.text("City", indexed=True),
        ],
    )


def test_shorthands_expand_into_canonical_indexes() -> None:
    spec = customers().normalized()
    assert [(i.name, i.field_names, i.primary, i.unique) for i in spec.indexes] == [
        ("PrimaryKey", ("CustomerID",), True, True),
        ("City", ("City",), False, False),
        ("Email", ("Email",), False, True),
    ]
    assert all(not (c.primary_key or c.unique or c.indexed) for c in spec.columns)
    assert spec.primary_key is None
    assert spec.normalized() == spec


def test_primary_index_implies_unique_and_required() -> None:
    index = IndexSpec.primary_key("A", "B")
    assert (index.primary, index.unique, index.required) == (True, True, True)
    assert index.field_names == ("A", "B")


def test_composite_primary_key_shorthand() -> None:
    spec = TableSpec(
        name="OrderLines",
        columns=[Column.number("OrderID"), Column.number("LineNo"), Column.text("Item")],
        primary_key=("OrderID", "LineNo"),
    ).normalized()
    assert spec.indexes[0].field_names == ("OrderID", "LineNo")
    assert spec.primary_index is not None


def test_index_field_coercion_and_descending() -> None:
    index = IndexSpec.on("ix_Recent", ("CreatedAt", "desc"), "Name")
    assert index.fields == (IndexField(name="CreatedAt", descending=True), IndexField(name="Name"))
    assert IndexSpec(name="ix", fields="Name").field_names == ("Name",)  # type: ignore[arg-type]


@pytest.mark.parametrize(
    ("kwargs", "message"),
    [
        ({"columns": []}, "at least 1"),
        ({"columns": [Column.text("A"), Column.text("a")]}, "duplicate column name"),
        ({"columns": [Column.autonumber("A"), Column.autonumber("B")]}, "only one AutoNumber"),
        (
            {"columns": [Column.text("A")], "indexes": [IndexSpec(name="ix", fields=("Nope",))]},  # type: ignore[arg-type]
            "unknown column",
        ),
        (
            {
                "columns": [Column.text("A", primary_key=True)],
                "indexes": [IndexSpec.primary_key("A")],
            },
            "declared more than once",
        ),
        (
            {
                "columns": [Column.text("A"), Column.text("B")],
                "indexes": [IndexSpec.on("A", "B"), IndexSpec.on("a", "A")],
            },
            "duplicate index name",
        ),
        (
            {"columns": [Column.text("A", unique=True)], "indexes": [IndexSpec.on("A", "A")]},
            "duplicate index name",
        ),
        (
            {"columns": [Column.ole_object("Blob")], "indexes": [IndexSpec.on("ix", "Blob")]},
            "cannot be indexed",
        ),
        ({"columns": [Column.text("A")], "validation_text": "x"}, "requires a validation_rule"),
    ],
)
def test_invalid_tables(kwargs: dict[str, object], message: str) -> None:
    with pytest.raises(ValidationError, match=message):
        TableSpec(name="T", **kwargs)  # type: ignore[arg-type]


def test_index_limit() -> None:
    columns = [Column.number(f"C{i}", indexed=True) for i in range(33)]
    with pytest.raises(ValidationError, match="at most 32 indexes"):
        TableSpec(name="T", columns=columns)


def test_index_field_limit() -> None:
    with pytest.raises(SpecError):
        IndexSpec.on("ix", *[f"C{i}" for i in range(11)])


def test_column_lookup_is_case_insensitive() -> None:
    spec = customers()
    assert spec.column("customername").name == "CustomerName"
    with pytest.raises(KeyError):
        spec.column("missing")
    assert spec.column_names == ("CustomerID", "CustomerName", "Email", "City")


def test_json_round_trip() -> None:
    spec = customers().normalized()
    assert TableSpec.model_validate_json(spec.model_dump_json()) == spec


def test_unique_on_single_primary_key_column_is_not_duplicated() -> None:
    spec = TableSpec(
        name="T", columns=[Column.number("ID", primary_key=True, unique=True)]
    ).normalized()
    assert [i.name for i in spec.indexes] == ["PrimaryKey"]
