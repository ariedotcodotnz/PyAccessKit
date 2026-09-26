from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal

import pytest
from pydantic import TypeAdapter, ValidationError

from pyaccesskit.enums import DataType, NumberSize
from pyaccesskit.errors import SpecError
from pyaccesskit.schema import (
    AutoNumberColumn,
    Column,
    ColumnSpec,
    CurrencyColumn,
    DateTimeColumn,
    DecimalColumn,
    Expr,
    NumberColumn,
    TextColumn,
    YesNoColumn,
)

COLUMN_ADAPTER: TypeAdapter[ColumnSpec] = TypeAdapter(ColumnSpec)


def test_factory_builds_the_right_classes() -> None:
    assert isinstance(Column.text("A"), TextColumn)
    assert isinstance(Column.number("B"), NumberColumn)
    assert isinstance(Column.decimal("C", precision=10, scale=2), DecimalColumn)
    assert isinstance(Column.currency("D"), CurrencyColumn)
    assert isinstance(Column.autonumber("E"), AutoNumberColumn)
    assert isinstance(Column.date_time("F"), DateTimeColumn)
    assert isinstance(Column.yes_no("G"), YesNoColumn)
    assert Column.long_text("H").data_type is DataType.LONG_TEXT
    assert Column.hyperlink("I").data_type is DataType.HYPERLINK
    assert Column.ole_object("J").data_type is DataType.OLE_OBJECT


def test_column_factory_is_not_instantiable() -> None:
    with pytest.raises(TypeError):
        Column()


def test_number_defaults_to_long_integer() -> None:
    column = Column.number("Qty")
    assert column.size is NumberSize.LONG_INTEGER
    assert Column.number("Small", size="integer").size is NumberSize.INTEGER


@pytest.mark.parametrize(
    ("factory", "default", "expected"),
    [
        (lambda d: Column.currency("X", default=d), 0, Decimal("0")),
        (lambda d: Column.currency("X", default=d), 1.25, Decimal("1.25")),
        (lambda d: Column.currency("X", default=d), "2.50", Decimal("2.50")),
        (lambda d: Column.number("X", default=d), 5.0, 5),
        (lambda d: Column.number("X", default=d), "7", 7),
        (lambda d: Column.number("X", size=NumberSize.DOUBLE, default=d), 2, 2.0),
        (lambda d: Column.yes_no("X", default=d), "No", False),
        (lambda d: Column.yes_no("X", default=d), -1, True),
        (lambda d: Column.date_time("X", default=d), "2026-01-31", date(2026, 1, 31)),
        (
            lambda d: Column.date_time("X", default=d),
            "2026-01-31T10:00:00",
            datetime(2026, 1, 31, 10),
        ),
        (lambda d: Column.date_time("X", default=d), Expr("Now()"), Expr("Now()")),
        (lambda d: Column.text("X", default=d), "Unknown", "Unknown"),
    ],
)
def test_defaults_are_coerced_to_canonical_types(
    factory, default: object, expected: object
) -> None:  # type: ignore[no-untyped-def]
    assert factory(default).default == expected


@pytest.mark.parametrize(
    ("build", "message"),
    [
        (lambda: Column.text("X", length=0), "greater than or equal to 1"),
        (lambda: Column.text("X", length=256), "less than or equal to 255"),
        (lambda: Column.decimal("X", precision=5, scale=6), "cannot exceed precision"),
        (lambda: Column.decimal("X", precision=29), "less than or equal to 28"),
        (lambda: Column.number("X", default="abc"), "not an integer"),
        (lambda: Column.number("X", default=1.5), "not an integer"),
        (lambda: Column.number("X", default=True), "booleans"),
        (lambda: Column.yes_no("X", default="maybe"), "Yes/No"),
        (lambda: Column.date_time("X", default="tomorrow"), "not a date"),
        (lambda: Column.number("X", size="replication_id", default=1), "Replication ID"),
        (lambda: Column.text("X", validation_text="oops"), "requires a validation_rule"),
        (lambda: Column.ole_object("X", unique=True), "cannot be indexed"),
        (lambda: Column.text("bad.name"), "cannot contain"),
    ],
)
def test_invalid_columns_raise_spec_error(build, message: str) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(SpecError) as info:
        build()
    assert message in str(info.value)


def test_autonumber_rules() -> None:
    with pytest.raises(ValidationError, match="default"):
        AutoNumberColumn(name="ID", default=1)
    with pytest.raises(ValidationError, match="always populated"):
        AutoNumberColumn(name="ID", required=True)
    assert Column.autonumber("ID", replication_id=True).replication_id
    with pytest.raises(ValidationError, match="Extra inputs"):
        AutoNumberColumn(name="ID", new_values="random")  # type: ignore[call-arg]


def test_specs_are_frozen() -> None:
    column = Column.text("Name")
    with pytest.raises(ValidationError):
        column.length = 10  # type: ignore[misc]


def test_normalized_clears_index_shorthands() -> None:
    column = Column.text("Email", unique=True, indexed=True)
    normalized = column.normalized()
    assert (normalized.unique, normalized.indexed, normalized.primary_key) == (False, False, False)
    plain = Column.text("Plain")
    assert plain.normalized() is plain


def test_discriminated_union_json_round_trip() -> None:
    columns = [
        Column.text("A", length=50, required=True, caption="Alpha"),
        Column.currency("B", default=Decimal("9.99")),
        Column.date_time("C", default=Expr("Now()")),
        Column.yes_no("D", default=False),
        Column.number("E", size=NumberSize.SINGLE, default=1.5),
        Column.decimal("F", precision=12, scale=3, default="1.125"),
    ]
    for column in columns:
        dumped = column.model_dump(mode="json")
        assert COLUMN_ADAPTER.validate_python(dumped) == column
        assert COLUMN_ADAPTER.validate_json(column.model_dump_json()) == column


def test_union_rejects_unknown_type() -> None:
    with pytest.raises(ValidationError):
        COLUMN_ADAPTER.validate_python({"type": "nonsense", "name": "X"})


def test_extra_fields_are_forbidden() -> None:
    with pytest.raises(ValidationError, match="Extra inputs"):
        TextColumn(name="X", lenght=10)  # type: ignore[call-arg]
