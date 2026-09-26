"""Spec ↔ DAO field mapping, round-tripped through a simulated DAO echo (pure Python)."""

from __future__ import annotations

from datetime import datetime
from decimal import Decimal
from typing import Any

import pytest

from pyaccesskit import Column, Expr, NumberSize, PropertyType, SpecError
from pyaccesskit._backends.dao import typemap as tm
from pyaccesskit.schema.columns import ColumnBase, DecimalColumn, UnsupportedColumn


def dao_echo(column: ColumnBase) -> tm.FieldRead:
    """What DAO reports for a field created from ``plan_field(column)``."""
    plan = tm.plan_field(column)
    properties: dict[str, Any] = {
        name.casefold(): value for name, _t, value in plan.post_properties
    }
    if isinstance(column, DecimalColumn):  # created by ADO DDL; DAO reads Precision/Scale back
        properties |= {"precision": column.precision, "scale": column.scale}
    return tm.FieldRead(
        name=column.name,
        dao_type=plan.dao_type,
        size=plan.size or 0,
        attributes=plan.attributes,
        required=column.required,
        allow_zero_length=bool(plan.allow_zero_length),
        default=plan.default or "",
        validation_rule=column.validation_rule or "",
        validation_text=column.validation_text or "",
        append_only=plan.append_only,
        properties=properties,
    )


ROUND_TRIPS = [
    Column.text(
        "Code",
        length=12,
        required=True,
        default='O"Brien',
        allow_zero_length=True,
        unicode_compression=False,
        input_mask=">LL000",
        caption="Code",
        description="Customer code",
        format=">",
    ),
    Column.long_text("Notes", rich_text=True, append_only=True),
    Column.hyperlink("Site", default="https://example.com"),
    Column.number("Tiny", size=NumberSize.BYTE, default=3, decimal_places=0),
    Column.number(
        "Small", size=NumberSize.INTEGER, validation_rule=">0", validation_text="positive"
    ),
    Column.number("Big", primary_key=True),
    Column.number("Ratio", size=NumberSize.SINGLE, default=0.5),
    Column.number("Precise", size=NumberSize.DOUBLE, default=1e-7),
    Column.number("Guid", size=NumberSize.REPLICATION_ID),
    Column.decimal("Money", precision=10, scale=2, default=Decimal("1.25")),
    Column.currency("Total", default=0, decimal_places=2, format="Currency"),
    Column.autonumber("ID", primary_key=True),
    Column.autonumber("RowGuid", replication_id=True),
    Column.date_time("CreatedAt", default=Expr("Now()"), format="Short Date"),
    Column.date_time("Epoch", default=datetime(2026, 1, 31, 12, 30)),
    Column.yes_no("Active", default=True),
    Column.ole_object("Blob"),
    Column.ole_object("Payload", required=True),
]


@pytest.mark.parametrize("column", ROUND_TRIPS, ids=lambda c: f"{c.type}-{c.name}")
def test_plan_and_read_round_trip(column: ColumnBase) -> None:
    assert tm.column_from_field(dao_echo(column)) == column.normalized()


def test_autonumber_plans() -> None:
    increment = tm.plan_field(Column.autonumber("ID"))
    assert (increment.dao_type, increment.default) == (tm.DB_LONG, None)
    assert increment.attributes & tm.DB_AUTOINCR_FIELD
    guid = tm.plan_field(Column.autonumber("G", replication_id=True))
    assert (guid.dao_type, guid.default) == (tm.DB_GUID, "GenGUID()")
    assert not guid.attributes & tm.DB_AUTOINCR_FIELD


def test_decimal_is_created_by_ado() -> None:
    assert tm.plan_field(Column.decimal("D", precision=5, scale=1)).ado_decimal


def test_access_ui_properties_are_planned() -> None:
    names = {name for name, _t, _v in tm.plan_field(Column.yes_no("Flag")).post_properties}
    assert "DisplayControl" in names, (
        "Yes/No columns show as check boxes, as in the Access designer"
    )
    text = {n: v for n, _t, v in tm.plan_field(Column.text("T")).post_properties}
    assert text["UnicodeCompression"] is True


def test_auto_decimal_places_read_as_none() -> None:
    read = dao_echo(Column.number("N"))
    read.properties["decimalplaces"] = tm.DECIMAL_PLACES_AUTO
    column = tm.column_from_field(read)
    assert column == Column.number("N")


def test_replication_id_number_defaulting_to_genguid_is_an_autonumber() -> None:
    with pytest.raises(SpecError, match="replication_id=True"):
        Column.number("G", size=NumberSize.REPLICATION_ID, default=Expr("GenGUID()"))


@pytest.mark.parametrize(
    ("dao_type", "expression", "detail"),
    [
        (tm.DB_ATTACHMENT, "", "Attachment"),
        (tm.DB_BIGINT, "", "Large Number"),
        (tm.DB_DATETIME_EXTENDED, "", "Date/Time Extended"),
        (104, "", "multi-valued field"),
        (tm.DB_TEXT, "[A] & [B]", "calculated field"),
    ],
)
def test_unsupported_fields_are_reported_not_hidden(
    dao_type: int, expression: str, detail: str
) -> None:
    read = tm.FieldRead(
        name="X",
        dao_type=dao_type,
        size=0,
        attributes=0,
        required=False,
        allow_zero_length=False,
        default="",
        validation_rule="",
        validation_text="",
        expression=expression,
    )
    column = tm.column_from_field(read)
    assert isinstance(column, UnsupportedColumn)
    assert (column.dao_type, column.detail) == (dao_type, detail)
    with pytest.raises(ValueError, match="can be read but not created"):
        tm.plan_field(column)


@pytest.mark.parametrize(
    ("name", "value", "explicit", "expected"),
    [
        ("Anything", 1, PropertyType.BYTE, tm.DB_BYTE),
        ("Caption", "short", None, tm.DB_TEXT),
        ("Description", "x" * 300, None, tm.DB_MEMO),
        ("DecimalPlaces", 2, None, tm.DB_BYTE),
        ("Custom", True, None, tm.DB_BOOLEAN),
        ("Custom", 7, None, tm.DB_LONG),
        ("Custom", 2**40, None, tm.DB_DOUBLE),
        ("Custom", 1.5, None, tm.DB_DOUBLE),
        ("Custom", Decimal("9.99"), None, tm.DB_CURRENCY),
        ("Custom", "y" * 256, None, tm.DB_MEMO),
        ("Custom", datetime(2026, 1, 1), None, tm.DB_DATE),
        ("Custom", "text", None, tm.DB_TEXT),
    ],
)
def test_property_types(
    name: str, value: object, explicit: PropertyType | None, expected: int
) -> None:
    assert tm.property_type_for(name, value, explicit) == expected
