"""Mapping between column specs and DAO field definitions (pure Python; no COM).

The DAO adapter reads each field into a plain :class:`FieldRead` and turns specs into a :class:`FieldPlan`,
so the whole mapping is unit-testable. Type codes and attribute flags are the ACEDAO type-library values
(see ``_com/constants.py`` and ADR 0001).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from typing import Any

from pyaccesskit.enums import NumberSize, PropertyType
from pyaccesskit.schema import (
    AutoNumberColumn,
    ColumnBase,
    CurrencyColumn,
    DateTimeColumn,
    DecimalColumn,
    Expr,
    HyperlinkColumn,
    LongTextColumn,
    NumberColumn,
    OleObjectColumn,
    TextColumn,
    YesNoColumn,
    parse_literal,
    render_literal,
)
from pyaccesskit.schema.columns import UnsupportedColumn

__all__ = [
    "DAO_PROPERTY_TYPES",
    "FieldPlan",
    "FieldRead",
    "column_from_field",
    "plan_field",
    "property_type_for",
]

# DataTypeEnum
DB_BOOLEAN, DB_BYTE, DB_INTEGER, DB_LONG, DB_CURRENCY, DB_SINGLE, DB_DOUBLE, DB_DATE = (
    1,
    2,
    3,
    4,
    5,
    6,
    7,
    8,
)
DB_BINARY, DB_TEXT, DB_LONGBINARY, DB_MEMO, DB_GUID, DB_BIGINT, DB_DECIMAL = (
    9,
    10,
    11,
    12,
    15,
    16,
    20,
)
DB_DATETIME_EXTENDED, DB_ATTACHMENT = 26, 101
# FieldAttributeEnum
DB_FIXED_FIELD, DB_VARIABLE_FIELD, DB_AUTOINCR_FIELD, DB_HYPERLINK_FIELD = 1, 2, 16, 32768
DB_DESCENDING = 1

REPLICATION_ID_DEFAULT = "GenGUID()"
AC_CHECKBOX = 106
DECIMAL_PLACES_AUTO = 255

_NUMBER_TYPES = {
    NumberSize.BYTE: DB_BYTE,
    NumberSize.INTEGER: DB_INTEGER,
    NumberSize.LONG_INTEGER: DB_LONG,
    NumberSize.SINGLE: DB_SINGLE,
    NumberSize.DOUBLE: DB_DOUBLE,
    NumberSize.REPLICATION_ID: DB_GUID,
}
_SIZE_OF = {code: size for size, code in _NUMBER_TYPES.items()}

DAO_PROPERTY_TYPES: dict[PropertyType, int] = {
    PropertyType.BOOLEAN: DB_BOOLEAN,
    PropertyType.BYTE: DB_BYTE,
    PropertyType.INTEGER: DB_INTEGER,
    PropertyType.LONG: DB_LONG,
    PropertyType.CURRENCY: DB_CURRENCY,
    PropertyType.SINGLE: DB_SINGLE,
    PropertyType.DOUBLE: DB_DOUBLE,
    PropertyType.DATE_TIME: DB_DATE,
    PropertyType.TEXT: DB_TEXT,
    PropertyType.MEMO: DB_MEMO,
}

# Access-defined properties PyAccessKit writes, with the DAO type Access itself uses for them.
KNOWN_PROPERTY_TYPES: dict[str, int] = {
    "description": DB_TEXT,
    "caption": DB_TEXT,
    "format": DB_TEXT,
    "inputmask": DB_TEXT,
    "decimalplaces": DB_BYTE,
    "unicodecompression": DB_BOOLEAN,
    "textformat": DB_BYTE,
    "displaycontrol": DB_INTEGER,
    "usemdimode": DB_BYTE,
    "webdesignmode": DB_BYTE,
    "showdocumenttabs": DB_BOOLEAN,
    "appicon": DB_TEXT,
    "apptitle": DB_TEXT,
    "startupform": DB_TEXT,
}


def property_type_for(name: str, value: Any, explicit: PropertyType | None = None) -> int:
    """The DAO type to create a property with: explicit, known Access property, or inferred."""
    if explicit is not None:
        return DAO_PROPERTY_TYPES[explicit]
    known = KNOWN_PROPERTY_TYPES.get(name.casefold())
    if known is not None:
        if known == DB_TEXT and isinstance(value, str) and len(value) > 255:
            return DB_MEMO
        return known
    if isinstance(value, bool):
        return DB_BOOLEAN
    if isinstance(value, int):
        return DB_LONG if -(2**31) <= value < 2**31 else DB_DOUBLE
    if isinstance(value, float):
        return DB_DOUBLE
    if isinstance(value, Decimal):
        return DB_CURRENCY
    if isinstance(value, str) and len(value) > 255:
        return DB_MEMO
    if value is not None and hasattr(value, "year"):
        return DB_DATE
    return DB_TEXT


@dataclass(frozen=True)
class FieldPlan:
    """How to create one column with DAO (``ado_decimal`` columns are created with ADO DDL instead)."""

    dao_type: int
    size: int | None
    attributes: int
    default: str | None
    allow_zero_length: bool | None
    append_only: bool
    ado_decimal: bool = False
    post_properties: tuple[tuple[str, int, Any], ...] = ()


def _post_properties(column: ColumnBase) -> list[tuple[str, int, Any]]:
    props: list[tuple[str, int, Any]] = []
    for name, value in (
        ("Description", column.description),
        ("Caption", column.caption),
        ("Format", column.format),
        ("InputMask", getattr(column, "input_mask", None)),
    ):
        if value is not None:
            props.append((name, property_type_for(name, value), value))
    decimal_places = getattr(column, "decimal_places", None)
    if decimal_places is not None:
        props.append(("DecimalPlaces", DB_BYTE, decimal_places))
    if isinstance(column, (TextColumn, LongTextColumn, HyperlinkColumn)):
        props.append(("UnicodeCompression", DB_BOOLEAN, column.unicode_compression))
    if isinstance(column, LongTextColumn) and column.rich_text:
        props.append(("TextFormat", DB_BYTE, 1))
    if isinstance(column, YesNoColumn):
        props.append(("DisplayControl", DB_INTEGER, AC_CHECKBOX))
    for name, value in column.properties.items():
        props.append((name, property_type_for(name, value), value))
    return props


def plan_field(column: ColumnBase) -> FieldPlan:
    """Plan the DAO definition of ``column``.

    Raises:
        ValueError: For columns that cannot be created (introspection-only types).
    """
    default = render_literal(column.default) if column.default is not None else None
    allow_zero: bool | None = getattr(column, "allow_zero_length", None)
    post = tuple(_post_properties(column))
    if isinstance(column, TextColumn):
        return FieldPlan(
            DB_TEXT,
            column.length,
            DB_VARIABLE_FIELD,
            default,
            allow_zero,
            False,
            post_properties=post,
        )
    if isinstance(column, LongTextColumn):
        return FieldPlan(
            DB_MEMO,
            None,
            DB_VARIABLE_FIELD,
            default,
            allow_zero,
            column.append_only,
            post_properties=post,
        )
    if isinstance(column, HyperlinkColumn):
        attributes = DB_VARIABLE_FIELD | DB_HYPERLINK_FIELD
        return FieldPlan(
            DB_MEMO, None, attributes, default, allow_zero, False, post_properties=post
        )
    if isinstance(column, NumberColumn):
        return FieldPlan(
            _NUMBER_TYPES[column.size],
            None,
            DB_FIXED_FIELD,
            default,
            None,
            False,
            post_properties=post,
        )
    if isinstance(column, DecimalColumn):
        return FieldPlan(
            DB_DECIMAL,
            None,
            DB_FIXED_FIELD,
            default,
            None,
            False,
            ado_decimal=True,
            post_properties=post,
        )
    if isinstance(column, CurrencyColumn):
        return FieldPlan(
            DB_CURRENCY, None, DB_FIXED_FIELD, default, None, False, post_properties=post
        )
    if isinstance(column, AutoNumberColumn):
        if column.replication_id:
            return FieldPlan(
                DB_GUID,
                None,
                DB_FIXED_FIELD,
                REPLICATION_ID_DEFAULT,
                None,
                False,
                post_properties=post,
            )
        attributes = DB_FIXED_FIELD | DB_AUTOINCR_FIELD
        return FieldPlan(DB_LONG, None, attributes, None, None, False, post_properties=post)
    if isinstance(column, DateTimeColumn):
        return FieldPlan(DB_DATE, None, DB_FIXED_FIELD, default, None, False, post_properties=post)
    if isinstance(column, YesNoColumn):
        return FieldPlan(
            DB_BOOLEAN, None, DB_FIXED_FIELD, default, None, False, post_properties=post
        )
    if isinstance(column, OleObjectColumn):
        return FieldPlan(
            DB_LONGBINARY, None, DB_VARIABLE_FIELD, None, None, False, post_properties=post
        )
    raise ValueError(
        f"columns of type {column.data_type.value!r} can be read but not created by PyAccessKit"
    )


@dataclass(frozen=True)
class FieldRead:
    """Plain data read from a DAO field (and its Access properties)."""

    name: str
    dao_type: int
    size: int
    attributes: int
    required: bool
    allow_zero_length: bool
    default: str
    validation_rule: str
    validation_text: str
    append_only: bool = False
    expression: str = ""
    properties: dict[str, Any] = field(default_factory=dict[str, Any])


def _prop(read: FieldRead, name: str) -> Any:
    return read.properties.get(name.casefold())


def _text(value: Any) -> str | None:
    return value if isinstance(value, str) and value != "" else None


def column_from_field(read: FieldRead) -> ColumnBase:
    """Rebuild the canonical column spec of a DAO field."""
    common: dict[str, Any] = {
        "name": read.name,
        "description": _text(_prop(read, "Description")),
        "caption": _text(_prop(read, "Caption")),
        "format": _text(_prop(read, "Format")),
        "validation_rule": _text(read.validation_rule),
        "validation_text": _text(read.validation_text) if _text(read.validation_rule) else None,
    }
    default_text = read.default.strip() if read.default else ""
    default: Any = parse_literal(default_text) if default_text else None
    input_mask = _text(_prop(read, "InputMask"))
    places = _prop(read, "DecimalPlaces")
    decimal_places = None if places is None or int(places) == DECIMAL_PLACES_AUTO else int(places)
    unicode_value = _prop(read, "UnicodeCompression")
    unicode_compression = True if unicode_value is None else bool(unicode_value)
    auto = bool(read.attributes & DB_AUTOINCR_FIELD)

    if read.expression:
        return UnsupportedColumn(**common, dao_type=read.dao_type, detail="calculated field")
    code = read.dao_type
    if code == DB_TEXT:
        return TextColumn(
            **common,
            required=read.required,
            default=default,
            length=read.size,
            allow_zero_length=read.allow_zero_length,
            unicode_compression=unicode_compression,
            input_mask=input_mask,
        )
    if code == DB_MEMO:
        if read.attributes & DB_HYPERLINK_FIELD:
            return HyperlinkColumn(
                **common,
                required=read.required,
                default=default,
                allow_zero_length=read.allow_zero_length,
                unicode_compression=unicode_compression,
            )
        text_format = _prop(read, "TextFormat")
        return LongTextColumn(
            **common,
            required=read.required,
            default=default,
            rich_text=bool(text_format) and int(text_format) == 1,
            append_only=read.append_only,
            allow_zero_length=read.allow_zero_length,
            unicode_compression=unicode_compression,
        )
    if code == DB_LONG and auto:
        return AutoNumberColumn(**common)
    if code == DB_GUID and default_text.casefold() == REPLICATION_ID_DEFAULT.casefold():
        return AutoNumberColumn(**common, replication_id=True)
    if code in _SIZE_OF:
        size = _SIZE_OF[code]
        return NumberColumn(
            **common,
            required=read.required,
            default=default
            if size is not NumberSize.REPLICATION_ID or isinstance(default, Expr)
            else None,
            size=size,
            decimal_places=decimal_places,
            input_mask=input_mask,
        )
    if code == DB_DECIMAL:
        precision = _prop(read, "Precision")
        scale = _prop(read, "Scale")
        return DecimalColumn(
            **common,
            required=read.required,
            default=default,
            precision=int(precision) if precision else 18,
            scale=int(scale) if scale is not None else 0,
            decimal_places=decimal_places,
            input_mask=input_mask,
        )
    if code == DB_CURRENCY:
        return CurrencyColumn(
            **common,
            required=read.required,
            default=default,
            decimal_places=decimal_places,
            input_mask=input_mask,
        )
    if code == DB_DATE:
        return DateTimeColumn(
            **common, required=read.required, default=default, input_mask=input_mask
        )
    if code == DB_BOOLEAN:
        return YesNoColumn(**common, required=read.required, default=default)
    if code == DB_LONGBINARY:
        return OleObjectColumn(**common)
    details = {
        DB_BIGINT: "Large Number",
        DB_DATETIME_EXTENDED: "Date/Time Extended",
        DB_ATTACHMENT: "Attachment",
        DB_BINARY: "Binary",
    }
    detail = details.get(code, "multi-valued field" if 102 <= code <= 109 else f"DAO type {code}")
    return UnsupportedColumn(**common, dao_type=code, detail=detail)
