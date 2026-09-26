"""Column (field) specifications and the :class:`Column` factory.

Each Access data type has its own immutable spec class (``TextColumn``, ``NumberColumn``...) so that only
options meaningful for that type are accepted. The classes form a discriminated union on ``type``
(:data:`ColumnSpec`), which is also how columns appear in JSON/YAML::

    {"type": "text", "name": "Email", "length": 255}

Most code uses the :class:`Column` factory, which mirrors the Access table designer::

    Column.autonumber("CustomerID", primary_key=True)
    Column.text("CustomerName", length=200, required=True)
    Column.number("Quantity", size=NumberSize.INTEGER)
    Column.currency("UnitPrice", default=0)
"""

from __future__ import annotations

from collections.abc import Mapping
from datetime import date, datetime, time
from decimal import Decimal, InvalidOperation
from typing import TYPE_CHECKING, Annotated, Any, ClassVar, Literal, Self, TypedDict, cast

from pydantic import Field, field_validator, model_validator

from pyaccesskit.enums import DataType, NumberSize
from pyaccesskit.schema._base import PropertyValue, SpecModel, build
from pyaccesskit.schema.expressions import DefaultValue, Expr, parse_literal
from pyaccesskit.schema.names import check_name

if TYPE_CHECKING:
    from typing_extensions import Unpack

__all__ = [
    "COLUMN_CLASSES",
    "AutoNumberColumn",
    "Column",
    "ColumnBase",
    "ColumnOptions",
    "ColumnSpec",
    "CurrencyColumn",
    "DateTimeColumn",
    "DecimalColumn",
    "HyperlinkColumn",
    "LongTextColumn",
    "NumberColumn",
    "OleObjectColumn",
    "OleObjectOptions",
    "TextColumn",
    "UnsupportedColumn",
    "YesNoColumn",
]


# ------------------------------------------------------------------------------------ default coercion
# Strings (e.g. from YAML/JSON) are accepted as *literals* of the column's type only. Anything that is not
# a literal must be an explicit Expr(...) so that a typo never silently becomes an expression.
def _literal(value: Any, expected: str) -> Any:
    if not isinstance(value, str):
        return value
    parsed = parse_literal(value)
    if isinstance(parsed, Expr) or parsed is None:
        raise ValueError(f"default {value!r} is not {expected}; wrap expressions in Expr(...)")
    return parsed


def _as_int(value: Any) -> int | Expr:
    value = _literal(value, "an integer")
    if isinstance(value, Expr):
        return value
    if isinstance(value, bool):
        raise ValueError("booleans are not valid defaults for integer columns (use 0 or 1)")
    if isinstance(value, int):
        return value
    if isinstance(value, (float, Decimal)) and value == int(value):
        return int(value)
    raise ValueError(f"default {value!r} is not an integer; wrap expressions in Expr(...)")


def _as_float(value: Any) -> float | Expr:
    value = _literal(value, "a number")
    if isinstance(value, Expr):
        return value
    if isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
        raise ValueError(f"default {value!r} is not a number; wrap expressions in Expr(...)")
    return float(value)


def _as_decimal(value: Any) -> Decimal | Expr:
    value = _literal(value, "a number")
    if isinstance(value, Expr):
        return value
    if isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
        raise ValueError(f"default {value!r} is not a number; wrap expressions in Expr(...)")
    try:
        return value if isinstance(value, Decimal) else Decimal(str(value))
    except InvalidOperation as exc:  # pragma: no cover - finite floats always convert
        raise ValueError(f"default {value!r} is not a valid decimal") from exc


def _as_bool(value: Any) -> bool | Expr:
    value = _literal(value, "a Yes/No value (True/False)")
    if isinstance(value, (bool, Expr)):
        return value
    if isinstance(value, int) and value in (0, 1, -1):
        return value != 0
    raise ValueError(f"default {value!r} is not a Yes/No value; use True/False")


def _as_datetime(value: Any) -> datetime | date | time | Expr:
    if isinstance(value, str):
        text = value.strip()
        try:
            parsed = datetime.fromisoformat(text)
        except ValueError:
            try:
                value = time.fromisoformat(text)
            except ValueError:
                value = _literal(text, "a date/time (use an ISO date, #date#, or Expr('Now()'))")
        else:
            value = parsed.date() if len(text) <= 10 else parsed
    if isinstance(value, (datetime, date, time, Expr)):
        return value
    raise ValueError(f"default {value!r} is not a date/time; use datetime/date or Expr('Now()')")


# ------------------------------------------------------------------------------------------ base class
class ColumnBase(SpecModel):
    """Options shared by every column type.

    Attributes:
        name: Field name (≤64 characters; see Access naming rules).
        required: Disallow Null values (*Required* property).
        default: Default value: a Python literal (rendered as an Access literal) or an ``Expr``.
        validation_rule: *Validation Rule* expression, e.g. ``">0"``.
        validation_text: Message shown when the validation rule fails.
        description: *Description* shown in the table designer.
        caption: *Caption* used as the default label text on forms and datasheet headers.
        format: *Format* property, e.g. ``"Short Date"`` or ``"Currency"``.
        primary_key: Shorthand: include this column in the table's primary key.
        unique: Shorthand: create a unique single-column index named after the column.
        indexed: Shorthand: create a non-unique single-column index named after the column.
        properties: Other Access/DAO field properties to set verbatim (escape hatch).
    """

    data_type: ClassVar[DataType]
    supports_default: ClassVar[bool] = True
    indexable: ClassVar[bool] = True

    name: str
    required: bool = False
    default: DefaultValue | None = None
    validation_rule: str | None = None
    validation_text: str | None = None
    description: str | None = None
    caption: str | None = None
    format: str | None = None
    primary_key: bool = False
    unique: bool = False
    indexed: bool = False
    properties: dict[str, PropertyValue] = Field(default_factory=dict)

    @field_validator("name")
    @classmethod
    def _check_name(cls, value: str) -> str:
        return check_name(value, what="column name")

    @model_validator(mode="before")
    @classmethod
    def _coerce_default_input(cls, data: Any) -> Any:
        if not isinstance(data, Mapping):
            return data
        fields = cast("Mapping[str, Any]", data)
        value = fields.get("default")
        if value is None or isinstance(value, Expr):
            return fields
        if isinstance(value, Mapping):
            payload = cast("Mapping[str, Any]", value)
            if set(payload) == {"expr"}:  # the JSON form of an Expr
                return {**fields, "default": Expr(str(payload["expr"]))}
        return {**fields, "default": cls._coerce_default(value, fields)}

    @classmethod
    def _coerce_default(cls, value: Any, data: Mapping[str, Any]) -> Any:
        """Convert a user-supplied default to this column type's canonical Python type."""
        return value

    @model_validator(mode="after")
    def _check_common(self) -> Self:
        if self.default is not None and not self.supports_default:
            raise ValueError(f"{self.data_type.value} columns cannot have a default value")
        if (self.primary_key or self.unique or self.indexed) and not self.indexable:
            raise ValueError(f"{self.data_type.value} columns cannot be indexed")
        if self.validation_text is not None and self.validation_rule is None:
            raise ValueError("validation_text requires a validation_rule")
        return self

    def normalized(self) -> Self:
        """Canonical form: the index shorthands are cleared (they live in ``TableSpec.indexes``)."""
        if not (self.primary_key or self.unique or self.indexed):
            return self
        return self.model_copy(update={"primary_key": False, "unique": False, "indexed": False})


# ------------------------------------------------------------------------------------------ text types
class TextColumn(ColumnBase):
    """Short Text: up to 255 characters."""

    data_type: ClassVar[DataType] = DataType.TEXT
    type: Literal["text"] = "text"
    length: int = Field(default=255, ge=1, le=255)
    allow_zero_length: bool = False
    unicode_compression: bool = True
    input_mask: str | None = None


class LongTextColumn(ColumnBase):
    """Long Text (Memo): up to ~1 GB; optionally rich text and append-only."""

    data_type: ClassVar[DataType] = DataType.LONG_TEXT
    type: Literal["long_text"] = "long_text"
    rich_text: bool = False
    append_only: bool = False
    allow_zero_length: bool = False
    unicode_compression: bool = True


class HyperlinkColumn(ColumnBase):
    """Hyperlink: a Long Text field flagged as a hyperlink."""

    data_type: ClassVar[DataType] = DataType.HYPERLINK
    type: Literal["hyperlink"] = "hyperlink"
    allow_zero_length: bool = False
    unicode_compression: bool = True


# ---------------------------------------------------------------------------------------- number types
class NumberColumn(ColumnBase):
    """Number with a *Field Size* of Byte, Integer (16-bit), Long Integer, Single, Double or Replication ID."""

    data_type: ClassVar[DataType] = DataType.NUMBER
    type: Literal["number"] = "number"
    size: NumberSize = NumberSize.LONG_INTEGER
    decimal_places: int | None = Field(default=None, ge=0, le=15)
    """``None`` means *Auto*."""
    input_mask: str | None = None

    @classmethod
    def _coerce_default(cls, value: Any, data: Mapping[str, Any]) -> Any:
        size = NumberSize(data.get("size", NumberSize.LONG_INTEGER))
        if size in (NumberSize.SINGLE, NumberSize.DOUBLE):
            return _as_float(value)
        if size is NumberSize.REPLICATION_ID:
            raise ValueError("Replication ID columns only accept Expr(...) defaults")
        return _as_int(value)

    @model_validator(mode="after")
    def _check_replication_default(self) -> Self:
        # Access defines an AutoNumber (Replication ID) as exactly this; one spelling keeps specs canonical.
        if (
            self.size is NumberSize.REPLICATION_ID
            and isinstance(self.default, Expr)
            and self.default.expr.replace(" ", "").casefold() == "genguid()"
        ):
            raise ValueError(
                "a Replication ID number defaulting to GenGUID() is an AutoNumber; "
                "use Column.autonumber(name, replication_id=True)"
            )
        return self


class DecimalColumn(ColumnBase):
    """Number with *Field Size* = Decimal (exact, with precision and scale)."""

    data_type: ClassVar[DataType] = DataType.DECIMAL
    type: Literal["decimal"] = "decimal"
    precision: int = Field(default=18, ge=1, le=28)
    scale: int = Field(default=0, ge=0, le=28)
    decimal_places: int | None = Field(default=None, ge=0, le=15)
    input_mask: str | None = None

    @classmethod
    def _coerce_default(cls, value: Any, data: Mapping[str, Any]) -> Any:
        return _as_decimal(value)

    @model_validator(mode="after")
    def _check_scale(self) -> Self:
        if self.scale > self.precision:
            raise ValueError(f"scale ({self.scale}) cannot exceed precision ({self.precision})")
        return self


class CurrencyColumn(ColumnBase):
    """Currency: fixed-point, 4 decimal places, no rounding surprises."""

    data_type: ClassVar[DataType] = DataType.CURRENCY
    type: Literal["currency"] = "currency"
    decimal_places: int | None = Field(default=None, ge=0, le=15)
    input_mask: str | None = None

    @classmethod
    def _coerce_default(cls, value: Any, data: Mapping[str, Any]) -> Any:
        return _as_decimal(value)


class AutoNumberColumn(ColumnBase):
    """AutoNumber: an incrementing Long Integer, or a Replication ID (GUID).

    *Random* AutoNumbers are not offered: DAO silently ignores the setting, so PyAccessKit cannot create
    them reliably (see ADR 0002).
    """

    data_type: ClassVar[DataType] = DataType.AUTONUMBER
    supports_default: ClassVar[bool] = False
    type: Literal["autonumber"] = "autonumber"
    replication_id: bool = False

    @model_validator(mode="after")
    def _check_autonumber(self) -> Self:
        if self.required:
            raise ValueError("AutoNumber columns are always populated; 'required' does not apply")
        return self


# ------------------------------------------------------------------------------------------ other types
class DateTimeColumn(ColumnBase):
    """Date/Time."""

    data_type: ClassVar[DataType] = DataType.DATE_TIME
    type: Literal["date_time"] = "date_time"
    input_mask: str | None = None

    @classmethod
    def _coerce_default(cls, value: Any, data: Mapping[str, Any]) -> Any:
        return _as_datetime(value)


class YesNoColumn(ColumnBase):
    """Yes/No (Boolean). Shown as a check box in datasheets and bound forms."""

    data_type: ClassVar[DataType] = DataType.YES_NO
    type: Literal["yes_no"] = "yes_no"

    @classmethod
    def _coerce_default(cls, value: Any, data: Mapping[str, Any]) -> Any:
        return _as_bool(value)


class OleObjectColumn(ColumnBase):
    """OLE Object (long binary data)."""

    data_type: ClassVar[DataType] = DataType.OLE_OBJECT
    supports_default: ClassVar[bool] = False
    indexable: ClassVar[bool] = False
    type: Literal["ole_object"] = "ole_object"


class UnsupportedColumn(ColumnBase):
    """A column PyAccessKit can read but not create yet (Attachment, Calculated, Large Number...).

    It appears when introspecting existing databases so that ``to_spec()`` never fails; creating a table
    with one raises :class:`~pyaccesskit.errors.SpecError`.
    """

    data_type: ClassVar[DataType] = DataType.UNKNOWN
    type: Literal["unsupported"] = "unsupported"
    dao_type: int
    detail: str = ""


ColumnSpec = Annotated[
    TextColumn
    | LongTextColumn
    | NumberColumn
    | DecimalColumn
    | CurrencyColumn
    | AutoNumberColumn
    | DateTimeColumn
    | YesNoColumn
    | HyperlinkColumn
    | OleObjectColumn
    | UnsupportedColumn,
    Field(discriminator="type"),
]
"""Any column spec (a Pydantic discriminated union on ``type``)."""

COLUMN_CLASSES: tuple[type[ColumnBase], ...] = (
    TextColumn,
    LongTextColumn,
    NumberColumn,
    DecimalColumn,
    CurrencyColumn,
    AutoNumberColumn,
    DateTimeColumn,
    YesNoColumn,
    HyperlinkColumn,
    OleObjectColumn,
    UnsupportedColumn,
)


# ---------------------------------------------------------------------------------------------- factory
class ColumnOptions(TypedDict, total=False):
    """Keyword options accepted by every :class:`Column` constructor (see :class:`ColumnBase`)."""

    required: bool
    default: DefaultValue | None
    validation_rule: str | None
    validation_text: str | None
    description: str | None
    caption: str | None
    format: str | None
    primary_key: bool
    unique: bool
    indexed: bool
    properties: dict[str, PropertyValue]


class AutoNumberOptions(TypedDict, total=False):
    """Keyword options accepted by :meth:`Column.autonumber` (no default/required)."""

    validation_rule: str | None
    validation_text: str | None
    description: str | None
    caption: str | None
    format: str | None
    primary_key: bool
    unique: bool
    indexed: bool
    properties: dict[str, PropertyValue]


class OleObjectOptions(AutoNumberOptions, total=False):
    """Keyword options accepted by :meth:`Column.ole_object` (no default value)."""

    required: bool


class Column:
    """Factory for column specs, mirroring the data types of the Access table designer.

    Every constructor raises :class:`~pyaccesskit.errors.SpecError` for invalid options.
    """

    def __init__(self) -> None:
        raise TypeError(
            "Column is a factory namespace; call Column.text(...), Column.number(...), etc."
        )

    @staticmethod
    def text(
        name: str,
        *,
        length: int = 255,
        allow_zero_length: bool = False,
        unicode_compression: bool = True,
        input_mask: str | None = None,
        **options: Unpack[ColumnOptions],
    ) -> TextColumn:
        """Short Text (``length`` 1-255, default 255)."""
        return build(
            TextColumn,
            f"text column {name!r}",
            name=name,
            length=length,
            allow_zero_length=allow_zero_length,
            unicode_compression=unicode_compression,
            input_mask=input_mask,
            **options,
        )

    @staticmethod
    def long_text(
        name: str,
        *,
        rich_text: bool = False,
        append_only: bool = False,
        allow_zero_length: bool = False,
        unicode_compression: bool = True,
        **options: Unpack[ColumnOptions],
    ) -> LongTextColumn:
        """Long Text (Memo)."""
        return build(
            LongTextColumn,
            f"long text column {name!r}",
            name=name,
            rich_text=rich_text,
            append_only=append_only,
            allow_zero_length=allow_zero_length,
            unicode_compression=unicode_compression,
            **options,
        )

    @staticmethod
    def number(
        name: str,
        *,
        size: NumberSize | str = NumberSize.LONG_INTEGER,
        decimal_places: int | None = None,
        input_mask: str | None = None,
        **options: Unpack[ColumnOptions],
    ) -> NumberColumn:
        """Number. ``size`` defaults to Long Integer (32-bit), exactly like Access."""
        return build(
            NumberColumn,
            f"number column {name!r}",
            name=name,
            size=size,
            decimal_places=decimal_places,
            input_mask=input_mask,
            **options,
        )

    @staticmethod
    def decimal(
        name: str,
        *,
        precision: int = 18,
        scale: int = 0,
        decimal_places: int | None = None,
        input_mask: str | None = None,
        **options: Unpack[ColumnOptions],
    ) -> DecimalColumn:
        """Number with Field Size = Decimal(``precision``, ``scale``)."""
        return build(
            DecimalColumn,
            f"decimal column {name!r}",
            name=name,
            precision=precision,
            scale=scale,
            decimal_places=decimal_places,
            input_mask=input_mask,
            **options,
        )

    @staticmethod
    def currency(
        name: str,
        *,
        decimal_places: int | None = None,
        input_mask: str | None = None,
        **options: Unpack[ColumnOptions],
    ) -> CurrencyColumn:
        """Currency."""
        return build(
            CurrencyColumn,
            f"currency column {name!r}",
            name=name,
            decimal_places=decimal_places,
            input_mask=input_mask,
            **options,
        )

    @staticmethod
    def autonumber(
        name: str,
        *,
        replication_id: bool = False,
        **options: Unpack[AutoNumberOptions],
    ) -> AutoNumberColumn:
        """AutoNumber (Long Integer by default; ``replication_id=True`` for a GUID)."""
        return build(
            AutoNumberColumn,
            f"autonumber column {name!r}",
            name=name,
            replication_id=replication_id,
            **options,
        )

    @staticmethod
    def date_time(
        name: str, *, input_mask: str | None = None, **options: Unpack[ColumnOptions]
    ) -> DateTimeColumn:
        """Date/Time."""
        return build(
            DateTimeColumn,
            f"date/time column {name!r}",
            name=name,
            input_mask=input_mask,
            **options,
        )

    @staticmethod
    def yes_no(name: str, **options: Unpack[ColumnOptions]) -> YesNoColumn:
        """Yes/No (Boolean)."""
        return build(YesNoColumn, f"yes/no column {name!r}", name=name, **options)

    @staticmethod
    def hyperlink(
        name: str,
        *,
        allow_zero_length: bool = False,
        unicode_compression: bool = True,
        **options: Unpack[ColumnOptions],
    ) -> HyperlinkColumn:
        """Hyperlink."""
        return build(
            HyperlinkColumn,
            f"hyperlink column {name!r}",
            name=name,
            allow_zero_length=allow_zero_length,
            unicode_compression=unicode_compression,
            **options,
        )

    @staticmethod
    def ole_object(name: str, **options: Unpack[OleObjectOptions]) -> OleObjectColumn:
        """OLE Object (long binary)."""
        return build(OleObjectColumn, f"OLE object column {name!r}", name=name, **options)
