"""Storage-type rules shared by pre-validation and the in-memory backend."""

from __future__ import annotations

from pyaccesskit.enums import NumberSize
from pyaccesskit.schema.columns import (
    AutoNumberColumn,
    ColumnBase,
    CurrencyColumn,
    DateTimeColumn,
    DecimalColumn,
    HyperlinkColumn,
    LongTextColumn,
    NumberColumn,
    OleObjectColumn,
    TextColumn,
    YesNoColumn,
)

__all__ = ["relationship_compatible", "storage_type"]

_NUMBER_STORAGE = {
    NumberSize.BYTE: "byte",
    NumberSize.INTEGER: "integer",
    NumberSize.LONG_INTEGER: "long",
    NumberSize.SINGLE: "single",
    NumberSize.DOUBLE: "double",
    NumberSize.REPLICATION_ID: "guid",
}


def storage_type(column: ColumnBase) -> str:
    """The engine-level storage type of a column (AutoNumber stores a Long Integer, etc.)."""
    if isinstance(column, AutoNumberColumn):
        return "guid" if column.replication_id else "long"
    if isinstance(column, NumberColumn):
        return _NUMBER_STORAGE[column.size]
    simple: tuple[tuple[type[ColumnBase], str], ...] = (
        (TextColumn, "text"),
        (LongTextColumn, "memo"),
        (HyperlinkColumn, "memo"),
        (CurrencyColumn, "currency"),
        (DecimalColumn, "decimal"),
        (DateTimeColumn, "datetime"),
        (YesNoColumn, "boolean"),
        (OleObjectColumn, "binary"),
    )
    for cls, kind in simple:
        if isinstance(column, cls):
            return kind
    return "unknown"  # pragma: no cover - every column class is listed above


def relationship_compatible(primary: ColumnBase, foreign: ColumnBase) -> bool:
    """Whether two columns can be paired in a relationship (same storage type, indexable)."""
    kind = storage_type(primary)
    return kind == storage_type(foreign) and kind not in ("memo", "binary", "unknown")
