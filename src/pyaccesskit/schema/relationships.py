# pyright: reportUnnecessaryIsInstance=false
# (parse_column_ref validates untyped user input at runtime)
"""Relationship specifications."""

from __future__ import annotations

from collections.abc import Sequence
from typing import Any, Self

from pydantic import Field, field_validator, model_validator

from pyaccesskit.enums import JoinType
from pyaccesskit.errors import SpecError
from pyaccesskit.schema._base import Items, SpecModel, build
from pyaccesskit.schema.names import check_name

__all__ = ["ColumnRef", "RelationshipSpec", "parse_column_ref"]

ColumnRef = str | tuple[str, str | Sequence[str]]
"""A relationship end: ``"Table.Column"``, ``("Table", "Column")`` or ``("Table", ["Col1", "Col2"])``."""


def parse_column_ref(ref: ColumnRef, *, what: str) -> tuple[str, tuple[str, ...]]:
    """Split a :data:`ColumnRef` into ``(table, columns)``.

    Raises:
        SpecError: If the reference is malformed.
    """
    if isinstance(ref, str):
        table, sep, column = ref.partition(".")
        if not sep or not table or not column or "." in column:
            raise SpecError(f"{what} must look like 'Table.Column', got {ref!r}")
        return check_name(table, what=f"{what} table"), (check_name(column, what=f"{what} column"),)
    if isinstance(ref, tuple) and len(ref) == 2:
        table, columns = ref
        names = (columns,) if isinstance(columns, str) else tuple(columns)
        if not names:
            raise SpecError(f"{what} needs at least one column")
        return check_name(table, what=f"{what} table"), tuple(
            check_name(c, what=f"{what} column") for c in names
        )
    raise SpecError(f"{what} must be 'Table.Column' or (table, columns), got {ref!r}")


class RelationshipSpec(SpecModel):
    """A relationship between a primary ("one") table and a foreign ("many") table.

    Attributes:
        name: Relationship name. Defaults to Access's convention: primary table name + foreign table name.
        primary_table: The table on the "one" side (its columns need a primary key or unique index).
        primary_columns: Referenced columns of the primary table.
        foreign_table: The table on the "many" side.
        foreign_columns: Referencing columns of the foreign table (same count and compatible types).
        enforce_integrity: Enforce referential integrity (the default; creates a hidden index).
        cascade_update: Cascade updates of the primary key to related rows.
        cascade_delete: Cascade deletes to related rows.
        one_to_one: Declare a one-to-one relationship.
        join: Default join type used by the query designer.
    """

    name: str | None = None
    primary_table: str
    primary_columns: Items[str] = Field(min_length=1, max_length=10)
    foreign_table: str
    foreign_columns: Items[str] = Field(min_length=1, max_length=10)
    enforce_integrity: bool = True
    cascade_update: bool = False
    cascade_delete: bool = False
    one_to_one: bool = False
    join: JoinType = JoinType.INNER

    @field_validator("name")
    @classmethod
    def _check_name(cls, value: str | None) -> str | None:
        return None if value is None else check_name(value, what="relationship name")

    @field_validator("primary_table", "foreign_table")
    @classmethod
    def _check_table(cls, value: str) -> str:
        return check_name(value, what="table name")

    @field_validator("primary_columns", "foreign_columns", mode="before")
    @classmethod
    def _coerce_columns(cls, value: Any) -> Any:
        return (value,) if isinstance(value, str) else value

    @field_validator("primary_columns", "foreign_columns")
    @classmethod
    def _check_columns(cls, value: Sequence[str]) -> tuple[str, ...]:
        return tuple(check_name(column, what="column name") for column in value)

    @model_validator(mode="after")
    def _check_relationship(self) -> Self:
        if len(self.primary_columns) != len(self.foreign_columns):
            raise ValueError(
                f"primary side has {len(self.primary_columns)} column(s) but foreign side has "
                f"{len(self.foreign_columns)}; relationships pair columns one to one"
            )
        if (self.cascade_update or self.cascade_delete) and not self.enforce_integrity:
            raise ValueError("cascade_update/cascade_delete require enforce_integrity=True")
        return self

    @property
    def effective_name(self) -> str:
        """``name`` or Access's default name (primary table + foreign table)."""
        return self.name if self.name is not None else f"{self.primary_table}{self.foreign_table}"

    def normalized(self) -> RelationshipSpec:
        """Canonical form with the name filled in."""
        if self.name is not None:
            return self
        return self.model_copy(update={"name": self.effective_name})

    @classmethod
    def between(cls, primary: ColumnRef, foreign: ColumnRef, **options: Any) -> RelationshipSpec:
        """Build a relationship from ``"Table.Column"`` references.

        Example::

        RelationshipSpec.between("Customers.CustomerID", "Orders.CustomerID", cascade_delete=True)
        """
        primary_table, primary_columns = parse_column_ref(primary, what="primary")
        foreign_table, foreign_columns = parse_column_ref(foreign, what="foreign")
        return build(
            cls,
            f"relationship {primary_table}->{foreign_table}",
            primary_table=primary_table,
            primary_columns=primary_columns,
            foreign_table=foreign_table,
            foreign_columns=foreign_columns,
            **options,
        )
