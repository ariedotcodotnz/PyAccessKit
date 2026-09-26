"""Index specifications."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from typing import Any, Self, cast

from pydantic import Field, field_validator, model_validator

from pyaccesskit.schema._base import Items, SpecModel, build
from pyaccesskit.schema.names import check_name

__all__ = ["PRIMARY_KEY_NAME", "IndexField", "IndexSpec"]

PRIMARY_KEY_NAME = "PrimaryKey"
"""The name Access gives primary-key indexes."""

MAX_FIELDS_PER_INDEX = 10


class IndexField(SpecModel):
    """One column of an index, optionally in descending order."""

    name: str
    descending: bool = False

    @field_validator("name")
    @classmethod
    def _check_name(cls, value: str) -> str:
        return check_name(value, what="index column name")


def _coerce_field(value: Any) -> Any:
    if isinstance(value, str):
        return {"name": value}
    if isinstance(value, tuple):
        pair = cast("tuple[Any, ...]", value)
        if len(pair) == 2:
            return {"name": pair[0], "descending": str(pair[1]).lower() in ("desc", "descending")}
        return pair
    return value


class IndexSpec(SpecModel):
    """An index on one or more columns.

    Attributes:
        name: Index name (Access names the primary key ``PrimaryKey``).
        fields: Indexed columns; plain strings are accepted (``["LastName", "FirstName"]``) as are
            ``("Col", "desc")`` pairs.
        primary: This is the table's primary key (implies ``unique`` and ``required``).
        unique: No two rows may have the same key.
        required: Null values are not allowed in the indexed columns.
        ignore_nulls: Rows with Null keys are left out of the index.
    """

    name: str
    fields: Items[IndexField] = Field(min_length=1, max_length=MAX_FIELDS_PER_INDEX)
    primary: bool = False
    unique: bool = False
    required: bool = False
    ignore_nulls: bool = False

    @field_validator("name")
    @classmethod
    def _check_name(cls, value: str) -> str:
        return check_name(value, what="index name")

    @field_validator("fields", mode="before")
    @classmethod
    def _coerce_fields(cls, value: Any) -> Any:
        if isinstance(value, str):
            return (_coerce_field(value),)
        if isinstance(value, Sequence) and not isinstance(value, (str, bytes)):
            items = cast("Sequence[Any]", value)
            return tuple(_coerce_field(item) for item in items)
        return value

    @model_validator(mode="before")
    @classmethod
    def _primary_implies_unique(cls, data: Any) -> Any:
        if isinstance(data, Mapping):
            fields = cast("Mapping[str, Any]", data)
            if fields.get("primary"):
                return {**fields, "unique": True, "required": True}
            return fields
        return data

    @model_validator(mode="after")
    def _check_duplicates(self) -> Self:
        seen: set[str] = set()
        for field in self.fields:
            key = field.name.casefold()
            if key in seen:
                raise ValueError(f"index {self.name!r} lists column {field.name!r} twice")
            seen.add(key)
        return self

    @property
    def field_names(self) -> tuple[str, ...]:
        """The indexed column names, in order."""
        return tuple(field.name for field in self.fields)

    @classmethod
    def primary_key(cls, *columns: str, name: str = PRIMARY_KEY_NAME) -> IndexSpec:
        """A primary-key index on ``columns`` (named ``PrimaryKey`` like Access does)."""
        return build(cls, f"primary key {name!r}", name=name, fields=columns, primary=True)

    @classmethod
    def on(
        cls, name: str, *columns: str | tuple[str, str], unique: bool = False, **options: bool
    ) -> IndexSpec:
        """Convenience constructor: ``IndexSpec.on("ix_Name", "LastName", "FirstName", unique=True)``."""
        return build(cls, f"index {name!r}", name=name, fields=columns, unique=unique, **options)
