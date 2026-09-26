"""Table specifications."""

from __future__ import annotations

from typing import Self

from pydantic import Field, field_validator, model_validator

from pyaccesskit.enums import DataType
from pyaccesskit.schema._base import Items, PropertyValue, SpecModel
from pyaccesskit.schema.columns import ColumnBase, ColumnSpec
from pyaccesskit.schema.indexes import PRIMARY_KEY_NAME, IndexField, IndexSpec
from pyaccesskit.schema.names import check_name

__all__ = ["MAX_COLUMNS", "MAX_INDEXES", "TableSpec"]

MAX_COLUMNS = 255
MAX_INDEXES = 32
"""Access limit per table, *including* the hidden indexes created for enforced relationships."""


class TableSpec(SpecModel):
    """A local Access table.

    Indexes can be declared explicitly in ``indexes`` or with shorthands: ``primary_key=`` here, or
    ``primary_key=``/``unique=``/``indexed=`` on individual columns. :meth:`normalized` expands every
    shorthand into explicit :class:`IndexSpec` objects; that canonical form is what introspection returns.

    Attributes:
        name: Table name.
        columns: Columns in order (1-255).
        indexes: Explicit indexes.
        primary_key: Shorthand for a (possibly composite) primary key.
        description: Table *Description* property.
        validation_rule: Table-level *Validation Rule* (can reference several columns).
        validation_text: Message shown when the table validation rule fails.
        properties: Other Access/DAO table properties to set verbatim (escape hatch).
    """

    name: str
    columns: Items[ColumnSpec] = Field(min_length=1, max_length=MAX_COLUMNS)
    indexes: Items[IndexSpec] = ()
    primary_key: Items[str] | None = None
    description: str | None = None
    validation_rule: str | None = None
    validation_text: str | None = None
    properties: dict[str, PropertyValue] = Field(default_factory=dict)

    @field_validator("name")
    @classmethod
    def _check_name(cls, value: str) -> str:
        return check_name(value, what="table name")

    @field_validator("primary_key", mode="before")
    @classmethod
    def _coerce_primary_key(cls, value: object) -> object:
        return (value,) if isinstance(value, str) else value

    @model_validator(mode="after")
    def _check_table(self) -> Self:
        names: dict[str, str] = {}
        for column in self.columns:
            key = column.name.casefold()
            if key in names:
                raise ValueError(
                    f"duplicate column name {column.name!r} (names are case-insensitive)"
                )
            names[key] = column.name
        autonumbers = [c.name for c in self.columns if c.data_type is DataType.AUTONUMBER]
        if len(autonumbers) > 1:
            raise ValueError(f"a table can have only one AutoNumber column, got {autonumbers}")
        if self.validation_text is not None and self.validation_rule is None:
            raise ValueError("validation_text requires a validation_rule")

        indexes = self._expand_indexes()
        if len(indexes) > MAX_INDEXES:
            raise ValueError(f"a table can have at most {MAX_INDEXES} indexes, got {len(indexes)}")
        index_names: set[str] = set()
        for index in indexes:
            key = index.name.casefold()
            if key in index_names:
                raise ValueError(f"duplicate index name {index.name!r}")
            index_names.add(key)
            for field_name in index.field_names:
                column = self._find(field_name)
                if column is None:
                    raise ValueError(
                        f"index {index.name!r} refers to unknown column {field_name!r}"
                    )
                if not column.indexable:
                    raise ValueError(
                        f"{column.data_type.value} column {column.name!r} cannot be indexed"
                    )
        return self

    # ------------------------------------------------------------------------------------ helpers
    def _find(self, name: str) -> ColumnBase | None:
        key = name.casefold()
        return next((c for c in self.columns if c.name.casefold() == key), None)

    def column(self, name: str) -> ColumnBase:
        """Return the column called ``name`` (case-insensitive).

        Raises:
            KeyError: If there is no such column.
        """
        column = self._find(name)
        if column is None:
            raise KeyError(name)
        return column

    @property
    def column_names(self) -> tuple[str, ...]:
        """Column names in order."""
        return tuple(column.name for column in self.columns)

    def _expand_indexes(self) -> tuple[IndexSpec, ...]:
        explicit = list(self.indexes)
        primary_sources = [
            label
            for label, present in (
                ("indexes", any(index.primary for index in explicit)),
                ("primary_key=", self.primary_key is not None),
                ("column primary_key=True", any(column.primary_key for column in self.columns)),
            )
            if present
        ]
        if len(primary_sources) > 1:
            raise ValueError(f"primary key declared more than once ({', '.join(primary_sources)})")
        if sum(1 for index in explicit if index.primary) > 1:
            raise ValueError("a table can have only one primary key")

        derived: list[IndexSpec] = []
        pk_columns = tuple(self.primary_key or (c.name for c in self.columns if c.primary_key))
        if pk_columns:
            derived.append(
                IndexSpec(
                    name=PRIMARY_KEY_NAME,
                    fields=tuple(IndexField(name=name) for name in pk_columns),
                    primary=True,
                )
            )
        pk_keys = {name.casefold() for name in pk_columns}
        for primary in (index for index in explicit if index.primary):
            pk_keys = {name.casefold() for name in primary.field_names}
        single_pk = len(pk_keys) == 1
        for column in self.columns:
            if not (column.unique or column.indexed):
                continue
            if single_pk and column.name.casefold() in pk_keys:
                continue  # the primary key already indexes it uniquely
            derived.append(
                IndexSpec(
                    name=column.name, fields=(IndexField(name=column.name),), unique=column.unique
                )
            )

        combined = explicit + derived
        return tuple(sorted(combined, key=lambda index: (not index.primary, index.name.casefold())))

    def effective_indexes(self) -> tuple[IndexSpec, ...]:
        """All indexes with shorthands expanded, primary key first, then by name."""
        return self._expand_indexes()

    @property
    def primary_index(self) -> IndexSpec | None:
        """The primary-key index, if any (shorthands included)."""
        return next((index for index in self._expand_indexes() if index.primary), None)

    def normalized(self) -> TableSpec:
        """The canonical form: shorthands expanded into ``indexes`` and cleared from columns."""
        return self.model_copy(
            update={
                "columns": tuple(column.normalized() for column in self.columns),
                "indexes": self._expand_indexes(),
                "primary_key": None,
            }
        )
