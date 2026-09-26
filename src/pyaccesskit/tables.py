"""Tables, fields and indexes: the ``db.tables`` collection and its handles."""

from __future__ import annotations

from collections.abc import Iterator, Mapping, Sequence
from typing import TYPE_CHECKING, overload

from pyaccesskit._backends.protocols import PropertyTarget, TableInfo
from pyaccesskit._ops import schema as ops
from pyaccesskit.enums import DataType, ObjectKind
from pyaccesskit.errors import ObjectNotFoundError
from pyaccesskit.properties import PropertyBag
from pyaccesskit.schema import (
    ColumnBase,
    ColumnSpec,
    IndexSpec,
    PropertyValue,
    TableSpec,
    quote_identifier,
)
from pyaccesskit.schema._base import build

if TYPE_CHECKING:
    from pyaccesskit._session.session import Session

__all__ = ["Field", "FieldCollection", "Table", "TableCollection"]


class Field:
    """A column of a table (a live, name-based handle)."""

    def __init__(self, session: Session, table: Table, name: str) -> None:
        self._session = session
        self._table = table
        self._name = name

    @property
    def name(self) -> str:
        """The field name."""
        return self._name

    @property
    def spec(self) -> ColumnBase:
        """The current definition of the field."""
        return self._table.to_spec().column(self._name)

    @property
    def data_type(self) -> DataType:
        """The Access data type."""
        return self.spec.data_type

    @property
    def size(self) -> int | None:
        """Text length for Short Text fields, else ``None``."""
        length: int | None = getattr(self.spec, "length", None)
        return length

    @property
    def required(self) -> bool:
        """Whether Null is disallowed."""
        return self.spec.required

    @property
    def properties(self) -> PropertyBag:
        """The field's DAO properties."""
        return PropertyBag(self._session, PropertyTarget.column(self._table.name, self._name))

    def rename(self, new_name: str) -> None:
        """Rename the field."""
        self._session.check_writable(f"rename column {self._name!r}")
        ops.rename_column(self._session.schema(), self._table.name, self._name, new_name)
        self._name = new_name

    def drop(self) -> None:
        """Delete the field."""
        self._session.check_writable(f"drop column {self._name!r}")
        ops.drop_column(self._session.schema(), self._table.name, self._name)

    def __repr__(self) -> str:
        return f"<Field {self._table.name}.{self._name}>"


class FieldCollection:
    """The fields of a table, in order."""

    def __init__(self, session: Session, table: Table) -> None:
        self._session = session
        self._table = table

    def names(self) -> list[str]:
        """Field names in order."""
        return list(self._table.to_spec().column_names)

    def __iter__(self) -> Iterator[Field]:
        return iter([Field(self._session, self._table, name) for name in self.names()])

    def __len__(self) -> int:
        return len(self.names())

    def __contains__(self, name: object) -> bool:
        return isinstance(name, str) and any(n.casefold() == name.casefold() for n in self.names())

    def __getitem__(self, name: str) -> Field:
        for actual in self.names():
            if actual.casefold() == name.casefold():
                return Field(self._session, self._table, actual)
        raise ObjectNotFoundError(
            f"table {self._table.name!r} has no field {name!r}", kind=ObjectKind.FIELD, name=name
        )


class Table:
    """A table (a live, name-based handle). ``to_spec()`` returns an immutable snapshot."""

    def __init__(self, session: Session, name: str) -> None:
        self._session = session
        self._name = name

    @property
    def name(self) -> str:
        """The table name."""
        return self._name

    def _info(self) -> TableInfo:
        return ops.require_table(self._session.schema(), self._name)

    def to_spec(self) -> TableSpec:
        """The table's current definition as a normalized :class:`TableSpec`."""
        return self._session.schema().read_table(self._name)

    @property
    def fields(self) -> FieldCollection:
        """The table's fields."""
        return FieldCollection(self._session, self)

    @property
    def indexes(self) -> tuple[IndexSpec, ...]:
        """The table's indexes (excluding the hidden ones owned by relationships)."""
        return tuple(self.to_spec().indexes)

    @property
    def primary_key(self) -> IndexSpec | None:
        """The primary-key index, if any."""
        return self.to_spec().primary_index

    @property
    def description(self) -> str | None:
        """The table *Description*."""
        return self.to_spec().description

    @description.setter
    def description(self, value: str | None) -> None:
        if value is None:
            if "Description" in self.properties:
                self.properties.delete("Description")
        else:
            self.properties["Description"] = value

    @property
    def is_linked(self) -> bool:
        """Whether this is a linked table."""
        return self._info().is_linked

    @property
    def connect(self) -> str | None:
        """A linked table's connection string (``None`` for local tables). It may contain credentials."""
        return self._info().connect

    @property
    def source_table(self) -> str | None:
        """A linked table's name in its source database (``None`` for local tables)."""
        return self._info().source_table

    @property
    def is_system(self) -> bool:
        """Whether this is a system table (``MSys*``)."""
        return self._info().is_system

    @property
    def properties(self) -> PropertyBag:
        """The table's DAO properties."""
        return PropertyBag(self._session, PropertyTarget.table(self._name))

    def record_count(self) -> int:
        """Number of rows (runs ``SELECT COUNT(*)``)."""
        rows = self._session.schema().fetch(
            f"SELECT COUNT(*) AS N FROM {quote_identifier(self._name)}"
        )
        return int(rows.rows[0][0])

    # ------------------------------------------------------------------------------- mutation
    def add_column(self, column: ColumnSpec) -> Field:
        """Append a column (``unique=``/``indexed=``/``primary_key=`` also create the index)."""
        self._session.check_writable(f"add column to {self._name!r}")
        ops.add_column(self._session.schema(), self._name, column)
        return Field(self._session, self, column.name)

    def drop_column(self, name: str) -> None:
        """Delete a column (drop its indexes and relationships first)."""
        self.fields[name].drop()

    def rename_column(self, old: str, new: str) -> None:
        """Rename a column."""
        self.fields[old].rename(new)

    def create_index(self, index: IndexSpec) -> None:
        """Create an index."""
        self._session.check_writable(f"create index on {self._name!r}")
        ops.create_index(self._session.schema(), self._name, index)

    def drop_index(self, name: str) -> None:
        """Delete an index."""
        self._session.check_writable(f"drop index of {self._name!r}")
        ops.drop_index(self._session.schema(), self._name, name)

    def rename(self, new_name: str) -> None:
        """Rename the table."""
        self._session.check_writable(f"rename table {self._name!r}")
        self._name = ops.rename_table(self._session.schema(), self._name, new_name)

    def drop(self, *, drop_relationships: bool = False) -> None:
        """Delete the table (with ``drop_relationships=True``, its relationships first)."""
        self._session.check_writable(f"drop table {self._name!r}")
        ops.drop_table(self._session.schema(), self._name, drop_relationships=drop_relationships)

    def __repr__(self) -> str:
        return f"<Table {self._name!r}>"


class TableCollection:
    """``db.tables``: every local and linked table (system tables are hidden unless asked for)."""

    def __init__(self, session: Session) -> None:
        self._session = session

    def _infos(self, include_system: bool) -> list[TableInfo]:
        return [
            info
            for info in self._session.schema().list_tables()
            if include_system or not (info.is_system or info.is_hidden)
        ]

    def names(self, *, include_system: bool = False) -> list[str]:
        """Table names."""
        return [info.name for info in self._infos(include_system)]

    def __iter__(self) -> Iterator[Table]:
        return iter([Table(self._session, name) for name in self.names()])

    def __len__(self) -> int:
        return len(self.names())

    def __contains__(self, name: object) -> bool:
        return isinstance(name, str) and self.get(name) is not None

    def get(self, name: str) -> Table | None:
        """The table called ``name`` (case-insensitive), or ``None``."""
        info = ops.find_table(self._session.schema(), name)
        return Table(self._session, info.name) if info is not None else None

    def __getitem__(self, name: str) -> Table:
        table = self.get(name)
        if table is None:
            raise ObjectNotFoundError(
                f"table {name!r} does not exist", kind=ObjectKind.TABLE, name=name
            )
        return table

    @overload
    def create(self, spec: TableSpec, /) -> Table: ...

    @overload
    def create(
        self,
        name: str,
        /,
        *,
        columns: Sequence[ColumnSpec],
        indexes: Sequence[IndexSpec] = (),
        primary_key: Sequence[str] | str | None = None,
        description: str | None = None,
        validation_rule: str | None = None,
        validation_text: str | None = None,
        properties: Mapping[str, PropertyValue] | None = None,
    ) -> Table: ...

    def create(
        self,
        spec_or_name: TableSpec | str,
        /,
        *,
        columns: Sequence[ColumnSpec] | None = None,
        indexes: Sequence[IndexSpec] = (),
        primary_key: Sequence[str] | str | None = None,
        description: str | None = None,
        validation_rule: str | None = None,
        validation_text: str | None = None,
        properties: Mapping[str, PropertyValue] | None = None,
    ) -> Table:
        """Create a table from a :class:`TableSpec` or from keyword arguments.

        Creation is atomic: if any step fails, the partially created table is removed.
        """
        if isinstance(spec_or_name, TableSpec):
            spec = spec_or_name
        else:
            spec = build(
                TableSpec,
                f"table {spec_or_name!r}",
                name=spec_or_name,
                columns=tuple(columns or ()),
                indexes=tuple(indexes),
                primary_key=primary_key,
                description=description,
                validation_rule=validation_rule,
                validation_text=validation_text,
                properties=dict(properties or {}),
            )
        self._session.check_writable(f"create table {spec.name!r}")
        created = ops.create_table(self._session.schema(), spec)
        return Table(self._session, created.name)

    def drop(self, name: str, *, drop_relationships: bool = False) -> None:
        """Delete a table."""
        self[name].drop(drop_relationships=drop_relationships)

    def specs(self) -> list[TableSpec]:
        """Specs of every (non-system, non-linked) table."""
        return [
            self._session.schema().read_table(info.name)
            for info in self._infos(False)
            if not info.is_linked
        ]
