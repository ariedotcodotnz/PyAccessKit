"""The typed seam between PyAccessKit's COM-free core and its COM adapters.

Backends exchange **specs and plain data only** — never COM objects. Operations are deliberately sized
like the steps of a future ``plan``/``apply`` (create table, add column, create relationship...).

Implementations:

* ``_backends.dao``    — DAO over any transport (in-process, Access-hosted, or ``CurrentDb``);
* ``_backends.access`` — ``Access.Application`` design features;
* ``_backends.fake``   — in-memory, used by the test-suite and kept honest by the contract tests.

Backends *may* assume their inputs are normalized specs whose names were validated; semantic
pre-validation (does the table exist? are the relationship types compatible?) happens in ``_ops`` so that
every backend reports problems the same way. Backends still translate engine errors into
:mod:`pyaccesskit.errors` exceptions.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal, Protocol

from pyaccesskit.enums import (
    ControlKind,
    DataType,
    ObjectKind,
    PropertyType,
    QueryKind,
    Section,
    Transport,
)
from pyaccesskit.forms.layout import ResolvedForm
from pyaccesskit.schema import (
    ColumnSpec,
    IndexSpec,
    PropertyValue,
    QuerySpec,
    RelationshipSpec,
    TableSpec,
)
from pyaccesskit.units import Length

__all__ = [
    "ControlInfo",
    "DatabaseInfo",
    "DesignBackend",
    "FetchResult",
    "ParameterInfo",
    "PropertyTarget",
    "QueryInfo",
    "SchemaBackend",
    "TableInfo",
]


@dataclass(frozen=True)
class TableInfo:
    """Summary of a table (cheap to list)."""

    name: str
    is_linked: bool = False
    is_system: bool = False
    is_hidden: bool = False
    connect: str | None = None
    source_table: str | None = None


@dataclass(frozen=True)
class QueryInfo:
    """Summary of a saved query."""

    name: str
    kind: QueryKind
    is_hidden: bool = False


@dataclass(frozen=True)
class ParameterInfo:
    """A query parameter (name without brackets)."""

    name: str
    data_type: DataType


@dataclass(frozen=True)
class FetchResult:
    """Rows returned by a SELECT: column names plus row tuples."""

    columns: tuple[str, ...]
    rows: list[tuple[Any, ...]]

    def as_dicts(self) -> list[dict[str, Any]]:
        """Rows as ``{column: value}`` dictionaries."""
        return [dict(zip(self.columns, row, strict=True)) for row in self.rows]


@dataclass(frozen=True)
class ControlInfo:
    """A control read back from a saved form."""

    name: str
    kind: ControlKind
    section: Section | None
    left: Length
    top: Length
    width: Length
    height: Length
    control_source: str | None = None
    caption: str | None = None
    parent: str | None = None
    """For attached labels: the control the label belongs to."""


@dataclass(frozen=True)
class DatabaseInfo:
    """Facts about the open database."""

    path: Path
    version: str
    transport: Transport


@dataclass(frozen=True)
class PropertyTarget:
    """Addresses the object whose DAO ``Properties`` collection is used."""

    kind: Literal["database", "table", "field", "query"]
    name: str | None = None
    field: str | None = None

    @classmethod
    def database(cls) -> PropertyTarget:
        """The database itself (``AppTitle``, ``StartUpForm``...)."""
        return cls("database")

    @classmethod
    def table(cls, name: str) -> PropertyTarget:
        """A table."""
        return cls("table", name)

    @classmethod
    def column(cls, table: str, field: str) -> PropertyTarget:
        """A field of a table."""
        return cls("field", table, field)

    @classmethod
    def query(cls, name: str) -> PropertyTarget:
        """A saved query."""
        return cls("query", name)

    def describe(self) -> str:
        """Human-readable description for messages."""
        if self.kind == "database":
            return "the database"
        if self.kind == "field":
            return f"field {self.field!r} of table {self.name!r}"
        return f"{self.kind} {self.name!r}"


class SchemaBackend(Protocol):
    """Schema and data operations (implemented with DAO)."""

    def database_info(self) -> DatabaseInfo:
        """Facts about the open database."""
        ...

    # ------------------------------------------------------------------------------------ tables
    def list_tables(self) -> list[TableInfo]:
        """All tables, including system and linked tables."""
        ...

    def read_table(self, name: str) -> TableSpec:
        """The normalized spec of a local table (relationship-owned indexes excluded)."""
        ...

    def count_indexes(self, table: str) -> int:
        """Number of indexes on ``table`` *including* hidden relationship indexes (Access limit: 32)."""
        ...

    def create_table(self, spec: TableSpec) -> None:
        """Create a table from a normalized spec, atomically (all or nothing)."""
        ...

    def drop_table(self, name: str) -> None:
        """Delete a table."""
        ...

    def rename_table(self, old: str, new: str) -> None:
        """Rename a table."""
        ...

    def add_column(self, table: str, column: ColumnSpec) -> None:
        """Append a column (index shorthands are handled by the caller)."""
        ...

    def drop_column(self, table: str, column: str) -> None:
        """Delete a column."""
        ...

    def rename_column(self, table: str, old: str, new: str) -> None:
        """Rename a column."""
        ...

    def create_index(self, table: str, index: IndexSpec) -> None:
        """Create an index."""
        ...

    def drop_index(self, table: str, name: str) -> None:
        """Delete an index."""
        ...

    # ----------------------------------------------------------------------------- relationships
    def list_relationships(self) -> list[RelationshipSpec]:
        """All (non-inherited) relationships, normalized."""
        ...

    def create_relationship(self, spec: RelationshipSpec) -> None:
        """Create a relationship from a normalized, pre-validated spec."""
        ...

    def drop_relationship(self, name: str) -> None:
        """Delete a relationship."""
        ...

    # ----------------------------------------------------------------------------------- queries
    def list_queries(self) -> list[QueryInfo]:
        """All saved queries (including hidden ``~`` queries)."""
        ...

    def read_query(self, name: str) -> QuerySpec:
        """The spec of a saved query (SQL as Access stored it)."""
        ...

    def query_parameters(self, name: str) -> list[ParameterInfo]:
        """Declared and implicit parameters of a saved query."""
        ...

    def create_query(self, spec: QuerySpec) -> None:
        """Create a saved query."""
        ...

    def set_query_sql(self, name: str, sql: str) -> None:
        """Replace the SQL of a saved query."""
        ...

    def rename_query(self, old: str, new: str) -> None:
        """Rename a saved query."""
        ...

    def drop_query(self, name: str) -> None:
        """Delete a saved query."""
        ...

    # -------------------------------------------------------------------------------- properties
    def get_property(self, target: PropertyTarget, name: str) -> PropertyValue:
        """Read a DAO property. Raises ``ObjectNotFoundError`` (kind PROPERTY) if it does not exist."""
        ...

    def set_property(
        self,
        target: PropertyTarget,
        name: str,
        value: PropertyValue,
        type: PropertyType | None = None,
    ) -> None:
        """Set (creating if needed) a DAO property."""
        ...

    def delete_property(self, target: PropertyTarget, name: str) -> None:
        """Delete a user-defined DAO property."""
        ...

    def list_properties(self, target: PropertyTarget) -> dict[str, PropertyValue]:
        """All readable properties of the target."""
        ...

    # -------------------------------------------------------------------------------------- data
    def execute(self, sql: str, params: Mapping[str, Any] | None = None) -> int:
        """Run an action statement; returns the number of records affected."""
        ...

    def fetch(
        self, sql: str, params: Mapping[str, Any] | None = None, *, limit: int | None = None
    ) -> FetchResult:
        """Run a SELECT and return its rows."""
        ...

    def execute_saved(self, name: str, params: Mapping[str, Any] | None = None) -> int:
        """Run a saved action query."""
        ...

    def fetch_saved(
        self, name: str, params: Mapping[str, Any] | None = None, *, limit: int | None = None
    ) -> FetchResult:
        """Return the rows of a saved select query."""
        ...

    # --------------------------------------------------------------------------------- documents
    def list_documents(self, kind: ObjectKind) -> list[str]:
        """Names of forms, reports, macros or modules (read through DAO containers; no Access needed)."""
        ...


class DesignBackend(Protocol):
    """Operations that need Microsoft Access itself (``Access.Application``)."""

    def list_objects(self, kind: ObjectKind) -> list[str]:
        """Names of the forms, reports, macros or modules in the database."""
        ...

    def delete_object(self, kind: ObjectKind, name: str) -> None:
        """Delete a form, report, macro or module."""
        ...

    def rename_object(self, kind: ObjectKind, old: str, new: str) -> None:
        """Rename a form, report, macro or module."""
        ...

    def export_text(self, kind: ObjectKind, name: str) -> bytes:
        """``SaveAsText`` bytes of an object (native encoding; see ``_text.codec``)."""
        ...

    def import_text(self, kind: ObjectKind, name: str, data: bytes) -> None:
        """``LoadFromText`` from native-encoded bytes (replaces an existing object)."""
        ...

    def build_form(self, form: ResolvedForm, *, replace: bool) -> None:
        """Create (or atomically replace) a form from a resolved layout."""
        ...

    def form_controls(self, name: str) -> list[ControlInfo]:
        """Read the controls of a saved form."""
        ...

    def check_form_opens(self, name: str) -> None:
        """Open the form in Form view (hidden) and close it again; raises if Access reports an error."""
        ...
