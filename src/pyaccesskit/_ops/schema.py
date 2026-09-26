"""Schema operations: semantic pre-validation + linting, then one backend call.

Every public schema mutation goes through here, so users get the same precise errors regardless of the
backend (in-process DAO, Access-hosted DAO, or the in-memory fake) — and most mistakes are reported
before Access is touched at all.
"""

from __future__ import annotations

import contextlib
import uuid
from collections.abc import Iterable

from pyaccesskit._backends.protocols import SchemaBackend, TableInfo
from pyaccesskit.enums import DataType, ObjectKind
from pyaccesskit.errors import (
    ObjectExistsError,
    ObjectNotFoundError,
    RelationshipError,
    SchemaError,
    SpecError,
)
from pyaccesskit.schema import (
    ColumnBase,
    ColumnSpec,
    IndexField,
    IndexSpec,
    QuerySpec,
    RelationshipSpec,
    TableSpec,
    UnsupportedColumn,
    sql_equivalent,
)
from pyaccesskit.schema.compat import relationship_compatible, storage_type
from pyaccesskit.schema.names import check_name, warn_name
from pyaccesskit.schema.tables import MAX_INDEXES

__all__ = [
    "add_column",
    "create_index",
    "create_or_replace_query",
    "create_relationship",
    "create_table",
    "drop_column",
    "drop_index",
    "drop_query",
    "drop_relationship",
    "drop_table",
    "find_query",
    "find_table",
    "rename_column",
    "rename_query",
    "rename_table",
]


def _key(name: str) -> str:
    return name.casefold()


# ------------------------------------------------------------------------------------- lookups
def find_table(schema: SchemaBackend, name: str) -> TableInfo | None:
    """The table called ``name`` (case-insensitive), or ``None``."""
    key = _key(name)
    return next((t for t in schema.list_tables() if _key(t.name) == key), None)


def require_table(schema: SchemaBackend, name: str) -> TableInfo:
    """Like :func:`find_table` but raises :class:`ObjectNotFoundError`."""
    info = find_table(schema, name)
    if info is None:
        raise ObjectNotFoundError(
            f"table {name!r} does not exist", kind=ObjectKind.TABLE, name=name
        )
    return info


def find_query(schema: SchemaBackend, name: str) -> str | None:
    """The actual name of the query called ``name`` (case-insensitive), or ``None``."""
    key = _key(name)
    return next((q.name for q in schema.list_queries() if _key(q.name) == key), None)


def ensure_name_free(schema: SchemaBackend, name: str, *, ignore: str | None = None) -> None:
    """Raise :class:`ObjectExistsError` if a table or query already uses ``name``.

    Tables and queries share one namespace in Access. ``ignore`` skips an object being renamed.
    """
    key = _key(name)
    if ignore is not None and _key(ignore) == key:
        return
    table = find_table(schema, name)
    if table is not None:
        raise ObjectExistsError(
            f"a table named {table.name!r} already exists", kind=ObjectKind.TABLE, name=name
        )
    query = find_query(schema, name)
    if query is not None:
        raise ObjectExistsError(
            f"a query named {query!r} already exists (tables and queries share one namespace)",
            kind=ObjectKind.QUERY,
            name=name,
        )


def _column(spec: TableSpec, name: str) -> ColumnBase:
    try:
        return spec.column(name)
    except KeyError:
        raise ObjectNotFoundError(
            f"table {spec.name!r} has no column {name!r} (columns: {', '.join(spec.column_names)})",
            kind=ObjectKind.FIELD,
            name=name,
        ) from None


def _lint_table(spec: TableSpec) -> None:
    warn_name(spec.name, what="table name", stacklevel=5)
    for column in spec.columns:
        warn_name(column.name, what=f"column name in table {spec.name!r}", stacklevel=5)


# -------------------------------------------------------------------------------------- tables
def _check_creatable(column: ColumnBase, table: str) -> None:
    if isinstance(column, UnsupportedColumn):
        raise SpecError(
            f"cannot create column {column.name!r} of table {table!r}: {column.detail or 'this type'} "
            "columns can be read but not created by PyAccessKit yet"
        )


def create_table(schema: SchemaBackend, spec: TableSpec, *, lint: bool = True) -> TableSpec:
    """Create a table; returns the normalized spec that was created."""
    for column in spec.columns:
        _check_creatable(column, spec.name)
    normalized = spec.normalized()
    ensure_name_free(schema, normalized.name)
    if lint:
        _lint_table(normalized)
    schema.create_table(normalized)
    return normalized


def drop_table(schema: SchemaBackend, name: str, *, drop_relationships: bool = False) -> None:
    """Delete a table (optionally deleting the relationships it takes part in first)."""
    info = require_table(schema, name)
    related = [
        rel
        for rel in schema.list_relationships()
        if _key(rel.primary_table) == _key(info.name) or _key(rel.foreign_table) == _key(info.name)
    ]
    if related and not drop_relationships:
        names = ", ".join(repr(rel.effective_name) for rel in related)
        raise SchemaError(
            f"table {info.name!r} is part of relationship(s) {names}; delete them first or pass "
            "drop_relationships=True"
        )
    for rel in related:
        schema.drop_relationship(rel.effective_name)
    schema.drop_table(info.name)


def rename_table(schema: SchemaBackend, old: str, new: str) -> str:
    """Rename a table; returns the new name."""
    info = require_table(schema, old)
    check_name(new, what="table name")
    ensure_name_free(schema, new, ignore=info.name)
    warn_name(new, what="table name", stacklevel=4)
    schema.rename_table(info.name, new)
    return new


def _index_budget(schema: SchemaBackend, table: str, extra: int) -> None:
    used = schema.count_indexes(table)
    if used + extra > MAX_INDEXES:
        raise SchemaError(
            f"table {table!r} would need {used + extra} indexes; Access allows {MAX_INDEXES} per table "
            "(including hidden indexes of enforced relationships)"
        )


def add_column(schema: SchemaBackend, table: str, column: ColumnSpec, *, lint: bool = True) -> None:
    """Append a column, creating any index requested with ``unique=``/``indexed=``/``primary_key=``."""
    info = require_table(schema, table)
    _check_creatable(column, info.name)
    spec = schema.read_table(info.name)
    if any(_key(c.name) == _key(column.name) for c in spec.columns):
        raise ObjectExistsError(
            f"table {spec.name!r} already has a column {column.name!r}",
            kind=ObjectKind.FIELD,
            name=column.name,
        )
    if column.data_type is DataType.AUTONUMBER and any(
        c.data_type is DataType.AUTONUMBER for c in spec.columns
    ):
        raise SchemaError(f"table {spec.name!r} already has an AutoNumber column")
    new_indexes: list[IndexSpec] = []
    if column.primary_key:
        if spec.primary_index is not None:
            raise SchemaError(f"table {spec.name!r} already has a primary key")
        new_indexes.append(IndexSpec.primary_key(column.name))
    elif column.unique or column.indexed:
        new_indexes.append(
            IndexSpec(
                name=column.name, fields=(IndexField(name=column.name),), unique=column.unique
            )
        )
    for index in new_indexes:  # check before mutating: a clash would otherwise strand the column
        if any(_key(i.name) == _key(index.name) for i in spec.indexes):
            raise ObjectExistsError(
                f"table {spec.name!r} already has an index named {index.name!r}; add the column "
                "without the shorthand and create the index with an explicit name",
                kind=ObjectKind.INDEX,
                name=index.name,
            )
    if new_indexes:
        _index_budget(schema, spec.name, len(new_indexes))
    if lint:
        warn_name(column.name, what=f"column name in table {spec.name!r}", stacklevel=4)
    schema.add_column(spec.name, column.normalized())
    created: list[str] = []
    try:
        for index in new_indexes:
            schema.create_index(spec.name, index)
            created.append(index.name)
    except BaseException:
        with contextlib.suppress(Exception):
            for name in reversed(created):
                schema.drop_index(spec.name, name)
            schema.drop_column(spec.name, column.name)
        raise


def drop_column(schema: SchemaBackend, table: str, column: str) -> None:
    """Delete a column (it must not be part of an index or relationship)."""
    info = require_table(schema, table)
    spec = schema.read_table(info.name)
    target = _column(spec, column)
    indexes = [
        i.name for i in spec.indexes if any(_key(f) == _key(target.name) for f in i.field_names)
    ]
    if indexes:
        raise SchemaError(f"column {target.name!r} is part of index(es) {indexes}; drop them first")
    for rel in schema.list_relationships():
        sides: list[str] = []
        if _key(rel.primary_table) == _key(spec.name):
            sides.extend(rel.primary_columns)
        if _key(rel.foreign_table) == _key(spec.name):
            sides.extend(rel.foreign_columns)
        if any(_key(c) == _key(target.name) for c in sides):
            raise SchemaError(
                f"column {target.name!r} is part of relationship {rel.effective_name!r}; drop it first"
            )
    if len(spec.columns) == 1:
        raise SchemaError(f"cannot drop the last column of table {spec.name!r}")
    schema.drop_column(spec.name, target.name)


def rename_column(schema: SchemaBackend, table: str, old: str, new: str) -> None:
    """Rename a column."""
    info = require_table(schema, table)
    spec = schema.read_table(info.name)
    target = _column(spec, old)
    check_name(new, what="column name")
    if _key(old) != _key(new) and any(_key(c.name) == _key(new) for c in spec.columns):
        raise ObjectExistsError(
            f"table {spec.name!r} already has a column {new!r}", kind=ObjectKind.FIELD, name=new
        )
    warn_name(new, what=f"column name in table {spec.name!r}", stacklevel=4)
    schema.rename_column(spec.name, target.name, new)


def create_index(schema: SchemaBackend, table: str, index: IndexSpec) -> None:
    """Create an index after checking columns, name and the 32-index budget."""
    info = require_table(schema, table)
    spec = schema.read_table(info.name)
    for field_name in index.field_names:
        column = _column(spec, field_name)
        if not column.indexable:
            raise SpecError(f"{column.data_type.value} column {column.name!r} cannot be indexed")
    if index.primary and spec.primary_index is not None:
        raise SchemaError(
            f"table {spec.name!r} already has a primary key ({spec.primary_index.name!r})"
        )
    if any(_key(i.name) == _key(index.name) for i in spec.indexes):
        raise ObjectExistsError(
            f"table {spec.name!r} already has an index named {index.name!r}",
            kind=ObjectKind.INDEX,
            name=index.name,
        )
    _index_budget(schema, spec.name, 1)
    schema.create_index(spec.name, index)


def drop_index(schema: SchemaBackend, table: str, name: str) -> None:
    """Delete an index."""
    info = require_table(schema, table)
    spec = schema.read_table(info.name)
    match = next((i for i in spec.indexes if _key(i.name) == _key(name)), None)
    if match is None:
        raise ObjectNotFoundError(
            f"table {spec.name!r} has no index {name!r}", kind=ObjectKind.INDEX, name=name
        )
    for rel in schema.list_relationships():
        if (
            match.unique
            and _key(rel.primary_table) == _key(spec.name)
            and {_key(c) for c in rel.primary_columns} == {_key(f) for f in match.field_names}
        ):
            raise SchemaError(
                f"index {match.name!r} is required by relationship {rel.effective_name!r}; drop the relationship first"
            )
    schema.drop_index(spec.name, match.name)


# ------------------------------------------------------------------------------- relationships
def _canonical_columns(spec: TableSpec, names: Iterable[str]) -> tuple[ColumnBase, ...]:
    return tuple(_column(spec, name) for name in names)


def create_relationship(schema: SchemaBackend, spec: RelationshipSpec) -> RelationshipSpec:
    """Validate a relationship thoroughly, then create it. Returns the normalized spec used."""
    primary_info = require_table(schema, spec.primary_table)
    foreign_info = require_table(schema, spec.foreign_table)
    primary = schema.read_table(primary_info.name)
    foreign = schema.read_table(foreign_info.name)
    primary_cols = _canonical_columns(primary, spec.primary_columns)
    foreign_cols = _canonical_columns(foreign, spec.foreign_columns)

    normalized = spec.model_copy(
        update={
            "primary_table": primary.name,
            "foreign_table": foreign.name,
            "primary_columns": tuple(c.name for c in primary_cols),
            "foreign_columns": tuple(c.name for c in foreign_cols),
            "name": spec.name or f"{primary.name}{foreign.name}",
        }
    )
    name = normalized.effective_name
    existing = schema.list_relationships()
    if any(_key(rel.effective_name) == _key(name) for rel in existing):
        hint = (
            ""
            if spec.name
            else "; pass name=... to create another relationship between these tables"
        )
        raise ObjectExistsError(
            f"a relationship named {name!r} already exists{hint}",
            kind=ObjectKind.RELATIONSHIP,
            name=name,
        )
    if any(_key(i.name) == _key(name) for i in foreign.indexes):
        raise ObjectExistsError(
            f"relationship name {name!r} clashes with an index of table {foreign.name!r} (the relationship's "
            "hidden index takes its name)",
            kind=ObjectKind.RELATIONSHIP,
            name=name,
        )
    for p_col, f_col in zip(primary_cols, foreign_cols, strict=True):
        if not relationship_compatible(p_col, f_col):
            raise RelationshipError(
                f"cannot relate {primary.name}.{p_col.name} ({storage_type(p_col)}) to "
                f"{foreign.name}.{f_col.name} ({storage_type(f_col)}): related columns must have the same "
                "storage type (an AutoNumber pairs with a Number/Long Integer)"
            )
    wanted = {_key(c.name) for c in primary_cols}
    if not any(i.unique and {_key(f) for f in i.field_names} == wanted for i in primary.indexes):
        columns = ", ".join(c.name for c in primary_cols)
        raise RelationshipError(
            f"{primary.name}({columns}) must be the primary key or have a unique index to be referenced by "
            "a relationship"
        )
    if normalized.enforce_integrity:
        _index_budget(schema, foreign.name, 1)
    warn_name(name, what="relationship name", stacklevel=4)
    schema.create_relationship(normalized)
    return normalized


def drop_relationship(schema: SchemaBackend, name: str) -> None:
    """Delete a relationship."""
    match = next(
        (r for r in schema.list_relationships() if _key(r.effective_name) == _key(name)), None
    )
    if match is None:
        raise ObjectNotFoundError(
            f"relationship {name!r} does not exist", kind=ObjectKind.RELATIONSHIP, name=name
        )
    schema.drop_relationship(match.effective_name)


# ------------------------------------------------------------------------------------- queries
def create_or_replace_query(
    schema: SchemaBackend, spec: QuerySpec, *, replace: bool = False
) -> bool:
    """Create a saved query, or replace its SQL when ``replace=True``.

    Returns ``True`` if the query was created or changed, ``False`` if an identical query already existed.
    """
    existing = find_query(schema, spec.name)
    if existing is not None:
        if not replace:
            raise ObjectExistsError(
                f"a query named {existing!r} already exists", kind=ObjectKind.QUERY, name=spec.name
            )
        current = schema.read_query(existing)
        if current.pass_through != spec.pass_through:
            _swap_query(schema, existing, spec.normalized())
            return True
        if spec.pass_through is not None:
            # Server dialect: Access's SQL normalization (case folding...) does not apply.
            unchanged = _line_endings(current.sql) == _line_endings(spec.sql)
        else:
            unchanged = sql_equivalent(current.sql, spec.sql)
        if unchanged:
            return False
        schema.set_query_sql(existing, spec.normalized().sql)
        return True
    ensure_name_free(schema, spec.name)
    warn_name(spec.name, what="query name", stacklevel=4)
    schema.create_query(spec.normalized())
    return True


def _line_endings(sql: str) -> str:
    return sql.replace("\r\n", "\n").replace("\r", "\n").strip()


def _swap_query(schema: SchemaBackend, existing: str, replacement: QuerySpec) -> None:
    """Replace a query whose type changes: build the new one first so a failure keeps the original."""
    original = schema.read_query(existing)
    temporary = f"~pak_tmp_{uuid.uuid4().hex[:8]}"
    try:
        schema.create_query(replacement.model_copy(update={"name": temporary}))
    except BaseException:
        with contextlib.suppress(Exception):
            if find_query(schema, temporary) is not None:
                schema.drop_query(temporary)
        raise
    try:
        schema.drop_query(existing)
        schema.rename_query(temporary, replacement.name)
    except BaseException:
        with contextlib.suppress(Exception):
            if find_query(schema, existing) is None:
                schema.create_query(original)
        with contextlib.suppress(Exception):
            schema.drop_query(temporary)
        raise


def rename_query(schema: SchemaBackend, old: str, new: str) -> str:
    """Rename a saved query; returns the new name."""
    actual = find_query(schema, old)
    if actual is None:
        raise ObjectNotFoundError(f"query {old!r} does not exist", kind=ObjectKind.QUERY, name=old)
    check_name(new, what="query name")
    ensure_name_free(schema, new, ignore=actual)
    warn_name(new, what="query name", stacklevel=4)
    schema.rename_query(actual, new)
    return new


def drop_query(schema: SchemaBackend, name: str) -> None:
    """Delete a saved query."""
    actual = find_query(schema, name)
    if actual is None:
        raise ObjectNotFoundError(
            f"query {name!r} does not exist", kind=ObjectKind.QUERY, name=name
        )
    schema.drop_query(actual)
