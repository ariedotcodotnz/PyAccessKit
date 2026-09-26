# pyright: basic
"""DAO implementation of :class:`~pyaccesskit._backends.protocols.SchemaBackend`.

Works identically over every transport: the DAO ``Database`` is obtained through a callable, so the same
code runs against in-process DAO, Access-hosted DAO and ``CurrentDb()``.

Behaviours verified in ADR 0001:

* collections are refreshed before being read (``QueryDef.Type`` is 0 until ``QueryDefs.Refresh``);
* Decimal columns are created with ADO DDL (DAO cannot set precision/scale and silently creates a BigInt),
  after which the original column order is restored with ``OrdinalPosition``;
* relationships among ``MSys*`` system tables (created by Access for the navigation pane) are ignored.
"""

from __future__ import annotations

import contextlib
from collections.abc import Callable, Mapping
from datetime import datetime
from pathlib import Path
from typing import Any

import pywintypes

from pyaccesskit._backends.dao import typemap as tm
from pyaccesskit._backends.protocols import (
    DatabaseInfo,
    FetchResult,
    ParameterInfo,
    PropertyTarget,
    QueryInfo,
    TableInfo,
)
from pyaccesskit._com.gateway import Com, get
from pyaccesskit._com.variants import normalize, to_variant
from pyaccesskit.enums import DataType, JoinType, ObjectKind, PropertyType, QueryKind, Transport
from pyaccesskit.errors import MissingParameterError, ObjectNotFoundError, SpecError
from pyaccesskit.schema import (
    AutoNumberColumn,
    ColumnBase,
    ColumnSpec,
    DecimalColumn,
    IndexField,
    IndexSpec,
    PassThroughOptions,
    PropertyValue,
    QuerySpec,
    RelationshipSpec,
    TableSpec,
    quote_identifier,
)
from pyaccesskit.schema.queries import DAO_QUERY_KINDS

__all__ = ["DaoSchemaBackend"]

PROPERTY_NOT_FOUND = 3270
DB_FAIL_ON_ERROR = 128
DB_OPEN_SNAPSHOT = 4
DB_SYSTEM_OBJECT = 0x80000000
DB_HIDDEN_OBJECT = 1
DB_ATTACHED_TABLE = 0x40000000
DB_ATTACHED_ODBC = 0x20000000
REL_UNIQUE, REL_DONT_ENFORCE, REL_INHERITED = 1, 2, 4
REL_UPDATE_CASCADE, REL_DELETE_CASCADE = 256, 4096
REL_LEFT, REL_RIGHT = 16777216, 33554432
PLACEHOLDER = "__pak_placeholder__"
FETCH_BATCH = 500
_CONTAINERS = {
    ObjectKind.FORM: "Forms",
    ObjectKind.REPORT: "Reports",
    ObjectKind.MACRO: "Scripts",
    ObjectKind.MODULE: "Modules",
}
_PARAMETER_TYPES = {
    tm.DB_BOOLEAN: DataType.YES_NO,
    tm.DB_BYTE: DataType.NUMBER,
    tm.DB_INTEGER: DataType.NUMBER,
    tm.DB_LONG: DataType.NUMBER,
    tm.DB_SINGLE: DataType.NUMBER,
    tm.DB_DOUBLE: DataType.NUMBER,
    tm.DB_CURRENCY: DataType.CURRENCY,
    tm.DB_DATE: DataType.DATE_TIME,
    tm.DB_TEXT: DataType.TEXT,
    tm.DB_MEMO: DataType.LONG_TEXT,
    tm.DB_GUID: DataType.NUMBER,
    tm.DB_DECIMAL: DataType.DECIMAL,
}


def _error_number(exc: pywintypes.com_error) -> int | None:
    excepinfo = exc.args[2] if len(exc.args) > 2 else None
    scode = excepinfo[5] if excepinfo else None
    return None if scode is None else scode & 0xFFFF


def _prop(obj: Any, name: str) -> Any:
    """Value of a DAO property, or ``None`` if the object does not have it."""
    try:
        return obj.Properties(name).Value
    except pywintypes.com_error as exc:
        if _error_number(exc) == PROPERTY_NOT_FOUND:
            return None
        raise


def _set_prop(obj: Any, name: str, dao_type: int, value: Any) -> None:
    """Set a DAO property, creating it (Access-style) when it does not exist yet."""
    value = to_variant(value)
    try:
        obj.Properties(name).Value = value
    except pywintypes.com_error as exc:
        if _error_number(exc) != PROPERTY_NOT_FOUND:
            raise
        obj.Properties.Append(obj.CreateProperty(name, dao_type, value))


def _param_value(value: Any) -> Any:
    return to_variant(value)


def _item(collection: Any, key: Any) -> Any:
    """``collection.Item(key)`` through an exact ``IDispatch.Invoke``.

    pywin32's dynamic dispatch resolves ``Item`` as a property on some DAO collections (e.g. an index's
    fields) and returns the wrong thing; invoking the DISPID directly always works.
    """
    return get(collection, "Item", key)


def _index_fields(idx: Any) -> tuple[IndexField, ...]:
    """Fields of a DAO index.

    For an *appended* index, late-bound ``Index.Fields`` returns a string such as ``"+CustomerID;-OrderDate"``
    (``+`` ascending, ``-`` descending) rather than a collection; unappended indexes return the collection.
    """
    fields = idx.Fields
    if isinstance(fields, str):
        result: list[IndexField] = []
        for part in fields.split(";"):
            item = part.strip()
            if not item:
                continue
            descending = item.startswith("-")
            name = item[1:] if item[0] in "+-" else item
            result.append(IndexField(name=_strip_brackets(name), descending=descending))
        return tuple(result)
    return tuple(
        IndexField(
            name=str(_item(fields, i).Name),
            descending=bool(int(_item(fields, i).Attributes) & tm.DB_DESCENDING),
        )
        for i in range(fields.Count)
    )


def _strip_brackets(name: str) -> str:
    return name[1:-1] if name.startswith("[") and name.endswith("]") else name


class DaoSchemaBackend:
    """Schema operations over a DAO ``Database``."""

    def __init__(
        self,
        *,
        com: Com,
        database: Callable[[], Any],
        path: Path,
        transport: Callable[[], Transport],
        run_ddl: Callable[[str], None],
    ) -> None:
        self._com_ref = com
        self._database = database
        self._path = path
        self._transport = transport
        self._run_ddl = run_ddl

    @property
    def _com(self) -> Com:
        return self._com_ref

    def _db(self) -> Any:
        return self._database()

    def database_info(self) -> DatabaseInfo:
        with self._com.op("read database information", path=self._path):
            return DatabaseInfo(self._path, str(self._db().Version), self._transport())

    # ------------------------------------------------------------------------------------ tables
    def list_tables(self) -> list[TableInfo]:
        with self._com.op("list tables"):
            tabledefs = self._db().TableDefs
            tabledefs.Refresh()
            tables: list[TableInfo] = []
            for index in range(tabledefs.Count):
                tdf = _item(tabledefs, index)
                name = str(tdf.Name)
                attributes = int(tdf.Attributes) & 0xFFFFFFFF
                connect = str(tdf.Connect or "")
                linked = bool(connect) or bool(attributes & (DB_ATTACHED_TABLE | DB_ATTACHED_ODBC))
                tables.append(
                    TableInfo(
                        name=name,
                        is_linked=linked,
                        is_system=bool(attributes & DB_SYSTEM_OBJECT)
                        or name.casefold().startswith("msys"),
                        is_hidden=bool(attributes & DB_HIDDEN_OBJECT) or name.startswith("~"),
                        connect=connect or None,
                        source_table=str(tdf.SourceTableName or "") or None if linked else None,
                    )
                )
            return tables

    @staticmethod
    def _read_field(fld: Any) -> tm.FieldRead:
        properties: dict[str, Any] = {}
        names = [
            "Description",
            "Caption",
            "Format",
            "InputMask",
            "DecimalPlaces",
            "UnicodeCompression",
            "TextFormat",
        ]
        dao_type = int(fld.Type)
        if dao_type == tm.DB_DECIMAL:
            names += ["Precision", "Scale"]
        for name in names:
            value = _prop(fld, name)
            if value is not None:
                properties[name.casefold()] = value
        expression = ""
        with contextlib.suppress(pywintypes.com_error, AttributeError):
            expression = str(fld.Expression or "")
        append_only = False
        with contextlib.suppress(pywintypes.com_error, AttributeError):
            append_only = bool(fld.AppendOnly)
        return tm.FieldRead(
            name=str(fld.Name),
            dao_type=dao_type,
            size=int(fld.Size),
            attributes=int(fld.Attributes),
            required=bool(fld.Required),
            allow_zero_length=bool(fld.AllowZeroLength),
            default=str(fld.DefaultValue or ""),
            validation_rule=str(fld.ValidationRule or ""),
            validation_text=str(fld.ValidationText or ""),
            append_only=append_only,
            expression=expression,
            properties=properties,
        )

    def read_table(self, name: str) -> TableSpec:
        with self._com.op(f"read table {name!r}", kind=ObjectKind.TABLE, name=name):
            db = self._db()
            db.TableDefs.Refresh()
            tdf = db.TableDefs(name)
            fields = tdf.Fields
            fields.Refresh()
            reads = []
            for index in range(fields.Count):
                fld = _item(fields, index)
                reads.append((int(fld.OrdinalPosition), index, self._read_field(fld)))
            columns = [
                tm.column_from_field(read)
                for _, _, read in sorted(reads, key=lambda item: item[:2])
            ]
            indexes: list[IndexSpec] = []
            dao_indexes = tdf.Indexes
            dao_indexes.Refresh()
            for index in range(dao_indexes.Count):
                idx = _item(dao_indexes, index)
                if bool(idx.Foreign):
                    continue  # hidden index owned by a relationship
                indexes.append(
                    IndexSpec(
                        name=str(idx.Name),
                        fields=_index_fields(idx),
                        primary=bool(idx.Primary),
                        unique=bool(idx.Unique),
                        required=bool(idx.Required),
                        ignore_nulls=bool(idx.IgnoreNulls),
                    )
                )
            rule = str(tdf.ValidationRule or "") or None
            spec = TableSpec(
                name=str(tdf.Name),
                columns=tuple(columns),  # type: ignore[arg-type]
                indexes=tuple(indexes),
                description=_prop(tdf, "Description") or None,
                validation_rule=rule,
                validation_text=(str(tdf.ValidationText or "") or None) if rule else None,
            )
            return spec.normalized()

    def count_indexes(self, table: str) -> int:
        with self._com.op(f"count indexes of {table!r}", kind=ObjectKind.TABLE, name=table):
            indexes = self._db().TableDefs(table).Indexes
            indexes.Refresh()
            return int(indexes.Count)

    def _make_field(self, tdf: Any, column: ColumnBase, plan: tm.FieldPlan) -> Any:
        if plan.size is not None:
            fld = tdf.CreateField(column.name, plan.dao_type, plan.size)
        else:
            fld = tdf.CreateField(column.name, plan.dao_type)
        fld.Attributes = plan.attributes
        self._configure_field(fld, column, plan)
        if plan.append_only:
            fld.AppendOnly = True
        return fld

    @staticmethod
    def _configure_field(fld: Any, column: ColumnBase, plan: tm.FieldPlan) -> None:
        if not isinstance(column, AutoNumberColumn):
            fld.Required = column.required
        if plan.allow_zero_length is not None:
            fld.AllowZeroLength = plan.allow_zero_length
        if plan.default is not None:
            fld.DefaultValue = plan.default
        if column.validation_rule:
            fld.ValidationRule = column.validation_rule
            if column.validation_text:
                fld.ValidationText = column.validation_text

    @staticmethod
    def _append_index(tdf: Any, index: IndexSpec) -> None:
        idx = tdf.CreateIndex(index.name)
        idx.Primary = index.primary
        idx.Unique = index.unique
        idx.Required = index.required
        idx.IgnoreNulls = index.ignore_nulls
        for field in index.fields:
            idx_field = idx.CreateField(field.name)
            if field.descending:
                idx_field.Attributes = tm.DB_DESCENDING
            idx.Fields.Append(idx_field)
        tdf.Indexes.Append(idx)

    @staticmethod
    def _set_table_validation(tdf: Any, spec: TableSpec) -> None:
        tdf.ValidationRule = spec.validation_rule
        if spec.validation_text:
            tdf.ValidationText = spec.validation_text

    def _decimal_ddl(self, table: str, column: DecimalColumn) -> None:
        self._run_ddl(
            f"ALTER TABLE {quote_identifier(table)} ADD COLUMN {quote_identifier(column.name)} "
            f"DECIMAL({column.precision},{column.scale})"
        )

    def create_table(self, spec: TableSpec) -> None:
        plans: list[tuple[ColumnBase, tm.FieldPlan]] = [
            (column, tm.plan_field(column)) for column in spec.columns
        ]
        dao_columns = [(column, plan) for column, plan in plans if not plan.ado_decimal]
        decimal_columns = [(column, plan) for column, plan in plans if plan.ado_decimal]
        with self._com.op(
            f"create table {spec.name!r}", kind=ObjectKind.TABLE, name=spec.name, path=self._path
        ):
            db = self._db()
            tdf = db.CreateTableDef(spec.name)
            if not dao_columns:
                tdf.Fields.Append(tdf.CreateField(PLACEHOLDER, tm.DB_LONG))
            for column, plan in dao_columns:
                tdf.Fields.Append(self._make_field(tdf, column, plan))
            if spec.validation_rule and not decimal_columns:
                # With Decimal columns the rule waits until the ADO DDL has added them (_finish_table).
                self._set_table_validation(tdf, spec)
            db.TableDefs.Append(tdf)
            del tdf, db
        try:
            self._finish_table(spec, plans, decimal_columns, placeholder=not dao_columns)
        except BaseException:
            with contextlib.suppress(Exception), self._com.op(f"roll back table {spec.name!r}"):
                self._db().TableDefs.Delete(spec.name)
            raise

    def _finish_table(
        self,
        spec: TableSpec,
        plans: list[tuple[ColumnBase, tm.FieldPlan]],
        decimal_columns: list[tuple[ColumnBase, tm.FieldPlan]],
        *,
        placeholder: bool,
    ) -> None:
        for column, _plan in decimal_columns:
            assert isinstance(column, DecimalColumn)
            self._decimal_ddl(spec.name, column)
        with self._com.op(f"configure table {spec.name!r}", kind=ObjectKind.TABLE, name=spec.name):
            db = self._db()
            db.TableDefs.Refresh()
            tdf = db.TableDefs(spec.name)
            tdf.Fields.Refresh()
            for column, plan in decimal_columns:
                self._configure_field(tdf.Fields(column.name), column, plan)
            if placeholder:
                tdf.Fields.Delete(PLACEHOLDER)
            if decimal_columns:
                for position, column in enumerate(spec.columns):
                    tdf.Fields(column.name).OrdinalPosition = position
            for index in spec.indexes:
                self._append_index(tdf, index)
            if spec.validation_rule and decimal_columns:
                self._set_table_validation(tdf, spec)
            if spec.description is not None:
                _set_prop(
                    tdf,
                    "Description",
                    tm.property_type_for("Description", spec.description),
                    spec.description,
                )
            for name, value in spec.properties.items():
                _set_prop(tdf, name, tm.property_type_for(name, value), value)
            for column, plan in plans:
                if plan.post_properties:
                    fld = tdf.Fields(column.name)
                    for name, dao_type, value in plan.post_properties:
                        _set_prop(fld, name, dao_type, value)

    def drop_table(self, name: str) -> None:
        with self._com.op(f"delete table {name!r}", kind=ObjectKind.TABLE, name=name):
            self._db().TableDefs.Delete(name)

    def rename_table(self, old: str, new: str) -> None:
        with self._com.op(f"rename table {old!r} to {new!r}", kind=ObjectKind.TABLE, name=old):
            self._db().TableDefs(old).Name = new

    def add_column(self, table: str, column: ColumnSpec) -> None:
        plan = tm.plan_field(column)
        if plan.ado_decimal:
            assert isinstance(column, DecimalColumn)
            self._decimal_ddl(table, column)
        with self._com.op(
            f"add column {column.name!r} to {table!r}", kind=ObjectKind.FIELD, name=column.name
        ):
            db = self._db()
            db.TableDefs.Refresh()
            tdf = db.TableDefs(table)
            if plan.ado_decimal:
                tdf.Fields.Refresh()
                self._configure_field(tdf.Fields(column.name), column, plan)
            else:
                tdf.Fields.Append(self._make_field(tdf, column, plan))
            if plan.post_properties:
                fld = tdf.Fields(column.name)
                for name, dao_type, value in plan.post_properties:
                    _set_prop(fld, name, dao_type, value)

    def drop_column(self, table: str, column: str) -> None:
        with self._com.op(
            f"delete column {column!r} of {table!r}", kind=ObjectKind.FIELD, name=column
        ):
            self._db().TableDefs(table).Fields.Delete(column)

    def rename_column(self, table: str, old: str, new: str) -> None:
        with self._com.op(f"rename column {old!r} of {table!r}", kind=ObjectKind.FIELD, name=old):
            self._db().TableDefs(table).Fields(old).Name = new

    def create_index(self, table: str, index: IndexSpec) -> None:
        with self._com.op(
            f"create index {index.name!r} on {table!r}", kind=ObjectKind.INDEX, name=index.name
        ):
            self._append_index(self._db().TableDefs(table), index)

    def drop_index(self, table: str, name: str) -> None:
        with self._com.op(f"delete index {name!r} of {table!r}", kind=ObjectKind.INDEX, name=name):
            self._db().TableDefs(table).Indexes.Delete(name)

    # ----------------------------------------------------------------------------- relationships
    def list_relationships(self) -> list[RelationshipSpec]:
        with self._com.op("list relationships"):
            relations = self._db().Relations
            relations.Refresh()
            result: list[RelationshipSpec] = []
            for index in range(relations.Count):
                rel = _item(relations, index)
                attributes = int(rel.Attributes)
                primary, foreign = str(rel.Table), str(rel.ForeignTable)
                if (
                    attributes & REL_INHERITED
                    or primary.casefold().startswith("msys")
                    or foreign.casefold().startswith("msys")
                ):
                    continue
                fields = rel.Fields
                pairs = [
                    (str(_item(fields, i).Name), str(_item(fields, i).ForeignName))
                    for i in range(fields.Count)
                ]
                join = (
                    JoinType.LEFT
                    if attributes & REL_LEFT
                    else JoinType.RIGHT
                    if attributes & REL_RIGHT
                    else JoinType.INNER
                )
                result.append(
                    RelationshipSpec(
                        name=str(rel.Name),
                        primary_table=primary,
                        primary_columns=tuple(p for p, _ in pairs),
                        foreign_table=foreign,
                        foreign_columns=tuple(f for _, f in pairs),
                        enforce_integrity=not attributes & REL_DONT_ENFORCE,
                        cascade_update=bool(attributes & REL_UPDATE_CASCADE),
                        cascade_delete=bool(attributes & REL_DELETE_CASCADE),
                        one_to_one=bool(attributes & REL_UNIQUE),
                        join=join,
                    )
                )
            return result

    def create_relationship(self, spec: RelationshipSpec) -> None:
        attributes = 0
        if spec.one_to_one:
            attributes |= REL_UNIQUE
        if not spec.enforce_integrity:
            attributes |= REL_DONT_ENFORCE
        if spec.cascade_update:
            attributes |= REL_UPDATE_CASCADE
        if spec.cascade_delete:
            attributes |= REL_DELETE_CASCADE
        if spec.join is JoinType.LEFT:
            attributes |= REL_LEFT
        elif spec.join is JoinType.RIGHT:
            attributes |= REL_RIGHT
        name = spec.effective_name
        with self._com.op(f"create relationship {name!r}", kind=ObjectKind.RELATIONSHIP, name=name):
            db = self._db()
            rel = db.CreateRelation(name, spec.primary_table, spec.foreign_table, attributes)
            for primary, foreign in zip(spec.primary_columns, spec.foreign_columns, strict=True):
                fld = rel.CreateField(primary)
                fld.ForeignName = foreign
                rel.Fields.Append(fld)
            db.Relations.Append(rel)

    def drop_relationship(self, name: str) -> None:
        with self._com.op(f"delete relationship {name!r}", kind=ObjectKind.RELATIONSHIP, name=name):
            self._db().Relations.Delete(name)

    # ----------------------------------------------------------------------------------- queries
    def list_queries(self) -> list[QueryInfo]:
        with self._com.op("list queries"):
            querydefs = self._db().QueryDefs
            querydefs.Refresh()
            result: list[QueryInfo] = []
            for index in range(querydefs.Count):
                qd = _item(querydefs, index)
                name = str(qd.Name)
                kind = DAO_QUERY_KINDS.get(int(qd.Type), QueryKind.UNKNOWN)
                result.append(QueryInfo(name, kind, name.startswith("~")))
            return result

    def read_query(self, name: str) -> QuerySpec:
        with self._com.op(f"read query {name!r}", kind=ObjectKind.QUERY, name=name):
            qd = self._db().QueryDefs(name)
            connect = str(qd.Connect or "")
            pass_through = None
            if connect.upper().startswith("ODBC;"):
                pass_through = PassThroughOptions(
                    connect=connect,
                    returns_records=bool(qd.ReturnsRecords),
                    timeout=int(qd.ODBCTimeout),
                )
            return QuerySpec(
                name=str(qd.Name),
                sql=str(qd.SQL),
                description=_prop(qd, "Description") or None,
                pass_through=pass_through,
            )

    def query_parameters(self, name: str) -> list[ParameterInfo]:
        with self._com.op(f"read parameters of query {name!r}", kind=ObjectKind.QUERY, name=name):
            params = self._db().QueryDefs(name).Parameters
            return [
                ParameterInfo(
                    _strip_brackets(str(_item(params, i).Name)),
                    _PARAMETER_TYPES.get(int(_item(params, i).Type), DataType.UNKNOWN),
                )
                for i in range(params.Count)
            ]

    def create_query(self, spec: QuerySpec) -> None:
        with self._com.op(
            f"create query {spec.name!r}", kind=ObjectKind.QUERY, name=spec.name, sql=spec.sql
        ):
            db = self._db()
            if spec.pass_through is not None:
                qd = db.CreateQueryDef(spec.name)
                qd.Connect = spec.pass_through.connect
                qd.SQL = spec.sql
                qd.ReturnsRecords = spec.pass_through.returns_records
                qd.ODBCTimeout = spec.pass_through.timeout
            else:
                qd = db.CreateQueryDef(spec.name, spec.sql)
            if spec.description is not None:
                _set_prop(
                    qd,
                    "Description",
                    tm.property_type_for("Description", spec.description),
                    spec.description,
                )

    def set_query_sql(self, name: str, sql: str) -> None:
        with self._com.op(f"update query {name!r}", kind=ObjectKind.QUERY, name=name, sql=sql):
            self._db().QueryDefs(name).SQL = sql

    def rename_query(self, old: str, new: str) -> None:
        with self._com.op(f"rename query {old!r} to {new!r}", kind=ObjectKind.QUERY, name=old):
            self._db().QueryDefs(old).Name = new

    def drop_query(self, name: str) -> None:
        with self._com.op(f"delete query {name!r}", kind=ObjectKind.QUERY, name=name):
            self._db().QueryDefs.Delete(name)

    # -------------------------------------------------------------------------------- properties
    def _target(self, target: PropertyTarget) -> Any:
        db = self._db()
        if target.kind == "database":
            return db
        if target.kind == "table":
            return db.TableDefs(target.name)
        if target.kind == "field":
            return db.TableDefs(target.name).Fields(target.field)
        return db.QueryDefs(target.name)

    def get_property(self, target: PropertyTarget, name: str) -> PropertyValue:
        with self._com.op(
            f"read property {name!r} of {target.describe()}", kind=ObjectKind.PROPERTY, name=name
        ):
            obj = self._target(target)
            try:
                return normalize(obj.Properties(name).Value)
            except pywintypes.com_error as exc:
                if _error_number(exc) == PROPERTY_NOT_FOUND:
                    raise ObjectNotFoundError(
                        f"{target.describe()} has no property {name!r}",
                        kind=ObjectKind.PROPERTY,
                        name=name,
                    ) from None
                raise

    def set_property(
        self,
        target: PropertyTarget,
        name: str,
        value: PropertyValue,
        type: PropertyType | None = None,
    ) -> None:
        with self._com.op(
            f"set property {name!r} of {target.describe()}", kind=ObjectKind.PROPERTY, name=name
        ):
            _set_prop(self._target(target), name, tm.property_type_for(name, value, type), value)

    def delete_property(self, target: PropertyTarget, name: str) -> None:
        with self._com.op(
            f"delete property {name!r} of {target.describe()}", kind=ObjectKind.PROPERTY, name=name
        ):
            obj = self._target(target)
            try:
                obj.Properties.Delete(name)
            except pywintypes.com_error as exc:
                if _error_number(exc) in (PROPERTY_NOT_FOUND, 3265):
                    raise ObjectNotFoundError(
                        f"{target.describe()} has no property {name!r}",
                        kind=ObjectKind.PROPERTY,
                        name=name,
                    ) from None
                raise

    def list_properties(self, target: PropertyTarget) -> dict[str, PropertyValue]:
        with self._com.op(f"list properties of {target.describe()}"):
            properties = self._target(target).Properties
            result: dict[str, PropertyValue] = {}
            for index in range(properties.Count):
                prop = _item(properties, index)
                with contextlib.suppress(pywintypes.com_error):
                    value = normalize(prop.Value)
                    if value is None or isinstance(value, (str, bool, int, float, datetime)):
                        result[str(prop.Name)] = value
            return result

    # -------------------------------------------------------------------------------------- data
    @staticmethod
    def _bind(qd: Any, params: Mapping[str, Any] | None, sql: str) -> None:
        dao_params = qd.Parameters
        available = {
            _strip_brackets(str(_item(dao_params, i).Name)).casefold(): _item(dao_params, i)
            for i in range(dao_params.Count)
        }
        given = {key.strip("[]").casefold(): value for key, value in (params or {}).items()}
        unknown = sorted(set(given) - set(available))
        if unknown:
            known = ", ".join(sorted(available)) or "none"
            raise SpecError(
                f"unknown parameter(s) {unknown}; the statement's parameters are: {known}. (A [name] that matches "
                "a column is a column reference, not a parameter.)"
            )
        missing = sorted(set(available) - set(given))
        if missing:
            raise MissingParameterError(f"no value supplied for parameter(s) {missing}", sql=sql)
        for key, value in given.items():
            available[key].Value = _param_value(value)

    def _query(self, sql: str, name: str | None) -> Any:
        db = self._db()
        return db.QueryDefs(name) if name is not None else db.CreateQueryDef("", sql)

    def _run(self, sql: str, name: str | None, params: Mapping[str, Any] | None) -> int:
        label = f"run query {name!r}" if name else "run SQL statement"
        with self._com.op(label, kind=ObjectKind.QUERY if name else None, name=name, sql=sql):
            qd = self._query(sql, name)
            self._bind(qd, params, sql)
            qd.Execute(DB_FAIL_ON_ERROR)
            return int(qd.RecordsAffected)

    def _fetch(
        self, sql: str, name: str | None, params: Mapping[str, Any] | None, limit: int | None
    ) -> FetchResult:
        label = f"read rows of query {name!r}" if name else "read rows"
        with self._com.op(label, kind=ObjectKind.QUERY if name else None, name=name, sql=sql):
            qd = self._query(sql, name)
            self._bind(qd, params, sql)
            rs = qd.OpenRecordset(DB_OPEN_SNAPSHOT)
            try:
                fields = rs.Fields
                columns = tuple(str(_item(fields, i).Name) for i in range(fields.Count))
                rows: list[tuple[Any, ...]] = []
                while not rs.EOF and (limit is None or len(rows) < limit):
                    batch = FETCH_BATCH if limit is None else min(FETCH_BATCH, limit - len(rows))
                    data = rs.GetRows(batch)
                    if not data:
                        break
                    rows.extend(normalize(tuple(row)) for row in zip(*data, strict=True))
                return FetchResult(columns, rows)
            finally:
                with contextlib.suppress(pywintypes.com_error):
                    rs.Close()

    def execute(self, sql: str, params: Mapping[str, Any] | None = None) -> int:
        return self._run(sql, None, params)

    def fetch(
        self, sql: str, params: Mapping[str, Any] | None = None, *, limit: int | None = None
    ) -> FetchResult:
        return self._fetch(sql, None, params, limit)

    def execute_saved(self, name: str, params: Mapping[str, Any] | None = None) -> int:
        return self._run("", name, params)

    def fetch_saved(
        self, name: str, params: Mapping[str, Any] | None = None, *, limit: int | None = None
    ) -> FetchResult:
        return self._fetch("", name, params, limit)

    # --------------------------------------------------------------------------------- documents
    def list_documents(self, kind: ObjectKind) -> list[str]:
        container = _CONTAINERS.get(kind)
        if container is None:
            raise SpecError(f"{kind.value} objects are not stored in a DAO container")
        with self._com.op(f"list {kind.value}s"):
            documents = self._db().Containers(container).Documents
            documents.Refresh()
            names = [str(_item(documents, i).Name) for i in range(documents.Count)]
            return [name for name in names if not name.startswith("~")]
