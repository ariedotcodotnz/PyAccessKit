"""An in-memory implementation of both backend protocols.

It models the semantics that matter to callers — case-insensitive names, the shared table/query
namespace, index and relationship rules, DAO property behaviour — closely enough that the same contract
test-suite passes against it and against real Access. It cannot execute SQL.
"""

from __future__ import annotations

import re
from collections.abc import Mapping
from pathlib import Path
from typing import Any

from pyaccesskit._backends.protocols import (
    ControlInfo,
    DatabaseInfo,
    FetchResult,
    ParameterInfo,
    PropertyTarget,
    QueryInfo,
    TableInfo,
)
from pyaccesskit.enums import ControlKind, DataType, ObjectKind, PropertyType, Transport
from pyaccesskit.errors import (
    CapabilityError,
    ObjectExistsError,
    ObjectNotFoundError,
    RelationshipError,
    SchemaError,
)
from pyaccesskit.forms.layout import ResolvedForm
from pyaccesskit.schema import (
    ColumnBase,
    ColumnSpec,
    IndexField,
    IndexSpec,
    PropertyValue,
    QuerySpec,
    RelationshipSpec,
    TableSpec,
    detect_query_kind,
)
from pyaccesskit.schema.compat import relationship_compatible
from pyaccesskit.schema.expressions import parse_literal, render_literal
from pyaccesskit.schema.tables import MAX_COLUMNS, MAX_INDEXES

__all__ = ["FakeBackend"]

_DESIGN_KINDS = (ObjectKind.FORM, ObjectKind.REPORT, ObjectKind.MACRO, ObjectKind.MODULE)
_TABLE_PROPS = {
    "description": "description",
    "validationrule": "validation_rule",
    "validationtext": "validation_text",
}
_FIELD_PROPS = {
    "description": "description",
    "caption": "caption",
    "format": "format",
    "inputmask": "input_mask",
    "decimalplaces": "decimal_places",
    "validationrule": "validation_rule",
    "validationtext": "validation_text",
    "required": "required",
    "defaultvalue": "default",
    "allowzerolength": "allow_zero_length",
    "unicodecompression": "unicode_compression",
}
_QUERY_PROPS = {"description": "description"}
_DISPLAY_NAMES = {
    name.casefold(): name
    for name in (
        "Description",
        "Caption",
        "Format",
        "InputMask",
        "DecimalPlaces",
        "ValidationRule",
        "ValidationText",
        "Required",
        "DefaultValue",
        "AllowZeroLength",
        "UnicodeCompression",
    )
}
_PARAMETERS = re.compile(r"^\s*PARAMETERS\s+(?P<body>.*?);", re.IGNORECASE | re.DOTALL)
_PARAM_ITEM = re.compile(
    r"\s*(?:\[(?P<bracketed>[^\]]+)\]|(?P<plain>[^\s,\[\]]+))\s+(?P<type>[A-Za-z]+)(?:\s*\([^)]*\))?\s*(?:,|$)"
)
_PARAM_TYPES = {
    "long": DataType.NUMBER,
    "integer": DataType.NUMBER,
    "short": DataType.NUMBER,
    "byte": DataType.NUMBER,
    "single": DataType.NUMBER,
    "double": DataType.NUMBER,
    "currency": DataType.CURRENCY,
    "datetime": DataType.DATE_TIME,
    "bit": DataType.YES_NO,
    "yesno": DataType.YES_NO,
    "text": DataType.TEXT,
}


def _key(name: str) -> str:
    return name.casefold()


class FakeBackend:
    """In-memory :class:`SchemaBackend` + :class:`DesignBackend`."""

    def __init__(self, path: Path | str = "memory.accdb") -> None:
        self.path = Path(path)
        self._tables: dict[str, TableSpec] = {}
        self._relationships: dict[str, RelationshipSpec] = {}
        self._queries: dict[str, QuerySpec] = {}
        self._custom: dict[tuple[str, str, str], dict[str, tuple[str, PropertyValue]]] = {}
        self._objects: dict[ObjectKind, dict[str, tuple[str, bytes]]] = {
            kind: {} for kind in (*_DESIGN_KINDS, ObjectKind.QUERY)
        }
        self._forms: dict[str, ResolvedForm] = {}

    # ------------------------------------------------------------------------------------- basics
    def database_info(self) -> DatabaseInfo:
        return DatabaseInfo(self.path, "12.0", Transport.MEMORY)

    def _table(self, name: str) -> TableSpec:
        try:
            return self._tables[_key(name)]
        except KeyError:
            raise ObjectNotFoundError(
                f"table {name!r} does not exist", kind=ObjectKind.TABLE, name=name
            ) from None

    def _query(self, name: str) -> QuerySpec:
        try:
            return self._queries[_key(name)]
        except KeyError:
            raise ObjectNotFoundError(
                f"query {name!r} does not exist", kind=ObjectKind.QUERY, name=name
            ) from None

    def _ensure_free(self, name: str) -> None:
        key = _key(name)
        if key in self._tables:
            raise ObjectExistsError(
                f"a table named {self._tables[key].name!r} already exists",
                kind=ObjectKind.TABLE,
                name=name,
            )
        if key in self._queries:
            raise ObjectExistsError(
                f"a query named {self._queries[key].name!r} already exists",
                kind=ObjectKind.QUERY,
                name=name,
            )

    def _store(self, spec: TableSpec) -> None:
        self._tables[_key(spec.name)] = spec

    @staticmethod
    def _column(spec: TableSpec, name: str) -> ColumnBase:
        try:
            return spec.column(name)
        except KeyError:
            raise ObjectNotFoundError(
                f"table {spec.name!r} has no column {name!r}", kind=ObjectKind.FIELD, name=name
            ) from None

    # ------------------------------------------------------------------------------------ tables
    def list_tables(self) -> list[TableInfo]:
        return [TableInfo(spec.name) for spec in self._tables.values()]

    def read_table(self, name: str) -> TableSpec:
        return self._table(name)

    def count_indexes(self, table: str) -> int:
        spec = self._table(table)
        hidden = sum(
            1
            for rel in self._relationships.values()
            if rel.enforce_integrity and _key(rel.foreign_table) == _key(spec.name)
        )
        return len(spec.indexes) + hidden

    def create_table(self, spec: TableSpec) -> None:
        self._ensure_free(spec.name)
        self._store(spec.normalized())

    def _relationships_of(self, table: str) -> list[RelationshipSpec]:
        key = _key(table)
        return [
            rel
            for rel in self._relationships.values()
            if _key(rel.primary_table) == key or _key(rel.foreign_table) == key
        ]

    def drop_table(self, name: str) -> None:
        spec = self._table(name)
        related = self._relationships_of(spec.name)
        if related:
            names = ", ".join(repr(rel.effective_name) for rel in related)
            raise SchemaError(
                f"table {spec.name!r} is part of relationship(s) {names}; delete them first"
            )
        del self._tables[_key(spec.name)]
        for target in [
            t for t in self._custom if t[1] == _key(spec.name) and t[0] in ("table", "field")
        ]:
            del self._custom[target]

    def rename_table(self, old: str, new: str) -> None:
        spec = self._table(old)
        if _key(old) != _key(new):
            self._ensure_free(new)
        del self._tables[_key(spec.name)]
        self._store(spec.model_copy(update={"name": new}))
        for key, rel in list(self._relationships.items()):
            update: dict[str, Any] = {}
            if _key(rel.primary_table) == _key(old):
                update["primary_table"] = new
            if _key(rel.foreign_table) == _key(old):
                update["foreign_table"] = new
            if update:
                self._relationships[key] = rel.model_copy(update=update)
        for target in [t for t in self._custom if t[1] == _key(old) and t[0] in ("table", "field")]:
            self._custom[(target[0], _key(new), target[2])] = self._custom.pop(target)

    def add_column(self, table: str, column: ColumnSpec) -> None:
        spec = self._table(table)
        if any(_key(c.name) == _key(column.name) for c in spec.columns):
            raise ObjectExistsError(
                f"table {spec.name!r} already has a column {column.name!r}",
                kind=ObjectKind.FIELD,
                name=column.name,
            )
        if len(spec.columns) >= MAX_COLUMNS:
            raise SchemaError(f"table {spec.name!r} already has {MAX_COLUMNS} columns")
        if column.data_type is DataType.AUTONUMBER and any(
            c.data_type is DataType.AUTONUMBER for c in spec.columns
        ):
            raise SchemaError(f"table {spec.name!r} already has an AutoNumber column")
        self._store(spec.model_copy(update={"columns": (*spec.columns, column.normalized())}))

    def drop_column(self, table: str, column: str) -> None:
        spec = self._table(table)
        target = self._column(spec, column)
        indexes = [
            i.name for i in spec.indexes if any(_key(f) == _key(target.name) for f in i.field_names)
        ]
        if indexes:
            raise SchemaError(
                f"column {target.name!r} is part of index(es) {indexes}; drop them first"
            )
        for rel in self._relationships_of(spec.name):
            columns = (
                rel.primary_columns
                if _key(rel.primary_table) == _key(spec.name)
                else rel.foreign_columns
            )
            if any(_key(c) == _key(target.name) for c in columns):
                raise SchemaError(
                    f"column {target.name!r} is part of relationship {rel.effective_name!r}; delete it first"
                )
        remaining = tuple(c for c in spec.columns if _key(c.name) != _key(target.name))
        if not remaining:
            raise SchemaError(f"cannot drop the last column of table {spec.name!r}")
        self._store(spec.model_copy(update={"columns": remaining}))
        self._custom.pop(("field", _key(spec.name), _key(target.name)), None)

    def rename_column(self, table: str, old: str, new: str) -> None:
        spec = self._table(table)
        target = self._column(spec, old)
        if _key(old) != _key(new) and any(_key(c.name) == _key(new) for c in spec.columns):
            raise ObjectExistsError(
                f"table {spec.name!r} already has a column {new!r}", kind=ObjectKind.FIELD, name=new
            )
        columns = tuple(
            c.model_copy(update={"name": new}) if c is target else c for c in spec.columns
        )
        indexes = tuple(
            i.model_copy(
                update={
                    "fields": tuple(
                        IndexField(name=new, descending=f.descending)
                        if _key(f.name) == _key(old)
                        else f
                        for f in i.fields
                    )
                }
            )
            for i in spec.indexes
        )
        self._store(spec.model_copy(update={"columns": columns, "indexes": indexes}))
        for key, rel in list(self._relationships.items()):
            update: dict[str, Any] = {}
            if _key(rel.primary_table) == _key(spec.name):
                update["primary_columns"] = tuple(
                    new if _key(c) == _key(old) else c for c in rel.primary_columns
                )
            if _key(rel.foreign_table) == _key(spec.name):
                update["foreign_columns"] = tuple(
                    new if _key(c) == _key(old) else c for c in rel.foreign_columns
                )
            if update:
                self._relationships[key] = rel.model_copy(update=update)

    def create_index(self, table: str, index: IndexSpec) -> None:
        spec = self._table(table)
        taken = {_key(i.name) for i in spec.indexes} | {
            _key(rel.effective_name)
            for rel in self._relationships.values()
            if _key(rel.foreign_table) == _key(spec.name)
        }
        if _key(index.name) in taken:
            raise ObjectExistsError(
                f"table {spec.name!r} already has an index named {index.name!r}",
                kind=ObjectKind.INDEX,
                name=index.name,
            )
        if index.primary and any(i.primary for i in spec.indexes):
            raise SchemaError(f"table {spec.name!r} already has a primary key")
        for field_name in index.field_names:
            self._column(spec, field_name)
        if self.count_indexes(spec.name) >= MAX_INDEXES:
            raise SchemaError(
                f"table {spec.name!r} already has {MAX_INDEXES} indexes (the Access maximum)"
            )
        self._store(spec.model_copy(update={"indexes": (*spec.indexes, index)}).normalized())

    def drop_index(self, table: str, name: str) -> None:
        spec = self._table(table)
        remaining = tuple(i for i in spec.indexes if _key(i.name) != _key(name))
        if len(remaining) == len(spec.indexes):
            raise ObjectNotFoundError(
                f"table {spec.name!r} has no index {name!r}", kind=ObjectKind.INDEX, name=name
            )
        self._store(spec.model_copy(update={"indexes": remaining}))

    # ----------------------------------------------------------------------------- relationships
    def list_relationships(self) -> list[RelationshipSpec]:
        return list(self._relationships.values())

    def create_relationship(self, spec: RelationshipSpec) -> None:
        spec = spec.normalized()
        name = spec.effective_name
        if _key(name) in self._relationships:
            raise ObjectExistsError(
                f"a relationship named {name!r} already exists",
                kind=ObjectKind.RELATIONSHIP,
                name=name,
            )
        primary = self._table(spec.primary_table)
        foreign = self._table(spec.foreign_table)
        if any(_key(i.name) == _key(name) for i in foreign.indexes):
            raise ObjectExistsError(
                f"table {foreign.name!r} already has an index named {name!r}",
                kind=ObjectKind.INDEX,
                name=name,
            )
        for p_name, f_name in zip(spec.primary_columns, spec.foreign_columns, strict=True):
            p_col, f_col = self._column(primary, p_name), self._column(foreign, f_name)
            if not relationship_compatible(p_col, f_col):
                raise RelationshipError(
                    "Relationship must be on the same number of fields with the same data types."
                )
        wanted = {_key(c) for c in spec.primary_columns}
        if not any(
            i.unique and {_key(f) for f in i.field_names} == wanted for i in primary.indexes
        ):
            raise RelationshipError(
                "No unique index found for the referenced field of the primary table."
            )
        if spec.enforce_integrity and self.count_indexes(foreign.name) >= MAX_INDEXES:
            raise SchemaError(
                f"table {foreign.name!r} already has {MAX_INDEXES} indexes (the Access maximum)"
            )
        self._relationships[_key(name)] = spec

    def drop_relationship(self, name: str) -> None:
        if self._relationships.pop(_key(name), None) is None:
            raise ObjectNotFoundError(
                f"relationship {name!r} does not exist", kind=ObjectKind.RELATIONSHIP, name=name
            )

    # ----------------------------------------------------------------------------------- queries
    def list_queries(self) -> list[QueryInfo]:
        return [
            QueryInfo(
                q.name,
                detect_query_kind(q.sql, pass_through=q.pass_through is not None),
                q.name.startswith("~"),
            )
            for q in self._queries.values()
        ]

    def read_query(self, name: str) -> QuerySpec:
        return self._query(name)

    def query_parameters(self, name: str) -> list[ParameterInfo]:
        match = _PARAMETERS.match(self._query(name).sql)
        if match is None:
            return []
        return [
            ParameterInfo(
                item.group("bracketed") or item.group("plain"),
                _PARAM_TYPES.get(item.group("type").casefold(), DataType.UNKNOWN),
            )
            for item in _PARAM_ITEM.finditer(match.group("body"))
        ]

    def create_query(self, spec: QuerySpec) -> None:
        self._ensure_free(spec.name)
        self._queries[_key(spec.name)] = spec.normalized()

    def set_query_sql(self, name: str, sql: str) -> None:
        query = self._query(name)
        self._queries[_key(name)] = query.model_copy(update={"sql": sql}).normalized()

    def rename_query(self, old: str, new: str) -> None:
        query = self._query(old)
        if _key(old) != _key(new):
            self._ensure_free(new)
        del self._queries[_key(old)]
        self._queries[_key(new)] = query.model_copy(update={"name": new})

    def drop_query(self, name: str) -> None:
        self._query(name)
        del self._queries[_key(name)]

    # -------------------------------------------------------------------------------- properties
    def _resolve(self, target: PropertyTarget) -> tuple[tuple[str, str, str], Any, dict[str, str]]:
        if target.kind == "database":
            return ("database", "", ""), None, {}
        if target.kind == "table":
            spec = self._table(target.name or "")
            return ("table", _key(spec.name), ""), spec, _TABLE_PROPS
        if target.kind == "field":
            spec = self._table(target.name or "")
            column = self._column(spec, target.field or "")
            return ("field", _key(spec.name), _key(column.name)), column, _FIELD_PROPS
        query = self._query(target.name or "")
        return ("query", _key(query.name), ""), query, _QUERY_PROPS

    def _replace(self, target: PropertyTarget, obj: Any, attr: str, value: Any) -> None:
        updated = obj.model_validate({**obj.model_dump(), attr: value})
        if target.kind == "table":
            self._store(updated)
        elif target.kind == "field":
            spec = self._table(target.name or "")
            self._store(
                spec.model_copy(
                    update={"columns": tuple(updated if c is obj else c for c in spec.columns)}
                )
            )
        else:
            self._queries[_key(obj.name)] = updated

    @staticmethod
    def _missing(target: PropertyTarget, name: str) -> ObjectNotFoundError:
        return ObjectNotFoundError(
            f"{target.describe()} has no property {name!r}", kind=ObjectKind.PROPERTY, name=name
        )

    def get_property(self, target: PropertyTarget, name: str) -> PropertyValue:
        store_key, obj, mapped = self._resolve(target)
        attr = mapped.get(_key(name))
        if attr is not None and hasattr(obj, attr):
            value = getattr(obj, attr)
            if attr == "default":
                return render_literal(value) if value is not None else ""
            if attr in ("validation_rule", "validation_text"):
                return value or ""
            if value is None:
                raise self._missing(target, name)
            return value
        entry = self._custom.get(store_key, {}).get(_key(name))
        if entry is None:
            raise self._missing(target, name)
        return entry[1]

    def set_property(
        self,
        target: PropertyTarget,
        name: str,
        value: PropertyValue,
        type: PropertyType | None = None,
    ) -> None:
        store_key, obj, mapped = self._resolve(target)
        attr = mapped.get(_key(name))
        if attr is not None and hasattr(obj, attr):
            new_value: Any = value
            if attr == "default":
                new_value = parse_literal(str(value)) if value not in (None, "") else None
            elif (attr == "decimal_places" and value == 255) or (
                attr in ("validation_rule", "validation_text") and value == ""
            ):
                new_value = None
            self._replace(target, obj, attr, new_value)
            return
        self._custom.setdefault(store_key, {})[_key(name)] = (name, value)

    def delete_property(self, target: PropertyTarget, name: str) -> None:
        store_key, obj, mapped = self._resolve(target)
        attr = mapped.get(_key(name))
        if attr is not None and hasattr(obj, attr):
            if getattr(obj, attr) is None:
                raise self._missing(target, name)
            self._replace(target, obj, attr, None)
            return
        if self._custom.get(store_key, {}).pop(_key(name), None) is None:
            raise self._missing(target, name)

    def list_properties(self, target: PropertyTarget) -> dict[str, PropertyValue]:
        store_key, obj, mapped = self._resolve(target)
        result: dict[str, PropertyValue] = {}
        for prop_key, attr in mapped.items():
            if hasattr(obj, attr) and getattr(obj, attr) is not None:
                result[_DISPLAY_NAMES[prop_key]] = self.get_property(target, prop_key)
        result.update(dict(self._custom.get(store_key, {}).values()))
        return result

    # -------------------------------------------------------------------------------------- data
    def _no_sql(self) -> CapabilityError:
        return CapabilityError(
            "the in-memory backend cannot run SQL; open a real database for data access"
        )

    def execute(self, sql: str, params: Mapping[str, Any] | None = None) -> int:
        raise self._no_sql()

    def fetch(
        self, sql: str, params: Mapping[str, Any] | None = None, *, limit: int | None = None
    ) -> FetchResult:
        raise self._no_sql()

    def execute_saved(self, name: str, params: Mapping[str, Any] | None = None) -> int:
        raise self._no_sql()

    def fetch_saved(
        self, name: str, params: Mapping[str, Any] | None = None, *, limit: int | None = None
    ) -> FetchResult:
        raise self._no_sql()

    def list_documents(self, kind: ObjectKind) -> list[str]:
        return self.list_objects(kind)

    # ------------------------------------------------------------------------------------ design
    def _design_names(self, kind: ObjectKind) -> dict[str, str]:
        names = {key: name for key, (name, _) in self._objects[kind].items()}
        if kind is ObjectKind.FORM:
            names.update({key: form.spec.name for key, form in self._forms.items()})
        if kind is ObjectKind.QUERY:  # like CurrentData.AllQueries: saved QueryDefs are listed too
            names.update({key: query.name for key, query in self._queries.items()})
        return names

    def list_objects(self, kind: ObjectKind) -> list[str]:
        return list(self._design_names(kind).values())

    def _require_object(self, kind: ObjectKind, name: str) -> str:
        names = self._design_names(kind)
        if _key(name) not in names:
            raise ObjectNotFoundError(f"{kind.value} {name!r} does not exist", kind=kind, name=name)
        return names[_key(name)]

    def delete_object(self, kind: ObjectKind, name: str) -> None:
        self._require_object(kind, name)
        self._objects[kind].pop(_key(name), None)
        if kind is ObjectKind.FORM:
            self._forms.pop(_key(name), None)

    def rename_object(self, kind: ObjectKind, old: str, new: str) -> None:
        self._require_object(kind, old)
        if _key(old) != _key(new) and _key(new) in self._design_names(kind):
            raise ObjectExistsError(
                f"a {kind.value} named {new!r} already exists", kind=kind, name=new
            )
        if _key(old) in self._objects[kind]:
            _, data = self._objects[kind].pop(_key(old))
            self._objects[kind][_key(new)] = (new, data)
        if kind is ObjectKind.FORM and _key(old) in self._forms:
            form = self._forms.pop(_key(old))
            self._forms[_key(new)] = ResolvedForm(
                spec=form.spec.model_copy(update={"name": new}),
                width=form.width,
                detail_height=form.detail_height,
                header_height=form.header_height,
                footer_height=form.footer_height,
                controls=form.controls,
                events=form.events,
                module_text=form.module_text,
            )

    def export_text(self, kind: ObjectKind, name: str) -> bytes:
        actual = self._require_object(kind, name)
        if _key(name) in self._objects[kind]:
            return self._objects[kind][_key(name)][1]
        form = self._forms[_key(name)]
        lines = [
            "Version =21",
            "Begin Form",
            f'    Caption ="{form.spec.caption or actual}"',
            "End",
        ]
        return "\r\n".join(lines).encode("utf-16")

    def import_text(self, kind: ObjectKind, name: str, data: bytes) -> None:
        existing = self._design_names(kind).get(_key(name))
        if kind is ObjectKind.FORM:
            self._forms.pop(_key(name), None)
        self._objects[kind][_key(name)] = (existing or name, data)

    def build_form(self, form: ResolvedForm, *, replace: bool) -> None:
        name = form.spec.name
        if _key(name) in self._design_names(ObjectKind.FORM) and not replace:
            raise ObjectExistsError(
                f"a form named {name!r} already exists", kind=ObjectKind.FORM, name=name
            )
        self._objects[ObjectKind.FORM].pop(_key(name), None)
        self._forms[_key(name)] = form

    def form_controls(self, name: str) -> list[ControlInfo]:
        self._require_object(ObjectKind.FORM, name)
        form = self._forms.get(_key(name))
        if form is None:
            return []
        infos: list[ControlInfo] = []
        for control in form.controls:
            source = getattr(control.spec, "field", None) or getattr(
                control.spec, "control_source", None
            )
            caption = getattr(control.spec, "caption", None)
            r = control.rect
            infos.append(
                ControlInfo(
                    control.name,
                    control.spec.control_kind,
                    control.section,
                    r.left,
                    r.top,
                    r.width,
                    r.height,
                    source,
                    caption,
                    None,
                )
            )
            if control.label is not None:
                lr = control.label.rect
                infos.append(
                    ControlInfo(
                        control.label.name,
                        ControlKind.LABEL,
                        control.label.section,
                        lr.left,
                        lr.top,
                        lr.width,
                        lr.height,
                        None,
                        control.label.caption,
                        control.name,
                    )
                )
        return infos

    def check_form_opens(self, name: str) -> None:
        self._require_object(ObjectKind.FORM, name)
        form = self._forms.get(_key(name))
        if form is None or form.spec.record_source is None:
            return
        source = form.spec.record_source
        if source.lstrip().upper().startswith("SELECT"):
            return
        if _key(source) in self._queries:
            return
        table = self._table(source)
        for control in form.controls:
            field = control.spec.bound_field
            if field is not None:
                self._column(table, field)
