"""Design operations (forms, modules, raw text objects): validation first, then the design backend."""

from __future__ import annotations

from pyaccesskit._backends.protocols import DesignBackend, SchemaBackend
from pyaccesskit._ops.schema import find_query, find_table
from pyaccesskit._text import codec
from pyaccesskit.enums import ModuleKind, ObjectKind
from pyaccesskit.errors import ObjectExistsError, ObjectNotFoundError, SpecError
from pyaccesskit.forms.layout import ResolvedForm, layout_form
from pyaccesskit.forms.spec import FormSpec
from pyaccesskit.schema.names import check_name, warn_name

__all__ = [
    "DESIGN_KINDS",
    "build_form",
    "create_module",
    "export_object",
    "find_object",
    "import_object",
    "read_module",
]

DESIGN_KINDS = (ObjectKind.FORM, ObjectKind.REPORT, ObjectKind.MACRO, ObjectKind.MODULE)
_SQL_PREFIXES = ("SELECT", "PARAMETERS", "TRANSFORM", "(")


def _check_kind(kind: ObjectKind) -> None:
    if kind not in DESIGN_KINDS and kind is not ObjectKind.QUERY:
        raise SpecError(f"{kind.value} objects cannot be exported/imported as text")


def find_object(design: DesignBackend, kind: ObjectKind, name: str) -> str | None:
    """The actual name of a form/report/macro/module called ``name`` (case-insensitive), or ``None``."""
    key = name.casefold()
    return next((n for n in design.list_objects(kind) if n.casefold() == key), None)


def build_form(
    schema: SchemaBackend, design: DesignBackend, spec: FormSpec, *, replace: bool = False
) -> ResolvedForm:
    """Validate ``spec`` against the database, lay it out, and build it atomically."""
    resolved = layout_form(spec)
    existing = find_object(design, ObjectKind.FORM, spec.name)
    if existing is not None and not replace:
        raise ObjectExistsError(
            f"a form named {existing!r} already exists (pass replace=True to rebuild it)",
            kind=ObjectKind.FORM,
            name=spec.name,
        )
    source = spec.record_source
    if source is not None and not source.lstrip().upper().startswith(_SQL_PREFIXES):
        table = find_table(schema, source)
        if table is None and find_query(schema, source) is None:
            raise ObjectNotFoundError(
                f"record source {source!r} of form {spec.name!r} is neither a table nor a query",
                kind=ObjectKind.TABLE,
                name=source,
            )
        if table is not None:
            columns = schema.read_table(table.name).column_names
            known = {c.casefold() for c in columns}
            for control in resolved.controls:
                field = control.spec.bound_field
                if field is not None and field.casefold() not in known:
                    raise SpecError(
                        f"control {control.name!r} of form {spec.name!r} is bound to {field!r}, which is not a "
                        f"column of {table.name!r} (columns: {', '.join(columns)})"
                    )
    warn_name(spec.name, what="form name", stacklevel=4)
    design.build_form(resolved, replace=existing is not None)
    return resolved


def create_module(
    design: DesignBackend,
    name: str,
    code: str,
    kind: ModuleKind = ModuleKind.STANDARD,
    *,
    replace: bool = False,
) -> None:
    """Create (or replace) a VBA module from source code."""
    check_name(name, what="module name")
    existing = find_object(design, ObjectKind.MODULE, name)
    if existing is not None and not replace:
        raise ObjectExistsError(
            f"a module named {existing!r} already exists (pass replace=True to overwrite it)",
            kind=ObjectKind.MODULE,
            name=name,
        )
    text = codec.module_import_text(code, kind)
    data = codec.encode_import(ObjectKind.MODULE, text)
    design.import_text(ObjectKind.MODULE, existing or name, data)


def read_module(design: DesignBackend, name: str) -> tuple[ModuleKind, str]:
    """Return ``(kind, code)`` of a VBA module."""
    actual = find_object(design, ObjectKind.MODULE, name)
    if actual is None:
        raise ObjectNotFoundError(
            f"module {name!r} does not exist", kind=ObjectKind.MODULE, name=name
        )
    text = codec.decode_export(ObjectKind.MODULE, design.export_text(ObjectKind.MODULE, actual))
    return codec.split_module_export(text)


def export_object(design: DesignBackend, kind: ObjectKind, name: str) -> str:
    """``SaveAsText`` an object and return its text (LF line endings)."""
    _check_kind(kind)
    return codec.decode_export(kind, design.export_text(kind, name))


def import_object(
    design: DesignBackend, kind: ObjectKind, name: str, text: str, *, replace: bool = False
) -> None:
    """``LoadFromText`` an object from text (encoded the way Access expects for ``kind``)."""
    _check_kind(kind)
    check_name(name, what=f"{kind.value} name")
    if kind is not ObjectKind.QUERY:
        existing = find_object(design, kind, name)
        if existing is not None and not replace:
            raise ObjectExistsError(
                f"a {kind.value} named {existing!r} already exists (pass replace=True to overwrite it)",
                kind=kind,
                name=name,
            )
    design.import_text(kind, name, codec.encode_import(kind, text))
