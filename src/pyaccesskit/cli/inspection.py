"""``pyaccesskit inspect``: describe a database without changing it."""

from __future__ import annotations

from pathlib import Path
from typing import Annotated, Any

import typer
from rich.markup import escape
from rich.table import Table

from pyaccesskit.cli._output import EXIT_OK, handle_errors, print_json, redact, stdout
from pyaccesskit.database import AccessDatabase
from pyaccesskit.enums import DataType, Engine, ObjectKind
from pyaccesskit.errors import PyAccessKitError
from pyaccesskit.schema import (
    AutoNumberColumn,
    ColumnBase,
    DecimalColumn,
    Expr,
    LongTextColumn,
    NumberColumn,
    TableSpec,
    TextColumn,
    render_literal,
)
from pyaccesskit.schema.columns import UnsupportedColumn

OTHER_OBJECTS = (ObjectKind.FORM, ObjectKind.REPORT, ObjectKind.MACRO, ObjectKind.MODULE)
_TYPE_NAMES = {
    DataType.TEXT: "Short Text",
    DataType.LONG_TEXT: "Long Text",
    DataType.NUMBER: "Number",
    DataType.DECIMAL: "Decimal",
    DataType.CURRENCY: "Currency",
    DataType.AUTONUMBER: "AutoNumber",
    DataType.DATE_TIME: "Date/Time",
    DataType.YES_NO: "Yes/No",
    DataType.HYPERLINK: "Hyperlink",
    DataType.OLE_OBJECT: "OLE Object",
}


def inspect_database(
    database: Annotated[Path, typer.Argument(help="The .accdb or .mdb file.", show_default=False)],
    json_output: Annotated[bool, typer.Option("--json", help="Print the report as JSON.")] = False,
    counts: Annotated[
        bool, typer.Option("--counts", help="Also count the rows of every table.")
    ] = False,
    password: Annotated[
        str | None,
        typer.Option(
            "--password",
            envvar="PYACCESSKIT_PASSWORD",
            help="Database password (prefer the PYACCESSKIT_PASSWORD environment variable).",
            show_default=False,
        ),
    ] = None,
    engine: Annotated[Engine, typer.Option(help="Engine to use.")] = Engine.AUTO,
) -> None:
    """Describe tables, relationships, queries and other objects (read-only; no startup code runs)."""
    with handle_errors():
        db = AccessDatabase.open(database, readonly=True, password=password, engine=engine)
        with db:
            report = describe(db, counts=counts)
    if json_output:
        print_json(report)
    else:
        render(report)
    raise typer.Exit(EXIT_OK)


def describe(db: AccessDatabase, *, counts: bool = False) -> dict[str, Any]:
    """Everything ``inspect`` reports, as JSON-ready data. Problems with single objects become warnings.

    Passwords in linked-table and pass-through connection strings are redacted.
    """
    warnings: list[str] = []
    tables: list[dict[str, Any]] = []
    for table in db.tables:
        entry: dict[str, Any] = {"name": table.name, "linked": None, "rows": None, "spec": None}
        try:
            if table.is_linked:
                entry["linked"] = {
                    "connect": redact(table.connect),
                    "source_table": table.source_table,
                }
            entry["spec"] = table.to_spec().model_dump(mode="json")
            if counts:
                entry["rows"] = table.record_count()
        except PyAccessKitError as exc:
            warnings.append(f"table {table.name!r}: {exc}")
        tables.append(entry)

    queries: list[dict[str, Any]] = []
    for query in db.queries:
        try:
            spec = query.to_spec().model_dump(mode="json")
            passthrough = spec.get("pass_through")
            if isinstance(passthrough, dict):
                passthrough["connect"] = redact(str(passthrough.get("connect")))  # pyright: ignore[reportUnknownMemberType, reportUnknownArgumentType]
            queries.append({"kind": query.kind.value, **spec})
        except PyAccessKitError as exc:
            warnings.append(f"query {query.name!r}: {exc}")

    return {
        "path": str(db.path),
        "format_version": db.format_version,
        "transport": db.transport.value,
        "tables": tables,
        "relationships": [spec.model_dump(mode="json") for spec in db.relationships.specs()],
        "queries": queries,
        **{f"{kind.value}s": db.objects.names(kind) for kind in OTHER_OBJECTS},
        "warnings": warnings,
    }


# ------------------------------------------------------------------------------------ rendering
def column_type(column: ColumnBase) -> str:
    """How the Access table designer would describe a column's type."""
    if isinstance(column, UnsupportedColumn):
        return f"{column.detail} (unsupported)"
    name = _TYPE_NAMES.get(column.data_type, column.data_type.value)
    if isinstance(column, TextColumn):
        return f"{name}({column.length})"
    if isinstance(column, NumberColumn):
        return f"{name} ({column.size.value.replace('_', ' ')})"
    if isinstance(column, DecimalColumn):
        return f"{name}({column.precision},{column.scale})"
    if isinstance(column, AutoNumberColumn) and column.replication_id:
        return f"{name} (replication ID)"
    if isinstance(column, LongTextColumn):
        flags = [
            flag
            for flag, on in (("rich text", column.rich_text), ("append only", column.append_only))
            if on
        ]
        return f"{name} ({', '.join(flags)})" if flags else name
    return name


def _default_text(column: ColumnBase) -> str:
    default = column.default
    if default is None:
        return ""
    return default.expr if isinstance(default, Expr) else render_literal(default)


def _render_table(entry: dict[str, Any]) -> None:
    console = stdout()
    details: list[str] = []
    if entry["linked"]:
        link = entry["linked"]
        details.append(f"linked to {link['source_table']} ({link['connect']})")
    if entry["rows"] is not None:
        details.append(f"{entry['rows']} row{'' if entry['rows'] == 1 else 's'}")
    suffix = f"  [dim]{escape('; '.join(details))}[/]" if details else ""
    console.print(f"\n[bold cyan]Table {escape(entry['name'])}[/]{suffix}")
    if entry["spec"] is None:
        console.print("  [yellow]could not be read (see warnings)[/]")
        return
    spec = TableSpec.model_validate(entry["spec"])
    grid = Table(box=None, padding=(0, 2), show_edge=False, header_style="dim")
    for heading in ("Field", "Type", "Required", "Default", "Description"):
        grid.add_column(heading)
    for column in spec.columns:
        grid.add_row(
            escape(column.name),
            escape(column_type(column)),
            "yes" if column.required else "",
            escape(_default_text(column)),
            escape(column.description or ""),
        )
    console.print(grid)
    for index in spec.indexes:
        kind = "primary key" if index.primary else ("unique" if index.unique else "")
        columns = ", ".join(f.name + (" DESC" if f.descending else "") for f in index.fields)
        console.print(
            f"  [dim]index[/] {escape(index.name)} ({escape(columns)})"
            + (f" [dim]{kind}[/]" if kind else "")
        )


def render(report: dict[str, Any]) -> None:
    """Print a human-readable report."""
    console = stdout()
    console.print(
        f"[bold]{escape(report['path'])}[/]  "
        f"[dim](format {report['format_version']}, via {report['transport']})[/]"
    )
    for entry in report["tables"]:
        _render_table(entry)
    if report["relationships"]:
        console.print("\n[bold cyan]Relationships[/]")
        for rel in report["relationships"]:
            rules = [
                rule
                for rule, on in (
                    ("enforced", rel["enforce_integrity"]),
                    ("cascade update", rel["cascade_update"]),
                    ("cascade delete", rel["cascade_delete"]),
                    ("one-to-one", rel["one_to_one"]),
                )
                if on
            ]
            primary = f"{rel['primary_table']}({', '.join(rel['primary_columns'])})"
            foreign = f"{rel['foreign_table']}({', '.join(rel['foreign_columns'])})"
            console.print(
                f"  {escape(rel['name'] or '')}: {escape(primary)} -> {escape(foreign)}"
                + (f"  [dim]{', '.join(rules)}[/]" if rules else "")
            )
    if report["queries"]:
        console.print("\n[bold cyan]Queries[/]")
        for query in report["queries"]:
            lines = query["sql"].strip().splitlines()
            first = lines[0][:100] if lines else ""
            console.print(f"  {escape(query['name'])} [dim]({query['kind']})[/]  {escape(first)}")
    for kind in OTHER_OBJECTS:
        names = report[f"{kind.value}s"]
        if names:
            console.print(f"\n[bold cyan]{kind.value.title()}s[/]  {escape(', '.join(names))}")
    for warning in report["warnings"]:
        console.print(f"\n[yellow]warning:[/] {escape(warning)}")
