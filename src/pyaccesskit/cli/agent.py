"""``pyaccesskit guide`` and ``pyaccesskit schema``: material for AI coding agents (and humans)."""

from __future__ import annotations

import sys
from enum import StrEnum
from typing import Annotated, Any

import typer
from pydantic import TypeAdapter

from pyaccesskit.cli._output import EXIT_OK, print_json
from pyaccesskit.forms import FormSpec
from pyaccesskit.schema import ColumnSpec, IndexSpec, QuerySpec, RelationshipSpec, TableSpec

GUIDE_RESOURCE = "AGENT_GUIDE.md"


class SchemaKind(StrEnum):
    """Spec types whose JSON Schema can be printed."""

    ALL = "all"
    TABLE = "table"
    COLUMN = "column"
    INDEX = "index"
    RELATIONSHIP = "relationship"
    QUERY = "query"
    FORM = "form"


_MODELS: dict[SchemaKind, Any] = {
    SchemaKind.TABLE: TableSpec,
    SchemaKind.COLUMN: TypeAdapter(ColumnSpec),
    SchemaKind.INDEX: IndexSpec,
    SchemaKind.RELATIONSHIP: RelationshipSpec,
    SchemaKind.QUERY: QuerySpec,
    SchemaKind.FORM: FormSpec,
}


def guide_text() -> str:
    """The agent guide shipped inside the package."""
    from importlib.resources import files  # noqa: PLC0415 - only needed here

    return files("pyaccesskit").joinpath(GUIDE_RESOURCE).read_text(encoding="utf-8")


def guide_path() -> str:
    """Where the installed agent guide lives on disk."""
    from importlib.resources import files  # noqa: PLC0415

    return str(files("pyaccesskit").joinpath(GUIDE_RESOURCE))


def json_schema(kind: SchemaKind) -> dict[str, Any]:
    """The JSON Schema of one spec type, or of all of them keyed by kind."""
    if kind is SchemaKind.ALL:
        return {k.value: json_schema(k) for k in _MODELS}
    model = _MODELS[kind]
    if isinstance(model, TypeAdapter):
        return model.json_schema()  # pyright: ignore[reportUnknownMemberType, reportUnknownVariableType]
    return model.model_json_schema()


def guide(
    path: Annotated[
        bool, typer.Option("--path", help="Print where the guide file is instead of its text.")
    ] = False,
) -> None:
    """Print the guide for AI agents that write Access software with PyAccessKit."""
    if path:
        typer.echo(guide_path())
    else:
        sys.stdout.write(guide_text())
    raise typer.Exit(EXIT_OK)


def schema(
    kind: Annotated[
        SchemaKind, typer.Argument(help="Which spec to describe.", show_default=True)
    ] = SchemaKind.ALL,
) -> None:
    """Print the JSON Schema of PyAccessKit specs (tables, columns, queries, forms...)."""
    print_json(json_schema(kind))
    raise typer.Exit(EXIT_OK)
