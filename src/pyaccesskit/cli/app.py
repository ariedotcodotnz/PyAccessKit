"""The ``pyaccesskit`` command-line application.

Exit codes: 0 success, 1 error, 2 usage error, 3 environment unusable (no engine, missing Access...).
"""

from __future__ import annotations

from typing import Annotated

import typer

from pyaccesskit._version import __version__
from pyaccesskit.cli.agent import guide, schema
from pyaccesskit.cli.cleanup import cleanup
from pyaccesskit.cli.doctor import doctor
from pyaccesskit.cli.inspection import inspect_database

app = typer.Typer(
    name="pyaccesskit",
    help="Create, inspect and maintain Microsoft Access databases from Python.",
    no_args_is_help=True,
    add_completion=False,
    pretty_exceptions_enable=False,
)


def _show_version(value: bool) -> None:
    if value:
        typer.echo(f"pyaccesskit {__version__}")
        raise typer.Exit


@app.callback()
def _root(
    version: Annotated[
        bool,
        typer.Option(
            "--version", callback=_show_version, is_eager=True, help="Show the version and exit."
        ),
    ] = False,
) -> None:
    """Create, inspect and maintain Microsoft Access databases from Python."""


app.command("doctor")(doctor)
app.command("inspect")(inspect_database)
app.command("cleanup")(cleanup)
app.command("guide")(guide)
app.command("schema")(schema)


def run() -> None:
    """Run the CLI (used by the ``pyaccesskit`` console script and ``python -m pyaccesskit``)."""
    app(prog_name="pyaccesskit")
