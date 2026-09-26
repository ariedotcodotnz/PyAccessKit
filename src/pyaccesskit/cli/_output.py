"""Console helpers shared by the commands: exit codes, error reporting, JSON output."""

from __future__ import annotations

import json
import re
import sys
from collections.abc import Generator
from contextlib import contextmanager
from typing import Any

import typer
from rich.console import Console
from rich.markup import escape

from pyaccesskit.errors import EnvironmentProblem, PyAccessKitError

EXIT_OK = 0
EXIT_ERROR = 1
EXIT_USAGE = 2
EXIT_ENVIRONMENT = 3

_SECRET = re.compile(r"(?i)\b(PWD|PASSWORD)=[^;]*")


def stdout() -> Console:
    """A console for normal output (created per call so tests can capture it)."""
    return Console(highlight=False, soft_wrap=False)


def stderr() -> Console:
    """A console for errors."""
    return Console(stderr=True, highlight=False)


def redact(connect: str | None) -> str | None:
    """Hide passwords in a connection string."""
    return None if connect is None else _SECRET.sub(r"\1=***", connect)


def print_json(data: Any) -> None:
    """Write ``data`` as ASCII-safe JSON (works whatever the console or pipe encoding)."""
    sys.stdout.write(json.dumps(data, indent=2, default=str) + "\n")


@contextmanager
def handle_errors() -> Generator[None, None, None]:
    """Report PyAccessKit errors without a traceback and exit with the documented code."""
    try:
        yield
    except EnvironmentProblem as exc:
        _report(exc)
        raise typer.Exit(EXIT_ENVIRONMENT) from None
    except PyAccessKitError as exc:
        _report(exc)
        raise typer.Exit(EXIT_ERROR) from None


def _report(exc: PyAccessKitError) -> None:
    console = stderr()
    console.print(f"[bold red]error:[/] {escape(str(exc))}")  # includes any engine diagnosis
    for note in getattr(exc, "__notes__", ()):
        console.print(f"  [dim]{escape(str(note))}[/]")
