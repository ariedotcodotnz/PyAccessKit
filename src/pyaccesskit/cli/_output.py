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

_SECRET_KEYS = re.compile(r"(?i)PWD|PASSWORD")


def stdout() -> Console:
    """A console for normal output (created per call so tests can capture it)."""
    return Console(highlight=False, soft_wrap=False)


def stderr() -> Console:
    """A console for errors."""
    return Console(stderr=True, highlight=False)


def _segments(connect: str) -> list[str]:
    """Split a connection string at the ``;`` separators, honouring ODBC ``{...}`` values (``}}`` escapes)."""
    segments: list[str] = []
    current: list[str] = []
    index, braced, at_value_start = 0, False, False
    while index < len(connect):
        char = connect[index]
        if braced:
            current.append(char)
            if char == "}":
                if connect[index + 1 : index + 2] == "}":  # escaped closing brace
                    current.append("}")
                    index += 1
                else:
                    braced = False
        elif char == ";":
            segments.append("".join(current))
            current = []
            at_value_start = False
        else:
            current.append(char)
            if char == "{" and at_value_start:
                braced = True
            at_value_start = char == "=" and "=" not in "".join(current[:-1])
        index += 1
    segments.append("".join(current))
    return segments


def redact(connect: str | None) -> str | None:
    """Hide password values (``PWD=``/``Password=``, including ``{brace;quoted}`` ones) in a connection string."""
    if connect is None:
        return None
    parts: list[str] = []
    for segment in _segments(connect):
        key, sep, _value = segment.partition("=")
        if sep and _SECRET_KEYS.fullmatch(key.strip()):
            parts.append(f"{key}=***")
        else:
            parts.append(segment)
    return ";".join(parts)


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
