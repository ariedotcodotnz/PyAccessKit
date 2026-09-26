# pyright: reportUnnecessaryIsInstance=false
# (dataclass __post_init__ and render_literal validate untyped input at runtime)
"""VBA snippets for event procedures and helpers that assemble form modules."""

from __future__ import annotations

import re
import textwrap
from collections.abc import Sequence
from dataclasses import dataclass

from pyaccesskit.errors import SpecError

__all__ = [
    "VBA_IDENTIFIER",
    "EventBinding",
    "Vba",
    "build_module",
    "event_procedure",
    "is_vba_identifier",
]

VBA_IDENTIFIER = re.compile(r"^[A-Za-z][A-Za-z0-9_]{0,254}$")


def is_vba_identifier(name: str) -> bool:
    """Whether ``name`` can be used as-is in a VBA procedure name (``cmdSave_Click``)."""
    return bool(VBA_IDENTIFIER.match(name))


@dataclass(frozen=True)
class Vba:
    """The body of a VBA event procedure, e.g. ``Vba("DoCmd.Close acForm, Me.Name")``.

    Indentation is normalized; write the body only (no ``Sub``/``End Sub``).
    """

    code: str

    def __post_init__(self) -> None:
        if not isinstance(self.code, str) or not self.code.strip():
            raise SpecError("Vba requires non-empty code")
        if re.search(r"^\s*(Private\s+|Public\s+)?Sub\s", self.code, re.IGNORECASE | re.MULTILINE):
            raise SpecError("Vba holds a procedure *body*; do not include 'Sub ... End Sub'")

    def lines(self) -> list[str]:
        """The body lines, dedented, without trailing blank lines."""
        text = textwrap.dedent(self.code.replace("\r\n", "\n").replace("\r", "\n")).strip("\n")
        return [line.rstrip() for line in text.split("\n")]


@dataclass(frozen=True)
class EventBinding:
    """An event wired to an ``[Event Procedure]``.

    Attributes:
        object_name: ``"Form"`` or a control name.
        property_name: The Access event property, e.g. ``"OnClick"``.
        event: The VBA event name, e.g. ``"Click"``.
        body: The procedure body.
    """

    object_name: str
    property_name: str
    event: str
    body: Vba

    @property
    def procedure_name(self) -> str:
        """The VBA procedure name, e.g. ``cmdClose_Click``."""
        return f"{self.object_name}_{self.event}"


def event_procedure(binding: EventBinding) -> str:
    """Render ``Private Sub <object>_<event>() ... End Sub`` (CRLF line endings)."""
    body = "\r\n".join(f"    {line}" if line else "" for line in binding.body.lines())
    return f"Private Sub {binding.procedure_name}()\r\n{body}\r\nEnd Sub\r\n"


def build_module(
    bindings: Sequence[EventBinding], extra_code: str | None = None, *, option_explicit: bool = True
) -> str:
    """Assemble a complete class-module text for a form: options, event procedures, extra code."""
    header = ["Option Compare Database"]
    if option_explicit:
        header.append("Option Explicit")
    declarations, procedures = split_declarations(extra_code or "")
    parts = ["\r\n".join(header) + "\r\n"]
    # VBA only accepts module-level declarations before the first procedure.
    if declarations:
        parts.append(_crlf(declarations))
    parts.extend(event_procedure(binding) for binding in bindings)
    if procedures:
        parts.append(_crlf(procedures))
    return "\r\n".join(parts)


_PROCEDURE_START = re.compile(
    r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s",
    re.IGNORECASE,
)
_DIRECTIVE_IF = re.compile(r"^\s*#If\b", re.IGNORECASE)
_DIRECTIVE_END_IF = re.compile(r"^\s*#End\s*If\b", re.IGNORECASE)
_OPTION_LINE = re.compile(r"^\s*Option\s+(?:Compare|Explicit)\b", re.IGNORECASE)


def split_declarations(code: str) -> tuple[str, str]:
    """Split VBA into its declarations section and its procedures (``Option`` lines are dropped).

    Everything before the first ``Sub``/``Function``/``Property`` is the declarations section; comments
    directly above that procedure stay with it. A ``#If ... #End If`` block containing that procedure is
    never split: it moves to the procedures as a whole.
    """
    lines = [
        line
        for line in code.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        if not _OPTION_LINE.match(line)
    ]
    start = len(lines)
    depth, block_start = 0, 0
    for index, line in enumerate(lines):
        if _DIRECTIVE_IF.match(line):
            if depth == 0:
                block_start = index
            depth += 1
        elif _DIRECTIVE_END_IF.match(line):
            depth = max(depth - 1, 0)
        elif _PROCEDURE_START.match(line):
            # Never split inside #If ... #End If: the whole conditional block goes with the procedures.
            start = block_start if depth else index
            break
    while start > 0 and lines[start - 1].lstrip().startswith("'"):
        start -= 1
    return "\n".join(lines[:start]).strip("\n"), "\n".join(lines[start:]).strip("\n")


def _crlf(code: str) -> str:
    return code.replace("\n", "\r\n") + "\r\n"
