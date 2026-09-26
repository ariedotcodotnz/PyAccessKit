"""COM-free interfaces between the session and the engines."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal, Protocol

from pyaccesskit._backends.protocols import DesignBackend, SchemaBackend
from pyaccesskit.enums import Transport
from pyaccesskit.options import SessionOptions

__all__ = ["EngineHandle", "EnginePlan", "EnvironmentProbe", "RawKind"]

RawKind = Literal["dao", "dbengine", "access"]


@dataclass(frozen=True)
class EnginePlan:
    """What engine to open, on which file, and how."""

    kind: Literal["dao", "access"]
    path: Path
    create: bool
    readonly: bool
    exclusive: bool
    password: str | None
    design: bool
    options: SessionOptions


class EngineHandle(Protocol):
    """An open engine: owns every COM root reference (and, for Access, the process)."""

    @property
    def transport(self) -> Transport:
        """How the database is currently reached."""
        ...

    @property
    def supports_design(self) -> bool:
        """Whether design features are possible without switching engines."""
        ...

    def schema(self) -> SchemaBackend:
        """The schema backend."""
        ...

    def design(self) -> DesignBackend:
        """The design backend (may upgrade to a design session, or raise ``CapabilityError``)."""
        ...

    def raw(self, which: str) -> Any:
        """Raw COM object for the escape hatch."""
        ...

    def close(self) -> list[BaseException]:
        """Release everything; returns the errors of failed cleanup steps."""
        ...

    def terminate(self) -> None:
        """Emergency stop from any thread (terminates an owned Access process)."""
        ...


class EnvironmentProbe(Protocol):
    """Answers "which engines can this process use?" (replaceable in tests)."""

    def inproc_dao(self) -> tuple[bool, str]:
        """``(available, reason_if_not)`` for in-process DAO."""
        ...

    def access(self, progid: str) -> bool:
        """Whether Microsoft Access is registered."""
        ...
