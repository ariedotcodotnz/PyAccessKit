"""Exception and warning hierarchy.

Every exception PyAccessKit raises derives from :class:`PyAccessKitError`. Low-level COM failures are
translated into specific subclasses; the original ``pywintypes.com_error`` is always kept as
``__cause__`` and summarised in :attr:`PyAccessKitError.details`, so nothing is lost for debugging.

Several classes also inherit a matching built-in exception (``FileNotFoundError``, ``LookupError``,
``ValueError``...) so idiomatic ``except`` clauses keep working.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from pyaccesskit.enums import ObjectKind

__all__ = [
    "AccessApplicationError",
    "AccessDialogError",
    "AccessDialogWarning",
    "AccessNameWarning",
    "AccessNotInstalledError",
    "AccessProcessDiedError",
    "AccessRuntimeOnlyError",
    "AccessTimeoutError",
    "CapabilityError",
    "CleanupError",
    "ComError",
    "DaoError",
    "DaoNotAvailableError",
    "DatabaseError",
    "DatabaseExistsError",
    "DatabaseLockedError",
    "DatabaseNotFoundError",
    "DialogInfo",
    "EngineUnavailableError",
    "EnvironmentProblem",
    "ErrorDetails",
    "IntegrityViolationError",
    "InvalidPasswordError",
    "MissingParameterError",
    "ObjectError",
    "ObjectExistsError",
    "ObjectInUseError",
    "ObjectNotFoundError",
    "PyAccessKitError",
    "PyAccessKitWarning",
    "QueryError",
    "ReadOnlyError",
    "RelationshipError",
    "SchemaError",
    "SessionClosedError",
    "SessionError",
    "SpecError",
    "SqlSyntaxError",
    "UnrecognizedFormatError",
    "WrongThreadError",
]


# --------------------------------------------------------------------------------------------- details
@dataclass(frozen=True)
class DaoError:
    """One entry of DAO's ``DBEngine.Errors`` collection."""

    number: int
    description: str
    source: str


@dataclass(frozen=True)
class ErrorDetails:
    """Low-level details of the COM error behind a PyAccessKit exception.

    Attributes:
        hresult: The HRESULT returned by ``IDispatch::Invoke`` (usually ``DISP_E_EXCEPTION``).
        scode: The exception's SCODE (for Access/DAO errors ``0x800A0000 | number``).
        number: The Access/DAO error number (e.g. ``3010``), when the error came from Access or DAO.
        source: Error source, e.g. ``"DAO.TableDefs"`` (``None`` for Access application errors).
        description: The error text reported by Access/DAO.
        dao_errors: Snapshot of ``DBEngine.Errors`` when it matched this error.
    """

    hresult: int | None = None
    scode: int | None = None
    number: int | None = None
    source: str | None = None
    description: str | None = None
    dao_errors: tuple[DaoError, ...] = ()

    def summary(self) -> str:
        """Return a one-line, human-readable summary."""
        parts: list[str] = []
        if self.number is not None:
            kind = "DAO" if (self.source or "").startswith("DAO") else "Access"
            parts.append(f"{kind} error {self.number}")
        elif self.scode is not None:
            parts.append(f"SCODE 0x{self.scode & 0xFFFFFFFF:08X}")
        elif self.hresult is not None:
            parts.append(f"HRESULT 0x{self.hresult & 0xFFFFFFFF:08X}")
        if self.description:
            parts.append(self.description.strip())
        return ": ".join(parts)


@dataclass(frozen=True)
class DialogInfo:
    """A modal dialog that Access showed during an automated call."""

    title: str
    text: str
    buttons: tuple[str, ...] = ()
    action: str = ""
    """How PyAccessKit dealt with it (e.g. ``"closed"``, ``"clicked 'No'"``, ``"terminated Access"``)."""

    def summary(self) -> str:
        """Return a one-line summary such as ``'Microsoft Access: The expression ... [OK]'``."""
        buttons = f" [{', '.join(self.buttons)}]" if self.buttons else ""
        text = f": {self.text}" if self.text else ""
        return f"{self.title or '(untitled dialog)'}{text}{buttons}"


# ------------------------------------------------------------------------------------------------ base
class PyAccessKitError(Exception):
    """Base class of every exception raised by PyAccessKit.

    Attributes:
        message: Human-readable description.
        operation: What PyAccessKit was doing, e.g. ``"create table 'Customers'"``.
        details: COM-level details when the error originated in Access/DAO.
    """

    def __init__(
        self, message: str, *, operation: str | None = None, details: ErrorDetails | None = None
    ) -> None:
        super().__init__(message)
        self.message = message
        self.operation = operation
        self.details = details

    def __str__(self) -> str:
        return self.message


class PyAccessKitWarning(UserWarning):
    """Base class of PyAccessKit warnings."""


class AccessNameWarning(PyAccessKitWarning):
    """A name is legal but likely to cause trouble (reserved word, spaces, special characters...)."""


class AccessDialogWarning(PyAccessKitWarning):
    """Access showed a dialog that was dismissed automatically (``dialog_policy="warn"``)."""


# ------------------------------------------------------------------------------------------ environment
class EnvironmentProblem(PyAccessKitError):
    """The machine cannot do what was asked (missing software, bitness mismatch...)."""


class EngineUnavailableError(EnvironmentProblem):
    """No usable automation engine for the requested operation.

    Attributes:
        diagnosis: Explanation and advice (installed products, bitness, what to try).
    """

    def __init__(
        self,
        message: str,
        *,
        diagnosis: str | None = None,
        operation: str | None = None,
        details: ErrorDetails | None = None,
    ) -> None:
        super().__init__(message, operation=operation, details=details)
        self.diagnosis = diagnosis

    def __str__(self) -> str:
        return f"{self.message}\n{self.diagnosis}" if self.diagnosis else self.message


class AccessNotInstalledError(EngineUnavailableError):
    """Microsoft Access (``Access.Application``) is not installed or cannot be started."""


class DaoNotAvailableError(EngineUnavailableError):
    """In-process DAO (``DAO.DBEngine.120``) cannot be loaded into this Python process."""


class AccessRuntimeOnlyError(EngineUnavailableError):
    """Only the Access Runtime is installed; design features need full Microsoft Access."""


# ------------------------------------------------------------------------------------------- session
class SessionError(PyAccessKitError):
    """The session cannot perform the operation in its current state."""


class SessionClosedError(SessionError):
    """The database session has been closed."""


class WrongThreadError(SessionError):
    """A session was used from a thread other than the one that created it (COM apartment rule)."""


class CapabilityError(SessionError):
    """The operation needs a capability the session does not have (e.g. forms with ``engine="dao"``)."""


class ReadOnlyError(SessionError):
    """A modification was attempted on a session opened with ``readonly=True``."""


# ------------------------------------------------------------------------------------------ database
class DatabaseError(PyAccessKitError):
    """A problem with the database file itself.

    Attributes:
        path: The database path involved, when known.
    """

    def __init__(
        self,
        message: str,
        *,
        path: Path | str | None = None,
        operation: str | None = None,
        details: ErrorDetails | None = None,
    ) -> None:
        super().__init__(message, operation=operation, details=details)
        self.path = Path(path) if path is not None else None


class DatabaseNotFoundError(DatabaseError, FileNotFoundError):
    """The database file does not exist."""


class DatabaseExistsError(DatabaseError, FileExistsError):
    """The database file already exists (pass ``overwrite=True`` to replace it)."""


class DatabaseLockedError(DatabaseError):
    """The database is in use (opened exclusively elsewhere, or locked)."""


class InvalidPasswordError(DatabaseError):
    """The database password is wrong or missing."""


class UnrecognizedFormatError(DatabaseError):
    """The file is not an Access database (or is damaged / from an unsupported version)."""


# -------------------------------------------------------------------------------------------- objects
class ObjectError(PyAccessKitError):
    """A problem with a named database object.

    Attributes:
        kind: Kind of object (table, query, form...).
        name: Object name.
    """

    def __init__(
        self,
        message: str,
        *,
        kind: ObjectKind | None = None,
        name: str | None = None,
        operation: str | None = None,
        details: ErrorDetails | None = None,
    ) -> None:
        super().__init__(message, operation=operation, details=details)
        self.kind = kind
        self.name = name


class ObjectNotFoundError(ObjectError, LookupError):
    """The named object does not exist."""


class ObjectExistsError(ObjectError):
    """An object with that name already exists (tables and queries share one namespace)."""


class ObjectInUseError(ObjectError):
    """The object is locked or open elsewhere."""


# ----------------------------------------------------------------------------------- specs & schema
class SpecError(PyAccessKitError, ValueError):
    """A specification is invalid. Raised by PyAccessKit *before* anything is sent to Access.

    Attributes:
        problems: Individual problems (e.g. from Pydantic validation), each ``"location: message"``.
    """

    def __init__(
        self, message: str, *, problems: tuple[str, ...] = (), operation: str | None = None
    ) -> None:
        super().__init__(message, operation=operation)
        self.problems = problems

    def __str__(self) -> str:
        if not self.problems:
            return self.message
        return self.message + "".join(f"\n  - {problem}" for problem in self.problems)


class SchemaError(PyAccessKitError):
    """The database engine rejected a schema change."""


class RelationshipError(SchemaError):
    """A relationship cannot be created as specified (types, keys, index budget...)."""


class IntegrityViolationError(SchemaError):
    """Existing data violates the referential integrity rule being created."""


# -------------------------------------------------------------------------------------------- queries
class QueryError(PyAccessKitError):
    """A SQL statement or saved query failed.

    Attributes:
        sql: The SQL text involved, when known.
    """

    def __init__(
        self,
        message: str,
        *,
        sql: str | None = None,
        operation: str | None = None,
        details: ErrorDetails | None = None,
    ) -> None:
        super().__init__(message, operation=operation, details=details)
        self.sql = sql


class SqlSyntaxError(QueryError):
    """Access SQL could not parse the statement."""


class MissingParameterError(QueryError):
    """The statement has parameters that were not supplied (DAO error 3061)."""


# ---------------------------------------------------------------------------------- Access application
class AccessApplicationError(PyAccessKitError):
    """A problem with the Microsoft Access application process."""


class AccessDialogError(AccessApplicationError):
    """Access showed a modal dialog that would have blocked automation.

    Attributes:
        dialogs: The dialogs that were observed (title, text, buttons, how they were handled).
    """

    def __init__(
        self,
        message: str,
        *,
        dialogs: tuple[DialogInfo, ...] = (),
        operation: str | None = None,
        details: ErrorDetails | None = None,
    ) -> None:
        super().__init__(message, operation=operation, details=details)
        self.dialogs = dialogs

    def __str__(self) -> str:
        return self.message + "".join(f"\n  - {dialog.summary()}" for dialog in self.dialogs)


class AccessTimeoutError(AccessApplicationError):
    """An Access call exceeded ``call_timeout``; the owned Access process was terminated."""


class AccessProcessDiedError(AccessApplicationError):
    """The Access process exited or was terminated while PyAccessKit was using it."""


# ------------------------------------------------------------------------------------------ cleanup
class CleanupError(PyAccessKitError):
    """One or more steps failed while closing a session (all remaining steps still ran).

    Attributes:
        errors: The exceptions raised by the failing cleanup steps.
    """

    def __init__(self, message: str, *, errors: tuple[BaseException, ...] = ()) -> None:
        super().__init__(message)
        self.errors = errors

    def __str__(self) -> str:
        return self.message + "".join(f"\n  - {type(e).__name__}: {e}" for e in self.errors)


class ComError(PyAccessKitError):
    """A COM error PyAccessKit has no more specific translation for (details are preserved)."""
