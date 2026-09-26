"""Access naming rules, name comparison, identifier quoting and name linting.

Hard rules (violations raise :class:`~pyaccesskit.errors.SpecError`) follow Microsoft's *Guidelines for
naming fields, controls, and objects*: at most 64 characters; no period, exclamation point, accent grave or
brackets; no leading space; no control characters.

Soft rules (violations emit :class:`~pyaccesskit.errors.AccessNameWarning`) flag names that are legal but
cause trouble later: reserved words, spaces/special characters, names that shadow common form properties.
"""

from __future__ import annotations

import re
import warnings

from pyaccesskit.errors import AccessNameWarning, SpecError
from pyaccesskit.schema._reserved_words import RESERVED_WORDS

__all__ = [
    "MAX_NAME_LENGTH",
    "check_name",
    "lint_name",
    "name_key",
    "names_equal",
    "quote_identifier",
    "warn_name",
]

MAX_NAME_LENGTH = 64
_FORBIDDEN_CHARACTERS = frozenset(".!`[]")
_PLAIN_IDENTIFIER = re.compile(r"^[A-Za-z][A-Za-z0-9_]*$")
_FORM_PROPERTY_NAMES = frozenset(
    name.casefold()
    for name in (
        "Name",
        "Caption",
        "Text",
        "Value",
        "Section",
        "Tag",
        "Visible",
        "Width",
        "Height",
        "Top",
        "Left",
    )
)


def check_name(name: object, *, what: str = "name") -> str:
    """Validate ``name`` against Access's hard naming rules and return it unchanged.

    Args:
        name: The candidate name.
        what: What is being named, used in error messages (``"table name"``...).

    Raises:
        SpecError: If the name is not a legal Access object name.
    """
    if not isinstance(name, str):
        raise SpecError(f"{what} must be a string, got {type(name).__name__}")
    problem = _hard_rule_violation(name)
    if problem is not None:
        raise SpecError(f"invalid {what} {name!r}: {problem}")
    return name


def _hard_rule_violation(name: str) -> str | None:
    if not name:
        return "names cannot be empty"
    if len(name) > MAX_NAME_LENGTH:
        return f"names are limited to {MAX_NAME_LENGTH} characters (this one has {len(name)})"
    if name[0] == " ":
        return "names cannot start with a space"
    bad = sorted({ch for ch in name if ch in _FORBIDDEN_CHARACTERS})
    if bad:
        return "names cannot contain " + ", ".join(repr(ch) for ch in bad)
    if any(ord(ch) < 32 for ch in name):
        return "names cannot contain control characters"
    return None


def lint_name(name: str, *, what: str = "name") -> list[str]:
    """Return human-readable warnings for a *legal* name that is likely to cause problems."""
    messages: list[str] = []
    key = name.casefold()
    if key in RESERVED_WORDS:
        messages.append(
            f"{what} {name!r} is an Access reserved word; it must be written as [{name}] in SQL and "
            f"expressions and can clash with built-in properties"
        )
    elif key in _FORM_PROPERTY_NAMES:
        messages.append(f"{what} {name!r} shadows a common form/control property of the same name")
    if name != name.rstrip():
        messages.append(f"{what} {name!r} ends with whitespace")
    elif not _PLAIN_IDENTIFIER.match(name):
        messages.append(
            f"{what} {name!r} contains spaces or special characters; it must be bracketed in SQL and VBA"
        )
    return messages


def warn_name(name: str, *, what: str = "name", stacklevel: int = 3) -> None:
    """Emit an :class:`AccessNameWarning` for every soft-rule problem found by :func:`lint_name`."""
    for message in lint_name(name, what=what):
        warnings.warn(message, AccessNameWarning, stacklevel=stacklevel)


def name_key(name: str) -> str:
    """Key for case-insensitive name comparison (Access object names are case-insensitive)."""
    return name.casefold()


def names_equal(left: str, right: str) -> bool:
    """Whether two Access names refer to the same object."""
    return left.casefold() == right.casefold()


def quote_identifier(name: str) -> str:
    """Bracket-quote a name for Access SQL and expressions: ``Order Details`` → ``[Order Details]``."""
    check_name(name)
    return f"[{name}]"
