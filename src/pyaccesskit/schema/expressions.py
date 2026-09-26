# pyright: reportUnnecessaryIsInstance=false
# (dataclass __post_init__ and render_literal validate untyped input at runtime)
"""Access expression literals: rendering Python values and parsing them back.

Access stores properties such as ``DefaultValue`` as *expression text*. A classic mistake is setting a text
default to ``Unknown`` instead of ``"Unknown"`` (Access then looks for a field or function called Unknown).
PyAccessKit takes Python values and renders the right literal; raw expressions are passed with
:class:`Expr`::

    Column.text("Status", default="Unknown")          # stored as  "Unknown"
    Column.date_time("CreatedAt", default=Expr("Now()"))
    Column.currency("Total", default=0)                # stored as  0
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date, datetime, time
from decimal import Decimal, InvalidOperation

from pyaccesskit.errors import SpecError

__all__ = [
    "DefaultValue",
    "Expr",
    "parse_literal",
    "render_literal",
]


@dataclass(frozen=True)
class Expr:
    """A raw Access expression, passed through verbatim (e.g. ``Expr("Now()")``, ``Expr("[Qty]*2")``)."""

    expr: str

    def __post_init__(self) -> None:
        if not isinstance(self.expr, str) or not self.expr.strip():
            raise SpecError("Expr requires a non-empty expression string")

    def __str__(self) -> str:
        return self.expr


DefaultValue = Expr | bool | int | float | Decimal | datetime | date | time | str
"""Types accepted as a column default: a Python literal or an :class:`Expr`."""


def render_literal(value: DefaultValue) -> str:
    """Render a Python value as Access expression text.

    Raises:
        SpecError: For values Access cannot represent (NaN, time-zone aware datetimes...).
    """
    if isinstance(value, Expr):
        return value.expr
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if not math.isfinite(value):
            raise SpecError(f"cannot store non-finite number {value!r} in Access")
        return repr(value)
    if isinstance(value, Decimal):
        if not value.is_finite():
            raise SpecError(f"cannot store non-finite number {value!r} in Access")
        return format(value, "f")
    if isinstance(value, datetime):
        if value.tzinfo is not None:
            raise SpecError("Access date/times have no time zone; pass a naive datetime")
        if value.microsecond:
            raise SpecError("Access date/time literals have whole-second precision")
        return f"#{value.year:04d}-{value:%m-%d %H:%M:%S}#"  # %Y is unpadded on Linux
    if isinstance(value, date):
        return f"#{value.year:04d}-{value:%m-%d}#"
    if isinstance(value, time):
        if value.tzinfo is not None or value.microsecond:
            raise SpecError("Access time literals are naive and have whole-second precision")
        return f"#{value:%H:%M:%S}#"
    if isinstance(value, str):
        return '"' + value.replace('"', '""') + '"'
    raise SpecError(f"unsupported literal type {type(value).__name__}")  # pragma: no cover


_NUMBER = re.compile(r"^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$")
_BOOLEAN_WORDS = {"true": True, "yes": True, "on": True, "false": False, "no": False, "off": False}
_DATE_FORMATS = (
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y-%m-%d",
    "%m/%d/%Y %H:%M:%S",
    "%m/%d/%Y %I:%M:%S %p",
    "%m/%d/%Y %H:%M",
    "%m/%d/%Y",
)
_TIME_FORMATS = ("%H:%M:%S", "%H:%M", "%I:%M:%S %p", "%I:%M %p")


def _unquote(text: str) -> str | None:
    """Return the value of a single quoted string literal, or ``None`` if ``text`` is not exactly one."""
    if len(text) < 2 or text[0] not in "\"'" or text[-1] != text[0]:
        return None
    quote = text[0]
    body = text[1:-1]
    # Every quote inside the body must be doubled; otherwise this is an expression like "a" & "b".
    stripped = body.replace(quote * 2, "")
    if quote in stripped:
        return None
    return body.replace(quote * 2, quote)


def _parse_date_literal(body: str) -> datetime | date | time | None:
    body = body.strip()
    for fmt in _DATE_FORMATS:
        try:
            parsed = datetime.strptime(body, fmt)
        except ValueError:
            continue
        has_time = "%H" in fmt or "%I" in fmt
        return parsed if has_time else parsed.date()
    for fmt in _TIME_FORMATS:
        try:
            return datetime.strptime(body, fmt).time()
        except ValueError:
            continue
    return None


def parse_literal(text: str | None) -> DefaultValue | None:
    """Parse Access expression text into a Python literal, or an :class:`Expr` if it is not a literal.

    Returns ``None`` for empty text. Understands quoted strings, numbers, ``True``/``False``/``Yes``/``No``/
    ``On``/``Off`` and ``#date#`` literals (ISO or US ``m/d/yyyy`` order).
    """
    if text is None:
        return None
    stripped = text.strip()
    if not stripped:
        return None
    unquoted = _unquote(stripped)
    if unquoted is not None:
        return unquoted
    if _NUMBER.match(stripped):
        if re.search(r"[.eE]", stripped) is None:
            return int(stripped)
        if "e" in stripped.lower():
            return float(stripped)
        try:
            return Decimal(stripped)
        except InvalidOperation:  # pragma: no cover - guarded by the regex
            return Expr(stripped)
    boolean = _BOOLEAN_WORDS.get(stripped.lower())
    if boolean is not None:
        return boolean
    if len(stripped) >= 2 and stripped[0] == "#" and stripped[-1] == "#":
        parsed = _parse_date_literal(stripped[1:-1])
        if parsed is not None:
            return parsed
    return Expr(stripped)
