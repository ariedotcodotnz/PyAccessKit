# pyright: basic
"""Conversion of values between Python and COM VARIANTs.

Access date/times are *wall-clock* values without a time zone. pywin32, however, treats a naive ``datetime``
as local time and converts it to UTC when it builds a ``VT_DATE`` — shifting the stored value by the UTC
offset (an integration test caught a 13-hour shift). :func:`to_variant` therefore passes wall-clock values
as UTC-tagged datetimes, which pywin32 stores unchanged; :func:`normalize` reads the wall-clock back.
"""

from __future__ import annotations

from datetime import UTC, date, datetime, time
from typing import Any

import pywintypes

__all__ = ["normalize", "to_variant"]

_OLE_EPOCH = date(1899, 12, 30)


def to_variant(value: Any) -> Any:
    """Prepare a Python value for COM without time-zone shifts.

    * naive ``datetime`` → the same wall-clock time, tagged UTC;
    * aware ``datetime`` → converted to local wall-clock time first (Access has no time zones);
    * ``date`` → midnight; ``time`` → a time on the OLE epoch day (how Access stores times of day).
    """
    if isinstance(value, (bytearray, memoryview)):
        return bytes(value)
    if isinstance(value, datetime):
        if value.tzinfo is not None:
            value = value.astimezone().replace(tzinfo=None)
        return value.replace(tzinfo=UTC)
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day, tzinfo=UTC)
    if isinstance(value, time):
        return datetime(
            _OLE_EPOCH.year,
            _OLE_EPOCH.month,
            _OLE_EPOCH.day,
            value.hour,
            value.minute,
            value.second,
            tzinfo=UTC,
        )
    return value


def normalize(value: Any) -> Any:
    """Convert pywin32 return values to plain Python values.

    * ``pywintypes.datetime`` (VT_DATE) → naive :class:`datetime` with the stored wall-clock time;
    * byte arrays (OLE Object / binary fields arrive as ``memoryview``) → :class:`bytes`;
    * tuples (SAFEARRAYs) are normalised recursively;
    * everything else (``Decimal`` for currency/decimal, ``None`` for Null/Empty...) is returned unchanged.
    """
    if isinstance(value, memoryview):
        return value.tobytes()
    if isinstance(value, pywintypes.TimeType):
        return datetime(
            value.year,
            value.month,
            value.day,
            value.hour,
            value.minute,
            value.second,
            value.microsecond,
        )
    if isinstance(value, tuple):
        return tuple(normalize(item) for item in value)
    return value
