# pyright: basic
"""The "native defaults" profile: database properties Access writes to new databases but DAO does not.

Recorded in spike S3 (ADR 0001) by comparing ``NewCurrentDatabase`` with DAO ``CreateDatabase``. Applying
them to DAO-created files makes them behave exactly like Access-created ones (e.g. tabbed documents
instead of legacy overlapping windows). ``AccessVersion`` is omitted: Access maintains it itself.
"""

from __future__ import annotations

from typing import Any

from pyaccesskit._backends.dao.schema import _prop, _set_prop
from pyaccesskit._backends.dao.typemap import DB_BOOLEAN, DB_BYTE, DB_LONG

__all__ = ["NATIVE_DEFAULTS", "apply_native_defaults"]

NATIVE_DEFAULTS: tuple[tuple[str, int, Any], ...] = (
    ("ANSI Query Mode", DB_LONG, 0),
    ("CheckTruncatedNumFields", DB_LONG, 1),
    ("Clear Cache on Close", DB_LONG, 0),
    ("Default Zoom Level", DB_LONG, 100),
    ("NavPane Category", DB_LONG, 0),
    ("Never Cache", DB_LONG, 0),
    ("Option to enable Monaco SQL Editor", DB_LONG, 1),
    ("Picture Property Storage Format", DB_LONG, 0),
    ("Show Navigation Pane Search Bar", DB_LONG, 1),
    ("ShowDocumentTabs", DB_BOOLEAN, True),
    ("Themed Form Controls", DB_LONG, 1),
    ("Use Microsoft Access 2007 compatible cache", DB_LONG, 0),
    ("UseMDIMode", DB_BYTE, 0),
    ("WebDesignMode", DB_BYTE, 0),
)


def apply_native_defaults(db: Any) -> None:
    """Create any missing native-default property on a DAO ``Database`` (existing values are kept)."""
    for name, dao_type, value in NATIVE_DEFAULTS:
        if _prop(db, name) is None:
            _set_prop(db, name, dao_type, value)
