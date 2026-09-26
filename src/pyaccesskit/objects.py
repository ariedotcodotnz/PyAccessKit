"""Generic management of Access objects and ``SaveAsText``/``LoadFromText`` text I/O (``db.objects``)."""

from __future__ import annotations

import os
from pathlib import Path
from typing import TYPE_CHECKING

from pyaccesskit._ops import design as design_ops
from pyaccesskit.enums import ObjectKind
from pyaccesskit.errors import SpecError

if TYPE_CHECKING:
    from pyaccesskit._session.session import Session

__all__ = ["AccessObjects"]


class AccessObjects:
    """Forms, reports, macros and modules as Access sees them, including raw text import/export.

    Text is exchanged as ``str`` with LF line endings and stored on disk as UTF-8; PyAccessKit converts to and
    from the encodings Access requires (UTF-16 for forms/reports/queries/macros, ANSI for VBA modules).
    """

    def __init__(self, session: Session) -> None:
        self._session = session

    @staticmethod
    def _kind(kind: ObjectKind | str) -> ObjectKind:
        value = ObjectKind(kind)
        if value not in (*design_ops.DESIGN_KINDS, ObjectKind.QUERY):
            raise SpecError(f"{value.value} objects are not supported here")
        return value

    def names(self, kind: ObjectKind | str) -> list[str]:
        """Names of all objects of ``kind`` (read through DAO where possible)."""
        value = self._kind(kind)
        if value is ObjectKind.QUERY:
            return [q.name for q in self._session.schema().list_queries()]
        return self._session.schema().list_documents(value)

    def export_text(self, kind: ObjectKind | str, name: str) -> str:
        """``SaveAsText`` an object."""
        return design_ops.export_object(self._session.design(), self._kind(kind), name)

    def import_text(
        self, kind: ObjectKind | str, name: str, text: str, *, replace: bool = False
    ) -> None:
        """``LoadFromText`` an object (Access replaces an existing object only when ``replace=True``)."""
        self._session.check_writable(f"import {name!r}")
        design_ops.import_object(
            self._session.design(), self._kind(kind), name, text, replace=replace
        )

    def save_text(self, kind: ObjectKind | str, name: str, path: str | os.PathLike[str]) -> Path:
        """Export an object to a UTF-8 text file; returns the path."""
        target = Path(path)
        target.write_text(self.export_text(kind, name), encoding="utf-8", newline="\n")
        return target

    def load_text(
        self,
        kind: ObjectKind | str,
        name: str,
        path: str | os.PathLike[str],
        *,
        replace: bool = False,
    ) -> None:
        """Import an object from a UTF-8 text file."""
        self.import_text(kind, name, Path(path).read_text(encoding="utf-8"), replace=replace)

    def delete(self, kind: ObjectKind | str, name: str) -> None:
        """Delete a form, report, macro or module."""
        self._session.check_writable(f"delete {name!r}")
        self._session.design().delete_object(self._kind(kind), name)

    def rename(self, kind: ObjectKind | str, old: str, new: str) -> None:
        """Rename a form, report, macro or module."""
        self._session.check_writable(f"rename {old!r}")
        self._session.design().rename_object(self._kind(kind), old, new)
