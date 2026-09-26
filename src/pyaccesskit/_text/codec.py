"""Encodings used by Access's ``SaveAsText``/``LoadFromText`` (verified in ADR 0001, spike S7).

* Forms, reports, queries and macros are exported as **UTF-16LE with a BOM** and must be imported the same
  way (UTF-8 input silently loses non-ASCII text).
* VBA modules are read and written in the **ANSI code page** (e.g. cp1252); a UTF-8 BOM would become literal
  ``ï»¿`` characters in the code.
* Class modules are recognised by four leading ``Attribute`` lines (Access's own export format); a VB6-style
  ``VERSION 1.0 CLASS`` header produces an *empty standard* module instead.

This module converts between those native byte formats and ordinary Python ``str`` (LF line endings),
which is how PyAccessKit stores text on disk (UTF-8, git-friendly).
"""

from __future__ import annotations

import codecs
import locale
import re

from pyaccesskit.enums import ModuleKind, ObjectKind
from pyaccesskit.errors import SpecError

__all__ = [
    "CLASS_MODULE_HEADER",
    "ansi_encoding",
    "decode_export",
    "encode_import",
    "module_import_text",
    "split_module_export",
]

CLASS_MODULE_HEADER = (
    "Attribute VB_GlobalNameSpace = False",
    "Attribute VB_Creatable = False",
    "Attribute VB_PredeclaredId = False",
    "Attribute VB_Exposed = False",
)
"""The lines Access writes at the top of an exported class module (and uses to recognise one)."""

_UNICODE_KINDS = frozenset({ObjectKind.FORM, ObjectKind.REPORT, ObjectKind.QUERY, ObjectKind.MACRO})
_VB6_HEADER = re.compile(
    r"\AVERSION\s+1\.0\s+CLASS\s*\nBEGIN\s*\n.*?\nEND\s*\n", re.IGNORECASE | re.DOTALL
)
_ATTRIBUTE_LINE = re.compile(r"^Attribute\s+VB_\w+\s*=.*$", re.IGNORECASE)


def ansi_encoding() -> str:
    """The Windows ANSI code page Access uses for module text (e.g. ``cp1252``)."""
    return locale.getencoding()


def _to_lf(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n")


def decode_export(kind: ObjectKind, raw: bytes) -> str:
    """Decode bytes written by ``SaveAsText`` into text with LF line endings."""
    if raw.startswith(codecs.BOM_UTF16_LE):
        text = raw.decode("utf-16")
    elif raw.startswith(codecs.BOM_UTF8):
        text = raw.decode("utf-8-sig")
    elif kind in _UNICODE_KINDS:
        text = raw.decode("utf-16-le")
    else:
        text = raw.decode(ansi_encoding())
    return _to_lf(text)


def encode_import(kind: ObjectKind, text: str) -> bytes:
    """Encode text for ``LoadFromText`` (CRLF line endings, UTF-16LE+BOM or ANSI depending on ``kind``).

    Raises:
        SpecError: If module text contains characters the ANSI code page cannot represent.
    """
    crlf = _to_lf(text).replace("\n", "\r\n")
    if kind in _UNICODE_KINDS:
        return codecs.BOM_UTF16_LE + crlf.encode("utf-16-le")
    encoding = ansi_encoding()
    try:
        return crlf.encode(encoding)
    except UnicodeEncodeError as exc:
        bad = crlf[exc.start : exc.end]
        raise SpecError(
            f"VBA module text contains {bad!r}, which the ANSI code page {encoding} cannot represent; "
            "VBA source is stored in the system code page (use ChrW(...) for other characters)"
        ) from exc


def split_module_export(text: str) -> tuple[ModuleKind, str]:
    """Split exported module text into its kind and the code (class header removed)."""
    lines = _to_lf(text).split("\n")
    header = [line.strip() for line in lines[: len(CLASS_MODULE_HEADER)]]
    if [h.casefold() for h in header] == [h.casefold() for h in CLASS_MODULE_HEADER]:
        return ModuleKind.CLASS, "\n".join(lines[len(CLASS_MODULE_HEADER) :])
    return ModuleKind.STANDARD, "\n".join(lines)


def module_import_text(code: str, kind: ModuleKind) -> str:
    """Prepare module code for ``LoadFromText``.

    * converts VB6 ``.cls`` files (``VERSION 1.0 CLASS`` header) and strips ``Attribute`` lines;
    * adds ``Option Compare Database`` if missing (as Access does for new modules);
    * prefixes class modules with :data:`CLASS_MODULE_HEADER`.
    """
    text = _to_lf(code).lstrip("﻿")
    text = _VB6_HEADER.sub("", text)
    body = [line for line in text.split("\n") if not _ATTRIBUTE_LINE.match(line.strip())]
    while body and not body[0].strip():
        body.pop(0)
    if not any(line.strip().casefold().startswith("option compare") for line in body):
        body.insert(0, "Option Compare Database")
    lines = list(CLASS_MODULE_HEADER) + body if kind is ModuleKind.CLASS else body
    result = "\n".join(lines)
    return result if result.endswith("\n") else result + "\n"
