"""Encoding rules verified against real Access in spike S7 (docs/adr/0001-spike-findings.md)."""

from __future__ import annotations

import codecs

import pytest

from pyaccesskit._text import codec
from pyaccesskit.enums import ModuleKind, ObjectKind
from pyaccesskit.errors import SpecError


@pytest.mark.parametrize(
    "kind", [ObjectKind.FORM, ObjectKind.REPORT, ObjectKind.QUERY, ObjectKind.MACRO]
)
def test_unicode_kinds_round_trip_as_utf16(kind: ObjectKind) -> None:
    text = 'Version =21\nBegin Form\n    Caption ="Formulaire é漢"\nEnd\n'
    raw = codec.encode_import(kind, text)
    assert raw.startswith(codecs.BOM_UTF16_LE)
    assert b"\r\x00\n\x00" in raw
    assert codec.decode_export(kind, raw) == text


def test_modules_use_the_ansi_code_page() -> None:
    text = 'Option Compare Database\nPublic Function Hello() As String\n    Hello = "héllo"\nEnd Function\n'
    raw = codec.encode_import(ObjectKind.MODULE, text)
    assert not raw.startswith((codecs.BOM_UTF8, codecs.BOM_UTF16_LE))
    assert b"\r\n" in raw
    assert codec.decode_export(ObjectKind.MODULE, raw) == text


def test_modules_reject_unrepresentable_characters() -> None:
    if codec.ansi_encoding().lower() in ("utf-8", "cp65001"):
        pytest.skip("system code page is UTF-8")
    with pytest.raises(SpecError, match="ANSI code page"):
        codec.encode_import(ObjectKind.MODULE, 'x = "漢"')


def test_decode_handles_bom_variants() -> None:
    assert codec.decode_export(ObjectKind.MODULE, codecs.BOM_UTF8 + b"a\r\nb") == "a\nb"
    assert codec.decode_export(ObjectKind.FORM, "a\r\nb".encode("utf-16-le")) == "a\nb"


def test_class_module_header_round_trip() -> None:
    exported = "\n".join(
        [*codec.CLASS_MODULE_HEADER, "Option Compare Database", "Public Title As String", ""]
    )
    kind, code = codec.split_module_export(exported)
    assert kind is ModuleKind.CLASS
    assert code.startswith("Option Compare Database")
    prepared = codec.module_import_text(code, ModuleKind.CLASS)
    assert prepared.splitlines()[:4] == list(codec.CLASS_MODULE_HEADER)
    assert codec.split_module_export(prepared)[0] is ModuleKind.CLASS


def test_standard_module_preparation_adds_option_compare() -> None:
    prepared = codec.module_import_text(
        "Public Function Two()\n    Two = 2\nEnd Function", ModuleKind.STANDARD
    )
    assert prepared.splitlines()[0] == "Option Compare Database"
    assert codec.split_module_export(prepared)[0] is ModuleKind.STANDARD
    assert (
        codec.module_import_text("Option Compare Text\n", ModuleKind.STANDARD).count(
            "Option Compare"
        )
        == 1
    )


def test_vb6_class_files_are_converted() -> None:
    vb6 = (
        "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  'True\nEND\n"
        'Attribute VB_Name = "clsThing"\nAttribute VB_Exposed = False\n'
        "Option Explicit\nPublic Title As String\n"
    )
    prepared = codec.module_import_text(vb6, ModuleKind.CLASS)
    assert "VERSION" not in prepared
    assert "VB_Name" not in prepared
    assert prepared.splitlines()[:4] == list(codec.CLASS_MODULE_HEADER)
    assert "Public Title As String" in prepared


def test_utf8_bom_in_source_is_stripped() -> None:
    prepared = codec.module_import_text(
        "﻿Option Compare Database\nSub X()\nEnd Sub", ModuleKind.STANDARD
    )
    assert "﻿" not in prepared
