from __future__ import annotations

import re
import warnings

import pytest

from pyaccesskit.errors import AccessNameWarning, SpecError
from pyaccesskit.schema.names import (
    check_name,
    lint_name,
    name_key,
    names_equal,
    quote_identifier,
    warn_name,
)


@pytest.mark.parametrize(
    "name", ["Customers", "Order Details", "tbl_2024", "Ünïcødé", "a" * 64, "x#y"]
)
def test_legal_names(name: str) -> None:
    assert check_name(name) == name


@pytest.mark.parametrize(
    ("name", "message"),
    [
        ("", "empty"),
        ("a" * 65, "64 characters"),
        (" Leading", "start with a space"),
        ("dot.name", "'.'"),
        ("bang!", "'!'"),
        ("grave`", "'`'"),
        ("[bracket]", "'['"),
        ("tab\tname", "control characters"),
    ],
)
def test_illegal_names(name: str, message: str) -> None:
    with pytest.raises(SpecError, match=re.escape(message)):
        check_name(name, what="table name")


def test_non_string_name() -> None:
    with pytest.raises(SpecError, match="must be a string"):
        check_name(42)


def test_lint_reserved_words_and_special_characters() -> None:
    assert any("reserved word" in m for m in lint_name("Date"))
    assert any("reserved word" in m for m in lint_name("name"))
    assert any("special characters" in m for m in lint_name("Order Details"))
    assert any("whitespace" in m for m in lint_name("Trailing "))
    assert any("shadows" in m for m in lint_name("Caption"))
    assert lint_name("CustomerName") == []


def test_warn_name_emits_warnings() -> None:
    with pytest.warns(AccessNameWarning, match="reserved word"):
        warn_name("Date", what="column name")
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        warn_name("CustomerID")


def test_case_insensitive_helpers() -> None:
    assert name_key("Customers") == name_key("CUSTOMERS")
    assert names_equal("Straße", "STRASSE")
    assert quote_identifier("Order Details") == "[Order Details]"
    with pytest.raises(SpecError):
        quote_identifier("bad]name")
