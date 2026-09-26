from __future__ import annotations

from datetime import UTC, date, datetime, time
from decimal import Decimal

import pytest
from hypothesis import given
from hypothesis import strategies as st

from pyaccesskit.errors import SpecError
from pyaccesskit.schema.expressions import Expr, parse_literal, render_literal


@pytest.mark.parametrize(
    ("value", "text"),
    [
        (True, "True"),
        (False, "False"),
        (42, "42"),
        (-7, "-7"),
        (1.5, "1.5"),
        (Decimal("12.50"), "12.50"),
        ("Unknown", '"Unknown"'),
        ('He said "hi"', '"He said ""hi"""'),
        ("", '""'),
        (date(2026, 1, 31), "#2026-01-31#"),
        (datetime(2026, 1, 31, 13, 45, 0), "#2026-01-31 13:45:00#"),
        (time(8, 30), "#08:30:00#"),
        (Expr("Now()"), "Now()"),
    ],
)
def test_render_literal(value: object, text: str) -> None:
    assert render_literal(value) == text  # type: ignore[arg-type]


@pytest.mark.parametrize(
    "value",
    [
        float("nan"),
        float("inf"),
        Decimal("NaN"),
        datetime(2026, 1, 1, tzinfo=UTC),
        datetime(2026, 1, 1, 0, 0, 0, 5),
    ],
)
def test_render_rejects_unrepresentable(value: object) -> None:
    with pytest.raises(SpecError):
        render_literal(value)  # type: ignore[arg-type]


@pytest.mark.parametrize(
    ("text", "value"),
    [
        ('"Unknown"', "Unknown"),
        ("'single'", "single"),
        ('"He said ""hi"""', 'He said "hi"'),
        ("42", 42),
        ("-3.25", Decimal("-3.25")),
        ("1E-05", 1e-05),
        ("True", True),
        ("no", False),
        ("Yes", True),
        ("#2026-01-31#", date(2026, 1, 31)),
        ("#1/31/2026#", date(2026, 1, 31)),
        ("#2026-01-31 13:45:00#", datetime(2026, 1, 31, 13, 45)),
        ("#08:30:00#", time(8, 30)),
        ("Now()", Expr("Now()")),
        ('"a" & "b"', Expr('"a" & "b"')),
        ("=Date()", Expr("=Date()")),
        ("", None),
        ("   ", None),
        (None, None),
    ],
)
def test_parse_literal(text: str | None, value: object) -> None:
    assert parse_literal(text) == value


def test_expr_requires_text() -> None:
    with pytest.raises(SpecError):
        Expr("  ")


@given(st.text(alphabet=st.characters(blacklist_categories=("Cs",)), max_size=40))
def test_string_round_trip(value: str) -> None:
    assert parse_literal(render_literal(value)) == value


@given(st.integers(min_value=-(2**40), max_value=2**40))
def test_int_round_trip(value: int) -> None:
    assert parse_literal(render_literal(value)) == value


@given(
    st.datetimes(min_value=datetime(100, 1, 1), max_value=datetime(9999, 12, 31)).map(
        lambda d: d.replace(microsecond=0)
    )
)
def test_datetime_round_trip(value: datetime) -> None:
    parsed = parse_literal(render_literal(value))
    assert parsed == value or (value.time() == time(0) and parsed == value)
