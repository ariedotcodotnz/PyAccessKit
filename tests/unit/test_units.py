from __future__ import annotations

import pytest
from hypothesis import given
from hypothesis import strategies as st
from pydantic import BaseModel

from pyaccesskit.units import TWIPS_PER_INCH, Length, cm, inch, mm, pt, twips


def test_constructors_convert_to_whole_twips() -> None:
    assert inch(1).twips == TWIPS_PER_INCH
    assert pt(12).twips == 240
    assert cm(2.54).twips == 1440
    assert mm(25.4).twips == 1440
    assert twips(7).twips == 7


def test_arithmetic_and_comparison() -> None:
    assert cm(2) + mm(5) == Length(cm(2).twips + mm(5).twips)
    assert inch(2) - inch(1) == inch(1)
    assert inch(1) * 2 == inch(2)
    assert 2 * inch(1) == inch(2)
    assert inch(2) / 2 == inch(1)
    assert inch(2) / inch(1) == 2.0
    assert -inch(1) == Length(-1440)
    assert abs(Length(-5)) == Length(5)
    assert inch(1) > cm(2) > mm(10)
    assert sorted([cm(3), cm(1), cm(2)]) == [cm(1), cm(2), cm(3)]
    assert not Length(0)
    assert Length(1)


def test_length_is_immutable_and_hashable() -> None:
    length = cm(1)
    with pytest.raises(AttributeError):
        length._twips = 5  # type: ignore[misc]
    assert len({cm(1), cm(1), cm(2)}) == 2


def test_rejects_non_integer_twips() -> None:
    with pytest.raises(TypeError):
        Length(1.5)  # type: ignore[arg-type]
    with pytest.raises(TypeError):
        Length(True)
    with pytest.raises(ValueError, match="finite"):
        cm(float("nan"))


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("2cm", cm(2)),
        ("2 cm", cm(2)),
        ("10mm", mm(10)),
        ("1.5in", inch(1.5)),
        ('1"', inch(1)),
        ("12pt", pt(12)),
        ("300tw", twips(300)),
        ("300", twips(300)),
        (".5in", inch(0.5)),
        ("-1cm", cm(-1)),
        (720, twips(720)),
    ],
)
def test_parse(text: str | int, expected: Length) -> None:
    assert Length.parse(text) == expected


@pytest.mark.parametrize("text", ["", "abc", "2 furlongs", "1.5", "cm"])
def test_parse_rejects_garbage(text: str) -> None:
    with pytest.raises(ValueError):
        Length.parse(text)


def test_format() -> None:
    assert cm(2).format() == "2cm"
    assert inch(1).format("in") == "1in"
    assert Length(0).format() == "0cm"
    assert Length(1134).format("tw") == "1134tw"
    assert str(mm(15)) == "1.5cm"
    assert cm(1.5).format() == "1.5cm"
    assert inch(0.25).format("in") == "0.25in"


@given(
    st.integers(min_value=-10_000_000, max_value=10_000_000),
    st.sampled_from(["cm", "mm", "in", "pt", "tw"]),
)
def test_format_parse_round_trip_is_exact(value: int, unit: str) -> None:
    length = Length(value)
    assert Length.parse(length.format(unit)) == length


class _Model(BaseModel):
    width: Length


def test_pydantic_integration() -> None:
    model = _Model.model_validate({"width": "2cm"})
    assert model.width == cm(2)
    assert model.model_dump(mode="json") == {"width": "2cm"}
    assert _Model.model_validate_json(model.model_dump_json()) == model
    assert _Model(width=cm(1)).width == cm(1)
    assert _Model.model_validate({"width": 1440}).width == inch(1)
    with pytest.raises(ValueError):
        _Model.model_validate({"width": 1.5})
    schema = _Model.model_json_schema()
    assert "anyOf" in schema["properties"]["width"]
