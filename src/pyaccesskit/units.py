# pyright: reportUnnecessaryIsInstance=false
# (operators and parse() deliberately type-check arguments at runtime)
"""Physical lengths for form and report layout.

Access measures positions and sizes in **twips** (1/1440 inch, 1/20 point). PyAccessKit lets you work in
familiar units instead and converts exactly once, when talking to Access::

    from pyaccesskit.units import cm, mm, inch

    width = cm(6) + mm(5)       # Length(twips=3685)
    width.cm                    # 6.5 (approximately; lengths are whole twips)

A :class:`Length` is an immutable whole number of twips. In specs it serializes to a readable string such
as ``"6.5cm"`` and parses from strings (``"2cm"``, ``"10mm"``, ``"1.5in"``, ``"12pt"``, ``"300tw"``) or plain
integers (twips).
"""

from __future__ import annotations

import math
import re
from typing import TYPE_CHECKING, Any, SupportsFloat

from pydantic_core import core_schema

if TYPE_CHECKING:
    from pydantic import GetCoreSchemaHandler, GetJsonSchemaHandler
    from pydantic.json_schema import JsonSchemaValue

__all__ = [
    "TWIPS_PER_CM",
    "TWIPS_PER_INCH",
    "TWIPS_PER_MM",
    "TWIPS_PER_POINT",
    "Length",
    "cm",
    "inch",
    "mm",
    "pt",
    "twips",
]

TWIPS_PER_INCH = 1440
TWIPS_PER_POINT = 20
TWIPS_PER_CM = TWIPS_PER_INCH / 2.54
TWIPS_PER_MM = TWIPS_PER_CM / 10

_UNIT_FACTORS: dict[str, float] = {
    "tw": 1,
    "twip": 1,
    "twips": 1,
    "pt": TWIPS_PER_POINT,
    "point": TWIPS_PER_POINT,
    "points": TWIPS_PER_POINT,
    "mm": TWIPS_PER_MM,
    "cm": TWIPS_PER_CM,
    "in": TWIPS_PER_INCH,
    "inch": TWIPS_PER_INCH,
    "inches": TWIPS_PER_INCH,
    '"': TWIPS_PER_INCH,
}
_PATTERN = re.compile(r"^\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+))\s*([a-zA-Z\"]*)\s*$")


def _to_twips(value: SupportsFloat, factor: float) -> int:
    number = float(value)
    if not math.isfinite(number):
        raise ValueError(f"length must be finite, got {value!r}")
    return round(number * factor)


class Length:
    """An immutable length, stored as a whole number of twips.

    Construct lengths with the helper functions :func:`twips`, :func:`pt`, :func:`mm`, :func:`cm` and
    :func:`inch`, or parse text with :meth:`parse`.
    """

    __slots__ = ("_twips",)

    _twips: int

    def __init__(self, twips: int) -> None:
        if isinstance(twips, bool) or not isinstance(twips, int):
            raise TypeError(
                f"Length expects an integer number of twips, got {type(twips).__name__}"
            )
        object.__setattr__(self, "_twips", twips)

    # --------------------------------------------------------------------------------------- access
    @property
    def twips(self) -> int:
        """The length in twips (the unit Access uses)."""
        return self._twips

    @property
    def points(self) -> float:
        """The length in typographic points."""
        return self._twips / TWIPS_PER_POINT

    @property
    def mm(self) -> float:
        """The length in millimetres."""
        return self._twips / TWIPS_PER_MM

    @property
    def cm(self) -> float:
        """The length in centimetres."""
        return self._twips / TWIPS_PER_CM

    @property
    def inches(self) -> float:
        """The length in inches."""
        return self._twips / TWIPS_PER_INCH

    # ---------------------------------------------------------------------------------- conversion
    @classmethod
    def parse(cls, value: str | int | Length) -> Length:
        """Parse ``"2cm"``, ``"10 mm"``, ``"1.5in"``, ``"12pt"``, ``"300tw"`` or an integer number of twips.

        Raises:
            ValueError: If the text is not a recognised length.
        """
        if isinstance(value, Length):
            return value
        if isinstance(value, bool):
            raise TypeError("a boolean is not a length")
        if isinstance(value, int):
            return cls(value)
        match = _PATTERN.match(value)
        if match is None:
            raise ValueError(
                f"not a length: {value!r} (expected e.g. '2cm', '10mm', '1.5in', '12pt')"
            )
        number, unit = match.group(1), match.group(2).lower()
        if not unit:
            if "." in number:
                raise ValueError(
                    f"a unitless length must be a whole number of twips, got {value!r}"
                )
            return cls(int(number))
        factor = _UNIT_FACTORS.get(unit)
        if factor is None:
            raise ValueError(f"unknown length unit {unit!r} in {value!r}")
        return cls(_to_twips(float(number), factor))

    def format(self, unit: str = "cm") -> str:
        """Format in ``unit`` (``"cm"``, ``"mm"``, ``"in"``, ``"pt"`` or ``"tw"``).

        The result is the *shortest* decimal that parses back to exactly the same number of twips:
        ``cm(1.5).format() == "1.5cm"`` and ``Length.parse(x.format(unit)) == x`` for every length and unit.
        """
        key = unit.lower()
        factor = _UNIT_FACTORS.get(key)
        if factor is None:
            raise ValueError(f"unknown length unit {unit!r}")
        if factor == 1:
            return f"{self._twips}tw"
        value = self._twips / factor
        text = f"{value:.6f}"
        for decimals in range(7):
            text = f"{value:.{decimals}f}"
            if _to_twips(float(text), factor) == self._twips:
                break
        text = text.rstrip("0").rstrip(".") if "." in text else text
        return f"{'0' if text in ('', '-0') else text}{key}"

    # ---------------------------------------------------------------------------------- arithmetic
    def __add__(self, other: Length) -> Length:
        if not isinstance(other, Length):
            return NotImplemented
        return Length(self._twips + other._twips)

    def __sub__(self, other: Length) -> Length:
        if not isinstance(other, Length):
            return NotImplemented
        return Length(self._twips - other._twips)

    def __mul__(self, factor: float) -> Length:
        if isinstance(factor, bool) or not isinstance(factor, (int, float)):
            return NotImplemented
        return Length(_to_twips(self._twips, factor))

    __rmul__ = __mul__

    def __truediv__(self, other: float | Length) -> Any:
        if isinstance(other, Length):
            return self._twips / other._twips
        if isinstance(other, bool) or not isinstance(other, (int, float)):
            return NotImplemented
        return Length(_to_twips(self._twips, 1 / other))

    def __neg__(self) -> Length:
        return Length(-self._twips)

    def __pos__(self) -> Length:
        return self

    def __abs__(self) -> Length:
        return Length(abs(self._twips))

    def __bool__(self) -> bool:
        return self._twips != 0

    # ---------------------------------------------------------------------------------- comparison
    def __eq__(self, other: object) -> bool:
        if not isinstance(other, Length):
            return NotImplemented
        return self._twips == other._twips

    def __lt__(self, other: Length) -> bool:
        if not isinstance(other, Length):
            return NotImplemented
        return self._twips < other._twips

    def __le__(self, other: Length) -> bool:
        if not isinstance(other, Length):
            return NotImplemented
        return self._twips <= other._twips

    def __gt__(self, other: Length) -> bool:
        if not isinstance(other, Length):
            return NotImplemented
        return self._twips > other._twips

    def __ge__(self, other: Length) -> bool:
        if not isinstance(other, Length):
            return NotImplemented
        return self._twips >= other._twips

    def __hash__(self) -> int:
        return hash(("Length", self._twips))

    def __setattr__(self, name: str, value: object) -> None:
        raise AttributeError("Length is immutable")

    def __repr__(self) -> str:
        return f"Length(twips={self._twips})"

    def __str__(self) -> str:
        return self.format("cm")

    # ------------------------------------------------------------------------------------- pydantic
    @classmethod
    def __get_pydantic_core_schema__(
        cls, _source: Any, _handler: GetCoreSchemaHandler
    ) -> core_schema.CoreSchema:
        def validate(value: Any) -> Length:
            if isinstance(value, (Length, str, int)):
                return cls.parse(value)
            # ValueError (not TypeError) so Pydantic reports a normal ValidationError.
            raise ValueError(
                f"expected a Length, a string like '2cm', or an integer of twips; got {value!r}"
            )

        return core_schema.no_info_plain_validator_function(
            validate,
            serialization=core_schema.plain_serializer_function_ser_schema(
                lambda length: length.format("cm"), when_used="json"
            ),
        )

    @classmethod
    def __get_pydantic_json_schema__(
        cls, _schema: core_schema.CoreSchema, _handler: GetJsonSchemaHandler
    ) -> JsonSchemaValue:
        return {
            "anyOf": [
                {
                    "type": "string",
                    "pattern": r"^\s*[+-]?(\d+(\.\d*)?|\.\d+)\s*(tw|twips?|pt|points?|mm|cm|in|inch|inches)\s*$",
                },
                {"type": "integer", "description": "twips"},
            ],
            "description": "A length such as '2cm', '10mm', '1.5in', '12pt', or an integer number of twips.",
        }


def twips(value: int) -> Length:
    """A length of ``value`` twips (1/1440 inch)."""
    return Length(value)


def pt(value: float) -> Length:
    """A length of ``value`` typographic points (20 twips each)."""
    return Length(_to_twips(value, TWIPS_PER_POINT))


def mm(value: float) -> Length:
    """A length of ``value`` millimetres."""
    return Length(_to_twips(value, TWIPS_PER_MM))


def cm(value: float) -> Length:
    """A length of ``value`` centimetres."""
    return Length(_to_twips(value, TWIPS_PER_CM))


def inch(value: float) -> Length:
    """A length of ``value`` inches."""
    return Length(_to_twips(value, TWIPS_PER_INCH))
