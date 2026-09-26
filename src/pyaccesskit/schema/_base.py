"""Shared base model and helpers for specification models."""

from __future__ import annotations

from collections.abc import Callable, Sequence
from datetime import datetime
from decimal import Decimal
from typing import Annotated, Any, TypeVar

from pydantic import AfterValidator, BaseModel, ConfigDict, ValidationError

from pyaccesskit.errors import SpecError

PropertyValue = str | bool | int | float | Decimal | datetime | None
"""Values accepted in ``properties={...}`` escape-hatch dictionaries (raw Access/DAO properties)."""

_T = TypeVar("_T")


def _as_tuple(value: Sequence[Any]) -> tuple[Any, ...]:
    return tuple(value)


Items = Annotated[Sequence[_T], AfterValidator(_as_tuple)]
"""A list-like spec field: any sequence is accepted (lists included), and it is stored as a tuple.

Type checkers derive a model's ``__init__`` from its annotations, so a ``tuple[...]`` field would reject the
lists people naturally write; this keeps the input type friendly and the stored value immutable.
"""


class SpecModel(BaseModel):
    """Base class of every PyAccessKit spec: immutable, strict about unknown fields."""

    model_config = ConfigDict(frozen=True, extra="forbid", validate_default=True)


def spec_error(exc: ValidationError, what: str) -> SpecError:
    """Convert a Pydantic ``ValidationError`` into a :class:`SpecError` with readable problems."""
    problems: list[str] = []
    for error in exc.errors(include_url=False):
        location = ".".join(str(part) for part in error["loc"]) or what
        message = error["msg"]
        if message.startswith("Value error, "):
            message = message.removeprefix("Value error, ")
        problems.append(f"{location}: {message}")
    return SpecError(f"invalid {what}", problems=tuple(problems))


def build(factory: Callable[..., _T], what: str, /, **kwargs: Any) -> _T:
    """Call ``factory(**kwargs)``, converting validation failures into :class:`SpecError`."""
    try:
        return factory(**kwargs)
    except ValidationError as exc:
        raise spec_error(exc, what) from exc
