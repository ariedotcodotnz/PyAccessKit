"""Specifications of form controls.

Each control kind is its own spec class; together they form the discriminated union :data:`ControlSpec`
(``kind`` is the discriminator). Positions and sizes are :class:`~pyaccesskit.units.Length` values and are
optional: the layout engine places controls that have no explicit ``at=``.
"""

from __future__ import annotations

from typing import Annotated, Literal, Self

from pydantic import Field, field_validator, model_validator

from pyaccesskit.enums import ControlKind, RowSourceType, Section
from pyaccesskit.forms.vba import Vba
from pyaccesskit.schema._base import Items, PropertyValue, SpecModel
from pyaccesskit.schema.names import check_name
from pyaccesskit.units import Length

__all__ = [
    "ButtonSpec",
    "CheckBoxSpec",
    "ComboBoxSpec",
    "ControlBase",
    "ControlSpec",
    "LabelSpec",
    "TextBoxSpec",
]


class ControlBase(SpecModel):
    """Options shared by all controls.

    Attributes:
        name: Control name. Bound controls default to their field name.
        section: Form section holding the control.
        at: Explicit ``(left, top)`` position; ``None`` lets the layout engine place it.
        width: Explicit width (defaults depend on the control kind).
        height: Explicit height.
        visible: Whether the control is visible in Form view.
        properties: Other Access control properties to set verbatim (escape hatch).
    """

    name: str | None = None
    section: Section = Section.DETAIL
    at: tuple[Length, Length] | None = None
    width: Length | None = None
    height: Length | None = None
    visible: bool = True
    properties: dict[str, PropertyValue] = Field(default_factory=dict)

    @field_validator("name")
    @classmethod
    def _check_name(cls, value: str | None) -> str | None:
        return None if value is None else check_name(value, what="control name")

    @property
    def control_kind(self) -> ControlKind:
        """The kind of control."""
        return ControlKind(getattr(self, "kind"))  # noqa: B009 - defined by each subclass

    @property
    def bound_field(self) -> str | None:
        """The bound field name, if the control is bound to a field."""
        return getattr(self, "field", None)

    @property
    def attached_label(self) -> str | None:
        """The attached label caption, or ``None`` if the control has no attached label."""
        return None


class _BoundControl(ControlBase):
    field: str | None = None
    label: str | Literal[False] | None = None
    """Attached label: ``None`` = the field name, ``False`` = no label, or explicit text."""
    enabled: bool = True
    locked: bool = False
    after_update: Vba | None = None

    @field_validator("field")
    @classmethod
    def _check_field(cls, value: str | None) -> str | None:
        return None if value is None else check_name(value, what="field name")

    @property
    def attached_label(self) -> str | None:
        """Label caption (defaults to the bound field or control name)."""
        if self.label is False:
            return None
        if self.label is not None:
            return self.label
        return self.field or self.name


class LabelSpec(ControlBase):
    """A free-standing label."""

    kind: Literal["label"] = "label"
    caption: str


class TextBoxSpec(_BoundControl):
    """A text box bound to ``field`` or showing a calculated ``control_source`` (``"=[Qty]*[Price]"``)."""

    kind: Literal["textbox"] = "textbox"
    control_source: str | None = None
    format: str | None = None

    @model_validator(mode="after")
    def _check_source(self) -> Self:
        if self.field is not None and self.control_source is not None:
            raise ValueError("give either field= or control_source=, not both")
        if self.control_source is not None and not self.control_source.startswith("="):
            raise ValueError("control_source expressions start with '=' (e.g. '=[Qty]*[Price]')")
        return self


class CheckBoxSpec(_BoundControl):
    """A check box (typically bound to a Yes/No field)."""

    kind: Literal["checkbox"] = "checkbox"


class ComboBoxSpec(_BoundControl):
    """A combo box whose rows come from a table/query/SQL statement or a value list."""

    kind: Literal["combobox"] = "combobox"
    row_source: str
    row_source_type: RowSourceType = RowSourceType.TABLE_QUERY
    bound_column: int = Field(default=1, ge=0)
    column_count: int = Field(default=1, ge=1, le=255)
    column_widths: Items[Length] | None = None
    limit_to_list: bool = True

    @model_validator(mode="after")
    def _check_columns(self) -> Self:
        if self.bound_column > self.column_count:
            raise ValueError("bound_column cannot exceed column_count")
        if self.column_widths is not None and len(self.column_widths) > self.column_count:
            raise ValueError("more column_widths than column_count")
        return self


class ButtonSpec(ControlBase):
    """A command button, optionally running VBA when clicked."""

    kind: Literal["button"] = "button"
    caption: str
    on_click: Vba | None = None


ControlSpec = Annotated[
    LabelSpec | TextBoxSpec | CheckBoxSpec | ComboBoxSpec | ButtonSpec,
    Field(discriminator="kind"),
]
"""Any control spec (a Pydantic discriminated union on ``kind``)."""
