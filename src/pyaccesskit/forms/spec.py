"""Form specifications."""

from __future__ import annotations

from typing import Self

from pydantic import Field, field_validator, model_validator

from pyaccesskit.enums import FormView, LayoutKind, ScrollBars, Section
from pyaccesskit.forms.controls import ButtonSpec, ControlSpec, LabelSpec
from pyaccesskit.forms.vba import EventBinding, Vba, build_module, is_vba_identifier
from pyaccesskit.schema._base import Items, PropertyValue, SpecModel
from pyaccesskit.schema.names import check_name
from pyaccesskit.units import Length

__all__ = ["FormSpec", "label_name_for"]


def label_name_for(control_name: str) -> str:
    """The name Access gives a control's attached label (``CustomerName`` → ``CustomerName_Label``)."""
    return f"{control_name}_Label"


class FormSpec(SpecModel):
    """A form built by PyAccessKit.

    Attributes:
        name: Form name.
        record_source: Table, query or SQL the form is bound to (``None`` for an unbound form).
        caption: Window caption.
        default_view: Single, continuous, datasheet or split form.
        layout: How controls without ``at=`` are placed (``AUTO`` picks stacked or tabular).
        header: Show the form header/footer sections; ``None`` = automatically when needed.
        width: Minimum form width (grown to fit the controls).
        controls: Controls in tab order.
        on_load: VBA run by the form's Load event.
        on_current: VBA run by the form's Current event.
        module_code: Extra VBA appended to the form's module (helper procedures...).
        option_explicit: Put ``Option Explicit`` at the top of the generated module.
        properties: Other Access form properties to set verbatim (escape hatch).
    """

    name: str
    record_source: str | None = None
    caption: str | None = None
    default_view: FormView = FormView.SINGLE
    layout: LayoutKind = LayoutKind.AUTO
    header: bool | None = None
    width: Length | None = None
    allow_additions: bool = True
    allow_edits: bool = True
    allow_deletions: bool = True
    data_entry: bool = False
    navigation_buttons: bool = True
    record_selectors: bool = True
    dividing_lines: bool = False
    scroll_bars: ScrollBars = ScrollBars.BOTH
    auto_center: bool = True
    pop_up: bool = False
    modal: bool = False
    controls: Items[ControlSpec] = ()
    on_load: Vba | None = None
    on_current: Vba | None = None
    module_code: str | None = None
    option_explicit: bool = True
    properties: dict[str, PropertyValue] = Field(default_factory=dict)

    @field_validator("name")
    @classmethod
    def _check_name(cls, value: str) -> str:
        return check_name(value, what="form name")

    @model_validator(mode="after")
    def _check_form(self) -> Self:
        seen: dict[str, str] = {}

        def claim(name: str, what: str) -> None:
            key = name.casefold()
            if key in seen:
                raise ValueError(f"{what} name {name!r} is already used by {seen[key]}")
            seen[key] = what

        for index, control in enumerate(self.resolved_controls()):
            name = control.name or ""
            claim(name, f"control #{index + 1}")
            if control.attached_label is not None:
                claim(label_name_for(name), f"the label of {name!r}")
        uses_sections = any(c.section is not Section.DETAIL for c in self.controls)
        if uses_sections and self.header is False:
            raise ValueError("controls are placed in the header/footer but header=False")
        for binding in self.event_bindings():
            if binding.object_name != "Form" and not is_vba_identifier(binding.object_name):
                raise ValueError(
                    f"control {binding.object_name!r} has an event procedure, so its name must be a plain "
                    "VBA identifier (letters, digits, underscores)"
                )
        return self

    # --------------------------------------------------------------------------------- derived data
    def resolved_controls(self) -> tuple[ControlSpec, ...]:
        """Controls with default names filled in (bound field name, else ``<Kind><n>``)."""
        resolved: list[ControlSpec] = []
        counters: dict[str, int] = {}
        for control in self.controls:
            if control.name is not None:
                resolved.append(control)
                continue
            field = control.bound_field
            if field is not None:
                name = field
            elif isinstance(control, ButtonSpec):
                raise ValueError(
                    "buttons need an explicit name (it is used by their event procedures)"
                )
            else:
                prefix = (
                    "Label"
                    if isinstance(control, LabelSpec)
                    else control.control_kind.value.capitalize()
                )
                counters[prefix] = counters.get(prefix, 0) + 1
                name = f"{prefix}{counters[prefix]}"
            resolved.append(control.model_copy(update={"name": name}))
        return tuple(resolved)

    def event_bindings(self) -> tuple[EventBinding, ...]:
        """Every event procedure the form needs, form events first."""
        bindings: list[EventBinding] = []
        if self.on_load is not None:
            bindings.append(EventBinding("Form", "OnLoad", "Load", self.on_load))
        if self.on_current is not None:
            bindings.append(EventBinding("Form", "OnCurrent", "Current", self.on_current))
        for control in self.resolved_controls():
            name = control.name or ""
            if isinstance(control, ButtonSpec):
                if control.on_click is not None:
                    bindings.append(EventBinding(name, "OnClick", "Click", control.on_click))
            elif not isinstance(control, LabelSpec) and control.after_update is not None:
                bindings.append(
                    EventBinding(name, "AfterUpdate", "AfterUpdate", control.after_update)
                )
        return tuple(bindings)

    @property
    def has_module(self) -> bool:
        """Whether the form needs a class module (event procedures or extra code)."""
        return bool(self.event_bindings()) or bool(self.module_code and self.module_code.strip())

    def module_text(self) -> str | None:
        """The complete VBA text of the form's module, or ``None`` if it needs none."""
        if not self.has_module:
            return None
        return build_module(
            self.event_bindings(), self.module_code, option_explicit=self.option_explicit
        )
