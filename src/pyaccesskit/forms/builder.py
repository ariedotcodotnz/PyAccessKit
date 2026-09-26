"""Fluent builder for :class:`FormSpec` objects.

The builder never talks to Access. It accumulates controls and produces a validated :class:`FormSpec`;
when obtained from ``db.forms.create(...)`` it also knows how to *save* itself, which builds the whole form
in one atomic step (nothing is created in Access until then)::

    with db.forms.create("frmCustomers", record_source="Customers") as form:
        form.textbox("CustomerName", label="Customer name")
        form.checkbox("IsActive")
        form.button("cmdClose", caption="Close", on_click="DoCmd.Close acForm, Me.Name")
    # saved here; if the block raises, nothing is created
"""

from __future__ import annotations

import warnings
from collections.abc import Callable, Sequence
from types import TracebackType
from typing import TYPE_CHECKING, Any, Generic, Literal, Self, TypeVar

from pyaccesskit.enums import RowSourceType, Section
from pyaccesskit.errors import PyAccessKitError
from pyaccesskit.forms.controls import (
    ButtonSpec,
    CheckBoxSpec,
    ComboBoxSpec,
    ControlSpec,
    LabelSpec,
    TextBoxSpec,
)
from pyaccesskit.forms.spec import FormSpec
from pyaccesskit.forms.vba import Vba
from pyaccesskit.schema._base import PropertyValue, build
from pyaccesskit.units import Length

if TYPE_CHECKING:
    from typing_extensions import TypedDict, Unpack

    class _ControlOptions(TypedDict, total=False):
        name: str | None
        section: Section
        at: tuple[Length, Length] | None
        width: Length | None
        height: Length | None
        visible: bool
        properties: dict[str, PropertyValue]


__all__ = ["FormBuilder"]

_R = TypeVar("_R")


def _vba(code: str | Vba | None) -> Vba | None:
    if code is None or isinstance(code, Vba):
        return code
    return Vba(code)


class FormBuilder(Generic[_R]):
    """Accumulates controls for a form and produces a :class:`FormSpec`.

    Args:
        name: Form name.
        on_save: Called with the finished spec by :meth:`save` (set by ``db.forms.create``); its return
            value is returned by :meth:`save`.
        **form_options: Any other :class:`FormSpec` field (``record_source``, ``caption``,
            ``default_view``...).
    """

    def __init__(
        self, name: str, *, on_save: Callable[[FormSpec], _R] | None = None, **form_options: Any
    ) -> None:
        self._name = name
        self._options: dict[str, Any] = dict(form_options)
        self._controls: list[ControlSpec] = []
        self._on_save = on_save
        self._state: Literal["open", "saved", "discarded"] = "open"
        # Validate the form-level options immediately so mistakes surface at the call site.
        build(FormSpec, f"form {name!r}", name=name, **self._options)

    # ------------------------------------------------------------------------------------ controls
    def add(self, control: ControlSpec) -> Self:
        """Append an already-built control spec."""
        self._check_open()
        self._controls.append(control)
        return self

    def textbox(
        self,
        field: str | None = None,
        *,
        label: str | Literal[False] | None = None,
        control_source: str | None = None,
        format: str | None = None,
        enabled: bool = True,
        locked: bool = False,
        after_update: str | Vba | None = None,
        **options: Unpack[_ControlOptions],
    ) -> Self:
        """Add a text box bound to ``field`` (or showing ``control_source="=..."``)."""
        return self.add(
            build(
                TextBoxSpec,
                f"text box {options.get('name') or field!r}",
                field=field,
                label=label,
                control_source=control_source,
                format=format,
                enabled=enabled,
                locked=locked,
                after_update=_vba(after_update),
                **options,
            )
        )

    def checkbox(
        self,
        field: str | None = None,
        *,
        label: str | Literal[False] | None = None,
        enabled: bool = True,
        locked: bool = False,
        after_update: str | Vba | None = None,
        **options: Unpack[_ControlOptions],
    ) -> Self:
        """Add a check box bound to ``field``."""
        return self.add(
            build(
                CheckBoxSpec,
                f"check box {options.get('name') or field!r}",
                field=field,
                label=label,
                enabled=enabled,
                locked=locked,
                after_update=_vba(after_update),
                **options,
            )
        )

    def combobox(
        self,
        field: str | None = None,
        *,
        row_source: str,
        row_source_type: RowSourceType = RowSourceType.TABLE_QUERY,
        bound_column: int = 1,
        column_count: int = 1,
        column_widths: Sequence[Length] | None = None,
        limit_to_list: bool = True,
        label: str | Literal[False] | None = None,
        enabled: bool = True,
        locked: bool = False,
        after_update: str | Vba | None = None,
        **options: Unpack[_ControlOptions],
    ) -> Self:
        """Add a combo box bound to ``field`` with rows from ``row_source``."""
        return self.add(
            build(
                ComboBoxSpec,
                f"combo box {options.get('name') or field!r}",
                field=field,
                row_source=row_source,
                row_source_type=row_source_type,
                bound_column=bound_column,
                column_count=column_count,
                column_widths=column_widths,
                limit_to_list=limit_to_list,
                label=label,
                enabled=enabled,
                locked=locked,
                after_update=_vba(after_update),
                **options,
            )
        )

    def label(self, caption: str, **options: Unpack[_ControlOptions]) -> Self:
        """Add a free-standing label."""
        return self.add(build(LabelSpec, f"label {caption!r}", caption=caption, **options))

    def button(
        self,
        name: str,
        *,
        caption: str,
        on_click: str | Vba | None = None,
        section: Section = Section.DETAIL,
        at: tuple[Length, Length] | None = None,
        width: Length | None = None,
        height: Length | None = None,
        visible: bool = True,
        properties: dict[str, PropertyValue] | None = None,
    ) -> Self:
        """Add a command button; ``on_click`` is the VBA body of its Click event procedure."""
        return self.add(
            build(
                ButtonSpec,
                f"button {name!r}",
                name=name,
                caption=caption,
                on_click=_vba(on_click),
                section=section,
                at=at,
                width=width,
                height=height,
                visible=visible,
                properties=properties or {},
            )
        )

    # ---------------------------------------------------------------------------- form-level code
    def on_load(self, code: str | Vba) -> Self:
        """Set the VBA body of the form's Load event."""
        self._check_open()
        self._options["on_load"] = _vba(code)
        return self

    def on_current(self, code: str | Vba) -> Self:
        """Set the VBA body of the form's Current event."""
        self._check_open()
        self._options["on_current"] = _vba(code)
        return self

    def module_code(self, code: str) -> Self:
        """Append extra VBA (helper procedures, module-level declarations) to the form's module."""
        self._check_open()
        existing = self._options.get("module_code") or ""
        self._options["module_code"] = f"{existing}\n{code}" if existing else code
        return self

    # ------------------------------------------------------------------------------------ results
    def to_spec(self) -> FormSpec:
        """Validate and return the finished :class:`FormSpec`."""
        return build(
            FormSpec,
            f"form {self._name!r}",
            name=self._name,
            controls=tuple(self._controls),
            **self._options,
        )

    def save(self) -> _R:
        """Build the form in Access (atomically) and return the result of ``on_save``.

        Raises:
            PyAccessKitError: If the builder has no save target or was already saved/discarded.
        """
        self._check_open()
        if self._on_save is None:
            raise PyAccessKitError(
                "this FormBuilder is standalone; use to_spec() or db.forms.create(...)"
            )
        spec = self.to_spec()
        result = self._on_save(spec)
        self._state = "saved"
        return result

    def discard(self) -> None:
        """Abandon the builder without creating anything."""
        self._state = "discarded"

    @property
    def is_saved(self) -> bool:
        """Whether :meth:`save` completed."""
        return self._state == "saved"

    def _check_open(self) -> None:
        if self._state != "open":
            raise PyAccessKitError(f"form builder {self._name!r} was already {self._state}")

    # ------------------------------------------------------------------------------ context manager
    def __enter__(self) -> Self:
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: TracebackType | None,
    ) -> None:
        if exc_type is not None:
            self.discard()
        elif self._state == "open":
            self.save()

    def __del__(self) -> None:
        if getattr(self, "_state", None) == "open" and getattr(self, "_on_save", None) is not None:
            warnings.warn(
                f"form {self._name!r} was never saved; call .save() or use 'with db.forms.create(...)'",
                ResourceWarning,
                stacklevel=2,
            )

    def __repr__(self) -> str:
        return f"<FormBuilder {self._name!r} controls={len(self._controls)} state={self._state}>"
