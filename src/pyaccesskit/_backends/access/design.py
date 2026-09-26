# pyright: basic
"""``Access.Application`` implementation of :class:`~pyaccesskit._backends.protocols.DesignBackend`.

Form building follows the flow verified in spike S6:

``CreateForm`` (an unsaved, auto-named form in Design view) → header/footer via
``RunCommand(acCmdFormHdrFtr)`` when needed → properties and section heights (``Section`` is an *indexed*
property, so it goes through the exact-argument gateway) → ``CreateControl`` for every control and attached
label → event properties + form module → ``DoCmd.Close(acSaveYes)`` → atomic swap with ``DoCmd.Rename``.

If anything fails before the save, the unsaved form is closed with ``acSaveNo`` and nothing is left behind.
When replacing, the old form is renamed to a hidden ``~pak_bak_…`` name first and restored on failure.
"""

from __future__ import annotations

import contextlib
import tempfile
import uuid
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pywintypes

from pyaccesskit._backends.protocols import ControlInfo
from pyaccesskit._com import constants as c
from pyaccesskit._com.gateway import Com, call, get, put
from pyaccesskit.enums import ControlKind, FormView, ObjectKind, RowSourceType, ScrollBars, Section
from pyaccesskit.errors import SpecError
from pyaccesskit.forms.controls import (
    ButtonSpec,
    CheckBoxSpec,
    ComboBoxSpec,
    LabelSpec,
    TextBoxSpec,
)
from pyaccesskit.forms.layout import ResolvedForm
from pyaccesskit.units import Length

__all__ = ["AccessDesignBackend"]

AC_CMD_FORM_HDR_FTR = 36
_AC_TYPES = {
    ObjectKind.FORM: c.AcObjectType.acForm,
    ObjectKind.REPORT: c.AcObjectType.acReport,
    ObjectKind.MACRO: c.AcObjectType.acMacro,
    ObjectKind.MODULE: c.AcObjectType.acModule,
    ObjectKind.QUERY: c.AcObjectType.acQuery,
}
_ALL_COLLECTIONS = {
    ObjectKind.FORM: ("CurrentProject", "AllForms"),
    ObjectKind.REPORT: ("CurrentProject", "AllReports"),
    ObjectKind.MACRO: ("CurrentProject", "AllMacros"),
    ObjectKind.MODULE: ("CurrentProject", "AllModules"),
    ObjectKind.QUERY: ("CurrentData", "AllQueries"),
}
_CONTROL_TYPES = {
    ControlKind.LABEL: c.AcControlType.acLabel,
    ControlKind.TEXTBOX: c.AcControlType.acTextBox,
    ControlKind.CHECKBOX: c.AcControlType.acCheckBox,
    ControlKind.COMBOBOX: c.AcControlType.acComboBox,
    ControlKind.BUTTON: c.AcControlType.acCommandButton,
}
_CONTROL_KINDS = {
    int(c.AcControlType.acLabel): ControlKind.LABEL,
    int(c.AcControlType.acTextBox): ControlKind.TEXTBOX,
    int(c.AcControlType.acCheckBox): ControlKind.CHECKBOX,
    int(c.AcControlType.acComboBox): ControlKind.COMBOBOX,
    int(c.AcControlType.acListBox): ControlKind.LISTBOX,
    int(c.AcControlType.acCommandButton): ControlKind.BUTTON,
    int(c.AcControlType.acOptionGroup): ControlKind.OPTION_GROUP,
    int(c.AcControlType.acOptionButton): ControlKind.OPTION_BUTTON,
    int(c.AcControlType.acToggleButton): ControlKind.TOGGLE_BUTTON,
    int(c.AcControlType.acSubform): ControlKind.SUBFORM,
    int(c.AcControlType.acImage): ControlKind.IMAGE,
    int(c.AcControlType.acLine): ControlKind.LINE,
    int(c.AcControlType.acRectangle): ControlKind.RECTANGLE,
    int(c.AcControlType.acTabCtl): ControlKind.TAB_CONTROL,
    int(c.AcControlType.acPage): ControlKind.PAGE,
    int(c.AcControlType.acAttachment): ControlKind.ATTACHMENT,
}
_SECTIONS = {Section.DETAIL: 0, Section.HEADER: 1, Section.FOOTER: 2}
_SECTION_OF = {value: key for key, value in _SECTIONS.items()}
_DEFAULT_VIEWS = {
    FormView.SINGLE: 0,
    FormView.CONTINUOUS: 1,
    FormView.DATASHEET: 2,
    FormView.SPLIT: 5,
}
_SCROLL_BARS = {
    ScrollBars.NEITHER: 0,
    ScrollBars.HORIZONTAL: 1,
    ScrollBars.VERTICAL: 2,
    ScrollBars.BOTH: 3,
}
_ROW_SOURCE_TYPES = {
    RowSourceType.TABLE_QUERY: "Table/Query",
    RowSourceType.VALUE_LIST: "Value List",
}


def _optional(obj: Any, name: str) -> Any:
    """Read a property that not every control has (``None`` when it is missing)."""
    try:
        return get(obj, name)
    except pywintypes.com_error:
        return None


class AccessDesignBackend:
    """Design operations through a live ``Access.Application`` with the database open."""

    def __init__(self, *, com: Callable[[], Com], app: Callable[[], Any]) -> None:
        self._com_getter = com
        self._app_getter = app
        self._temp: Path | None = None

    @property
    def _com(self) -> Com:
        return self._com_getter()

    def _app(self) -> Any:
        return self._app_getter()

    def _docmd(self) -> Any:
        return get(self._app(), "DoCmd")

    def _temp_file(self, suffix: str) -> Path:
        if self._temp is None:
            self._temp = Path(tempfile.mkdtemp(prefix="pyaccesskit-"))
        return self._temp / f"{uuid.uuid4().hex}{suffix}"

    def cleanup(self) -> None:
        """Delete the private temp folder used for text import/export."""
        if self._temp is not None:
            for path in self._temp.glob("*"):
                with contextlib.suppress(OSError):
                    path.unlink()
            with contextlib.suppress(OSError):
                self._temp.rmdir()
            self._temp = None

    # ------------------------------------------------------------------------------------ objects
    def list_objects(self, kind: ObjectKind) -> list[str]:
        owner, collection_name = _ALL_COLLECTIONS[kind]
        with self._com.op(f"list {kind.value}s"):
            collection = get(get(self._app(), owner), collection_name)
            return [
                str(get(get(collection, "Item", i), "Name"))
                for i in range(int(get(collection, "Count")))
            ]

    def delete_object(self, kind: ObjectKind, name: str) -> None:
        with self._com.op(f"delete {kind.value} {name!r}", kind=kind, name=name):
            call(self._docmd(), "DeleteObject", int(_AC_TYPES[kind]), name)

    def rename_object(self, kind: ObjectKind, old: str, new: str) -> None:
        with self._com.op(f"rename {kind.value} {old!r} to {new!r}", kind=kind, name=old):
            call(self._docmd(), "Rename", new, int(_AC_TYPES[kind]), old)

    def export_text(self, kind: ObjectKind, name: str) -> bytes:
        path = self._temp_file(".txt")
        try:
            with self._com.op(f"export {kind.value} {name!r} as text", kind=kind, name=name):
                call(self._app(), "SaveAsText", int(_AC_TYPES[kind]), name, str(path))
            return path.read_bytes()
        finally:
            with contextlib.suppress(OSError):
                path.unlink()

    def import_text(self, kind: ObjectKind, name: str, data: bytes) -> None:
        path = self._temp_file(".txt")
        path.write_bytes(data)
        try:
            with self._com.op(f"import {kind.value} {name!r} from text", kind=kind, name=name):
                call(self._app(), "LoadFromText", int(_AC_TYPES[kind]), name, str(path))
        finally:
            with contextlib.suppress(OSError):
                path.unlink()

    # -------------------------------------------------------------------------------------- forms
    def build_form(self, form: ResolvedForm, *, replace: bool) -> None:
        spec = form.spec
        app = self._app()
        docmd = self._docmd()
        with self._com.op(f"build form {spec.name!r}", kind=ObjectKind.FORM, name=spec.name):
            frm = call(app, "CreateForm")
            auto = str(get(frm, "Name"))
            saved = False
            try:
                self._populate(app, frm, auto, form)
                del frm
                call(docmd, "Close", int(c.AcObjectType.acForm), auto, int(c.AcCloseSave.acSaveYes))
                saved = True
            finally:
                if not saved:
                    with contextlib.suppress(pywintypes.com_error):
                        call(
                            docmd,
                            "Close",
                            int(c.AcObjectType.acForm),
                            auto,
                            int(c.AcCloseSave.acSaveNo),
                        )
        self._install(auto, spec.name, replace=replace)

    def _populate(self, app: Any, frm: Any, auto: str, form: ResolvedForm) -> None:
        spec = form.spec
        if form.has_header:
            call(get(app, "DoCmd"), "RunCommand", AC_CMD_FORM_HDR_FTR)
        if spec.record_source is not None:
            put(frm, "RecordSource", spec.record_source)
        if spec.caption is not None:
            put(frm, "Caption", spec.caption)
        for prop, value in (
            ("DefaultView", _DEFAULT_VIEWS[spec.default_view]),
            ("AllowAdditions", spec.allow_additions),
            ("AllowEdits", spec.allow_edits),
            ("AllowDeletions", spec.allow_deletions),
            ("DataEntry", spec.data_entry),
            ("NavigationButtons", spec.navigation_buttons),
            ("RecordSelectors", spec.record_selectors),
            ("DividingLines", spec.dividing_lines),
            ("ScrollBars", _SCROLL_BARS[spec.scroll_bars]),
            ("AutoCenter", spec.auto_center),
            ("PopUp", spec.pop_up),
            ("Modal", spec.modal),
            ("Width", form.width.twips),
        ):
            put(frm, prop, value)
        put(get(frm, "Section", _SECTIONS[Section.DETAIL]), "Height", form.detail_height.twips)
        if form.has_header:
            put(
                get(frm, "Section", _SECTIONS[Section.HEADER]),
                "Height",
                (form.header_height or Length(0)).twips,
            )
            put(
                get(frm, "Section", _SECTIONS[Section.FOOTER]),
                "Height",
                (form.footer_height or Length(0)).twips,
            )

        created: dict[str, Any] = {}
        for resolved in form.controls:
            control = resolved.spec
            kind = control.control_kind
            if kind not in _CONTROL_TYPES:
                raise SpecError(f"control kind {kind.value!r} is not supported by the form builder")
            rect = resolved.rect
            column = control.bound_field or ""
            ctl = call(
                app,
                "CreateControl",
                auto,
                int(_CONTROL_TYPES[kind]),
                _SECTIONS[resolved.section],
                "",
                column,
                rect.left.twips,
                rect.top.twips,
                rect.width.twips,
                rect.height.twips,
            )
            put(ctl, "Name", resolved.name)
            self._configure_control(ctl, control)
            created[resolved.name] = ctl
            label = resolved.label
            if label is not None:
                # Tabular layouts put labels in the header: those are free-standing, not attached.
                parent = resolved.name if label.section is resolved.section else ""
                lbl = call(
                    app,
                    "CreateControl",
                    auto,
                    int(c.AcControlType.acLabel),
                    _SECTIONS[label.section],
                    parent,
                    "",
                    label.rect.left.twips,
                    label.rect.top.twips,
                    label.rect.width.twips,
                    label.rect.height.twips,
                )
                put(lbl, "Name", label.name)
                put(lbl, "Caption", label.caption)

        for binding in form.events:
            target = frm if binding.object_name == "Form" else created[binding.object_name]
            put(target, binding.property_name, "[Event Procedure]")
        if form.module_text is not None:
            put(frm, "HasModule", True)
            module = get(frm, "Module")
            lines = int(get(module, "CountOfLines"))
            if lines:
                call(module, "DeleteLines", 1, lines)
            call(module, "AddFromString", form.module_text)
        for prop, value in spec.properties.items():
            put(get(get(frm, "Properties"), "Item", prop), "Value", value)

    @staticmethod
    def _configure_control(ctl: Any, control: Any) -> None:
        put(ctl, "Visible", control.visible)
        if isinstance(control, (LabelSpec, ButtonSpec)):
            put(ctl, "Caption", control.caption)
        elif isinstance(control, (TextBoxSpec, CheckBoxSpec, ComboBoxSpec)):
            put(ctl, "Enabled", control.enabled)
            put(ctl, "Locked", control.locked)
            if isinstance(control, TextBoxSpec):
                if control.control_source is not None:
                    put(ctl, "ControlSource", control.control_source)
                if control.format is not None:
                    put(ctl, "Format", control.format)
            if isinstance(control, ComboBoxSpec):
                put(ctl, "RowSourceType", _ROW_SOURCE_TYPES[control.row_source_type])
                put(ctl, "RowSource", control.row_source)
                put(ctl, "BoundColumn", control.bound_column)
                put(ctl, "ColumnCount", control.column_count)
                put(ctl, "LimitToList", control.limit_to_list)
                if control.column_widths is not None:
                    put(
                        ctl,
                        "ColumnWidths",
                        ";".join(str(width.twips) for width in control.column_widths),
                    )
        for prop, value in control.properties.items():
            put(get(get(ctl, "Properties"), "Item", prop), "Value", value)

    def _install(self, auto: str, name: str, *, replace: bool) -> None:
        form_type = int(c.AcObjectType.acForm)
        with self._com.op(f"install form {name!r}", kind=ObjectKind.FORM, name=name):
            docmd = self._docmd()
            backup: str | None = None
            try:
                if replace:
                    backup = f"~pak_bak_{uuid.uuid4().hex[:12]}"
                    call(docmd, "Rename", backup, form_type, name)
                call(docmd, "Rename", name, form_type, auto)
            except BaseException:
                with contextlib.suppress(pywintypes.com_error):
                    call(docmd, "DeleteObject", form_type, auto)
                if backup is not None:
                    with contextlib.suppress(pywintypes.com_error):
                        call(docmd, "Rename", name, form_type, backup)
                raise
            if backup is not None:
                call(docmd, "DeleteObject", form_type, backup)

    def _open_form(self, name: str, view: int) -> Any:
        call(
            self._docmd(),
            "OpenForm",
            name,
            view,
            "",
            "",
            int(c.AcFormOpenDataMode.acFormPropertySettings),
            int(c.AcWindowMode.acHidden),
        )
        return get(get(self._app(), "Forms"), "Item", name)

    def _close_form(self, name: str) -> None:
        with contextlib.suppress(pywintypes.com_error):
            call(
                self._docmd(),
                "Close",
                int(c.AcObjectType.acForm),
                name,
                int(c.AcCloseSave.acSaveNo),
            )

    def form_controls(self, name: str) -> list[ControlInfo]:
        with self._com.op(f"read controls of form {name!r}", kind=ObjectKind.FORM, name=name):
            frm = self._open_form(name, int(c.AcFormView.acDesign))
            try:
                controls = get(frm, "Controls")
                infos: list[ControlInfo] = []
                for index in range(int(get(controls, "Count"))):
                    ctl = get(controls, "Item", index)
                    parent_obj = _optional(ctl, "Parent")
                    parent = str(get(parent_obj, "Name")) if parent_obj is not None else None
                    if parent is not None and parent.casefold() == name.casefold():
                        parent = None
                    control_source = _optional(ctl, "ControlSource")
                    caption = _optional(ctl, "Caption")
                    infos.append(
                        ControlInfo(
                            name=str(get(ctl, "Name")),
                            kind=_CONTROL_KINDS.get(
                                int(get(ctl, "ControlType")), ControlKind.OTHER
                            ),
                            section=_SECTION_OF.get(int(get(ctl, "Section"))),
                            left=Length(int(get(ctl, "Left"))),
                            top=Length(int(get(ctl, "Top"))),
                            width=Length(int(get(ctl, "Width"))),
                            height=Length(int(get(ctl, "Height"))),
                            control_source=str(control_source) if control_source else None,
                            caption=str(caption) if caption is not None else None,
                            parent=parent,
                        )
                    )
                return infos
            finally:
                del frm
                self._close_form(name)

    def check_form_opens(self, name: str) -> None:
        with self._com.op(f"open form {name!r}", kind=ObjectKind.FORM, name=name):
            frm = self._open_form(name, int(c.AcFormView.acNormal))
            try:
                get(frm, "CurrentRecord")
            finally:
                del frm
                self._close_form(name)
