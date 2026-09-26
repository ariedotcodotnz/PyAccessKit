from __future__ import annotations

import warnings
from typing import Any

import pytest
from pydantic import ValidationError

from pyaccesskit.enums import FormView, LayoutKind, Section
from pyaccesskit.errors import PyAccessKitError, SpecError
from pyaccesskit.forms import (
    ButtonSpec,
    FormBuilder,
    FormSpec,
    LabelSpec,
    TextBoxSpec,
    Vba,
    layout_form,
)
from pyaccesskit.forms.layout import DEFAULT_METRICS, MAX_FORM_EXTENT
from pyaccesskit.forms.vba import EventBinding, build_module, event_procedure
from pyaccesskit.units import cm, inch


def customer_form(**options: Any) -> FormSpec:
    builder = FormBuilder[None](
        "frmCustomers", record_source="Customers", caption="Customers", **options
    )
    builder.textbox("CustomerName", label="Customer name", width=cm(8))
    builder.textbox("Email")
    builder.checkbox("IsActive")
    builder.combobox("Region", row_source="SELECT RegionName FROM Regions")
    builder.button("cmdClose", caption="Close", on_click="DoCmd.Close acForm, Me.Name")
    return builder.to_spec()


# --------------------------------------------------------------------------------------------- vba
def test_vba_body_is_dedented() -> None:
    body = Vba("""
        If IsNull(Me.Email) Then
            MsgBox "Email required"
        End If
    """)
    assert body.lines() == ["If IsNull(Me.Email) Then", '    MsgBox "Email required"', "End If"]


def test_vba_rejects_full_procedures() -> None:
    with pytest.raises(SpecError, match="body"):
        Vba("Private Sub X()\nEnd Sub")
    with pytest.raises(SpecError):
        Vba("   ")


def test_event_procedure_and_module() -> None:
    binding = EventBinding("cmdSave", "OnClick", "Click", Vba("DoCmd.Save"))
    assert (
        event_procedure(binding) == "Private Sub cmdSave_Click()\r\n    DoCmd.Save\r\nEnd Sub\r\n"
    )
    module = build_module([binding], "Private Function Helper() As Long\nEnd Function")
    assert module.startswith("Option Compare Database\r\nOption Explicit\r\n")
    assert "Private Sub cmdSave_Click()" in module
    assert module.rstrip().endswith("End Function")
    assert "\n" not in module.replace("\r\n", "")
    assert "Option Explicit" not in build_module([binding], option_explicit=False)


# -------------------------------------------------------------------------------------------- spec
def test_default_names_and_labels() -> None:
    spec = customer_form()
    names = [c.name for c in spec.resolved_controls()]
    assert names == ["CustomerName", "Email", "IsActive", "Region", "cmdClose"]
    assert spec.resolved_controls()[1].attached_label == "Email"
    assert spec.resolved_controls()[0].attached_label == "Customer name"


def test_unbound_controls_get_numbered_names() -> None:
    spec = FormSpec(
        name="f",
        controls=(
            TextBoxSpec(control_source="=Now()"),
            TextBoxSpec(control_source="=Date()"),
            LabelSpec(caption="Hi"),
        ),
    )
    assert [c.name for c in spec.resolved_controls()] == ["Textbox1", "Textbox2", "Label1"]


def test_duplicate_names_are_rejected() -> None:
    with pytest.raises(ValidationError, match="already used"):
        FormSpec(name="f", controls=(TextBoxSpec(field="A"), TextBoxSpec(field="a")))
    with pytest.raises(ValidationError, match="already used"):
        FormSpec(
            name="f", controls=(TextBoxSpec(field="A"), LabelSpec(name="A_Label", caption="x"))
        )


def test_event_controls_need_vba_identifiers() -> None:
    with pytest.raises(ValidationError, match="VBA identifier"):
        FormSpec(
            name="f",
            controls=(ButtonSpec(name="Save Button", caption="Save", on_click=Vba("x = 1")),),
        )


def test_buttons_need_names() -> None:
    with pytest.raises(ValidationError, match="explicit name"):
        FormSpec(name="f", controls=(ButtonSpec(caption="Go"),))


def test_textbox_source_rules() -> None:
    with pytest.raises(ValidationError, match="either field"):
        TextBoxSpec(field="A", control_source="=1")
    with pytest.raises(ValidationError, match="start with '='"):
        TextBoxSpec(control_source="[A]*2")


def test_module_text_contains_all_events() -> None:
    spec = customer_form()
    spec = spec.model_copy(update={"on_load": Vba('Me.Caption = "Hi"')})
    text = spec.module_text()
    assert text is not None
    assert text.index("Form_Load") < text.index("cmdClose_Click")
    assert FormSpec(name="plain").module_text() is None


def test_form_spec_json_round_trip() -> None:
    spec = customer_form()
    assert FormSpec.model_validate_json(spec.model_dump_json()) == spec


# ------------------------------------------------------------------------------------------ builder
def test_builder_context_manager_saves_once() -> None:
    saved: list[FormSpec] = []
    with FormBuilder("frm", record_source="T", on_save=saved.append) as form:
        form.textbox("A")
    assert len(saved) == 1
    assert form.is_saved
    with pytest.raises(PyAccessKitError, match="already saved"):
        form.textbox("B")


def test_builder_discards_on_exception() -> None:
    saved: list[FormSpec] = []
    with pytest.raises(RuntimeError), FormBuilder("frm", on_save=saved.append) as form:  # noqa: PT012
        form.textbox("A")
        raise RuntimeError("boom")
    assert saved == []
    assert not form.is_saved


def test_builder_validates_eagerly() -> None:
    with pytest.raises(SpecError):
        FormBuilder("bad.name")
    builder = FormBuilder[None]("f")
    with pytest.raises(SpecError):
        builder.textbox("A", control_source="=1")
    with pytest.raises(PyAccessKitError, match="standalone"):
        builder.save()
    builder.discard()


def test_unsaved_builder_warns() -> None:
    builder = FormBuilder("frm", on_save=lambda spec: spec)
    builder.textbox("A")
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        del builder
        import gc

        gc.collect()
    assert any(issubclass(w.category, ResourceWarning) for w in caught)


# ------------------------------------------------------------------------------------------- layout
def test_stacked_layout_places_labels_left_of_controls() -> None:
    resolved = layout_form(customer_form())
    m = DEFAULT_METRICS
    assert not resolved.has_header
    tops = [c.rect.top for c in resolved.controls]
    assert tops == sorted(tops)
    for control in resolved.controls:
        if control.label is not None:
            assert control.label.rect.right <= control.rect.left
            assert control.label.rect.left == m.margin
            assert control.label.name == f"{control.name}_Label"
    rects = [c.rect for c in resolved.controls] + [
        c.label.rect for c in resolved.controls if c.label
    ]
    for i, a in enumerate(rects):
        for b in rects[i + 1 :]:
            assert not a.overlaps(b)
    assert resolved.width >= max(r.right for r in rects)
    assert resolved.detail_height >= max(r.bottom for r in rects)
    assert resolved.controls[0].rect.width == cm(8)
    assert resolved.module_text is not None
    assert [e.procedure_name for e in resolved.events] == ["cmdClose_Click"]


def test_tabular_layout_for_continuous_forms() -> None:
    resolved = layout_form(customer_form(default_view=FormView.CONTINUOUS))
    assert resolved.has_header
    detail = [c for c in resolved.controls if c.section is Section.DETAIL]
    lefts = [c.rect.left for c in detail]
    assert lefts == sorted(lefts)
    assert all(c.rect.top == detail[0].rect.top for c in detail)
    labelled = [c for c in detail if c.label is not None]
    assert all(c.label is not None and c.label.section is Section.HEADER for c in labelled)
    checkbox = next(c for c in detail if c.name == "IsActive")
    assert checkbox.label is not None
    assert checkbox.label.rect.width >= DEFAULT_METRICS.tabular_column_width
    assert checkbox.label.rect.left <= checkbox.rect.left < checkbox.label.rect.right


def test_explicit_positions_are_kept() -> None:
    spec = FormSpec(name="f", controls=(TextBoxSpec(field="A", at=(cm(5), cm(3))),))
    control = layout_form(spec).controls[0]
    assert (control.rect.left, control.rect.top) == (cm(5), cm(3))
    assert control.label is not None
    assert control.label.rect.right < control.rect.left


def test_explicit_position_without_room_for_label() -> None:
    spec = FormSpec(name="f", controls=(TextBoxSpec(field="A", at=(cm(1), cm(1))),))
    with pytest.raises(SpecError, match="no room"):
        layout_form(spec)
    spec = FormSpec(name="f", controls=(TextBoxSpec(field="A", label=False, at=(cm(1), cm(1))),))
    assert layout_form(spec).controls[0].label is None


def test_layout_none_requires_positions() -> None:
    spec = FormSpec(name="f", layout=LayoutKind.NONE, controls=(TextBoxSpec(field="A"),))
    with pytest.raises(SpecError, match="requires at="):
        layout_form(spec)


def test_form_size_limit() -> None:
    spec = FormSpec(name="f", controls=(TextBoxSpec(field="A", label=False, width=inch(23)),))
    with pytest.raises(SpecError, match="exceeds Access's limit"):
        layout_form(spec)
    assert inch(22) == MAX_FORM_EXTENT


def test_header_controls_enable_the_header() -> None:
    spec = FormSpec(name="f", controls=(LabelSpec(caption="Title", section=Section.HEADER),))
    assert layout_form(spec).has_header
    with pytest.raises(ValidationError, match="header=False"):
        FormSpec(
            name="f", header=False, controls=(LabelSpec(caption="Title", section=Section.HEADER),)
        )
