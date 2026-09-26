"""Behaviour every DesignBackend must share: modules, raw text objects, forms."""

from __future__ import annotations

import pytest

from pyaccesskit._backends.protocols import DesignBackend
from pyaccesskit._ops import design
from pyaccesskit._ops import schema as ops
from pyaccesskit.enums import ControlKind, ModuleKind, ObjectKind
from pyaccesskit.errors import ObjectExistsError, ObjectNotFoundError, SpecError
from pyaccesskit.forms import FormBuilder, FormSpec
from pyaccesskit.schema import Column, TableSpec

from .conftest import Backends

pytestmark = pytest.mark.filterwarnings("ignore::pyaccesskit.errors.AccessNameWarning")


@pytest.fixture
def backends(backends: Backends) -> Backends:
    if backends.design is None:
        pytest.skip(f"the {backends.name} backend has no design features")
    return backends


def dsg(backends: Backends) -> DesignBackend:
    assert backends.design is not None
    return backends.design


STANDARD = 'Option Compare Database\n\nPublic Function Greeting() As String\n    Greeting = "héllo"\nEnd Function\n'
CLASS = "Option Compare Database\n\nPublic Title As String\n"


def customers(backends: Backends) -> None:
    ops.create_table(
        backends.schema,
        TableSpec(
            name="Customers",
            columns=[
                Column.autonumber("CustomerID", primary_key=True),
                Column.text("CustomerName"),
                Column.yes_no("IsActive"),
            ],
        ),
    )


def test_standard_and_class_modules(backends: Backends) -> None:
    design.create_module(dsg(backends), "modUtils", STANDARD)
    design.create_module(dsg(backends), "clsThing", CLASS, ModuleKind.CLASS)
    kind, code = design.read_module(dsg(backends), "modutils")
    assert kind is ModuleKind.STANDARD
    assert 'Greeting = "héllo"' in code
    kind, code = design.read_module(dsg(backends), "clsThing")
    assert kind is ModuleKind.CLASS
    assert "Public Title As String" in code
    assert "Attribute" not in code
    assert sorted(dsg(backends).list_objects(ObjectKind.MODULE)) == ["clsThing", "modUtils"]

    with pytest.raises(ObjectExistsError):
        design.create_module(dsg(backends), "MODUTILS", STANDARD)
    design.create_module(dsg(backends), "modUtils", STANDARD.replace("héllo", "hi"), replace=True)
    assert 'Greeting = "hi"' in design.read_module(dsg(backends), "modUtils")[1]
    dsg(backends).delete_object(ObjectKind.MODULE, "modUtils")
    with pytest.raises(ObjectNotFoundError):
        design.read_module(dsg(backends), "modUtils")


def test_forms_build_replace_and_introspect(backends: Backends) -> None:
    customers(backends)
    builder = FormBuilder[None]("frmCustomers", record_source="Customers", caption="Customers")
    builder.textbox("CustomerName", label="Customer name")
    builder.checkbox("IsActive")
    builder.button("cmdClose", caption="Close", on_click="DoCmd.Close acForm, Me.Name")
    spec = builder.to_spec()

    resolved = design.build_form(backends.schema, dsg(backends), spec)
    assert dsg(backends).list_objects(ObjectKind.FORM) == ["frmCustomers"]
    controls = {c.name: c for c in dsg(backends).form_controls("frmCustomers")}
    assert controls["CustomerName"].kind is ControlKind.TEXTBOX
    assert controls["CustomerName"].control_source == "CustomerName"
    assert controls["CustomerName_Label"].caption == "Customer name"
    assert controls["CustomerName_Label"].parent == "CustomerName"
    assert controls["cmdClose"].kind is ControlKind.BUTTON
    by_name = {c.name: c for c in resolved.controls}
    assert (controls["CustomerName"].left, controls["CustomerName"].top) == (
        by_name["CustomerName"].rect.left,
        by_name["CustomerName"].rect.top,
    )
    dsg(backends).check_form_opens("frmCustomers")

    with pytest.raises(ObjectExistsError, match="replace=True"):
        design.build_form(backends.schema, dsg(backends), spec)
    design.build_form(
        backends.schema,
        dsg(backends),
        spec.model_copy(update={"caption": "Clients"}),
        replace=True,
    )
    assert dsg(backends).list_objects(ObjectKind.FORM) == ["frmCustomers"]


def test_form_validation_against_schema(backends: Backends) -> None:
    customers(backends)
    with pytest.raises(ObjectNotFoundError, match="neither a table nor a query"):
        design.build_form(backends.schema, dsg(backends), FormSpec(name="f", record_source="Nope"))
    bad = FormBuilder[None]("f2", record_source="Customers").textbox("Missing").to_spec()
    with pytest.raises(SpecError, match="not a column"):
        design.build_form(backends.schema, dsg(backends), bad)
    assert dsg(backends).list_objects(ObjectKind.FORM) == []


def test_text_export_import_round_trip(backends: Backends) -> None:
    design.create_module(dsg(backends), "modA", STANDARD)
    text = design.export_object(dsg(backends), ObjectKind.MODULE, "modA")
    design.import_object(dsg(backends), ObjectKind.MODULE, "modB", text)
    assert design.read_module(dsg(backends), "modB") == design.read_module(dsg(backends), "modA")
    with pytest.raises(ObjectExistsError):
        design.import_object(dsg(backends), ObjectKind.MODULE, "modB", text)
    with pytest.raises(SpecError):
        design.export_object(dsg(backends), ObjectKind.TABLE, "x")
