"""Tier 4: forms, modules and text import/export in real Access."""

from __future__ import annotations

from pathlib import Path

import pytest

from pyaccesskit import (
    AccessDatabase,
    Column,
    ControlKind,
    FormView,
    ModuleKind,
    ObjectExistsError,
    Section,
    SpecError,
    Vba,
    cm,
)
from pyaccesskit.errors import PyAccessKitError

pytestmark = [
    pytest.mark.integration,
    pytest.mark.access,
    pytest.mark.filterwarnings("ignore::pyaccesskit.errors.AccessNameWarning"),
]


def _schema(db: AccessDatabase) -> None:
    db.tables.create("Regions", columns=[Column.text("RegionName", length=50, primary_key=True)])
    db.execute("INSERT INTO Regions (RegionName) VALUES ('North')")
    db.tables.create(
        "Customers",
        columns=[
            Column.autonumber("CustomerID", primary_key=True),
            Column.text("CustomerName", length=100),
            Column.text("Email"),
            Column.yes_no("IsActive", default=True),
            Column.text("Region", length=50),
        ],
    )
    db.execute(
        "INSERT INTO Customers (CustomerName, IsActive, Region) VALUES ('Ann', True, 'North')"
    )


def test_single_form_builds_opens_and_round_trips(tmp_path: Path) -> None:
    path = tmp_path / "forms.accdb"
    with AccessDatabase.create(path, engine="access") as db:
        _schema(db)
        with db.forms.create(
            "frmCustomers", record_source="Customers", caption="Customers"
        ) as form:
            form.textbox("CustomerName", label="Customer name", width=cm(8))
            form.textbox("Email")
            form.checkbox("IsActive", label="Active")
            form.combobox(
                "Region", row_source="SELECT RegionName FROM Regions ORDER BY RegionName;"
            )
            form.textbox(control_source="=Now()", name="txtNow", label=False, format="Short Date")
            form.button("cmdClose", caption="Close", on_click=Vba("DoCmd.Close acForm, Me.Name"))
            form.on_load('Me.Caption = "Customers (" & Me.Recordset.RecordCount & ")"')
        db.forms["frmCustomers"].check_opens()
    with AccessDatabase.open(path, engine="access") as db:
        assert db.forms.names() == ["frmCustomers"]
        controls = {c.name: c for c in db.forms["frmCustomers"].controls()}
        assert controls["CustomerName"].kind is ControlKind.TEXTBOX
        assert controls["CustomerName"].control_source == "CustomerName"
        assert controls["CustomerName"].width == cm(8)
        assert controls["CustomerName_Label"].caption == "Customer name"
        assert controls["CustomerName_Label"].parent == "CustomerName"
        assert controls["IsActive"].kind is ControlKind.CHECKBOX
        assert controls["Region"].kind is ControlKind.COMBOBOX
        assert controls["txtNow"].control_source == "=Now()"
        assert controls["cmdClose"].kind is ControlKind.BUTTON
        text = db.forms["frmCustomers"].export_text()
        assert "Private Sub cmdClose_Click()" in text
        assert "Private Sub Form_Load()" in text


def test_continuous_form_uses_a_tabular_layout(tmp_path: Path) -> None:
    with AccessDatabase.create(tmp_path / "list.accdb", engine="access") as db:
        _schema(db)
        with db.forms.create(
            "frmList", record_source="Customers", default_view=FormView.CONTINUOUS
        ) as form:
            form.textbox("CustomerName")
            form.checkbox("IsActive")
        controls = {c.name: c for c in db.forms["frmList"].controls()}
        assert controls["CustomerName_Label"].section is Section.HEADER
        assert controls["CustomerName"].section is Section.DETAIL
        assert controls["CustomerName_Label"].parent is None
        db.forms["frmList"].check_opens()


def test_replace_is_atomic_and_failed_builds_leave_nothing(tmp_path: Path) -> None:
    with AccessDatabase.create(tmp_path / "swap.accdb", engine="access") as db:
        _schema(db)
        with db.forms.create("frmA", record_source="Customers", caption="v1") as form:
            form.textbox("CustomerName")
        with pytest.raises(ObjectExistsError):
            db.forms.create("frmA", record_source="Customers").save()
        with db.forms.create("frmA", record_source="Customers", caption="v2", replace=True) as form:
            form.textbox("Email")
        assert db.forms.names() == ["frmA"]
        assert [c.name for c in db.forms["frmA"].controls() if c.kind is ControlKind.TEXTBOX] == [
            "Email"
        ]

        # an invalid Access property fails *inside* Access after controls were created: nothing may remain
        with pytest.raises(PyAccessKitError):
            db.forms.create(
                "frmBroken", record_source="Customers", properties={"NoSuchProperty": 1}
            ).save()
        assert db.forms.names() == ["frmA"]
        with (
            pytest.raises(PyAccessKitError),
            db.forms.create(
                "frmA", record_source="Customers", replace=True, properties={"NoSuchProperty": 1}
            ) as form,
        ):
            form.textbox("CustomerName")
        assert db.forms.names() == ["frmA"], "a failed replace keeps the original form"
        assert [c.name for c in db.forms["frmA"].controls() if c.kind is ControlKind.TEXTBOX] == [
            "Email"
        ]


def test_form_validation_happens_before_access(tmp_path: Path) -> None:
    with AccessDatabase.create(tmp_path / "v.accdb", engine="access") as db:
        _schema(db)
        with pytest.raises(SpecError, match="not a column"):
            db.forms.create("frmBad", record_source="Customers").textbox("Nope").save()
        assert db.forms.names() == []


def test_modules(tmp_path: Path) -> None:
    path = tmp_path / "vba.accdb"
    code = 'Public Function Greeting(ByVal who As String) As String\n    Greeting = "Héllo, " & who\nEnd Function\n'
    with AccessDatabase.create(path, engine="access") as db:
        db.modules.create("modUtils", code)
        db.modules.create(
            "clsCounter",
            "Private m As Long\nPublic Sub Add()\n    m = m + 1\nEnd Sub\n",
            kind=ModuleKind.CLASS,
        )
        with pytest.raises(SpecError, match="ANSI code page"):
            db.modules.create("modBad", 'x = "漢"')
    with AccessDatabase.open(path, engine="access") as db:
        assert sorted(db.modules.names()) == ["clsCounter", "modUtils"]
        assert db.modules["modUtils"].kind is ModuleKind.STANDARD
        assert db.modules["clsCounter"].kind is ModuleKind.CLASS
        assert 'Greeting = "Héllo, " & who' in db.modules["modUtils"].code
        db.modules["modUtils"].code = code.replace("Héllo", "Hi")
        assert '"Hi, "' in db.modules["modUtils"].code


def test_text_export_import_between_databases(tmp_path: Path) -> None:
    source = tmp_path / "source.accdb"
    with AccessDatabase.create(source, engine="access") as db:
        _schema(db)
        with db.forms.create("frmCustomers", record_source="Customers") as form:
            form.textbox("CustomerName", label="Name é漢")
        db.objects.save_text("form", "frmCustomers", tmp_path / "frmCustomers.form.txt")
    raw = (tmp_path / "frmCustomers.form.txt").read_text(encoding="utf-8")
    assert "Name é漢" in raw, "exported as UTF-8 text with non-ASCII intact"
    with AccessDatabase.create(tmp_path / "target.accdb", engine="access") as db:
        _schema(db)
        db.objects.load_text("form", "frmCopy", tmp_path / "frmCustomers.form.txt")
        labels = [c.caption for c in db.forms["frmCopy"].controls() if c.kind is ControlKind.LABEL]
        assert labels == ["Name é漢"]
        db.forms["frmCopy"].check_opens()


def test_constants_match_the_installed_type_libraries() -> None:
    """Drift test: the checked-in constants module equals a fresh render from the installed typelibs."""
    import importlib.util
    import sys

    script = Path(__file__).resolve().parents[2] / "scripts" / "gen_constants.py"
    spec = importlib.util.spec_from_file_location("gen_constants", script)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules["gen_constants"] = module
    spec.loader.exec_module(module)
    assert module.OUTPUT.read_text(encoding="utf-8") == module.render()
