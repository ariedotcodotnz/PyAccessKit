"""S6 â€” form automation in a hidden instance: CreateForm/CreateControl, attached labels, event code,
build-then-swap rename, header sections, rollback of unsaved forms, reading controls back.
"""

from __future__ import annotations

import pywintypes
from _harness import DialogWatcher, describe_com_error, fresh_dir, owned_access
from s06b_sections import get_indexed

acForm = 2
acLabel, acCommandButton, acCheckBox, acTextBox, acComboBox = 100, 104, 106, 109, 111
acDetail, acHeader, acFooter = 0, 1, 2
acSaveYes, acSaveNo = 1, 2
acNormal, acDesign = 0, 1
acHidden = 1
acFormPropertySettings = -1
acCmdFormHdrFtr = 36


def attempt(label, fn):
    try:
        out = fn()
        print(f"OK   {label}" + (f" -> {out!r}" if out is not None else ""))
        return out
    except pywintypes.com_error as exc:
        print(f"FAIL {label}: {describe_com_error(exc)}")
        return None


def all_forms(app) -> list[str]:
    forms = app.CurrentProject.AllForms
    return [forms.Item(i).Name for i in range(forms.Count)]


def build(app, target_name: str, caption: str) -> str:
    frm = app.CreateForm()
    auto = frm.Name
    print(f"CreateForm -> auto name {auto!r}")
    frm.RecordSource = "Customers"
    frm.Caption = caption
    frm.DefaultView = 0
    frm.Width = 7000
    get_indexed(frm, 'Section', acDetail).Height = 3200

    tb = app.CreateControl(auto, acTextBox, acDetail, "", "CustomerName", 2200, 200, 3600, 315)
    tb.Name = "CustomerName"
    lbl = app.CreateControl(auto, acLabel, acDetail, tb.Name, "", 200, 200, 1900, 315)
    lbl.Caption = "Customer name"
    lbl.Name = "CustomerName_Label"

    cb = app.CreateControl(auto, acCheckBox, acDetail, "", "IsActive", 2200, 700, 260, 240)
    cb.Name = "IsActive"
    cbl = app.CreateControl(auto, acLabel, acDetail, cb.Name, "", 200, 700, 1900, 315)
    cbl.Caption = "Active"
    cbl.Name = "IsActive_Label"

    combo = app.CreateControl(auto, acComboBox, acDetail, "", "Region", 2200, 1200, 3600, 315)
    combo.Name = "Region"
    combo.RowSourceType = "Table/Query"
    combo.RowSource = "SELECT RegionName FROM Regions ORDER BY RegionName;"
    combo.BoundColumn = 1
    combo.ColumnCount = 1
    combo.LimitToList = True
    col = app.CreateControl(auto, acLabel, acDetail, combo.Name, "", 200, 1200, 1900, 315)
    col.Caption = "Region"
    col.Name = "Region_Label"

    btn = app.CreateControl(auto, acCommandButton, acDetail, "", "", 2200, 1900, 1500, 400)
    btn.Name = "cmdClose"
    btn.Caption = "Close"
    btn.OnClick = "[Event Procedure]"
    frm.HasModule = True
    frm.Module.AddFromString(
        "Option Explicit\r\n\r\nPrivate Sub cmdClose_Click()\r\n    DoCmd.Close acForm, Me.Name\r\nEnd Sub\r\n"
    )
    frm.OnLoad = "[Event Procedure]"
    frm.Module.AddFromString("Private Sub Form_Load()\r\n    Me.Caption = Me.Caption\r\nEnd Sub\r\n")
    del tb, lbl, cb, cbl, combo, col, btn, frm
    app.DoCmd.Close(acForm, auto, acSaveYes)
    return auto


def main() -> None:
    work = fresh_dir("s06")
    with owned_access() as acc, DialogWatcher(acc.pid) as watch:
        app = acc.app
        app.NewCurrentDatabase(str(work / "forms.accdb"), 12)
        db = app.CurrentDb()
        db.Execute("CREATE TABLE Regions (RegionName TEXT(50) PRIMARY KEY)")
        db.Execute("INSERT INTO Regions VALUES ('North')")
        db.Execute(
            "CREATE TABLE Customers (CustomerID COUNTER PRIMARY KEY, CustomerName TEXT(100), IsActive YESNO, Region TEXT(50))"
        )
        db.Execute("INSERT INTO Customers (CustomerName, IsActive, Region) VALUES ('Ann', True, 'North')")
        del db

        auto = attempt("build form", lambda: build(app, "frmCustomers", "Customers"))
        attempt("rename auto -> frmCustomers", lambda: app.DoCmd.Rename("frmCustomers", acForm, auto))
        print(f"forms: {all_forms(app)}")

        def read_controls():
            app.DoCmd.OpenForm("frmCustomers", acDesign, "", "", acFormPropertySettings, acHidden)
            f = app.Forms("frmCustomers")
            rows = []
            for i in range(f.Controls.Count):
                c = f.Controls.Item(i)
                try:
                    source = get_indexed(c, "ControlSource")
                except pywintypes.com_error:
                    source = None
                try:
                    parent = get_indexed(c, "Parent").Name
                except pywintypes.com_error:
                    parent = None
                rows.append((c.Name, c.ControlType, source, c.Section, c.Left, c.Top, c.Width, c.Height, parent))
            info = (f.HasModule, f.RecordSource, f.Caption, f.Width, get_indexed(f, 'Section', acDetail).Height)
            del c, f
            app.DoCmd.Close(acForm, "frmCustomers", acSaveNo)
            return rows, info

        rows_info = attempt("read controls in hidden design view", read_controls)
        if rows_info:
            for row in rows_info[0]:
                print(f"     {row}")

        def open_normal():
            app.DoCmd.OpenForm("frmCustomers", acNormal, "", "", acFormPropertySettings, acHidden)
            f = app.Forms("frmCustomers")
            state = (f.CurrentRecord, f.Recordset.RecordCount, f.Controls("CustomerName").Value)
            del f
            app.DoCmd.Close(acForm, "frmCustomers", acSaveNo)
            return state

        attempt("open in normal view (hidden) and read bound value", open_normal)
        attempt("IsCompiled", lambda: app.IsCompiled)

        # rollback: unsaved form discarded
        def rollback():
            frm = app.CreateForm()
            name = frm.Name
            del frm
            try:
                app.CreateControl(name, 9999, acDetail, "", "", 0, 0, 100, 100)
            except pywintypes.com_error as exc:
                print(f"     (expected) bad control type: {describe_com_error(exc)}")
            app.DoCmd.Close(acForm, name, acSaveNo)
            return name, name in all_forms(app)

        attempt("rollback unsaved form (name, still exists?)", rollback)

        # header section creation strategies
        def header_via_create_control():
            frm = app.CreateForm()
            name = frm.Name
            del frm
            try:
                app.CreateControl(name, acLabel, acHeader, "", "", 100, 100, 1000, 300)
                result = "CreateControl in acHeader OK"
            except pywintypes.com_error as exc:
                result = f"CreateControl acHeader failed: {describe_com_error(exc)}"
            app.DoCmd.Close(acForm, name, acSaveNo)
            return result

        attempt("header via CreateControl(acHeader)", header_via_create_control)

        def header_via_runcommand():
            frm = app.CreateForm()
            name = frm.Name
            try:
                app.DoCmd.RunCommand(acCmdFormHdrFtr)
                h = get_indexed(frm, 'Section', acHeader).Height
                result = f"RunCommand acCmdFormHdrFtr OK, header height={h}"
            except pywintypes.com_error as exc:
                result = f"failed: {describe_com_error(exc)}"
            del frm
            app.DoCmd.Close(acForm, name, acSaveNo)
            return result

        attempt("header via RunCommand(acCmdFormHdrFtr)", header_via_runcommand)

        # build-then-swap replacement with a hidden '~' backup name
        def swap():
            new_auto = build(app, "frmCustomers", "Customers v2")
            app.DoCmd.Rename("~pak_bak_frmCustomers", acForm, "frmCustomers")
            app.DoCmd.Rename("frmCustomers", acForm, new_auto)
            app.DoCmd.DeleteObject(acForm, "~pak_bak_frmCustomers")
            return all_forms(app)

        attempt("swap replace", swap)
        attempt("rename onto existing name (expect failure)", lambda: app.DoCmd.Rename("frmCustomers", acForm, "frmCustomers"))
        print(f"dialogs seen: {len(watch.events)}")
        app.CloseCurrentDatabase()


if __name__ == "__main__":
    main()
