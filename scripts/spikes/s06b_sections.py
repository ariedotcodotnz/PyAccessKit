"""S6b — isolate the 'Member not found' failure: indexed properties (Form.Section) and header creation."""

from __future__ import annotations

import pythoncom
import pywintypes
from _harness import DialogWatcher, describe_com_error, fresh_dir, owned_access
from win32com.client import dynamic

acForm = 2
acLabel, acCommandButton, acTextBox = 100, 104, 109
acDetail, acHeader, acFooter = 0, 1, 2
acSaveYes, acSaveNo = 1, 2
acCmdFormHdrFtr = 36


def step(label, fn):
    try:
        out = fn()
        print(f"OK   {label}" + (f" -> {out!r}" if out is not None else ""))
        return out
    except pywintypes.com_error as exc:
        print(f"FAIL {label}: {describe_com_error(exc)}")
        return None


def get_indexed(obj, name: str, *args):
    dispid = obj._oleobj_.GetIDsOfNames(name)
    result = obj._oleobj_.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYGET | pythoncom.DISPATCH_METHOD, True, *args)
    return dynamic.Dispatch(result) if isinstance(result, pythoncom.TypeIIDs[pythoncom.IID_IDispatch]) else result


def main() -> None:
    work = fresh_dir("s06b")
    with owned_access() as acc, DialogWatcher(acc.pid) as watch:
        app = acc.app
        app.NewCurrentDatabase(str(work / "sections.accdb"), 12)
        app.CurrentDb().Execute("CREATE TABLE T (ID COUNTER PRIMARY KEY, Name1 TEXT(50))")
        frm = app.CreateForm()
        auto = frm.Name
        step("frm.RecordSource =", lambda: setattr(frm, "RecordSource", "T"))
        step("frm.Caption =", lambda: setattr(frm, "Caption", "Hello"))
        step("frm.DefaultView =", lambda: setattr(frm, "DefaultView", 1))
        step("frm.Width =", lambda: setattr(frm, "Width", 7000))
        step("frm.Section(0) [dynamic]", lambda: frm.Section(acDetail).Height)
        step("get_indexed(frm,'Section',0).Height", lambda: get_indexed(frm, "Section", acDetail).Height)
        step("set detail height via get_indexed", lambda: setattr(get_indexed(frm, "Section", acDetail), "Height", 3200))
        step("header exists? get_indexed Section(1)", lambda: get_indexed(frm, "Section", acHeader).Height)
        step("DoCmd.RunCommand(acCmdFormHdrFtr)", lambda: app.DoCmd.RunCommand(acCmdFormHdrFtr))
        step("header after RunCommand", lambda: get_indexed(frm, "Section", acHeader).Height)
        step("footer after RunCommand", lambda: get_indexed(frm, "Section", acFooter).Height)
        step("CreateControl label in header", lambda: app.CreateControl(auto, acLabel, acHeader, "", "", 100, 60, 2000, 300).Name)
        step("frm.HasModule = True", lambda: setattr(frm, "HasModule", True))
        step("frm.Module", lambda: frm.Module.Name)
        step("frm.Module.AddFromString", lambda: frm.Module.AddFromString("Option Explicit\r\n"))
        step("frm.OnLoad =", lambda: setattr(frm, "OnLoad", "[Event Procedure]"))
        tb = step("CreateControl textbox", lambda: app.CreateControl(auto, acTextBox, acDetail, "", "Name1", 2000, 200, 3000, 315))
        if tb is not None:
            step("tb.Name =", lambda: setattr(tb, "Name", "Name1"))
            step("tb.ControlSource", lambda: tb.ControlSource)
            step("tb.Properties('Format')", lambda: tb.Properties("Format").Value)
            step("tb.Properties('Format').Value = 'General Number'", lambda: setattr(tb.Properties("Format"), "Value", ""))
        btn = step("CreateControl button", lambda: app.CreateControl(auto, acCommandButton, acDetail, "", "", 2000, 700, 1500, 400))
        if btn is not None:
            step("btn.OnClick =", lambda: setattr(btn, "OnClick", "[Event Procedure]"))
        step("frm.Properties('ScrollBars').Value", lambda: frm.Properties("ScrollBars").Value)
        step("frm.Properties('ScrollBars').Value = 2", lambda: setattr(frm.Properties("ScrollBars"), "Value", 2))
        step("Close save", lambda: app.DoCmd.Close(acForm, auto, acSaveYes))
        step("Rename", lambda: app.DoCmd.Rename("frmT", acForm, auto))
        out = work / "frmT.txt"
        step("SaveAsText", lambda: app.SaveAsText(acForm, "frmT", str(out)))
        if out.exists():
            text = out.read_bytes().decode("utf-16")
            print("---- sections in exported form ----")
            for line in text.splitlines():
                if line.strip().startswith(("Begin FormHeader", "Begin FormFooter", "Begin Section", "Begin PageHeader")):
                    print("   ", line.strip())
        del frm
        print(f"dialogs seen: {len(watch.events)}")
        app.CloseCurrentDatabase()


if __name__ == "__main__":
    main()
