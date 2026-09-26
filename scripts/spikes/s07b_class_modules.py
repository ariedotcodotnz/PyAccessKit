"""S7b — class modules: what does Access's own SaveAsText produce, and can LoadFromText recreate it?

Also: do UTF-8 form files keep non-ASCII text, and are LoadFromText objects persisted after reopen?
"""

from __future__ import annotations

import codecs

import pywintypes
from _harness import describe_com_error, fresh_dir, owned_access

acForm, acModule = 2, 5
acSaveYes = 1
vbext_ct_ClassModule = 2


def decode(raw: bytes) -> str:
    if raw.startswith(codecs.BOM_UTF16_LE):
        return raw.decode("utf-16")
    if raw.startswith(codecs.BOM_UTF8):
        return raw.decode("utf-8-sig")
    return raw.decode("cp1252")


def main() -> None:
    work = fresh_dir("s07b")
    path = work / "cls.accdb"
    with owned_access() as acc:
        app = acc.app
        app.NewCurrentDatabase(str(path), 12)
        comps = app.VBE.ActiveVBProject.VBComponents
        comp = comps.Add(vbext_ct_ClassModule)
        comp.Name = "clsViaVbe"
        comp.CodeModule.AddFromString("Public Title As String\r\nPublic Function Greet() As String\r\n    Greet = \"hi \" & Title\r\nEnd Function")
        try:
            app.DoCmd.Close(acModule, "clsViaVbe", acSaveYes)
            print("DoCmd.Close acModule clsViaVbe acSaveYes -> OK")
        except pywintypes.com_error as exc:
            print(f"DoCmd.Close failed: {describe_com_error(exc)}")
            try:
                app.DoCmd.Save(acModule, "clsViaVbe")
                print("DoCmd.Save acModule -> OK")
            except pywintypes.com_error as exc2:
                print(f"DoCmd.Save failed: {describe_com_error(exc2)}")

        out = work / "clsViaVbe.out.cls"
        app.SaveAsText(acModule, "clsViaVbe", str(out))
        raw = out.read_bytes()
        print(f"export bytes={len(raw)} head={raw[:40]!r}")
        text = decode(raw)
        print("---- Access SaveAsText of a class module ----")
        print(text)
        print("---------------------------------------------")

        app.LoadFromText(acModule, "clsCopy", str(out))
        print(f"clsCopy Type after LoadFromText of Access's own export: {comps.Item('clsCopy').Type}")

        # standard module + form with UTF-8 non-ASCII, then persistence check
        std = work / "modStd.in.bas"
        std.write_bytes(b"Option Compare Database\r\nPublic Function Two() As Long\r\n    Two = 2\r\nEnd Function\r\n")
        app.LoadFromText(acModule, "modStd", str(std))
        frm = app.CreateForm()
        auto = frm.Name
        frm.Caption = "Caption é漢"
        app.DoCmd.Close(acForm, auto, acSaveYes)
        del frm
        exported = work / "frm.txt"
        app.SaveAsText(acForm, auto, str(exported))
        utf8 = work / "frmUtf8.txt"
        utf8.write_bytes(decode(exported.read_bytes()).encode("utf-8"))
        app.LoadFromText(acForm, "frmFromUtf8", str(utf8))
        back = work / "frmFromUtf8.out.txt"
        app.SaveAsText(acForm, "frmFromUtf8", str(back))
        print(f"UTF-8 (no BOM) form import keeps non-ASCII caption: {'Caption é漢' in decode(back.read_bytes())}")

        # overwrite semantics: LoadFromText onto an existing module name
        std2 = work / "modStd2.in.bas"
        std2.write_bytes(b"Option Compare Database\r\nPublic Function Three() As Long\r\n    Three = 3\r\nEnd Function\r\n")
        try:
            app.LoadFromText(acModule, "modStd", str(std2))
            print(f"LoadFromText over existing module -> OK; line3={comps.Item('modStd').CodeModule.Lines(3, 1)!r}")
        except pywintypes.com_error as exc:
            print(f"LoadFromText over existing module failed: {describe_com_error(exc)}")
        app.CloseCurrentDatabase()

        app.OpenCurrentDatabase(str(path), True)
        names = [app.CurrentProject.AllModules.Item(i).Name for i in range(app.CurrentProject.AllModules.Count)]
        forms = [app.CurrentProject.AllForms.Item(i).Name for i in range(app.CurrentProject.AllForms.Count)]
        print(f"after reopen: modules={names} forms={forms}")
        comps = app.VBE.ActiveVBProject.VBComponents
        for i in range(1, comps.Count + 1):
            c = comps.Item(i)
            print(f"  component {c.Name}: Type={c.Type} lines={c.CodeModule.CountOfLines}")
        app.CloseCurrentDatabase()


if __name__ == "__main__":
    main()
