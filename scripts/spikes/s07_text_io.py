"""S7 — SaveAsText / LoadFromText encodings, module headers, macro format, VBE access."""

from __future__ import annotations

import codecs
import locale

import pywintypes
from _harness import describe_com_error, fresh_dir, owned_access

acQuery, acForm, acMacro, acModule = 1, 2, 4, 5
acTextBox, acDetail = 109, 0
acSaveYes = 1


def sniff(path) -> str:
    raw = path.read_bytes()
    if raw.startswith(codecs.BOM_UTF16_LE):
        enc = "utf-16-le BOM"
    elif raw.startswith(codecs.BOM_UTF8):
        enc = "utf-8 BOM"
    else:
        enc = "no BOM"
    return f"{enc}, {len(raw)} bytes, head={raw[:24]!r}"


def decode(path) -> str:
    raw = path.read_bytes()
    if raw.startswith(codecs.BOM_UTF16_LE):
        try:
            return raw.decode("utf-16")
        except UnicodeDecodeError:
            return "<<not valid utf-16>> " + raw.decode("cp1252", errors="replace")
    if raw.startswith(codecs.BOM_UTF8):
        return raw.decode("utf-8-sig")
    return raw.decode(locale.getpreferredencoding(False), errors="replace")


def vbe_line(app, component: str, line: int) -> str:
    try:
        cm = app.VBE.ActiveVBProject.VBComponents(component).CodeModule
        return repr(cm.Lines(line, 1))
    except pywintypes.com_error as exc:
        return f"<vbe error {describe_com_error(exc)}>"


def try_call(label: str, fn) -> object:
    try:
        result = fn()
        print(f"  OK   {label}")
        return result
    except pywintypes.com_error as exc:
        print(f"  FAIL {label}: {describe_com_error(exc)}")
        return None


def main() -> None:
    work = fresh_dir("s07")
    print(f"ANSI code page: {locale.getpreferredencoding(False)}")
    with owned_access() as acc:
        app = acc.app
        app.NewCurrentDatabase(str(work / "text.accdb"), 12)
        db = app.CurrentDb()
        db.Execute("CREATE TABLE Customers (ID COUNTER PRIMARY KEY, CustomerName TEXT(100))")
        db.CreateQueryDef("qryTest", "SELECT ID, CustomerName, 'é漢' AS U FROM Customers;")

        frm = app.CreateForm()
        auto = frm.Name
        frm.RecordSource = "Customers"
        frm.Caption = "Formulaire é漢"
        app.CreateControl(auto, acTextBox, acDetail, "", "CustomerName", 1440, 360, 2880, 300)
        app.DoCmd.Close(acForm, auto, acSaveYes)
        app.DoCmd.Rename("frmTest", acForm, auto)
        del frm

        print("[export] SaveAsText encodings:")
        for kind, name in ((acQuery, "qryTest"), (acForm, "frmTest")):
            out = work / f"{name}.txt"
            try_call(f"SaveAsText {name}", lambda k=kind, n=name, o=out: app.SaveAsText(k, n, str(o)))
            print(f"    {name}: {sniff(out)}")

        text_mod = (
            "Option Compare Database\r\nOption Explicit\r\n\r\n"
            'Public Function Hello() As String\r\n    Hello = "héllo"\r\nEnd Function\r\n'
        )
        print("[import] standard module via LoadFromText in different encodings:")
        variants = {
            "modUtf8": text_mod.encode("utf-8"),
            "modUtf8Bom": codecs.BOM_UTF8 + text_mod.encode("utf-8"),
            "modAnsi": text_mod.encode("cp1252"),
            "modUtf16": codecs.BOM_UTF16_LE + text_mod.encode("utf-16-le"),
        }
        for name, payload in variants.items():
            src = work / f"{name}.in.bas"
            src.write_bytes(payload)
            if try_call(f"LoadFromText {name}", lambda n=name, s=src: app.LoadFromText(acModule, n, str(s))) is not None or True:
                out = work / f"{name}.out.bas"
                if try_call(f"  SaveAsText {name}", lambda n=name, o=out: app.SaveAsText(acModule, n, str(o))) is not None or out.exists():
                    if out.exists():
                        body = decode(out)
                        print(f"    {name} export: {sniff(out)}")
                        print(f"    contains 'héllo': {'héllo' in body}; first line: {body.splitlines()[0]!r}")
                        print(f"    VBE line1={vbe_line(app, name, 1)} line5={vbe_line(app, name, 5)}")

        print("[import] class module with VB header:")
        cls_text = (
            "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n"
            'Attribute VB_Name = "clsThing"\r\n'
            "Attribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\n"
            "Attribute VB_PredeclaredId = False\r\nAttribute VB_Exposed = False\r\n"
            "Option Compare Database\r\nOption Explicit\r\n\r\nPublic Title As String\r\n"
        )
        src = work / "clsThing.in.cls"
        src.write_bytes(cls_text.encode("cp1252"))
        try_call("LoadFromText clsThing", lambda: app.LoadFromText(acModule, "clsThing", str(src)))
        out = work / "clsThing.out.cls"
        try_call("SaveAsText clsThing", lambda: app.SaveAsText(acModule, "clsThing", str(out)))
        if out.exists():
            print(f"    export: {sniff(out)}")
            print("    ---- exported class text ----")
            print("    " + decode(out).replace("\r\n", "\n    "))
        try:
            comps = app.VBE.ActiveVBProject.VBComponents
            print(f"[vbe] VBComponents.Count={comps.Count}")
            for i in range(1, comps.Count + 1):
                c = comps.Item(i)
                print(f"    {c.Name}: Type={c.Type}")
        except pywintypes.com_error as exc:
            print(f"[vbe] VBE not accessible: {describe_com_error(exc)}")

        print("[export] standard module text:")
        out = work / "modAnsi.out.bas"
        if out.exists():
            print("    " + decode(out).replace("\r\n", "\n    "))

        print("[import] form from UTF-8 re-encoded export:")
        form_text = decode(work / "frmTest.txt")
        for name, payload in {
            "frmUtf8": form_text.encode("utf-8"),
            "frmUtf8Bom": codecs.BOM_UTF8 + form_text.encode("utf-8"),
            "frmUtf16": codecs.BOM_UTF16_LE + form_text.encode("utf-16-le"),
        }.items():
            src = work / f"{name}.in.txt"
            src.write_bytes(payload)
            if try_call(f"LoadFromText {name}", lambda n=name, s=src: app.LoadFromText(acForm, n, str(s))) is not None:
                out = work / f"{name}.out.txt"
                app.SaveAsText(acForm, name, str(out))
                print(f"    {name} caption survived: {'Formulaire é漢' in decode(out)}")

        print("[import] macro (legacy text format):")
        macro_text = (
            "Version =196611\r\nColumnsShown =0\r\nBegin\r\n"
            '    Action ="Beep"\r\nEnd\r\n'
        )
        src = work / "mcrTest.in.txt"
        src.write_bytes(macro_text.encode("cp1252"))
        try_call("LoadFromText mcrTest (ANSI)", lambda: app.LoadFromText(acMacro, "mcrTest", str(src)))
        src16 = work / "mcrTest16.in.txt"
        src16.write_bytes(codecs.BOM_UTF16_LE + macro_text.encode("utf-16-le"))
        try_call("LoadFromText mcrTest16 (UTF-16)", lambda: app.LoadFromText(acMacro, "mcrTest16", str(src16)))
        for name in ("mcrTest", "mcrTest16"):
            out = work / f"{name}.out.txt"
            if try_call(f"SaveAsText {name}", lambda n=name, o=out: app.SaveAsText(acMacro, n, str(o))) is not None or out.exists():
                if out.exists():
                    print(f"    {name} export: {sniff(out)}")
                    print("    " + decode(out).replace("\r\n", "\n    "))

        print("[export] query text:")
        print("    " + decode(work / "qryTest.txt").replace("\r\n", "\n    "))
        print("[export] form text (first 40 lines):")
        print("    " + "\n    ".join(decode(work / "frmTest.txt").splitlines()[:40]))
        del db
        app.CloseCurrentDatabase()


if __name__ == "__main__":
    main()
