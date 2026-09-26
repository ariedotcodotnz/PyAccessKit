"""S2 — AutomationSecurity vs. startup code (AutoExec) and design operations.

For a database containing VBA + an AutoExec macro that writes a marker file, compare opening it:
* via Access-hosted DAO (app.DBEngine.OpenDatabase) — expected: no code runs;
* via OpenCurrentDatabase under AutomationSecurity = ByUI(2), ForceDisable(3), Low(1).
Record: open success, marker written (code ran?), dialogs shown, design ops (CreateForm) possible,
Application.Run possible.
"""

from __future__ import annotations

import time

import pywintypes
from _harness import DialogWatcher, describe_com_error, fresh_dir, owned_access

acForm, acMacro, acModule = 2, 4, 5
acSaveNo = 2


def build_fixture(path, marker) -> None:
    with owned_access() as acc:
        app = acc.app
        app.NewCurrentDatabase(str(path), 12)
        mod = path.parent / "modSpike.bas"
        mod.write_bytes(
            (
                "Option Compare Database\r\nOption Explicit\r\n\r\n"
                "Public Function SpikeMarker() As Boolean\r\n"
                "    Dim f As Integer\r\n    f = FreeFile\r\n"
                f'    Open "{marker}" For Output As #f\r\n    Print #f, "ran"\r\n    Close #f\r\n'
                "    SpikeMarker = True\r\nEnd Function\r\n"
            ).encode("cp1252")
        )
        app.LoadFromText(acModule, "modSpike", str(mod))
        mac = path.parent / "AutoExec.txt"
        mac.write_bytes(
            'Version =196611\r\nColumnsShown =0\r\nBegin\r\n    Action ="RunCode"\r\n    Argument ="SpikeMarker()"\r\nEnd\r\n'.encode(
                "cp1252"
            )
        )
        app.LoadFromText(acMacro, "AutoExec", str(mac))
        app.CloseCurrentDatabase()


def main() -> None:
    work = fresh_dir("s02")
    db_path = work / "startup.accdb"
    marker = work / "marker.txt"
    build_fixture(db_path, marker)
    print(f"fixture built; marker exists after build? {marker.exists()}")

    # Hosted DAO open
    marker.unlink(missing_ok=True)
    with owned_access() as acc, DialogWatcher(acc.pid) as watch:
        db = acc.app.DBEngine.OpenDatabase(str(db_path), True, False)
        names = [db.TableDefs.Item(i).Name for i in range(db.TableDefs.Count)][:3]
        db.Close()
        del db
        time.sleep(1)
        print(f"[hosted DAO] opened OK (tables sample={names}); code ran={marker.exists()}; dialogs={len(watch.events)}")

    for setting, label in ((2, "ByUI"), (3, "ForceDisable"), (1, "Low")):
        marker.unlink(missing_ok=True)
        with owned_access() as acc, DialogWatcher(acc.pid) as watch:
            app = acc.app
            app.AutomationSecurity = setting
            started = time.perf_counter()
            try:
                app.OpenCurrentDatabase(str(db_path), True)
                opened = True
            except pywintypes.com_error as exc:
                opened = False
                print(f"[{label}] OpenCurrentDatabase FAILED: {describe_com_error(exc)}")
            elapsed = time.perf_counter() - started
            time.sleep(2)
            print(f"[{label}] opened={opened} in {elapsed:.2f}s; AutoExec code ran={marker.exists()}; dialogs={len(watch.events)}")
            if opened:
                try:
                    cur = app.CurrentProject.FullName
                    print(f"[{label}] CurrentProject.FullName={cur}")
                except pywintypes.com_error as exc:
                    print(f"[{label}] CurrentProject failed: {describe_com_error(exc)}")
                try:
                    frm = app.CreateForm()
                    name = frm.Name
                    del frm
                    app.DoCmd.Close(acForm, name, acSaveNo)
                    print(f"[{label}] design op CreateForm OK")
                except pywintypes.com_error as exc:
                    print(f"[{label}] design op failed: {describe_com_error(exc)}")
                marker.unlink(missing_ok=True)
                try:
                    result = app.Run("SpikeMarker")
                    print(f"[{label}] Application.Run -> {result!r}; marker={marker.exists()}")
                except pywintypes.com_error as exc:
                    print(f"[{label}] Application.Run failed: {describe_com_error(exc)}")
                try:
                    app.CloseCurrentDatabase()
                except pywintypes.com_error as exc:
                    print(f"[{label}] CloseCurrentDatabase failed: {describe_com_error(exc)}")


if __name__ == "__main__":
    main()
