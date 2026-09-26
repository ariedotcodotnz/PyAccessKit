"""S9 — what do Access/DAO errors look like through IDispatch (numbers, descriptions, sources)?"""

from __future__ import annotations

import pywintypes
from _harness import OwnedAccess, dao_errors, describe_com_error, fresh_dir, owned_access

dbLangGeneral = ";LANGID=0x0409;CP=1252;COUNTRY=0"
dbText, dbLong = 10, 4
dbRelationUpdateCascade = 256
acForm = 2


def attempt(app, label, fn):
    try:
        out = fn()
        print(f"OK   {label}" + (f" -> {out!r}" if out is not None else ""))
    except pywintypes.com_error as exc:
        print(f"FAIL {label}: {describe_com_error(exc)}")
        scode = (exc.args[2] or (None,) * 6)[5]
        if scode is not None:
            try:
                print(f"     AccessError({scode & 0xFFFF}) = {app.AccessError(scode & 0xFFFF)!r}")
            except pywintypes.com_error:
                pass
        errs = dao_errors(app.DBEngine)
        if errs:
            print(f"     DBEngine.Errors = {errs}")


def main() -> None:
    work = fresh_dir("s09")
    path = work / "errors.accdb"
    bogus = work / "not_a_db.accdb"
    bogus.write_text("this is not an Access database", encoding="ascii")
    with owned_access() as acc:
        app = acc.app
        app.NewCurrentDatabase(str(path), 12)
        db = app.CurrentDb()
        db.Execute("CREATE TABLE Parent (PID LONG, Label TEXT(20))")
        db.Execute("CREATE TABLE Child (CID LONG, PID LONG, PIDText TEXT(20))")
        db.Execute("INSERT INTO Child (CID, PID) VALUES (1, 99)")

        attempt(app, "TableDefs('Nope')", lambda: db.TableDefs("Nope").Name)
        def dup():
            t = db.CreateTableDef("Parent")
            t.Fields.Append(t.CreateField("X", dbLong))
            db.TableDefs.Append(t)
        attempt(app, "append duplicate table", dup)
        attempt(app, "Fields.Append duplicate field", lambda: db.TableDefs("Parent").Fields.Append(db.TableDefs("Parent").CreateField("PID", dbLong)))
        attempt(app, "property not found", lambda: db.TableDefs("Parent").Properties("Description").Value)
        def rel_no_unique():
            r = db.CreateRelation("ParentChild", "Parent", "Child", 0)
            f = r.CreateField("PID")
            f.ForeignName = "PID"
            r.Fields.Append(f)
            db.Relations.Append(r)
        attempt(app, "relation without unique index on primary", rel_no_unique)
        db.Execute("CREATE UNIQUE INDEX PrimaryKey ON Parent (PID) WITH PRIMARY")
        attempt(app, "relation with orphan child rows (RI violation)", rel_no_unique)
        def rel_type_mismatch():
            r = db.CreateRelation("ParentChildText", "Parent", "Child", 0)
            f = r.CreateField("PID")
            f.ForeignName = "PIDText"
            r.Fields.Append(f)
            db.Relations.Append(r)
        attempt(app, "relation with mismatched types", rel_type_mismatch)
        attempt(app, "invalid name with '!'", lambda: db.CreateQueryDef("bad!name", "SELECT 1;"))
        attempt(app, "Execute syntax error", lambda: db.Execute("SELECT FROM WHERE", 128))
        attempt(app, "Execute missing parameter", lambda: db.Execute("UPDATE Parent SET Label = [x]", 128))
        attempt(app, "DoCmd.OpenForm missing", lambda: app.DoCmd.OpenForm("frmNope"))
        attempt(app, "DoCmd.DeleteObject missing form", lambda: app.DoCmd.DeleteObject(acForm, "frmNope"))
        attempt(app, "DoCmd.Rename missing form", lambda: app.DoCmd.Rename("x", acForm, "frmNope"))
        attempt(app, "LoadFromText missing file", lambda: app.LoadFromText(5, "modX", str(work / "missing.bas")))
        del db
        # keep path open exclusively in this instance for the lock tests below
        app.CloseCurrentDatabase()
        app.OpenCurrentDatabase(str(path), True)

        other = OwnedAccess()
        try:
            o = other.app
            attempt(o, "hosted OpenDatabase on exclusively-open file", lambda: o.DBEngine.OpenDatabase(str(path)).Close())
            attempt(o, "OpenCurrentDatabase on exclusively-open file", lambda: o.OpenCurrentDatabase(str(path), False))
            attempt(o, "OpenCurrentDatabase missing file", lambda: o.OpenCurrentDatabase(str(work / "missing.accdb"), False))
            attempt(o, "hosted OpenDatabase missing file", lambda: o.DBEngine.OpenDatabase(str(work / "missing.accdb")))
            attempt(o, "hosted OpenDatabase non-database file", lambda: o.DBEngine.OpenDatabase(str(bogus)))
            attempt(o, "OpenCurrentDatabase non-database file", lambda: o.OpenCurrentDatabase(str(bogus), False))
            attempt(o, "CreateDatabase over existing file", lambda: o.DBEngine.CreateDatabase(str(path), dbLangGeneral, 128))
            attempt(o, "NewCurrentDatabase over existing file", lambda: o.NewCurrentDatabase(str(path), 12))
        finally:
            other.quit()
        app.CloseCurrentDatabase()


if __name__ == "__main__":
    main()
