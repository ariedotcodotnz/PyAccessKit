"""S3 — how do DAO-created and Access-created databases differ?

* Access: NewCurrentDatabase(acNewDatabaseFormatAccess2007)
* DAO:    DBEngine.CreateDatabase(dbLangGeneral, dbVersion120) (hosted in Access; same ACE as in-proc)
Dump Database.Properties + system tables for each, then open the DAO one in Access and see what gets added.
Also: which dbVersion options can Access open, and what db.Version reports.
"""

from __future__ import annotations

import pywintypes
from _harness import describe_com_error, fresh_dir, owned_access

dbLangGeneral = ";LANGID=0x0409;CP=1252;COUNTRY=0"


def props(db) -> dict[str, tuple[int, object]]:
    out: dict[str, tuple[int, object]] = {}
    for i in range(db.Properties.Count):
        p = db.Properties.Item(i)
        try:
            value = p.Value
        except pywintypes.com_error:
            value = "<unreadable>"
        out[p.Name] = (int(p.Type), value)
    return out


def tables(db) -> list[str]:
    return sorted(db.TableDefs.Item(i).Name for i in range(db.TableDefs.Count))


def main() -> None:
    work = fresh_dir("s03")
    access_path = work / "via_access.accdb"
    dao_path = work / "via_dao.accdb"
    with owned_access() as acc:
        app = acc.app
        app.NewCurrentDatabase(str(access_path), 12)
        db = app.CurrentDb()
        p_access = props(db)
        t_access = tables(db)
        del db
        app.CloseCurrentDatabase()

        db = app.DBEngine.CreateDatabase(str(dao_path), dbLangGeneral, 128)
        p_dao = props(db)
        t_dao = tables(db)
        print(f"DAO-created Version={db.Version}")
        db.Close()
        del db

        app.OpenCurrentDatabase(str(dao_path), True)
        db = app.CurrentDb()
        p_dao_after_open = props(db)
        t_dao_after_open = tables(db)
        del db
        app.CloseCurrentDatabase()

        db = app.DBEngine.OpenDatabase(str(dao_path))
        p_dao_persisted = props(db)
        db.Close()
        del db

        print("== Properties only in the Access-created DB ==")
        for k in sorted(p_access.keys() - p_dao.keys()):
            print(f"  {k}: {p_access[k]}")
        print("== Properties only in the DAO-created DB ==")
        for k in sorted(p_dao.keys() - p_access.keys()):
            print(f"  {k}: {p_dao[k]}")
        print("== Shared properties with different values ==")
        for k in sorted(p_access.keys() & p_dao.keys()):
            if p_access[k] != p_dao[k] and k not in ("Name",):
                print(f"  {k}: access={p_access[k]} dao={p_dao[k]}")
        print("== Properties added to the DAO DB by opening it in Access (in-session) ==")
        for k in sorted(p_dao_after_open.keys() - p_dao.keys()):
            print(f"  {k}: {p_dao_after_open[k]}")
        print("== ...and still present after reopening via DAO (persisted) ==")
        print(f"  {sorted(p_dao_persisted.keys() - p_dao.keys())}")
        print(f"== System tables: access={t_access}")
        print(f"                  dao={t_dao}")
        print(f"   dao after Access open={t_dao_after_open}")

        for option, label in ((128, "dbVersion120"), (256, "dbVersion140"), (512, "dbVersion150"), (1024, "dbVersion167")):
            path = work / f"{label}.accdb"
            try:
                db = app.DBEngine.CreateDatabase(str(path), dbLangGeneral, option)
                version = db.Version
                db.Close()
                del db
                app.OpenCurrentDatabase(str(path), True)
                app.CloseCurrentDatabase()
                print(f"{label}: created, Version={version}, opens in Access: yes")
            except pywintypes.com_error as exc:
                print(f"{label}: {describe_com_error(exc)}")


if __name__ == "__main__":
    main()
