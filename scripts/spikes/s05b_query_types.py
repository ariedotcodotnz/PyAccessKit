"""S5b — Is QueryDef.Type populated after reopening? (It read 0 right after creation.)

Also S4b: can DAO *write* Precision/Scale on a decimal field (avoiding ADO)?
"""

from __future__ import annotations

import pywintypes
from _harness import describe_com_error, fresh_dir, owned_access

QUERIES = {
    "q_select": "SELECT * FROM T;",
    "q_update": "UPDATE T SET X = X + 1;",
    "q_delete": "DELETE FROM T WHERE X < 0;",
    "q_append": "INSERT INTO T (X) SELECT X FROM T;",
    "q_maketable": "SELECT * INTO T2 FROM T;",
    "q_ddl": "CREATE TABLE Z (A LONG);",
    "q_crosstab": "TRANSFORM Sum(X) SELECT X FROM T GROUP BY X PIVOT X;",
    "q_union": "SELECT X FROM T UNION SELECT X FROM T;",
}


def types(db) -> dict[str, int]:
    return {db.QueryDefs.Item(i).Name: db.QueryDefs.Item(i).Type for i in range(db.QueryDefs.Count)}


def main() -> None:
    work = fresh_dir("s05b")
    path = work / "types.accdb"
    with owned_access() as acc:
        app = acc.app
        app.NewCurrentDatabase(str(path), 12)
        db = app.CurrentDb()
        db.Execute("CREATE TABLE T (X LONG)")
        for name, sql in QUERIES.items():
            db.CreateQueryDef(name, sql)
        qd = db.CreateQueryDef("q_pt")
        qd.Connect = "ODBC;DSN=NoSuchDsn;"
        qd.SQL = "SELECT 1"
        qd.ReturnsRecords = True
        del qd
        print(f"same handle, right after create: {types(db)}")
        db.QueryDefs.Refresh()
        print(f"after QueryDefs.Refresh():       {types(db)}")
        db2 = app.CurrentDb()
        print(f"fresh CurrentDb():               {types(db2)}")
        del db, db2

        # S4b: DAO Precision/Scale writes
        db = app.CurrentDb()
        tdf = db.CreateTableDef("D")
        fld = tdf.CreateField("Amount", 20)
        try:
            names = [fld.Properties.Item(i).Name for i in range(fld.Properties.Count)]
            print(f"unappended dbDecimal field props: {names}")
        except pywintypes.com_error as exc:
            print(f"unappended props: {describe_com_error(exc)}")
        for prop, value in (("Precision", 10), ("Scale", 2)):
            try:
                fld.Properties(prop).Value = value
                print(f"set unappended {prop}={value}: OK")
            except pywintypes.com_error as exc:
                print(f"set unappended {prop}: {describe_com_error(exc)}")
        tdf.Fields.Append(fld)
        db.TableDefs.Append(tdf)
        db.TableDefs.Refresh()
        f = db.TableDefs("D").Fields("Amount")
        print(f"appended: Type={f.Type} Size={f.Size} Precision={f.Properties('Precision').Value} Scale={f.Properties('Scale').Value}")
        try:
            f.Properties("Precision").Value = 12
            print("set appended Precision: OK")
        except pywintypes.com_error as exc:
            print(f"set appended Precision: {describe_com_error(exc)}")
        del f, fld, tdf, db
        app.CloseCurrentDatabase()

        app.OpenCurrentDatabase(str(path), True)
        db = app.CurrentDb()
        print(f"after reopen (CurrentDb):        {types(db)}")
        del db
        app.CloseCurrentDatabase()
        db = app.DBEngine.OpenDatabase(str(path))
        print(f"after reopen (hosted DAO):       {types(db)}")
        db.Close()


if __name__ == "__main__":
    main()
