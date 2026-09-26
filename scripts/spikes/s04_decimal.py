"""S4 — Decimal(p,s): which paths can create it, and how do we read precision/scale back?"""

from __future__ import annotations

import pywintypes
from _harness import describe_com_error, fresh_dir, owned_access

adSchemaColumns = 4
dbDecimal = 20


def main() -> None:
    work = fresh_dir("s04")
    path = work / "decimal.accdb"
    with owned_access() as acc:
        app = acc.app
        app.NewCurrentDatabase(str(path), 12)
        db = app.CurrentDb()
        db.Execute("CREATE TABLE T (ID LONG)")

        for label, sql in (
            ("DAO Execute ALTER ADD DECIMAL(10,2)", "ALTER TABLE T ADD COLUMN d_dao DECIMAL(10,2)"),
            ("DAO Execute ALTER ADD NUMERIC(10,2)", "ALTER TABLE T ADD COLUMN d_dao2 NUMERIC(10,2)"),
        ):
            try:
                db.Execute(sql, 128)
                print(f"OK   {label}")
            except pywintypes.com_error as exc:
                print(f"FAIL {label}: {describe_com_error(exc)}")

        conn = app.CurrentProject.Connection
        for label, sql in (
            ("ADO ALTER ADD DECIMAL(10,2)", "ALTER TABLE T ADD COLUMN d_ado DECIMAL(10,2)"),
            ("ADO ALTER ADD DECIMAL(28,6)", "ALTER TABLE T ADD COLUMN d_big DECIMAL(28,6)"),
            ("ADO ALTER ALTER COLUMN d_ado DECIMAL(12,3)", "ALTER TABLE T ALTER COLUMN d_ado DECIMAL(12,3)"),
            ("ADO CREATE TABLE with DECIMAL", "CREATE TABLE T2 (ID LONG, Amount DECIMAL(18,4) NOT NULL)"),
        ):
            try:
                conn.Execute(sql)
                print(f"OK   {label}")
            except pywintypes.com_error as exc:
                print(f"FAIL {label}: {describe_com_error(exc)}")

        # DAO-only decimal (no precision control)
        tdf = db.TableDefs("T")
        try:
            fld = tdf.CreateField("d_daofield", dbDecimal)
            tdf.Fields.Append(fld)
            print("OK   DAO CreateField(dbDecimal) append")
        except pywintypes.com_error as exc:
            print(f"FAIL DAO CreateField(dbDecimal): {describe_com_error(exc)}")
        db.TableDefs.Refresh()
        tdf = db.TableDefs("T")
        for i in range(tdf.Fields.Count):
            f = tdf.Fields.Item(i)
            pnames = [f.Properties.Item(j).Name for j in range(f.Properties.Count)]
            extra = [n for n in pnames if n.lower() in ("precision", "scale", "decimalplaces")]
            print(f"  DAO field {f.Name}: Type={f.Type} Size={f.Size} precision-ish props={extra}")
            for name in extra:
                try:
                    print(f"      {name}={f.Properties(name).Value}")
                except pywintypes.com_error as exc:
                    print(f"      {name}: {describe_com_error(exc)}")

        rs = conn.OpenSchema(adSchemaColumns)
        while not rs.EOF:
            if rs.Fields("TABLE_NAME").Value in ("T", "T2"):
                print(
                    "  ADO schema:",
                    rs.Fields("TABLE_NAME").Value,
                    rs.Fields("COLUMN_NAME").Value,
                    "type",
                    rs.Fields("DATA_TYPE").Value,
                    "prec",
                    rs.Fields("NUMERIC_PRECISION").Value,
                    "scale",
                    rs.Fields("NUMERIC_SCALE").Value,
                    "nullable",
                    rs.Fields("IS_NULLABLE").Value,
                )
            rs.MoveNext()
        rs.Close()
        del rs, conn, tdf, db
        app.CloseCurrentDatabase()

    # in-proc ADO + ACE OLEDB from this (64-bit) Python
    from win32com.client import dynamic
    import pythoncom

    for provider in ("Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0"):
        try:
            conn = dynamic.Dispatch(pythoncom.CoCreateInstance("ADODB.Connection", None, pythoncom.CLSCTX_INPROC_SERVER, pythoncom.IID_IDispatch))
            conn.Open(f"Provider={provider};Data Source={path}")
            print(f"in-proc ADO {provider}: opened OK")
            conn.Close()
        except pywintypes.com_error as exc:
            print(f"in-proc ADO {provider}: {describe_com_error(exc)}")


if __name__ == "__main__":
    main()
