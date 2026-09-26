"""S5 — QueryDef behaviour: SQL round-trip, reference validation, parameters, types, name collisions."""

from __future__ import annotations

import pywintypes
from _harness import dao_errors, describe_com_error, fresh_dir, owned_access

dbOpenSnapshot = 4
dbFailOnError = 128


def attempt(label, fn):
    try:
        out = fn()
        print(f"OK   {label}" + (f" -> {out!r}" if out is not None else ""))
        return out
    except pywintypes.com_error as exc:
        print(f"FAIL {label}: {describe_com_error(exc)}")
        return None


def main() -> None:
    work = fresh_dir("s05")
    with owned_access() as acc:
        app = acc.app
        app.NewCurrentDatabase(str(work / "q.accdb"), 12)
        db = app.CurrentDb()
        run_sql = db.Execute
        run_sql("CREATE TABLE Customers (ID COUNTER PRIMARY KEY, CustomerName TEXT(100), Score LONG)")
        run_sql("INSERT INTO Customers (CustomerName, Score) VALUES ('Ann', 5)")
        run_sql("INSERT INTO Customers (CustomerName, Score) VALUES ('Bob', 9)")

        samples = {
            "q_oneline": "SELECT * FROM Customers ORDER BY CustomerName;",
            "q_multiline": "SELECT *\nFROM Customers\nORDER BY CustomerName;",
            "q_crlf_nosemi": "SELECT ID, CustomerName\r\nFROM Customers\r\nWHERE Score > 1",
            "q_lower": "select id from customers where score>1",
            "q_union": "SELECT ID FROM Customers UNION SELECT Score FROM Customers;",
        }
        for name, sql in samples.items():
            attempt(f"create {name}", lambda n=name, s=sql: db.CreateQueryDef(n, s) and None)
            stored = db.QueryDefs(name).SQL
            print(f"     in ={sql!r}\n     out={stored!r}  Type={db.QueryDefs(name).Type}")

        attempt("create q_missing_table (refers to NoSuchTable)", lambda: db.CreateQueryDef("q_missing", "SELECT * FROM NoSuchTable;") and None)
        attempt("create q_forward (refers to q_later query not yet created)", lambda: db.CreateQueryDef("q_forward", "SELECT * FROM q_later;") and None)
        attempt("create q_syntax (SELEC)", lambda: db.CreateQueryDef("q_syntax", "SELEC * FROM Customers") and None)
        print(f"     DAO errors: {dao_errors(app.DBEngine)}")
        attempt("create q_syntax2 (missing operator)", lambda: db.CreateQueryDef("q_syntax2", "SELECT * FROM Customers WHERE Score > > 1") and None)
        attempt("create query named like a table ('Customers')", lambda: db.CreateQueryDef("Customers", "SELECT 1;") and None)
        attempt("create table named like a query ('q_oneline')", lambda: run_sql("CREATE TABLE q_oneline (X LONG)", dbFailOnError))

        for name, sql in {
            "q_update": "UPDATE Customers SET Score = Score + 1;",
            "q_delete": "DELETE FROM Customers WHERE Score < 0;",
            "q_append": "INSERT INTO Customers (CustomerName) SELECT CustomerName FROM Customers WHERE ID = 0;",
            "q_maketable": "SELECT * INTO CustomersCopy FROM Customers;",
            "q_ddl": "CREATE TABLE Z (A LONG);",
            "q_crosstab": "TRANSFORM Sum(Score) SELECT CustomerName FROM Customers GROUP BY CustomerName PIVOT ID;",
            "q_params": "PARAMETERS [pMin] Long, [pName] Text ( 255 ); SELECT * FROM Customers WHERE Score > [pMin] AND CustomerName <> [pName];",
            "q_implicit": "SELECT * FROM Customers WHERE Score > [MinScore];",
        }.items():
            attempt(f"create {name}", lambda n=name, s=sql: db.CreateQueryDef(n, s) and None)
            qd = db.QueryDefs(name)
            params = [(qd.Parameters.Item(i).Name, qd.Parameters.Item(i).Type) for i in range(qd.Parameters.Count)]
            print(f"     Type={qd.Type} params={params} sql={qd.SQL!r}")

        # pass-through: create empty, then set Connect before SQL
        def make_pt():
            qd = db.CreateQueryDef("q_pt")
            qd.Connect = "ODBC;DSN=NoSuchDsn;"
            qd.SQL = "SELECT TOP 5 * FROM dbo.Anything"
            qd.ReturnsRecords = True
            qd.ODBCTimeout = 30
            return (qd.Type, qd.Connect, qd.SQL)

        attempt("create pass-through q_pt", make_pt)

        # temp querydef with parameters + GetRows
        def temp_fetch():
            qd = db.CreateQueryDef("", "SELECT ID, CustomerName, Score FROM Customers WHERE Score >= [p] ORDER BY ID")
            qd.Parameters("p").Value = 5
            rs = qd.OpenRecordset(dbOpenSnapshot)
            cols = [rs.Fields.Item(i).Name for i in range(rs.Fields.Count)]
            data = rs.GetRows(1000)
            rs.Close()
            return cols, data

        attempt("temp querydef fetch with parameter", temp_fetch)

        def temp_action():
            qd = db.CreateQueryDef("", "UPDATE Customers SET Score = Score + [inc] WHERE ID = [id]")
            qd.Parameters("inc").Value = 10
            qd.Parameters("id").Value = 1
            qd.Execute(dbFailOnError)
            return qd.RecordsAffected

        attempt("temp querydef action with parameters", temp_action)
        attempt("QueryDefs names", lambda: sorted(db.QueryDefs.Item(i).Name for i in range(db.QueryDefs.Count)))
        attempt("rename query via Name", lambda: setattr(db.QueryDefs("q_lower"), "Name", "q_renamed"))
        attempt("delete query", lambda: db.QueryDefs.Delete("q_renamed"))
        del db
        app.CloseCurrentDatabase()


if __name__ == "__main__":
    main()
