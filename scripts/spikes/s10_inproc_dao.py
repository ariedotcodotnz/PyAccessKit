"""S10 — in-process DAO (and ADO+ACE OLEDB) from a bitness-matched (x86) Python under C2R Office."""

from __future__ import annotations

import struct
import time

import pythoncom
import pywintypes
from _harness import describe_com_error, fresh_dir
from win32com.client import dynamic

dbLangGeneral = ";LANGID=0x0409;CP=1252;COUNTRY=0"


def inproc(progid: str):
    return dynamic.Dispatch(pythoncom.CoCreateInstance(progid, None, pythoncom.CLSCTX_INPROC_SERVER, pythoncom.IID_IDispatch))


def main() -> None:
    print(f"python bits: {struct.calcsize('P') * 8}")
    work = fresh_dir("s10")
    path = work / "inproc.accdb"
    t0 = time.perf_counter()
    engine = inproc("DAO.DBEngine.120")
    print(f"DAO.DBEngine.120 created in-proc in {time.perf_counter() - t0:.3f}s; Version={engine.Version}")
    db = engine.CreateDatabase(str(path), dbLangGeneral, 128)
    db.Execute("CREATE TABLE T (ID COUNTER PRIMARY KEY, Label TEXT(20))")
    db.Execute("INSERT INTO T (Label) VALUES ('x')")
    print(f"created DB Version={db.Version}; tables={[db.TableDefs.Item(i).Name for i in range(db.TableDefs.Count) if not db.TableDefs.Item(i).Name.startswith('MSys')]}")
    db.Close()
    del db

    # exclusive DAO + ADO in the same process
    db = engine.OpenDatabase(str(path), True, False)
    for provider in ("Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0"):
        conn = inproc("ADODB.Connection")
        try:
            conn.Open(f"Provider={provider};Data Source={path}")
            print(f"ADO {provider} while DAO holds exclusive: opened (unexpected?)")
            conn.Close()
        except pywintypes.com_error as exc:
            print(f"ADO {provider} while DAO holds exclusive: {describe_com_error(exc)}")
    db.Close()
    del db

    conn = inproc("ADODB.Connection")
    try:
        conn.Open(f"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={path}")
        conn.Execute("ALTER TABLE T ADD COLUMN Amount DECIMAL(12,3)")
        conn.Close()
        print("ADO DECIMAL DDL with DAO closed: OK")
    except pywintypes.com_error as exc:
        print(f"ADO DDL: {describe_com_error(exc)}")
    del conn

    db = engine.OpenDatabase(str(path), False, False)
    f = db.TableDefs("T").Fields("Amount")
    print(f"DAO sees Amount: Type={f.Type} Precision={f.Properties('Precision').Value} Scale={f.Properties('Scale').Value}")
    del f
    # shared DAO + ADO concurrently
    conn = inproc("ADODB.Connection")
    try:
        conn.Open(f"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={path}")
        conn.Execute("ALTER TABLE T ADD COLUMN Amount2 DECIMAL(8,2)")
        conn.Close()
        db.TableDefs.Refresh()
        print(f"ADO DDL while DAO holds *shared*: OK; DAO sees Amount2={db.TableDefs('T').Fields('Amount2').Properties('Precision').Value}")
    except pywintypes.com_error as exc:
        print(f"ADO DDL while DAO shared: {describe_com_error(exc)}")
    db.Close()
    del db, engine


if __name__ == "__main__":
    main()
