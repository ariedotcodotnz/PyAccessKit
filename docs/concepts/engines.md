# Engines

PyAccessKit uses DAO for all schema and data work, and Microsoft Access (`Access.Application`) for design
work: forms, modules, and `SaveAsText`/`LoadFromText`. DAO can be reached in three ways, and the choice
matters for bitness, licensing and safety.

| Transport (`db.transport`) | How | Good for | Caveats |
|---|---|---|---|
| `dao-inproc` | `DAO.DBEngine.120` loaded into Python | Fastest. No `MSACCESS.EXE`. Works with only the Access Runtime installed. | Python and Office must have the **same bitness** |
| `access-hosted` | A hidden Access instance owned by PyAccessKit hosts DAO; the database is opened with `DBEngine.OpenDatabase` | Any bitness combination. The database is not opened in the Access UI, so startup code never runs. | Starts one `MSACCESS.EXE` per session |
| `access-design` | The database is Access's *current database* (`OpenCurrentDatabase`) | Forms, modules, text import/export | Needs full Access, not the Runtime. Startup code is governed by `macro_security`. |

## Choosing an engine

```python
AccessDatabase.open("app.accdb")  # engine="auto" (default)
AccessDatabase.open("app.accdb", engine="dao")  # in-process DAO only; never starts Access
AccessDatabase.open("app.accdb", engine="access")  # always through Microsoft Access
```

- **`auto`** uses in-process DAO when this Python can load it, and Microsoft Access otherwise. The first
  design feature (a form, a module, text I/O, `db.raw.access`) switches the session to a design session.
- **`dao`** never starts Access. Design features raise `CapabilityError`. If DAO cannot be loaded, the
  session raises `DaoNotAvailableError` with an explanation such as *"DAO is installed for 32-bit
  programs only, but this Python is 64-bit"*.
- **`access`** always goes through an owned Access instance, using Access-hosted DAO for schema work until a
  design feature needs a design session.

## The switch to a design session

When a session switches engines, PyAccessKit releases the in-process DAO connection, starts (or reuses) its
own Access instance, and opens the database as the current database. Two consequences:

- **Handles keep working.** `db.tables["Customers"]` holds only a name, so it is valid after the switch.
- **Raw objects are revoked.** A `db.raw.dao` obtained before the switch raises `SessionClosedError`.
  Get a fresh one afterwards.

Read-only sessions cannot switch: opening a database for design needs exclusive, writable access.

## Capabilities

| Capability | In-process DAO | Access-hosted DAO | Design session |
|---|---|---|---|
| Create / open / close | ✅ | ✅ | ✅ |
| Tables, fields, indexes, field properties | ✅ | ✅ | ✅ |
| Decimal columns | ✅ (ADO + ACE OLEDB) | switches to design | ✅ (`CurrentProject.Connection`) |
| Relationships, saved and pass-through queries, execute/fetch | ✅ | ✅ | ✅ |
| List forms, reports, macros and modules | ✅ | ✅ | ✅ |
| Create or modify forms and modules; text import/export | ❌ | switches to design | ✅ |
| Runs the database's startup code | never | never | per `macro_security` (off by default) |

## Bitness

| Python | Office / Access Runtime | In-process DAO | Access transports |
|---|---|---|---|
| 64-bit | 64-bit | ✅ | ✅ |
| 32-bit | 32-bit | ✅ | ✅ |
| 64-bit | 32-bit | ❌ | ✅ |
| 32-bit | 64-bit | ❌ | ✅ |

Microsoft 365 is often installed as 32-bit, while Python installers default to 64-bit. PyAccessKit works in
that combination through Access. For in-process speed, add a matching Python, for example
`uv python install cpython-3.12-windows-x86`.

## Same result from every engine

DAO's `CreateDatabase` and Access's `NewCurrentDatabase` write slightly different database properties.
When DAO creates a database, PyAccessKit adds the properties Access itself writes (tabbed documents, and so
on). A database looks the same in Access whichever engine created it. To disable this, set
`SessionOptions(apply_native_defaults=False)`.
