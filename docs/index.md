# PyAccessKit

PyAccessKit is a typed, Pythonic toolkit for creating and modifying Microsoft Access databases (`.accdb`)
and, progressively, whole Access applications. You work with databases, tables, columns, indexes,
relationships, queries, forms, controls and modules. Raw COM objects, DAO type codes and `MSACCESS.EXE`
processes stay out of your code.

```python
from pyaccesskit import AccessDatabase, Column

with AccessDatabase.create("crm.accdb") as db:
    db.tables.create(
        "Customers",
        columns=[
            Column.autonumber("CustomerID", primary_key=True),
            Column.text("CustomerName", length=200, required=True),
        ],
    )
```

## Why

Automating Access from Python usually means `win32com.client.Dispatch("Access.Application")`, then calls
like `CreateField("x", 10, 255)` and `CreateControl(frm, 109, 0, ...)`. After that you are catching
`com_error` tuples, and finding that `MSACCESS.EXE` is still running after the script has crashed. Worse,
`Dispatch` attaches to an Access window the user already has open, so the script's `Quit()` can close the
user's work.

PyAccessKit replaces that with:

- **Specs**: immutable, validated, serializable descriptions (`TableSpec`, `Column.text(...)`,
  `RelationshipSpec`, `FormSpec`…). Most mistakes are caught before Access is involved.
- **Handles and collections**: `db.tables["Customers"].fields["Email"]`, case-insensitive like Access.
  Handles are name-based and never hold COM objects.
- **An isolated, typed COM boundary**: late binding, exact-argument `IDispatch` calls, and COM errors
  translated into specific exceptions that carry the Access or DAO error number.
- **Process ownership**: every Access process is started fresh, tracked by identity, placed in a
  kill-on-close job object, and closed even when your code raises.

## Built for AI coding agents

Specs are plain, validated data with a JSON Schema, and errors explain themselves. Builds are atomic and
every Access process is cleaned up, so an agent can write, run and fix code in a loop without leaving
damage behind. `pyaccesskit guide` prints a compact, version-matched guide that tells an agent how to write
Access applications with the library. See [Building with AI agents](agents/index.md).

## Design principles

1. **The COM world is quarantined.** Only a few internal packages import `pywin32`. Everything else,
   including every spec, is pure Python, importable on any OS, and strictly type-checked.
2. **Validate before COM.** Names, limits, types, relationships and form layouts are checked in Python
   first.
3. **Atomic where Access allows it.** Databases are built in a temporary file, tables are created
   all-or-nothing, and forms are built under a temporary name and swapped in.
4. **One owner.** A session owns the engine and the Access process. You never need to call `Quit()`.
5. **One model for code and files.** The same spec models drive today's imperative API and tomorrow's
   declarative project format.

## Roadmap

| Version | Planned |
|---|---|
| **0.1** | Lifecycle and safety, tables, relationships, queries and data, small form API, VBA modules, text import/export, CLI (`doctor`, `inspect`, `cleanup`) |
| 0.2 | Reports, calculated/attachment/lookup columns, linked tables, database settings spec, VBA references and compile checks, `Application.Run`, `db.detach(show=True)`, public fake engine for tests |
| 0.3 | Project format (YAML/SQL/VBA files), JSON Schema, `pyaccesskit export` and `pyaccesskit build` |
| 0.4 | `pyaccesskit plan` / `apply`: desired-state schema changes, replace-if-changed forms and modules |
| 0.5+ | Native form/report specs as the main authoring path, seed data, ribbons, navigation pane, themes, data macros |

The project format will be built on the same specs you use today. `Table.to_spec()` already returns a
canonical, round-trippable `TableSpec`.
