# PyAccessKit

A modern, typed, Pythonic toolkit for building and modifying Microsoft Access databases (`.accdb`) and,
progressively, whole Access applications without raw `pywin32` calls, DAO magic numbers or orphaned
`MSACCESS.EXE` processes.

```python
from pyaccesskit import AccessDatabase, Column, Expr, Vba, cm

with AccessDatabase.create("crm.accdb") as db:
    db.tables.create(
        "Customers",
        columns=[
            Column.autonumber("CustomerID", primary_key=True),
            Column.text("CustomerName", length=200, required=True),
            Column.text("Email", length=255, unique=True),
            Column.date_time("CreatedAt", default=Expr("Now()")),
        ],
    )
    db.tables.create(
        "Orders",
        columns=[
            Column.autonumber("OrderID", primary_key=True),
            Column.number("CustomerID", required=True),  # Long Integer, as in Access
            Column.currency("Total", default=0),
        ],
    )
    db.relationships.create("Customers.CustomerID", "Orders.CustomerID", cascade_delete=True)
    db.queries.create("qryCustomers", "SELECT * FROM Customers ORDER BY CustomerName;")

    with db.forms.create("frmCustomers", record_source="Customers", caption="Customers") as form:
        form.textbox("CustomerName", label="Customer name", width=cm(8))
        form.textbox("Email")
        form.button("cmdClose", caption="Close", on_click=Vba("DoCmd.Close acForm, Me.Name"))
```

When the `with` block ends, the database file appears at `crm.accdb`, and any Access process PyAccessKit
started is closed. If the block raises, no file is created and nothing is left running.

> **Status:** early development (`0.1.0.dev0`). The API may still change between 0.x releases.

## Installation

```console
pip install pyaccesskit
```

Requirements:

- Windows 10 or 11 and CPython 3.11–3.14 (64- or 32-bit; free-threaded builds are not supported).
- Microsoft Access 2016 or later, including Microsoft 365. Alternatively, the Microsoft 365 Access Runtime
  with a Python of the same bitness, for schema and data work only.

Check your machine with:

```console
pyaccesskit doctor          # which engines work here, and why not
pyaccesskit doctor --probe  # also builds a scratch database end to end
```

The spec models (`TableSpec`, `Column`, `FormSpec`, units, layout…) are pure Python and also import on
Linux and macOS, so you can validate specs in CI without Access.

## What you can do in 0.1

| Area | Highlights | Stability |
|---|---|---|
| Lifecycle | Atomic `create()`, `open()` (read-only/shared, exclusive, password), guaranteed cleanup, `db.raw` escape hatch | stable-track |
| Tables | Every common column type, defaults as Python values or `Expr`, validation rules, captions, formats, input masks; indexes (composite, descending, unique, primary); add, drop and rename columns and indexes; `to_spec()` round-trips | stable-track |
| Relationships | Single and composite keys, referential integrity, cascades, join types, all checked before Access sees them | stable-track |
| Queries & data | Saved queries (select, action, union, crosstab, pass-through), parameters bound by name, `execute()` / `fetch_all()` | stable-track |
| Forms | Labels, text boxes, check boxes, combo boxes and buttons; stacked or tabular layout; VBA event procedures; atomic build-then-swap replacement | provisional |
| Modules & text I/O | Standard and class modules; `SaveAsText`/`LoadFromText` for forms, reports, macros, queries and modules, stored as UTF-8 | provisional |
| Decimal columns | `Column.decimal(precision=…, scale=…)` via ADO DDL | provisional |

Reports, linked tables, VBA references, and the declarative `build`/`plan`/`apply` workflow are planned
(see the [roadmap](docs/index.md#roadmap)).

## Engines

PyAccessKit reaches the database in one of three ways and picks one for you (`engine="auto"`):

| Engine | When | Notes |
|---|---|---|
| In-process DAO | Python and Office have the **same bitness** | Fastest; never starts `MSACCESS.EXE`; enough for tables, relationships, queries and data |
| Access-hosted DAO | Bitness differs, or `engine="access"` | A hidden Access instance that PyAccessKit owns hosts DAO; the database is **not** opened in the Access UI, so no startup code runs |
| Design session | First use of forms, modules or text I/O | The database becomes Access's current database; handles such as `db.tables["X"]` keep working across the switch |

| Python | Office | In-process DAO | Access transports |
|---|---|---|---|
| 64-bit | 64-bit | ✅ | ✅ |
| 32-bit | 32-bit | ✅ | ✅ |
| 64-bit | 32-bit (common with Microsoft 365) | ❌ | ✅ |
| 32-bit | 64-bit | ❌ | ✅ |

`pyaccesskit doctor` explains what applies to your machine.

## Safety guarantees

- **Never touches your own Access windows.** PyAccessKit always starts a *new* Access instance
  (`CoCreateInstanceEx`, local server). It never attaches to a running Access (which `Dispatch()` does),
  so it can never close your work.
- **No orphaned `MSACCESS.EXE`.** Every Access process it starts is identified by PID, creation time and
  image, and placed in a kill-on-close job object. If Python crashes, Windows ends the process. The process
  is also recorded in an ownership ledger: `pyaccesskit cleanup` ends leftovers from crashed sessions, and
  only when their Python owner is gone. Processes it did not start are never touched.
- **Exceptions clean up too.** Closing runs every cleanup step even if one fails. Cleanup failures are
  attached as notes to the exception that caused the close, never masking it.
- **Atomic creation.** `AccessDatabase.create()` builds in a hidden sibling file and moves it into place
  only on success. Tables are created all-or-nothing, and forms are built under a temporary name and
  swapped in.
- **No hangs on hidden dialogs.** A watchdog watches the windows of *our* Access process. Unexpected
  dialogs are dismissed and reported as `AccessDialogError` with their text. Calls that exceed
  `call_timeout` end the owned process, and Ctrl+C works during long calls.
- **No startup code by default.** Databases are opened with macros disabled. Read-only sessions never open
  the database in the Access UI, so `AutoExec` and startup forms do not run.

## Errors you can act on

COM errors are translated into specific exceptions. Each one carries what PyAccessKit was doing, the
object involved, and the original Access or DAO error:

```python
from pyaccesskit import AccessDatabase, IntegrityViolationError

with AccessDatabase.open("crm.accdb") as db:
    try:
        db.execute("INSERT INTO Orders (CustomerID, Total) VALUES (999, 10)")
    except IntegrityViolationError as exc:
        print(exc)  # what failed, with Access's own explanation
        print(exc.details.number)  # 3201: there is no related record in Customers
```

Most mistakes are caught before Access is involved at all. A duplicate table name, an unknown column in a
relationship, or incompatible key types each raise a precise `ObjectExistsError`, `SpecError` or
`RelationshipError`.

## Command line

```console
pyaccesskit doctor [--json] [--probe]      # environment report (exit code 3 if nothing works)
pyaccesskit inspect DB [--json] [--counts] # tables, relationships, queries, objects; read-only
pyaccesskit cleanup [--dry-run] [--json]   # end orphaned Access processes started by PyAccessKit
```

Exit codes: `0` success, `1` error, `2` usage error, `3` environment unusable.

## Documentation

- [Getting started](docs/getting-started.md)
- Concepts: [engines](docs/concepts/engines.md), [lifecycle](docs/concepts/lifecycle.md),
  [specs](docs/concepts/specs.md)
- Guides: [tables](docs/guides/tables.md), [relationships](docs/guides/relationships.md),
  [queries](docs/guides/queries.md), [forms](docs/guides/forms.md), [modules](docs/guides/modules.md),
  [text I/O](docs/guides/text-io.md)
- [Environment & troubleshooting](docs/environment.md), [contributing](docs/contributing.md),
  [architecture decisions](docs/adr/)
- Examples: [`examples/`](examples/)

## License

MIT. See [LICENSE](LICENSE).
