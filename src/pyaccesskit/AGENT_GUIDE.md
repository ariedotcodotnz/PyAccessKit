# PyAccessKit — guide for AI coding agents

You are writing Python that builds or changes a Microsoft Access application with PyAccessKit. This guide
is the authoritative, compact reference for that job. Print it any time with `pyaccesskit guide`. Follow the
rules below exactly; they encode Access behaviour that is not obvious and that you cannot observe directly.

## 1. Workflow

1. **Check the machine** once: `pyaccesskit doctor --json`. Continue only if `"usable": true`.
   `"auto_engine"` says what `engine="auto"` uses (`"dao"` or `"access"`). Forms, modules and text
   import/export need full Microsoft Access (`engines[].engine == "access"` available).
2. **Understand an existing database** before changing it: `pyaccesskit inspect DB --json` (read-only, runs
   no startup code). Its `tables[].spec` entries are valid `TableSpec` JSON.
3. **Write the program** in three parts (see the complete example in section 9):
   *specs as data* → *one build function inside one `with` block* → *verification after reopening*.
4. **Run it**, read errors (section 7), fix, rerun. Building with `AccessDatabase.create(...)` is atomic,
   so a failed run leaves nothing behind and can simply be repeated.
5. **Verify**: `to_spec() == spec.normalized()` for tables, `fetch()` for queries, `check_opens()` for every
   form (this also compiles VBA).

Validate specs without Access: `pyaccesskit schema table` (JSON Schema), or construct the spec in Python —
invalid specs raise `SpecError` listing every problem.

## 2. Rules (MUST / NEVER)

Sessions and processes
- MUST open databases only with `with AccessDatabase.create(path, ...) as db:` or
  `with AccessDatabase.open(path, ...) as db:`. Everything is cleaned up on exit, also on exceptions.
- NEVER use `win32com.client.Dispatch`, `GetObject`, `DispatchEx` or `pythoncom` for Access. NEVER kill
  `MSACCESS.EXE` processes. PyAccessKit owns and closes its own Access processes; if a crash left one
  behind, run `pyaccesskit cleanup`.
- NEVER call `Quit`, `Close`, `CloseCurrentDatabase` on `db.raw.*` objects. Avoid `db.raw` unless the API
  truly lacks a feature.
- Use a session only on the thread that opened it. Do not open the same file in two sessions at once.
- `create()` refuses to replace an existing file unless `overwrite=True`.

Schema
- `Column.number(...)` is a Long Integer. Foreign keys that point to an `autonumber` use `Column.number`.
- Defaults: Python values (`default=0`, `default=True`, `default="pcs"`, `default=date(2026, 1, 1)`).
  Access expressions MUST be wrapped: `default=Expr("Now()")`, `default=Expr("Date()")`.
- Text length is 1–255 (`Column.text`); use `Column.long_text` for more. Decimal needs
  `precision` 1–28 and `scale <= precision`.
- At most one AutoNumber per table; name it `<Table>ID` by convention. Primary key index is named
  `PrimaryKey`; `unique=True`/`indexed=True` create an index named after the column.
- Names: ≤ 64 chars; no `.`, `!`, `` ` ``, `[`, `]`; no leading space. Names are case-insensitive:
  `Name` and `name` clash. AVOID reserved words and names like `Name`, `Date`, `Time`, `Value`, `Order`,
  `Level`, `Section`, `Note`, `Description`, `Type`, `Year`, `Month`, `User`, `Text` — use
  `CustomerName`, `OrderDate`, … PyAccessKit warns (`AccessNameWarning`); treat that warning as an error.
- Create tables before relationships, relationships before data, data before forms that need it.

SQL
- Access SQL (Jet/ACE), not ANSI/T-SQL: `TOP n` (no `LIMIT`), `IIf()` (no `CASE`), `&` concatenates strings,
  dates are `#2026-01-31#`, booleans `True`/`False`, `LIKE` wildcards are `*` and `?`, several joins need
  parentheses: `FROM (A INNER JOIN B ON ...) INNER JOIN C ON ...`.
- NEVER build SQL with f-strings or `+`. Pass values as parameters: SQL `[name]`, Python
  `db.execute(sql, {"name": value})`. Saved queries may declare `PARAMETERS [pX] Long;`.
- SQL that Python runs (`db.execute`, `db.fetch_all`, `Query.fetch`) goes through DAO outside the Access UI:
  use only engine functions (`IIf`, `IsNull`, `Sum`, `Format`, `Left`, `DateAdd`, `Year`…). `Nz()` and your
  own VBA functions are NOT available there (DAO error 3085). They do work in form control sources.

Forms and VBA
- Forms need full Access. Build them with `with db.forms.create(name, record_source=..., ...) as form:`;
  the form is saved when the block ends. Reuse a name only with `replace=True`.
- Buttons MUST have a name (`form.button("cmdSave", caption=..., on_click=Vba(...))`). Controls with events
  need plain VBA identifier names.
- `Vba("...")` holds a procedure **body** only — never `Sub`/`End Sub`. Put helpers and module-level
  variables in `form.module_code("...")` or in a standard module (`db.modules.create`).
- VBA is case-insensitive: a constant `PLATFORM` and a function `Platform` clash and the module will not
  compile (reported later as Access error 7960 or an `AccessDialogError`). Keep all names in a module
  unique ignoring case. `Option Compare Database` / `Option Explicit` are added for you.
- VBA text must be representable in the Windows ANSI code page (no emoji); otherwise `SpecError`.
- `default_view=FormView.CONTINUOUS` uses a tabular layout: labels go in the form header, so do not combine
  it with `header=False` unless every control has `label=False`.

## 3. Program skeleton

```python
from pathlib import Path

from pyaccesskit import AccessDatabase, Column, TableSpec

CUSTOMERS = TableSpec(
    name="Customers",
    columns=[
        Column.autonumber("CustomerID", primary_key=True),
        Column.text("CustomerName", length=120, required=True),
    ],
)


def build(db: AccessDatabase) -> None:
    db.tables.create(CUSTOMERS)


def verify(path: Path) -> None:
    with AccessDatabase.open(path, readonly=True) as db:
        assert db.tables["Customers"].to_spec() == CUSTOMERS.normalized()


def main(path: Path) -> None:
    with AccessDatabase.create(path, overwrite=True) as db:
        build(db)
    verify(path)
```

To change an existing database, open it writable and apply the change, then verify:

```python
from pyaccesskit import AccessDatabase, Column

with AccessDatabase.open("app.accdb") as db:
    if "Phone" not in db.tables["Customers"].fields:
        db.tables["Customers"].add_column(Column.text("Phone", length=30))
```

Open changes are not atomic across steps: back up the file first (copy it) when a change is risky.

## 4. API cheat sheet

```text
AccessDatabase.create(path, *, overwrite=False, atomic=True, engine="auto", options=None)
AccessDatabase.open(path, *, readonly=False, exclusive=None, password=None, engine="auto", options=None)
db.path  db.transport  db.format_version  db.readonly  db.is_open  db.access_pid
db.execute(sql, params=None) -> int                      # action SQL, returns affected rows
db.fetch_all(sql, params=None, *, limit=None) -> list[dict]
db.properties["AppTitle"] = "..." ; db.properties["StartUpForm"] = "frmMain" ; .get(name, default)

db.tables: names() | [name] | get(name) | in | len | iter | specs()
    create(TableSpec) | create(name, *, columns, indexes=(), primary_key=None, description=None,
                               validation_rule=None, validation_text=None, properties=None)
    drop(name, *, drop_relationships=False)
Table: name, fields, indexes, primary_key, description (settable), is_linked, properties, record_count(),
    to_spec(), add_column(col), drop_column(name), rename_column(old, new), create_index(IndexSpec),
    drop_index(name), rename(new), drop()
Field: name, data_type, size, required, spec, properties, rename(new), drop()

Column.text(name, *, length=255, allow_zero_length=False, unicode_compression=True, input_mask=None, **common)
Column.long_text(name, *, rich_text=False, append_only=False, **common)
Column.number(name, *, size=NumberSize.LONG_INTEGER, decimal_places=None, input_mask=None, **common)
Column.decimal(name, *, precision=18, scale=0, **common)      Column.currency(name, **common)
Column.autonumber(name, *, replication_id=False, primary_key=..., description=..., caption=...)
Column.date_time(name, **common)  Column.yes_no(name, **common)  Column.hyperlink(name, **common)
Column.ole_object(name, *, required=False, description=..., caption=...)
    common = required, default, validation_rule, validation_text, description, caption, format,
             primary_key, unique, indexed, properties={...}
IndexSpec.on(name, *columns, unique=False, ignore_nulls=False, required=False)   # ("Col", "desc") pairs
IndexSpec.primary_key(*columns)
TableSpec(name, columns, indexes=(), primary_key=None|"Col"|["A","B"], description, validation_rule, ...)

db.relationships.create("Primary.Col", "Foreign.Col", *, name=None, enforce_integrity=True,
    cascade_update=False, cascade_delete=False, one_to_one=False, join=JoinType.INNER)
db.relationships.create(RelationshipSpec.between("A.ID", "B.AID", cascade_delete=True))
    composite: create(("A", ["K1", "K2"]), ("B", ["K1", "K2"]))

db.queries.create(name, sql, *, description=None, replace=False) | create(QuerySpec, replace=False)
db.queries.create_pass_through(name, sql, *, connect="ODBC;...", returns_records=True, timeout=60)
Query: name, sql (settable), kind, parameters, description, execute(params) -> int,
    fetch(params=None, *, limit=None) -> list[dict], rename(new), drop(), to_spec()

db.forms.create(name, *, replace=False, **FormSpec options) -> FormBuilder (use as context manager)
    FormSpec options: record_source, caption, default_view (FormView.SINGLE|CONTINUOUS|DATASHEET|SPLIT),
    layout (LayoutKind.AUTO|STACKED|TABULAR|NONE), header, width, allow_additions, allow_edits,
    allow_deletions, data_entry, navigation_buttons, record_selectors, dividing_lines, scroll_bars,
    auto_center, pop_up, modal, option_explicit, properties={...}
FormBuilder: textbox(field=None, *, label=None|False|"text", control_source=None, format=None, enabled=True,
        locked=False, after_update=None, name=None, section=Section.DETAIL, at=(left, top), width, height,
        visible=True, properties={})
    checkbox(field, *, label, enabled, locked, after_update, ...)
    combobox(field, *, row_source, row_source_type=RowSourceType.TABLE_QUERY|VALUE_LIST, bound_column=1,
        column_count=1, column_widths=[cm(0), cm(4)], limit_to_list=True, label, ...)
    label(caption, ...)      button(name, *, caption, on_click=None, section, at, width, height)
    on_load(Vba) on_current(Vba) module_code(str) to_spec() save() discard()
db.forms.build(FormSpec, replace=False) ; db.forms[name]: controls(), check_opens(), export_text(),
    rename(new), drop()

db.modules.create(name, code, *, kind=ModuleKind.STANDARD|CLASS, replace=False) ; Module: code (settable),
    kind, rename(new), drop()
db.objects: names(kind) export_text(kind, name) import_text(kind, name, text, *, replace=False)
    save_text(kind, name, path) load_text(kind, name, path, *, replace=False) delete(kind, name)
    rename(kind, old, new)        kind = "form" | "report" | "macro" | "module" | "query"

Units: cm(2), mm(5), inch(1), pt(12), twips(1440); Length supports + - * / and comparisons.
Enums: FormView single|continuous|datasheet|split; LayoutKind auto|stacked|tabular|none;
    NumberSize byte|integer|long_integer|single|double|replication_id; Section detail|header|footer;
    RowSourceType table_query|value_list; ModuleKind standard|class; JoinType inner|left|right;
    QueryKind select|crosstab|delete|update|append|make_table|ddl|pass_through|union|...
SessionOptions(visible=False, macro_security=MacroSecurity.DISABLE, dialog_policy=DialogPolicy.FAIL,
    call_timeout=600.0, quit_timeout=30.0, access_progid="Access.Application")
```

All names above are importable from `pyaccesskit` (form control specs from `pyaccesskit.forms`).

## 5. Values in and out

| Access type | Python value written | Python value read |
|---|---|---|
| Short/Long Text, Hyperlink | `str` | `str` |
| Number (Byte/Integer/Long) | `int` | `int` |
| Number (Single/Double) | `float` | `float` |
| Decimal, Currency | `Decimal` (or `int`) | `Decimal` |
| Date/Time | naive `datetime` or `date` (wall-clock) | naive `datetime` |
| Yes/No | `bool` | `bool` |
| OLE Object | `bytes` | `bytes` |
| Replication ID (GUID) | `str` | `str` in DAO's form `"{guid {…}}"` |
| Null | `None` | `None` |

A `time` reads back as a `datetime` on 1899-12-30 (Access's day zero). Aggregates can come back as `float`
(e.g. `Sum` over integers through `IIf`). `bytes` parameters work in `db.execute`/`db.fetch_all`; a saved
query that receives bytes must declare the parameter: `PARAMETERS [payload] LongBinary;`.

## 6. Access limits

| Limit | Value |
|---|---|
| Object and column names | 64 characters |
| Columns per table | 255 |
| Indexes per table (including relationship indexes) | 32 |
| Columns per index | 10 |
| Short Text length | 255 |
| Form width / section height | 22 inches (`inch(22)`) |
| Database size | 2 GB |
| SQL statement | ~64,000 characters |

## 7. Errors and what to do

All exceptions derive from `pyaccesskit.PyAccessKitError`; `str(exc)` says what failed and why;
`exc.details.number` holds the Access/DAO error number when there is one.

| Exception | Typical cause | Fix |
|---|---|---|
| `SpecError` | invalid spec, name, option or default | read `exc.problems`, correct the spec |
| `ObjectExistsError` | name already used (tables and queries share one namespace) | pick another name, or `replace=True` where offered |
| `ObjectNotFoundError` | wrong table/column/query/form name | check `names()`; lookups are case-insensitive |
| `RelationshipError` | incompatible key types, no unique index on the primary side | primary side needs a PK/unique index; FK type `Column.number` for AutoNumber keys |
| `IntegrityViolationError` | duplicate key (3022), missing related row (3201), rows still related (3200) | fix the data or insert parents first |
| `SqlSyntaxError` | invalid Access SQL (`exc.sql` holds it) | apply the SQL rules in section 2 |
| `MissingParameterError` | a `[name]` in SQL without a value (often a misspelled column) | pass the parameter or fix the column name |
| `ComError` 3085 "Undefined function" | `Nz`/VBA function in SQL run from Python | use `IIf(IsNull(x), 0, x)` |
| `ComError` 7960 / `AccessDialogError` naming VBA | VBA does not compile | fix the module (duplicate names, missing `End Sub`, typos) |
| `AccessDialogError` | Access showed a modal dialog; `exc.dialogs` has title/text | usually VBA or a broken expression; read the text |
| `CapabilityError` | design feature with `engine="dao"` or `readonly=True` | use `engine="auto"`/`"access"`, open writable |
| `AccessRuntimeOnlyError` | only the Access Runtime is installed | forms/modules impossible; schema and data still work |
| `DatabaseLockedError` | file open elsewhere (exclusive) | close it in Access, or open `readonly=True` |
| `DatabaseExistsError` | target exists | `overwrite=True`, or another path |
| `AccessTimeoutError` | one call exceeded `call_timeout`; the owned Access was ended | split the work, raise `SessionOptions(call_timeout=...)` |
| `EngineUnavailableError` | no usable engine | run `pyaccesskit doctor` |

Warnings: `AccessNameWarning` (troublesome name) — rename; `AccessDialogWarning` (with
`dialog_policy="warn"`).

## 8. Not supported yet (do not attempt through PyAccessKit)

Reports, subforms, tab controls, list boxes, option groups, attachment / calculated / multi-value / lookup
columns, creating linked tables, macros (other than raw text import), ribbons, VBA references, compiling on
demand, encrypted database creation, changing a column's type or order (create a new table and copy rows
with `INSERT INTO ... SELECT` instead). For anything essential, `db.raw.access` / `db.raw.dao` expose the
underlying objects; save your own raw design changes.

## 9. Complete example

A small inventory application: three related tables, two queries (one with a parameter), a VBA module using
conditional compilation, a single form with lookups and events, a continuous form, startup settings, and
verification. It is `examples/04_inventory_app.py` in the source tree and is run by the test-suite.

<!-- example:start -->
```python
"""A complete small Access application, written the way the agent guide recommends.

Run:  python examples/04_inventory_app.py [path/to/inventory.accdb]

Structure: (1) the schema as immutable specs, (2) one build function, (3) verification that reopens the
file and checks what was built. Everything happens in one atomic ``create()``: if any step fails, no file
is left behind and the Access process PyAccessKit started is closed.
"""

from __future__ import annotations

import sys
from decimal import Decimal
from pathlib import Path

from pyaccesskit import (
    AccessDatabase,
    Column,
    Expr,
    FormView,
    NumberSize,
    RelationshipSpec,
    RowSourceType,
    TableSpec,
    Vba,
    cm,
)

# --- 1. Schema as data ---------------------------------------------------------------------------
TABLES = [
    TableSpec(
        name="Categories",
        columns=[
            Column.autonumber("CategoryID", primary_key=True),
            Column.text("CategoryName", length=60, required=True, unique=True),
        ],
    ),
    TableSpec(
        name="Products",
        description="Everything we stock",
        columns=[
            Column.autonumber("ProductID", primary_key=True),
            Column.text("ProductName", length=100, required=True, indexed=True),
            Column.number("CategoryID", required=True),
            Column.text("Unit", length=10, required=True, default="pcs"),
            Column.currency("UnitCost", default=0, validation_rule=">=0"),
            Column.number("ReorderLevel", size=NumberSize.INTEGER, default=5),
            Column.yes_no("Discontinued", default=False),
        ],
    ),
    TableSpec(
        name="StockMoves",
        columns=[
            Column.autonumber("MoveID", primary_key=True),
            Column.number("ProductID", required=True),
            Column.date_time("MovedAt", default=Expr("Now()"), format="General Date"),
            Column.number(
                "Quantity",
                required=True,
                validation_rule="<>0",
                validation_text="Use a positive number for receipts, negative for issues",
            ),
            Column.text("Reference", length=50),
        ],
    ),
]
RELATIONSHIPS = [
    RelationshipSpec.between("Categories.CategoryID", "Products.CategoryID"),
    RelationshipSpec.between("Products.ProductID", "StockMoves.ProductID", cascade_delete=True),
]
QUERIES = {
    "qryStockLevels": (
        # Only database-engine functions here: Nz() and VBA functions exist inside Access, not in DAO.
        "SELECT p.ProductID, p.ProductName, p.ReorderLevel,\n"
        "IIf(IsNull(Sum(m.Quantity)), 0, Sum(m.Quantity)) AS OnHand\n"
        "FROM Products AS p LEFT JOIN StockMoves AS m ON p.ProductID = m.ProductID\n"
        "GROUP BY p.ProductID, p.ProductName, p.ReorderLevel;"
    ),
    "qryLowStock": (
        "PARAMETERS [pMargin] Long;\n"
        "SELECT ProductName, OnHand, ReorderLevel FROM qryStockLevels\n"
        "WHERE OnHand <= ReorderLevel + [pMargin] ORDER BY ProductName;"
    ),
}
INVENTORY_MODULE = """\
Private lastRefresh As Date

#If Win64 Then
Private Const PLATFORM_NAME As String = "64-bit Office"
#Else
Private Const PLATFORM_NAME As String = "32-bit Office"
#End If

Public Function OnHand(ByVal productID As Long) As Long
    OnHand = Nz(DSum("Quantity", "StockMoves", "ProductID=" & productID), 0)
    lastRefresh = Now()
End Function

Public Function Platform() As String
    Platform = PLATFORM_NAME
End Function
"""


# --- 2. Build ------------------------------------------------------------------------------------
def build(db: AccessDatabase) -> None:
    for table in TABLES:
        db.tables.create(table)
    for relationship in RELATIONSHIPS:
        db.relationships.create(relationship)
    for name, sql in QUERIES.items():
        db.queries.create(name, sql)

    for category in ("Stationery", "Hardware"):
        db.execute("INSERT INTO Categories (CategoryName) VALUES ([name])", {"name": category})
    products = [("Notebook", 1, Decimal("1.20")), ("Stapler", 2, Decimal("6.50"))]
    for name, category, cost in products:
        db.execute(
            "INSERT INTO Products (ProductName, CategoryID, UnitCost) VALUES ([n], [c], [cost])",
            {"n": name, "c": category, "cost": cost},
        )
    db.execute("INSERT INTO StockMoves (ProductID, Quantity, Reference) VALUES (1, 40, 'PO-1')")
    db.execute("INSERT INTO StockMoves (ProductID, Quantity, Reference) VALUES (2, 3, 'PO-2')")

    db.modules.create("modInventory", INVENTORY_MODULE)

    with db.forms.create(
        "frmProducts", record_source="Products", caption="Products", width=cm(16)
    ) as form:
        form.textbox("ProductName", label="Product", width=cm(8))
        form.combobox(
            "CategoryID",
            label="Category",
            row_source="SELECT CategoryID, CategoryName FROM Categories ORDER BY CategoryName",
            column_count=2,
            column_widths=[cm(0), cm(5)],
        )
        form.combobox(
            "Unit", row_source='"pcs";"box";"kg"', row_source_type=RowSourceType.VALUE_LIST
        )
        form.textbox("UnitCost", label="Unit cost", format="Currency")
        form.textbox(
            name="txtOnHand",
            label="On hand",
            control_source="=OnHand([ProductID])",
            locked=True,
            enabled=False,
        )
        form.checkbox("Discontinued")
        form.button("cmdClose", caption="Close", on_click=Vba("DoCmd.Close acForm, Me.Name"))
        form.on_current(Vba("Me.txtOnHand.Requery"))

    with db.forms.create(
        "frmStockMoves",
        record_source="SELECT * FROM StockMoves ORDER BY MovedAt DESC",
        caption="Stock moves",
        default_view=FormView.CONTINUOUS,
    ) as form:
        form.textbox("MovedAt", width=cm(4))
        form.textbox("ProductID", width=cm(2))
        form.textbox("Quantity", width=cm(2))
        form.textbox("Reference", width=cm(4))

    db.properties["AppTitle"] = "Inventory"
    db.properties["StartUpForm"] = "frmProducts"


# --- 3. Verify -----------------------------------------------------------------------------------
def verify(path: Path) -> None:
    with AccessDatabase.open(path) as db:
        for table in TABLES:  # what Access stored is exactly what was specified
            assert db.tables[table.name].to_spec() == table.normalized(), table.name
        low = db.queries["qryLowStock"].fetch({"pMargin": 0})
        assert [row["ProductName"] for row in low] == ["Stapler"], low
        for name in db.forms.names():
            db.forms[name].check_opens()  # compiles the form's module and opens it hidden
        print(f"verified {path}: {db.tables.names()}, forms {db.forms.names()}, low stock {low}")


def main(target: Path) -> None:
    with AccessDatabase.create(target, overwrite=True) as db:
        build(db)
    verify(target)


if __name__ == "__main__":
    main(Path(sys.argv[1]) if len(sys.argv) > 1 else Path.cwd() / "inventory.accdb")
```
<!-- example:end -->
