# Getting started

## Install

```console
pip install pyaccesskit
pyaccesskit doctor
```

`doctor` tells you which engines work on your machine and why (see [Engines](concepts/engines.md)). If it
reports that PyAccessKit is ready, everything below works.

## Create a database

```python
from datetime import date
from decimal import Decimal

from pyaccesskit import AccessDatabase, Column, Expr

with AccessDatabase.create("crm.accdb") as db:
    db.tables.create(
        "Customers",
        columns=[
            Column.autonumber("CustomerID", primary_key=True),
            Column.text("CustomerName", length=120, required=True),
            Column.text("Email", length=255, unique=True),
            Column.yes_no("IsActive", default=True),
            Column.date_time("CreatedAt", default=Expr("Now()")),
        ],
        description="People and companies we sell to",
    )
    db.tables.create(
        "Orders",
        columns=[
            Column.autonumber("OrderID", primary_key=True),
            Column.number("CustomerID", required=True),
            Column.date_time("OrderDate", default=Expr("Date()")),
            Column.currency("Amount", default=0, validation_rule=">=0"),
        ],
    )
    db.relationships.create("Customers.CustomerID", "Orders.CustomerID", cascade_delete=True)
    db.queries.create(
        "qryCustomerTotals",
        "SELECT c.CustomerName, Sum(o.Amount) AS Total "
        "FROM Customers AS c LEFT JOIN Orders AS o ON c.CustomerID = o.CustomerID "
        "GROUP BY c.CustomerName;",
    )

    db.execute(
        "INSERT INTO Customers (CustomerName, Email) VALUES ([name], [email])",
        {"name": "Ada Lovelace", "email": "ada@example.com"},
    )
    db.execute(
        "INSERT INTO Orders (CustomerID, OrderDate, Amount) VALUES (1, [day], [amount])",
        {"day": date(2026, 1, 31), "amount": Decimal("149.95")},
    )
    print(db.queries["qryCustomerTotals"].fetch())
```

A few things to notice:

- **The `with` block matters.** The file appears at `crm.accdb` only when the block ends without an
  exception (atomic creation). Any Access process PyAccessKit started is closed either way.
- **Defaults are Python values**, rendered as correct Access literals: `True`, `0`, or a quoted string.
  Expressions are wrapped in `Expr(...)`.
- **Parameters are bound by name.** `[name]` in the SQL is an Access parameter, and values are never
  pasted into the SQL text.
- **`Column.number` is a Long Integer by default**, like the Access table designer. Pass
  `size=NumberSize.INTEGER` for Access's 16-bit *Integer*.

## Open an existing database

```python
with AccessDatabase.open("crm.accdb", readonly=True) as db:
    customers = db.tables["customers"]  # names are case-insensitive, as in Access
    print(customers.fields["Email"].size)
    print(customers.to_spec().model_dump_json(indent=2))
    for query in db.queries:
        print(query.name, query.kind, query.sql)
```

Read-only sessions open the file in shared mode, so it can stay open in Access. They never run the
database's startup code.

## Add a form and some VBA

Forms and modules need Microsoft Access (the full product, not the Runtime). The session switches to
Access automatically the first time you use them:

```python
from pyaccesskit import AccessDatabase, Vba, cm

with AccessDatabase.open("crm.accdb") as db:
    db.modules.create(
        "modFormatting",
        'Public Function Money(ByVal v As Currency) As String\n    Money = Format$(v, "Currency")\nEnd Function\n',
    )
    with db.forms.create("frmCustomers", record_source="Customers", caption="Customers") as form:
        form.textbox("CustomerName", label="Customer name", width=cm(8))
        form.textbox("Email", width=cm(8))
        form.checkbox("IsActive")
        form.button("cmdClose", caption="Close", on_click=Vba("DoCmd.Close acForm, Me.Name"))
    db.forms["frmCustomers"].check_opens()  # opens it hidden in Form view, then closes it
```

## The escape hatch

Anything PyAccessKit doesn't cover yet is reachable through `db.raw`:

```python
with AccessDatabase.open("crm.accdb") as db:
    dao_database = db.raw.dao  # the DAO Database
    access = db.raw.access  # Access.Application (switches to a design session)
    print(access.Version)
```

Raw objects belong to the session. They stop working when it closes or switches engines, and you must not
call `Quit()` or `Close()` on them. Save your own raw design changes: PyAccessKit quits Access without
saving.

## Next steps

- [Engines](concepts/engines.md): what happens under the hood, and bitness.
- [Lifecycle & safety](concepts/lifecycle.md): what is cleaned up when, and the session options.
- The guides for [tables](guides/tables.md), [relationships](guides/relationships.md),
  [queries](guides/queries.md), [forms](guides/forms.md), [modules](guides/modules.md) and
  [text I/O](guides/text-io.md).
- The runnable scripts in the `examples/` folder of the source tree.
