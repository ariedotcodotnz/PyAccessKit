# Recipes

Short, complete answers to common tasks. Each snippet assumes `from pathlib import Path` and the imports
it shows.

## Load rows from a CSV file

```python
import csv
from decimal import Decimal
from pathlib import Path

from pyaccesskit import AccessDatabase

with (
    AccessDatabase.open("shop.accdb") as db,
    Path("products.csv").open(newline="", encoding="utf-8") as f,
):
    for row in csv.DictReader(f):
        db.execute(
            "INSERT INTO Products (ProductName, UnitPrice) VALUES ([name], [price])",
            {"name": row["name"], "price": Decimal(row["price"])},
        )
```

Convert CSV strings to real Python types first (`Decimal`, `int`, `date.fromisoformat`...). Parameters keep
their Python type; passing the string `"12.50"` would leave the conversion to Access, which uses the
Windows locale (a decimal comma on many systems).

## Change an existing database safely and repeatably

Write changes so that running them twice does no harm, and keep a backup:

```python
import shutil

from pyaccesskit import AccessDatabase, Column, IndexSpec

shutil.copy2("app.accdb", "app.backup.accdb")
with AccessDatabase.open("app.accdb") as db:
    customers = db.tables["Customers"]
    if "Phone" not in customers.fields:
        customers.add_column(Column.text("Phone", length=30))
    if not any(index.name == "ixCity" for index in customers.indexes):
        customers.create_index(IndexSpec.on("ixCity", "City"))
    if "qryCustomersByCity" not in db.queries:
        db.queries.create(
            "qryCustomersByCity", "SELECT * FROM Customers ORDER BY City, CustomerName;"
        )
```

Changes to an opened database are applied one by one (they are not a single transaction), which is why the
backup matters. `AccessDatabase.create()` is the atomic path: rebuilding from specs is often simpler.

## Snapshot a schema as JSON and rebuild it

```python
import json
from pathlib import Path

from pyaccesskit import AccessDatabase, RelationshipSpec, TableSpec

with AccessDatabase.open("app.accdb", readonly=True) as db:
    snapshot = {
        "tables": [spec.model_dump(mode="json") for spec in db.tables.specs()],
        "relationships": [spec.model_dump(mode="json") for spec in db.relationships.specs()],
    }
Path("schema.json").write_text(json.dumps(snapshot, indent=2), encoding="utf-8")

data = json.loads(Path("schema.json").read_text(encoding="utf-8"))
with AccessDatabase.create("copy.accdb", overwrite=True) as db:
    for table in data["tables"]:
        db.tables.create(TableSpec.model_validate(table))
    for relationship in data["relationships"]:
        db.relationships.create(RelationshipSpec.model_validate(relationship))
```

## Export every object as text (for source control)

```python
from pathlib import Path

from pyaccesskit import AccessDatabase

out = Path("src-access")
with AccessDatabase.open("app.accdb") as db:
    for kind in ("form", "report", "macro", "module", "query"):
        folder = out / f"{kind}s"
        folder.mkdir(parents=True, exist_ok=True)
        for name in db.objects.names(kind):
            db.objects.save_text(kind, name, folder / f"{name}.txt")
```

Text export needs Microsoft Access (the session switches to a design session), so the database is opened
writable and exclusively.

## Copy forms and modules between databases

```python
from pyaccesskit import AccessDatabase

with AccessDatabase.open("template.accdb") as source:
    texts = {name: source.objects.export_text("form", name) for name in source.forms.names()}
with AccessDatabase.open("app.accdb") as target:
    for name, text in texts.items():
        target.objects.import_text("form", name, text, replace=True)
```

Forms reference tables and queries by name: create those in the target first.

## Lookup lists on a form

```python
from pyaccesskit import AccessDatabase, RowSourceType, cm

with (
    AccessDatabase.open("app.accdb") as db,
    db.forms.create("frmOrders", record_source="Orders", replace=True) as form,
):
    form.combobox(
        "CustomerID",
        label="Customer",
        row_source="SELECT CustomerID, CustomerName FROM Customers ORDER BY CustomerName",
        column_count=2,
        column_widths=[cm(0), cm(6)],  # store the ID, show the name
    )
    form.combobox(
        "Status", row_source='"Open";"Shipped";"Closed"', row_source_type=RowSourceType.VALUE_LIST
    )
```

## Watch Access while debugging

```python
from pyaccesskit import AccessDatabase, DialogPolicy, SessionOptions

debug = SessionOptions(visible=True, dialog_policy=DialogPolicy.WARN, call_timeout=None)
with AccessDatabase.open("app.accdb", engine="access", options=debug) as db:
    ...
```

With `visible=True` you see what Access does; `DialogPolicy.WARN` dismisses dialogs but only warns.

## Treat questionable names as errors

```python
import warnings

from pyaccesskit import AccessNameWarning

warnings.simplefilter("error", AccessNameWarning)  # "Name", "Date", spaces... now raise immediately
```

## Password-protected databases

```python
import os

from pyaccesskit import AccessDatabase

with AccessDatabase.open("secret.accdb", password=os.environ["APP_DB_PASSWORD"]) as db:
    print(db.tables.names())
```

Keep passwords out of source code; `pyaccesskit inspect` reads `PYACCESSKIT_PASSWORD`.

## Check the environment from code

```python
from pyaccesskit import diagnose

report = diagnose()
if not report.usable:
    raise SystemExit("\n".join(report.problems))
print("engine='auto' will use", report.auto_engine)
```
