# Specs, handles & collections

PyAccessKit has three kinds of objects.

## Specs: what something should look like

Specs are immutable, validated [Pydantic](https://docs.pydantic.dev/) models: `TableSpec`, the column
types returned by `Column.*`, `IndexSpec`, `RelationshipSpec`, `QuerySpec` and `FormSpec`.

```python
from pyaccesskit import Column, TableSpec

customers = TableSpec(
    name="Customers",
    columns=[
        Column.autonumber("CustomerID", primary_key=True),
        Column.text("Email", length=255, unique=True),
    ],
)
```

- **Validated on construction.** Invalid names, a second AutoNumber, `length=300` for Short Text, a
  duplicate column, or options that do not apply to a type (`Column.yes_no("x", length=5)`) raise
  `SpecError` with every problem listed. Nothing touches Access.
- **Serializable.** `customers.model_dump_json()` and `TableSpec.model_validate_json(...)` round-trip.
  `TableSpec.model_json_schema()` describes the format for editors and AI tools. Lengths serialize as
  strings such as `"2cm"`.
- **Canonical.** `spec.normalized()` expands shorthands (`primary_key=True`, `unique=True` on columns)
  into explicit `IndexSpec`s. `Table.to_spec()` returns exactly that normal form, so
  `db.tables.create(spec)` followed by `db.tables[name].to_spec() == spec.normalized()` holds. Future
  `plan`/`apply` will be built on this.

## Handles: live, name-based views

`db.tables["Customers"]`, `table.fields["Email"]`, `db.queries["qryX"]`, `db.forms["frmX"]` and
`db.modules["modX"]` are handles. A handle holds only the session and the object's name, never a COM
object. It stays valid when the session switches engines, and it reads the current state each time you ask.

```python
table = db.tables["customers"]  # case-insensitive, like Access
table.add_column(Column.text("Phone", length=30))
table.rename("Clients")  # the handle follows the rename
spec = table.to_spec()  # an immutable snapshot
```

## Collections

`db.tables`, `db.relationships`, `db.queries`, `db.forms`, `db.modules` and `db.objects` support
`names()`, iteration, `len()`, `in`, `[]` and `get()`, plus `create(...)` and `drop(...)`.
`create` accepts either a spec or keyword arguments:

```python
db.tables.create(customers)  # a spec
db.tables.create("Customers", columns=[...], primary_key="ID")  # keyword arguments
```

## Names

Access compares names case-insensitively. PyAccessKit does the same when looking things up and when
checking for duplicates. Names are validated before Access sees them: at most 64 characters, none of
`` . ! ` [ ] ``, no leading space, no control characters. Legal but troublesome names, such as reserved
words (`Name`, `Date`, `Order`…) or names with spaces, trigger an `AccessNameWarning` pointing at your
code.

## Units

Access measures layouts in twips (1/1440 inch). PyAccessKit uses a `Length` type instead:

```python
from pyaccesskit import cm, inch, mm, pt

width = cm(8) + mm(5)
print(width.twips, width.cm, width.format("in"))
```

`Length.parse("2.5cm")` accepts `tw`, `pt`, `mm`, `cm` and `in`. Arithmetic and comparisons work as
expected.
