# Tables

## Column types

| Access designer | PyAccessKit | Notes |
|---|---|---|
| Short Text | `Column.text(name, length=255)` | `allow_zero_length=False`, `unicode_compression=True`, `input_mask=` |
| Long Text | `Column.long_text(name, rich_text=False, append_only=False)` | |
| Number | `Column.number(name, size=NumberSize.LONG_INTEGER)` | Sizes: `BYTE`, `INTEGER` (16-bit), `LONG_INTEGER`, `SINGLE`, `DOUBLE`, `REPLICATION_ID`; `decimal_places=` |
| Number (Decimal) | `Column.decimal(name, precision=18, scale=0)` | Created with ADO DDL (DAO cannot create decimals) |
| Currency | `Column.currency(name)` | Exact, 4 decimal places |
| AutoNumber | `Column.autonumber(name, replication_id=False)` | Incrementing Long Integer, or a GUID (Replication ID) |
| Date/Time | `Column.date_time(name)` | |
| Yes/No | `Column.yes_no(name)` | Shown as a check box, as in Access |
| Hyperlink | `Column.hyperlink(name)` | |
| OLE Object | `Column.ole_object(name)` | |

Every column also accepts `required`, `default`, `validation_rule`, `validation_text`, `description`,
`caption`, `format`, and `properties={...}` (other Access field properties, set verbatim). Options that do
not apply to a type are rejected.

Columns PyAccessKit cannot create yet (Attachment, calculated, multi-valued, Large Number, Date/Time
Extended) still appear in `to_spec()`, as `UnsupportedColumn` entries that name the Access type, so
introspection never silently drops a field.

## Defaults and expressions

```python
from datetime import date
from pyaccesskit import Column, Expr

Column.text("Country", default="Belgium")  # rendered as "Belgium"
Column.text("Name", default='O"Brien')  # quotes are escaped for you
Column.currency("Total", default=0)
Column.yes_no("IsActive", default=True)
Column.date_time("Since", default=date(2026, 1, 1))
Column.date_time("CreatedAt", default=Expr("Now()"))  # an Access expression, passed verbatim
```

On Text columns, a string default is always a *value*: `default="Now()"` means the four characters
`Now()`. On other columns, a string must be a literal of the column's type (`"42"`, `"True"`,
`"#2026-01-31#"`). Anything else, such as `default="Now()"` on a Date/Time column, is rejected with a
hint to use `Expr(...)`, which prevents the classic unquoted-default bug.

## Keys and indexes

```python
from pyaccesskit import Column, IndexSpec

db.tables.create(
    "People",
    columns=[
        Column.autonumber("PersonID", primary_key=True),
        Column.text("LastName", length=80, indexed=True),  # non-unique index "LastName"
        Column.text("FirstName", length=80),
        Column.text("Email", length=255, unique=True),  # unique index "Email"
    ],
    indexes=[IndexSpec.on("ixFullName", "LastName", "FirstName")],
)

db.tables.create(
    "Enrollment",
    columns=[Column.number("PersonID"), Column.number("CourseID")],
    primary_key=["PersonID", "CourseID"],  # composite primary key, named PrimaryKey as in Access
)
```

`IndexSpec.on(name, *columns, unique=False, ignore_nulls=False, required=False)` accepts
`("Column", "desc")` pairs for descending keys. Limits are checked up front: 32 indexes per table,
10 columns per index, one primary key.

## Changing tables

```python
table = db.tables["People"]
table.add_column(Column.text("Phone", length=30))
table.rename_column("Phone", "Mobile")
table.drop_column("Mobile")
table.create_index(IndexSpec.on("ixEmailLast", "Email", "LastName", unique=True))
table.drop_index("ixEmailLast")
table.description = "Everyone we know"
table.rename("Contacts")
db.tables.drop("Contacts", drop_relationships=True)
```

Table creation is atomic: the table, its fields and its indexes are built and appended in one step. If
anything fails afterwards (for example, a property), the table is removed again. Changing a column's type
or order is not supported yet. It needs copy-and-migrate, which will come with `plan`/`apply`.

## Properties

Access stores many settings as DAO properties that only exist once set. `properties` bags read and write
them:

```python
from pyaccesskit import PropertyType

field = db.tables["People"].fields["Email"]
field.properties["Caption"] = "E-mail address"
db.properties["AppTitle"] = "My application"
db.tables["People"].properties.set(
    "MyTag", 7, PropertyType.BYTE
)  # force the DAO type of a new property
print(db.properties.get("StartUpForm"))  # None if not set
```

## Reading a table back

```python
table = db.tables["People"]
[field.name for field in table.fields]
table.fields["Email"].data_type, table.fields["Email"].size
table.primary_key, table.indexes
table.record_count()
spec = table.to_spec()  # canonical TableSpec; relationship-owned indexes are left out
```
