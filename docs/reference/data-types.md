# Data types

How Access field types map to PyAccessKit column specs, DAO, and Python values.

| Access designer | Column spec (`type` in JSON) | DAO type | Python value in / out |
|---|---|---|---|
| Short Text | `Column.text` (`"text"`) | `dbText` (10) | `str` |
| Long Text | `Column.long_text` (`"long_text"`) | `dbMemo` (12) | `str` |
| Hyperlink | `Column.hyperlink` (`"hyperlink"`) | `dbMemo` + hyperlink flag | `str` (`display#address#`) |
| Number: Byte | `Column.number(size="byte")` | `dbByte` (2) | `int` 0–255 |
| Number: Integer | `Column.number(size="integer")` | `dbInteger` (3) | `int`, 16-bit |
| Number: Long Integer | `Column.number()` (default) | `dbLong` (4) | `int`, 32-bit |
| Number: Single | `Column.number(size="single")` | `dbSingle` (6) | `float` |
| Number: Double | `Column.number(size="double")` | `dbDouble` (7) | `float` |
| Number: Replication ID | `Column.number(size="replication_id")` | `dbGUID` (15) | `str`, DAO's form `"{guid {…}}"` |
| Number: Decimal | `Column.decimal(precision, scale)` (`"decimal"`) | `dbDecimal` (20), created with ADO | `Decimal` |
| Currency | `Column.currency` (`"currency"`) | `dbCurrency` (5) | `Decimal` (4 decimals) |
| AutoNumber | `Column.autonumber()` (`"autonumber"`) | `dbLong` + auto-increment | `int` (assigned by Access) |
| AutoNumber (Replication ID) | `Column.autonumber(replication_id=True)` | `dbGUID` + `GenGUID()` default | `str`, DAO's form `"{guid {…}}"` |
| Date/Time | `Column.date_time` (`"date_time"`) | `dbDate` (8) | naive `datetime`; `date` and `time` are accepted as input (a time reads back on 1899-12-30, Access's day zero) |
| Yes/No | `Column.yes_no` (`"yes_no"`) | `dbBoolean` (1) | `bool` |
| OLE Object | `Column.ole_object` (`"ole_object"`) | `dbLongBinary` (11) | `bytes` (`bytearray`/`memoryview` accepted) |
| Attachment, Calculated, multi-value/lookup, Large Number, Date/Time Extended | read only: `UnsupportedColumn` (`"unsupported"`) | various | — |

## Choosing types

- **Keys:** `Column.autonumber("XID", primary_key=True)` on the parent, `Column.number("XID")` (Long
  Integer) on the child. The types must match for a relationship.
- **Money:** `Column.currency` (exact, 4 decimals) or `Column.decimal(precision, scale)`. Do not use
  `double` for money.
- **Flags:** `Column.yes_no`. Access stores Yes as -1; PyAccessKit gives you `True`/`False`.
- **Long text:** `Column.long_text`. Short Text is limited to 255 characters.
- **Dates:** Access has no time zones. PyAccessKit writes and reads the wall-clock value of naive
  datetimes. Convert aware datetimes to local time yourself if that matters.

## Defaults

| Column | Literal defaults | Expression defaults |
|---|---|---|
| text, long_text, hyperlink | any `str` (stored as a quoted literal) | `Expr("...")` |
| number | `int` (`float` for single/double) | `Expr("...")` |
| decimal, currency | `int`, `Decimal`, `float` | `Expr("...")` |
| date_time | `date`, `datetime`, `time` | `Expr("Now()")`, `Expr("Date()")` |
| yes_no | `True` / `False` | — |
| autonumber, ole_object | not allowed | not allowed |

## Binary parameters

`bytes` values can be passed as parameters to `db.execute`/`db.fetch_all`; PyAccessKit declares those
parameters as `LongBinary` for you. A **saved** query must declare them itself, because DAO types undeclared
parameters as text, which would corrupt the bytes:

```sql
PARAMETERS [payload] LongBinary;
INSERT INTO Files (Payload) VALUES ([payload]);
```

Passing bytes to an undeclared parameter of a saved query raises `SpecError` instead of storing damaged data.
