# Relationships

```python
db.relationships.create("Customers.CustomerID", "Orders.CustomerID")
db.relationships.create(
    "Orders.OrderID",
    "OrderLines.OrderID",
    cascade_delete=True,
)
```

The first argument is the *primary* side (the "one" side, which needs a primary key or unique index on
those columns). The second is the *foreign* side (the "many" side).

## Options

| Option | Default | Meaning |
|---|---|---|
| `name` | primary table + foreign table (`CustomersOrders`), like Access | Relationship name |
| `enforce_integrity` | `True` | Enforce referential integrity (Access then adds a hidden index on the foreign side) |
| `cascade_update` | `False` | Cascade key updates to related rows |
| `cascade_delete` | `False` | Delete related rows along with the primary row |
| `one_to_one` | `False` | Declare a one-to-one relationship |
| `join` | `JoinType.INNER` | Default join type for the query designer (`INNER`, `LEFT`, `RIGHT`) |

## Composite keys

Pass `(table, [columns])` pairs, or build a `RelationshipSpec`:

```python
from pyaccesskit import RelationshipSpec

db.relationships.create(
    ("Enrollment", ["PersonID", "CourseID"]), ("Grades", ["PersonID", "CourseID"])
)

spec = RelationshipSpec.between("People.PersonID", "Enrollment.PersonID", cascade_delete=True)
db.relationships.create(spec)
```

## Checked before Access sees them

Before calling DAO, PyAccessKit verifies that:

- both tables and all columns exist, and the column counts match
- the column types are compatible (AutoNumber matches Long Integer, Replication ID matches a GUID number)
- the primary side has a primary key or unique index on exactly those columns
- the name is free, and the foreign table has index capacity left

Violations raise `RelationshipError` or `SpecError` with a readable explanation. If existing rows break the
rule you are enforcing, Access refuses the relationship. That surfaces as `IntegrityViolationError`.

## Reading and dropping

```python
for relationship in db.relationships:
    print(relationship.to_spec())
db.relationships.drop("CustomersOrders")
```

Tables involved in relationships cannot be dropped until the relationships are gone.
`db.tables.drop(name, drop_relationships=True)` removes them first.
