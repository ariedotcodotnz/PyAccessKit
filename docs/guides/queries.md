# Queries & data

## Saved queries

```python
db.queries.create("qryActive", "SELECT * FROM Customers WHERE IsActive = True;")
db.queries.create("qryActive", "SELECT * FROM Customers WHERE IsActive;", replace=True)

query = db.queries["qryActive"]
query.kind  # QueryKind.SELECT (also UNION, CROSSTAB, UPDATE, DELETE, APPEND, MAKE_TABLE, DDL, PASS_THROUGH…)
query.sql  # the SQL as Access stores it
query.sql = "SELECT CustomerName FROM Customers;"
query.parameters  # declared and implicit parameters
query.rename("qryCustomerNames")
query.drop()
```

Tables and queries share one namespace in Access. Creating a query named like a table raises
`ObjectExistsError`.

Access rewrites SQL when it saves a query: it adds brackets, normalizes whitespace and may reorder
clauses. `query.sql` returns what Access stored. Do not compare it character for character with what you
wrote.

## Running SQL

```python
count = db.execute(
    "UPDATE Orders SET Amount = Amount * [factor] WHERE CustomerID = [id]",
    {"factor": 1.1, "id": 42},
)
rows = db.fetch_all("SELECT * FROM Orders WHERE Amount > [minimum]", {"minimum": 100}, limit=50)
```

- **Parameters are bound by name.** Any `[name]` that is not a column becomes a parameter. Values are
  never interpolated into the SQL, so there is no injection and no quoting trouble. An unknown name in
  `params` raises `SpecError`. A missing value raises `MissingParameterError`.
- `execute` runs with DAO's `dbFailOnError`: a statement that fails part-way raises instead of silently
  applying half its changes.
- `fetch_all` returns a list of dictionaries. Dates come back as naive `datetime`, currency as `Decimal`,
  and Null as `None`.

Saved queries run the same way:

```python
db.queries["qryRaisePrices"].execute({"factor": 1.05})
db.queries["qryOrdersSince"].fetch({"since": date(2026, 1, 1)}, limit=10)
```

This data access is deliberately minimal: enough for seed data, verification and tests. For heavy data
work, use a DB-API driver such as `pyodbc` with the Access ODBC driver.

## Pass-through queries

```python
db.queries.create_pass_through(
    "qryServerVersion",
    "SELECT @@VERSION",
    connect="ODBC;Driver={ODBC Driver 18 for SQL Server};Server=db01;Trusted_Connection=Yes;",
    returns_records=True,
    timeout=30,
)
```

`pyaccesskit inspect` hides passwords (`PWD=…`) in connection strings when it prints them.
