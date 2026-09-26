"""Create a small CRM database: tables, a relationship, a saved query and some rows.

Run:  python examples/01_create_crm.py [path/to/crm.accdb]

Works with any engine: in-process DAO when your Python matches Office's bitness, otherwise a hidden Access
instance that PyAccessKit starts, owns and closes.
"""

from __future__ import annotations

import sys
from datetime import date
from decimal import Decimal
from pathlib import Path

from pyaccesskit import AccessDatabase, Column, Expr


def main(target: Path) -> None:
    # Atomic by default: the file appears only if the whole block succeeds.
    with AccessDatabase.create(target, overwrite=True) as db:
        db.tables.create(
            "Customers",
            columns=[
                Column.autonumber("CustomerID", primary_key=True),
                Column.text("CustomerName", length=120, required=True),
                Column.text("Email", length=255, unique=True),
                Column.yes_no("IsActive", default=True),
                Column.date_time("CreatedAt", default=Expr("Now()"), format="General Date"),
            ],
            description="People and companies we sell to",
        )
        db.tables.create(
            "Orders",
            columns=[
                Column.autonumber("OrderID", primary_key=True),
                Column.number("CustomerID", required=True),  # Long Integer, like Access's default
                Column.date_time("OrderDate", default=Expr("Date()"), format="Short Date"),
                Column.currency("Amount", default=0, validation_rule=">=0"),
            ],
        )
        db.relationships.create("Customers.CustomerID", "Orders.CustomerID", cascade_delete=True)
        db.queries.create(
            "qryCustomerTotals",
            "SELECT c.CustomerName, Count(o.OrderID) AS OrderCount, Sum(o.Amount) AS Total\n"
            "FROM Customers AS c LEFT JOIN Orders AS o ON c.CustomerID = o.CustomerID\n"
            "GROUP BY c.CustomerName;",
        )

        # Parameters are bound by name, never interpolated into the SQL.
        for name, email in [
            ("Ada Lovelace", "ada@example.com"),
            ("Alan Turing", "alan@example.com"),
        ]:
            db.execute(
                "INSERT INTO Customers (CustomerName, Email) VALUES ([name], [email])",
                {"name": name, "email": email},
            )
        db.execute(
            "INSERT INTO Orders (CustomerID, OrderDate, Amount) VALUES (1, [day], [amount])",
            {"day": date(2026, 1, 31), "amount": Decimal("149.95")},
        )

        for row in db.queries["qryCustomerTotals"].fetch():
            print(row)
    print(f"created {target}")


if __name__ == "__main__":
    main(Path(sys.argv[1]) if len(sys.argv) > 1 else Path.cwd() / "crm.accdb")
