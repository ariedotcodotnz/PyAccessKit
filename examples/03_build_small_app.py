"""Build a small Access application: tables, relationships, queries, a VBA module and two forms.

Run:  python examples/03_build_small_app.py [path/to/orders.accdb]

Forms and modules need Microsoft Access (the full product, not the Runtime). With engine="auto" the
session starts with DAO and switches to an owned, hidden Access instance at the first design feature.
Open the result in Access afterwards: the startup form shows customers, and the product list is a
continuous form.
"""

from __future__ import annotations

import sys
from decimal import Decimal
from pathlib import Path

from pyaccesskit import AccessDatabase, Column, Expr, FormView, IndexSpec, NumberSize, Vba, cm

FORMATTING_MODULE = """
Public Function FormatMoney(ByVal amount As Currency) As String
    FormatMoney = Format$(amount, "Currency")
End Function
"""


def build_schema(db: AccessDatabase) -> None:
    db.tables.create(
        "Customers",
        columns=[
            Column.autonumber("CustomerID", primary_key=True),
            Column.text("CustomerName", length=120, required=True, caption="Customer"),
            Column.text("Email", length=255, unique=True),
            Column.text("Phone", length=30, input_mask="!(999) 000-0000;;_"),
            Column.yes_no("IsActive", default=True, caption="Active"),
        ],
    )
    db.tables.create(
        "Products",
        columns=[
            Column.autonumber("ProductID", primary_key=True),
            Column.text("ProductName", length=100, required=True, indexed=True),
            Column.currency("UnitPrice", default=0, validation_rule=">=0"),
            Column.yes_no("Discontinued", default=False),
        ],
    )
    db.tables.create(
        "Orders",
        columns=[
            Column.autonumber("OrderID", primary_key=True),
            Column.number("CustomerID", required=True),
            Column.date_time("OrderDate", default=Expr("Date()"), format="Short Date"),
            Column.long_text("Notes"),
        ],
    )
    db.tables.create(
        "OrderLines",
        columns=[
            Column.number("OrderID", required=True),
            Column.number("ProductID", required=True),
            Column.number("Quantity", size=NumberSize.INTEGER, default=1, validation_rule=">0"),
            Column.currency("UnitPrice", default=0),
        ],
        indexes=[IndexSpec.primary_key("OrderID", "ProductID")],
    )
    db.relationships.create("Customers.CustomerID", "Orders.CustomerID", cascade_delete=True)
    db.relationships.create("Orders.OrderID", "OrderLines.OrderID", cascade_delete=True)
    db.relationships.create("Products.ProductID", "OrderLines.ProductID")

    db.queries.create(
        "qryOrderTotals",
        "SELECT o.OrderID, c.CustomerName, o.OrderDate, Sum(l.Quantity*l.UnitPrice) AS OrderTotal\n"
        "FROM (Customers AS c INNER JOIN Orders AS o ON c.CustomerID = o.CustomerID)\n"
        "INNER JOIN OrderLines AS l ON o.OrderID = l.OrderID\n"
        "GROUP BY o.OrderID, c.CustomerName, o.OrderDate;",
    )
    db.queries.create(
        "qryActiveCustomers",
        "SELECT * FROM Customers WHERE IsActive = True ORDER BY CustomerName;",
    )


def seed(db: AccessDatabase) -> None:
    customers = [("Ada Lovelace", "ada@example.com"), ("Grace Hopper", "grace@example.com")]
    for name, email in customers:
        db.execute(
            "INSERT INTO Customers (CustomerName, Email) VALUES ([name], [email])",
            {"name": name, "email": email},
        )
    for product, price in [("Notebook", Decimal("4.50")), ("Fountain pen", Decimal("39.00"))]:
        db.execute(
            "INSERT INTO Products (ProductName, UnitPrice) VALUES ([product], [price])",
            {"product": product, "price": price},
        )
    db.execute("INSERT INTO Orders (CustomerID, Notes) VALUES (1, 'First order')")
    db.execute(
        "INSERT INTO OrderLines (OrderID, ProductID, Quantity, UnitPrice) VALUES (1, 2, 1, 39)"
    )


def build_ui(db: AccessDatabase) -> None:
    db.modules.create("modFormatting", FORMATTING_MODULE)

    # A single-record form with stacked label/field pairs; saved when the block ends.
    with db.forms.create(
        "frmCustomers", record_source="Customers", caption="Customers", width=cm(14)
    ) as form:
        form.textbox("CustomerName", width=cm(8))
        form.textbox("Email", width=cm(8))
        form.textbox("Phone")
        form.checkbox("IsActive")
        form.button("cmdClose", caption="Close", on_click=Vba("DoCmd.Close acForm, Me.Name"))

    # A continuous form: labels in the header, one row of controls per record.
    with db.forms.create(
        "frmProducts",
        record_source="SELECT * FROM Products ORDER BY ProductName",
        caption="Products",
        default_view=FormView.CONTINUOUS,
    ) as form:
        form.textbox("ProductName", width=cm(6))
        form.textbox("UnitPrice", format="Currency", width=cm(3))
        form.checkbox("Discontinued")

    for name in db.forms.names():
        db.forms[name].check_opens()  # opens hidden in Form view, then closes

    db.properties["AppTitle"] = "PyAccessKit Orders"
    db.properties["StartUpForm"] = "frmCustomers"


def main(target: Path) -> None:
    with AccessDatabase.create(target, overwrite=True) as db:
        build_schema(db)
        seed(db)
        build_ui(db)
        totals = db.fetch_all("SELECT * FROM qryOrderTotals")
        print(f"built {len(db.tables)} tables, {len(db.queries)} queries, forms {db.forms.names()}")
        print(f"order totals: {totals}")
    print(f"created {target}")


if __name__ == "__main__":
    main(Path(sys.argv[1]) if len(sys.argv) > 1 else Path.cwd() / "orders.accdb")
