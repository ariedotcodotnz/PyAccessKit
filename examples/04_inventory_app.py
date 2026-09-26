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
