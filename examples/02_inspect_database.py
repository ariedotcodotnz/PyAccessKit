"""Inspect an existing database without changing it.

Run:  python examples/02_inspect_database.py path/to/database.accdb

Read-only sessions open the file in shared mode (so it can stay open in Access) and never run the
database's startup code: PyAccessKit reads it through DAO, not through the Access user interface.
(``pyaccesskit inspect DATABASE`` prints a similar report.)
"""

from __future__ import annotations

import sys
from pathlib import Path

from pyaccesskit import AccessDatabase


def main(path: Path) -> None:
    with AccessDatabase.open(path, readonly=True) as db:
        print(f"{db.path} (format {db.format_version}, via {db.transport.value})")
        for table in db.tables:
            spec = table.to_spec()  # an immutable, serializable snapshot
            kind = "linked table" if table.is_linked else "table"
            print(f"\n{kind} {table.name}: {len(spec.columns)} columns")
            for column in spec.columns:
                flags = " required" if column.required else ""
                print(f"  {column.name:<24} {column.type:<12}{flags}")
            for index in spec.indexes:
                print(f"  index {index.name} on {', '.join(index.field_names)}")
        for relationship in db.relationships.specs():
            print(
                f"\nrelationship {relationship.name}: "
                f"{relationship.primary_table}{list(relationship.primary_columns)} -> "
                f"{relationship.foreign_table}{list(relationship.foreign_columns)}"
            )
        for query in db.queries:
            print(f"\nquery {query.name} ({query.kind.value}):\n  {query.sql.strip()}")
        for kind in ("form", "report", "macro", "module"):
            names = db.objects.names(kind)
            if names:
                print(f"\n{kind}s: {', '.join(names)}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit("usage: python examples/02_inspect_database.py DATABASE")
    main(Path(sys.argv[1]))
