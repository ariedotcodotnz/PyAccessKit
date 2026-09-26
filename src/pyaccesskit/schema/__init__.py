"""Immutable, validated specifications of database schema objects.

Specs are plain Python values (Pydantic models): they never touch Access. Collections such as
``db.tables`` turn them into real objects, and introspection turns real objects back into specs.
"""

from pyaccesskit.schema._base import PropertyValue
from pyaccesskit.schema.columns import (
    AutoNumberColumn,
    Column,
    ColumnBase,
    ColumnSpec,
    CurrencyColumn,
    DateTimeColumn,
    DecimalColumn,
    HyperlinkColumn,
    LongTextColumn,
    NumberColumn,
    OleObjectColumn,
    TextColumn,
    UnsupportedColumn,
    YesNoColumn,
)
from pyaccesskit.schema.expressions import DefaultValue, Expr, parse_literal, render_literal
from pyaccesskit.schema.indexes import PRIMARY_KEY_NAME, IndexField, IndexSpec
from pyaccesskit.schema.names import check_name, lint_name, name_key, names_equal, quote_identifier
from pyaccesskit.schema.queries import (
    PassThroughOptions,
    QuerySpec,
    detect_query_kind,
    normalize_sql,
    sql_equivalent,
)
from pyaccesskit.schema.relationships import ColumnRef, RelationshipSpec
from pyaccesskit.schema.tables import TableSpec

__all__ = [
    "PRIMARY_KEY_NAME",
    "AutoNumberColumn",
    "Column",
    "ColumnBase",
    "ColumnRef",
    "ColumnSpec",
    "CurrencyColumn",
    "DateTimeColumn",
    "DecimalColumn",
    "DefaultValue",
    "Expr",
    "HyperlinkColumn",
    "IndexField",
    "IndexSpec",
    "LongTextColumn",
    "NumberColumn",
    "OleObjectColumn",
    "PassThroughOptions",
    "PropertyValue",
    "QuerySpec",
    "RelationshipSpec",
    "TableSpec",
    "TextColumn",
    "UnsupportedColumn",
    "YesNoColumn",
    "check_name",
    "detect_query_kind",
    "lint_name",
    "name_key",
    "names_equal",
    "normalize_sql",
    "parse_literal",
    "quote_identifier",
    "render_literal",
    "sql_equivalent",
]
