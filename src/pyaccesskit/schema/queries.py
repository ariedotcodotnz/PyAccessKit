"""Saved query specifications and Access SQL helpers.

Access rewrites SQL when it saves a query (keywords upper-cased, one clause per line, a trailing ``;``,
``DELETE`` becoming ``DELETE *``...). :func:`sql_equivalent` compares SQL modulo those cosmetic changes, which
is what change detection needs.
"""

from __future__ import annotations

import re
from typing import Self

from pydantic import Field, field_validator, model_validator

from pyaccesskit.enums import QueryKind
from pyaccesskit.schema._base import SpecModel
from pyaccesskit.schema.names import check_name

__all__ = [
    "DAO_QUERY_KINDS",
    "MAX_SQL_LENGTH",
    "PassThroughOptions",
    "QuerySpec",
    "detect_query_kind",
    "normalize_sql",
    "sql_equivalent",
]

MAX_SQL_LENGTH = 64_000
"""Approximate Access limit on the length of a SQL statement."""

DAO_QUERY_KINDS: dict[int, QueryKind] = {
    0: QueryKind.SELECT,
    16: QueryKind.CROSSTAB,
    32: QueryKind.DELETE,
    48: QueryKind.UPDATE,
    64: QueryKind.APPEND,
    80: QueryKind.MAKE_TABLE,
    96: QueryKind.DDL,
    112: QueryKind.PASS_THROUGH,
    128: QueryKind.UNION,
    144: QueryKind.PASS_THROUGH_BULK,
    160: QueryKind.COMPOUND,
    224: QueryKind.PROCEDURE,
    240: QueryKind.ACTION,
}
"""DAO ``QueryDefTypeEnum`` values (verified against the ACEDAO type library) → :class:`QueryKind`."""


class PassThroughOptions(SpecModel):
    """Settings of an ODBC pass-through query.

    Attributes:
        connect: ODBC connection string; must start with ``ODBC;``.
        returns_records: Whether the statement returns rows.
        timeout: ODBC timeout in seconds (0 = no timeout).
    """

    connect: str
    returns_records: bool = True
    timeout: int = Field(default=60, ge=0)

    @field_validator("connect")
    @classmethod
    def _check_connect(cls, value: str) -> str:
        if not value.upper().startswith("ODBC;"):
            raise ValueError("pass-through connect strings must start with 'ODBC;'")
        return value


class QuerySpec(SpecModel):
    """A saved query (Access QueryDef).

    Attributes:
        name: Query name (tables and queries share one namespace).
        sql: The SQL text (Access SQL, or the server's dialect for pass-through queries).
        description: Query *Description* property.
        pass_through: Set for ODBC pass-through queries.
    """

    name: str
    sql: str = Field(min_length=1, max_length=MAX_SQL_LENGTH)
    description: str | None = None
    pass_through: PassThroughOptions | None = None

    @field_validator("name")
    @classmethod
    def _check_name(cls, value: str) -> str:
        return check_name(value, what="query name")

    @model_validator(mode="after")
    def _check_sql(self) -> Self:
        if not self.sql.strip():
            raise ValueError("sql cannot be blank")
        return self

    @property
    def kind(self) -> QueryKind:
        """The query kind as implied by the SQL text (Access reports the authoritative kind after saving)."""
        return detect_query_kind(self.sql, pass_through=self.pass_through is not None)

    def normalized(self) -> QuerySpec:
        """Canonical form (SQL line endings normalized to CRLF, as Access stores them)."""
        sql = self.sql.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
        return self if sql == self.sql else self.model_copy(update={"sql": sql})


# -------------------------------------------------------------------------------------- SQL helpers
_TOKEN = re.compile(
    r"""
      (?P<string>'(?:[^']|'')*'|"(?:[^"]|"")*")
    | (?P<bracket>\[[^\]]*\])
    | (?P<date>\#[^#\r\n]*\#)
    | (?P<space>\s+)
    | (?P<punct>[(),;=<>+\-*/&^\\])
    | (?P<word>[^\s'"\[\#(),;=<>+\-*/&^\\]+)
    """,
    re.VERBOSE,
)


def _tokens(sql: str) -> list[tuple[str, str]]:
    tokens: list[tuple[str, str]] = []
    for match in _TOKEN.finditer(sql):
        kind = match.lastgroup or "word"
        if kind == "space":
            continue
        tokens.append((kind, match.group()))
    return tokens


def normalize_sql(sql: str) -> str:
    """A canonical single-line form of ``sql`` for comparisons.

    Whitespace is collapsed and dropped around punctuation, words and bracketed identifiers are
    upper-cased (Access identifiers are case-insensitive), string and date literals are preserved, trailing
    semicolons are removed, and Access's ``DELETE * FROM`` rewrite is undone.
    """
    parts: list[str] = []
    for kind, text in _tokens(sql):
        parts.append(text if kind in ("string", "date") else text.upper())
    while parts and parts[-1] == ";":
        parts.pop()
    if len(parts) >= 3 and parts[0] == "DELETE" and parts[1] == "*" and parts[2] == "FROM":
        del parts[1]
    return " ".join(parts)


def sql_equivalent(left: str, right: str) -> bool:
    """Whether two SQL texts differ only cosmetically (see :func:`normalize_sql`)."""
    return normalize_sql(left) == normalize_sql(right)


def detect_query_kind(sql: str, *, pass_through: bool = False) -> QueryKind:
    """Infer the kind of query from its SQL text (used before Access has classified a saved query)."""
    if pass_through:
        return QueryKind.PASS_THROUGH
    words = [text.upper() if kind == "word" else text for kind, text in _tokens(sql)]
    if words and words[0] == "PARAMETERS":
        words = words[words.index(";") + 1 :] if ";" in words else []
    if not words:
        return QueryKind.UNKNOWN
    first = words[0].lstrip("(")
    simple = {
        "TRANSFORM": QueryKind.CROSSTAB,
        "INSERT": QueryKind.APPEND,
        "UPDATE": QueryKind.UPDATE,
        "DELETE": QueryKind.DELETE,
        "CREATE": QueryKind.DDL,
        "ALTER": QueryKind.DDL,
        "DROP": QueryKind.DDL,
        "PROCEDURE": QueryKind.PROCEDURE,
    }
    if first in simple:
        return simple[first]
    if first not in ("SELECT", "("):
        return QueryKind.UNKNOWN
    depth = 0
    seen_from = False
    for word in words:
        if word == "(":
            depth += 1
        elif word == ")":
            depth -= 1
        elif depth == 0:
            if word == "UNION":
                return QueryKind.UNION
            if word == "FROM":
                seen_from = True
            elif word == "INTO" and not seen_from:
                return QueryKind.MAKE_TABLE
    return QueryKind.SELECT
