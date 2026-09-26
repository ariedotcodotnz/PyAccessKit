"""Saved queries: the ``db.queries`` collection and :class:`Query` handles."""

from __future__ import annotations

from collections.abc import Iterator, Mapping
from typing import TYPE_CHECKING, Any, overload

from pyaccesskit._backends.protocols import ParameterInfo, PropertyTarget
from pyaccesskit._ops import schema as ops
from pyaccesskit.enums import ObjectKind, QueryKind
from pyaccesskit.errors import ObjectNotFoundError
from pyaccesskit.properties import PropertyBag
from pyaccesskit.schema import PassThroughOptions, QuerySpec
from pyaccesskit.schema._base import build

if TYPE_CHECKING:
    from pyaccesskit._session.session import Session

__all__ = ["Query", "QueryCollection"]


class Query:
    """A saved query (a live, name-based handle)."""

    def __init__(self, session: Session, name: str) -> None:
        self._session = session
        self._name = name

    @property
    def name(self) -> str:
        """The query name."""
        return self._name

    def to_spec(self) -> QuerySpec:
        """The query's current definition (SQL as Access stored it)."""
        return self._session.schema().read_query(self._name)

    @property
    def sql(self) -> str:
        """The SQL text (Access reformats SQL when saving it)."""
        return self.to_spec().sql

    @sql.setter
    def sql(self, value: str) -> None:
        self._session.check_writable(f"update query {self._name!r}")
        spec = build(QuerySpec, f"query {self._name!r}", name=self._name, sql=value)
        self._session.schema().set_query_sql(self._name, spec.normalized().sql)

    @property
    def kind(self) -> QueryKind:
        """Select, update, append, crosstab... as classified by Access."""
        for info in self._session.schema().list_queries():
            if info.name.casefold() == self._name.casefold():
                return info.kind
        raise ObjectNotFoundError(
            f"query {self._name!r} does not exist", kind=ObjectKind.QUERY, name=self._name
        )

    @property
    def parameters(self) -> list[ParameterInfo]:
        """Declared and implicit parameters."""
        return self._session.schema().query_parameters(self._name)

    @property
    def description(self) -> str | None:
        """The query *Description*."""
        return self.to_spec().description

    @property
    def properties(self) -> PropertyBag:
        """The query's DAO properties."""
        return PropertyBag(self._session, PropertyTarget.query(self._name))

    def execute(self, params: Mapping[str, Any] | None = None) -> int:
        """Run an action query; returns the number of affected rows."""
        self._session.check_writable(f"run query {self._name!r}")
        return self._session.schema().execute_saved(self._name, params)

    def fetch(
        self, params: Mapping[str, Any] | None = None, *, limit: int | None = None
    ) -> list[dict[str, Any]]:
        """Return the rows of a select query as dictionaries."""
        return self._session.schema().fetch_saved(self._name, params, limit=limit).as_dicts()

    def rename(self, new_name: str) -> None:
        """Rename the query."""
        self._session.check_writable(f"rename query {self._name!r}")
        self._name = ops.rename_query(self._session.schema(), self._name, new_name)

    def drop(self) -> None:
        """Delete the query."""
        self._session.check_writable(f"drop query {self._name!r}")
        ops.drop_query(self._session.schema(), self._name)

    def __repr__(self) -> str:
        return f"<Query {self._name!r}>"


class QueryCollection:
    """``db.queries``: saved queries (hidden ``~`` queries are excluded unless asked for)."""

    def __init__(self, session: Session) -> None:
        self._session = session

    def names(self, *, include_hidden: bool = False) -> list[str]:
        """Query names."""
        return [
            q.name
            for q in self._session.schema().list_queries()
            if include_hidden or not q.is_hidden
        ]

    def __iter__(self) -> Iterator[Query]:
        return iter([Query(self._session, name) for name in self.names()])

    def __len__(self) -> int:
        return len(self.names())

    def __contains__(self, name: object) -> bool:
        return isinstance(name, str) and self.get(name) is not None

    def get(self, name: str) -> Query | None:
        """The query called ``name`` (case-insensitive), or ``None``."""
        actual = ops.find_query(self._session.schema(), name)
        return Query(self._session, actual) if actual is not None else None

    def __getitem__(self, name: str) -> Query:
        query = self.get(name)
        if query is None:
            raise ObjectNotFoundError(
                f"query {name!r} does not exist", kind=ObjectKind.QUERY, name=name
            )
        return query

    @overload
    def create(self, spec: QuerySpec, /, *, replace: bool = False) -> Query: ...

    @overload
    def create(
        self, name: str, sql: str, /, *, description: str | None = None, replace: bool = False
    ) -> Query: ...

    def create(
        self,
        spec_or_name: QuerySpec | str,
        sql: str | None = None,
        /,
        *,
        description: str | None = None,
        replace: bool = False,
    ) -> Query:
        """Create a saved query (``replace=True`` updates the SQL of an existing one)."""
        if isinstance(spec_or_name, QuerySpec):
            spec = spec_or_name
        else:
            spec = build(
                QuerySpec,
                f"query {spec_or_name!r}",
                name=spec_or_name,
                sql=sql or "",
                description=description,
            )
        self._session.check_writable(f"create query {spec.name!r}")
        ops.create_or_replace_query(self._session.schema(), spec, replace=replace)
        return self[spec.name]

    def create_pass_through(
        self,
        name: str,
        sql: str,
        *,
        connect: str,
        returns_records: bool = True,
        timeout: int = 60,
        replace: bool = False,
    ) -> Query:
        """Create an ODBC pass-through query (``connect`` must start with ``ODBC;``)."""
        options = build(
            PassThroughOptions,
            f"pass-through query {name!r}",
            connect=connect,
            returns_records=returns_records,
            timeout=timeout,
        )
        spec = build(QuerySpec, f"query {name!r}", name=name, sql=sql, pass_through=options)
        return self.create(spec, replace=replace)

    def drop(self, name: str) -> None:
        """Delete a saved query."""
        self[name].drop()
