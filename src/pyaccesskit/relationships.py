"""Relationships: the ``db.relationships`` collection."""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, overload

from pyaccesskit._ops import schema as ops
from pyaccesskit.enums import JoinType, ObjectKind
from pyaccesskit.errors import ObjectNotFoundError
from pyaccesskit.schema import ColumnRef, RelationshipSpec

if TYPE_CHECKING:
    from pyaccesskit._session.session import Session

__all__ = ["Relationship", "RelationshipCollection"]


class Relationship:
    """A relationship (a live, name-based handle)."""

    def __init__(self, session: Session, name: str) -> None:
        self._session = session
        self._name = name

    @property
    def name(self) -> str:
        """The relationship name."""
        return self._name

    def to_spec(self) -> RelationshipSpec:
        """The relationship's definition."""
        for spec in self._session.schema().list_relationships():
            if spec.effective_name.casefold() == self._name.casefold():
                return spec
        raise ObjectNotFoundError(
            f"relationship {self._name!r} does not exist",
            kind=ObjectKind.RELATIONSHIP,
            name=self._name,
        )

    def drop(self) -> None:
        """Delete the relationship."""
        self._session.check_writable(f"drop relationship {self._name!r}")
        ops.drop_relationship(self._session.schema(), self._name)

    def __repr__(self) -> str:
        return f"<Relationship {self._name!r}>"


class RelationshipCollection:
    """``db.relationships``."""

    def __init__(self, session: Session) -> None:
        self._session = session

    def specs(self) -> list[RelationshipSpec]:
        """Every relationship as a spec."""
        return self._session.schema().list_relationships()

    def names(self) -> list[str]:
        """Relationship names."""
        return [spec.effective_name for spec in self.specs()]

    def __iter__(self) -> Iterator[Relationship]:
        return iter([Relationship(self._session, name) for name in self.names()])

    def __len__(self) -> int:
        return len(self.names())

    def __contains__(self, name: object) -> bool:
        return isinstance(name, str) and any(n.casefold() == name.casefold() for n in self.names())

    def __getitem__(self, name: str) -> Relationship:
        for actual in self.names():
            if actual.casefold() == name.casefold():
                return Relationship(self._session, actual)
        raise ObjectNotFoundError(
            f"relationship {name!r} does not exist", kind=ObjectKind.RELATIONSHIP, name=name
        )

    @overload
    def create(self, spec: RelationshipSpec, /) -> Relationship: ...

    @overload
    def create(
        self,
        primary: ColumnRef,
        foreign: ColumnRef,
        /,
        *,
        name: str | None = None,
        enforce_integrity: bool = True,
        cascade_update: bool = False,
        cascade_delete: bool = False,
        one_to_one: bool = False,
        join: JoinType = JoinType.INNER,
    ) -> Relationship: ...

    def create(
        self,
        primary: RelationshipSpec | ColumnRef,
        foreign: ColumnRef | None = None,
        /,
        *,
        name: str | None = None,
        enforce_integrity: bool = True,
        cascade_update: bool = False,
        cascade_delete: bool = False,
        one_to_one: bool = False,
        join: JoinType = JoinType.INNER,
    ) -> Relationship:
        """Create a relationship from the primary ("one") side to the foreign ("many") side.

        Example::

            db.relationships.create("Customers.CustomerID", "Orders.CustomerID", cascade_delete=True)

        The primary columns need a primary key or unique index. The default name follows Access's
        convention (primary table + foreign table).
        """
        if isinstance(primary, RelationshipSpec):
            spec = primary
        else:
            if foreign is None:
                raise TypeError("create() needs both the primary and the foreign side")
            spec = RelationshipSpec.between(
                primary,
                foreign,
                name=name,
                enforce_integrity=enforce_integrity,
                cascade_update=cascade_update,
                cascade_delete=cascade_delete,
                one_to_one=one_to_one,
                join=join,
            )
        self._session.check_writable(f"create relationship {spec.effective_name!r}")
        created = ops.create_relationship(self._session.schema(), spec)
        return Relationship(self._session, created.effective_name)

    def drop(self, name: str) -> None:
        """Delete a relationship."""
        self[name].drop()
