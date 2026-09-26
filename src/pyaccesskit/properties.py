"""Access to DAO property collections (the escape hatch for properties PyAccessKit does not model)."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pyaccesskit._backends.protocols import PropertyTarget
from pyaccesskit.enums import PropertyType
from pyaccesskit.errors import ObjectNotFoundError
from pyaccesskit.schema import PropertyValue

if TYPE_CHECKING:
    from pyaccesskit._session.session import Session

__all__ = ["PropertyBag"]


class PropertyBag:
    """The DAO ``Properties`` of a database, table, field or query.

    Reading a property that does not exist raises :class:`~pyaccesskit.errors.ObjectNotFoundError` (a
    ``LookupError``); :meth:`get` returns a default instead. Setting a missing property creates it — with the
    DAO type Access itself uses for well-known properties (``Description``, ``AppTitle``, ``UseMDIMode``...)
    or one inferred from the value, unless ``type`` is given::

        db.properties["AppTitle"] = "Customer Manager"
        db.tables["Customers"].properties.set("SubdatasheetName", "[None]")
    """

    def __init__(self, session: Session, target: PropertyTarget) -> None:
        self._session = session
        self._target = target

    def __getitem__(self, name: str) -> PropertyValue:
        return self._session.schema().get_property(self._target, name)

    def get(self, name: str, default: PropertyValue = None) -> PropertyValue:
        """The property value, or ``default`` if it does not exist."""
        try:
            return self[name]
        except ObjectNotFoundError:
            return default

    def __contains__(self, name: object) -> bool:
        if not isinstance(name, str):
            return False
        try:
            self[name]
        except ObjectNotFoundError:
            return False
        return True

    def __setitem__(self, name: str, value: PropertyValue) -> None:
        self.set(name, value)

    def set(self, name: str, value: PropertyValue, type: PropertyType | None = None) -> None:
        """Set (creating if necessary) a property; ``type`` forces the DAO type of a new property."""
        self._session.check_writable(f"set property {name!r}")
        self._session.schema().set_property(self._target, name, value, type)

    def __delitem__(self, name: str) -> None:
        self.delete(name)

    def delete(self, name: str) -> None:
        """Delete a user-defined property."""
        self._session.check_writable(f"delete property {name!r}")
        self._session.schema().delete_property(self._target, name)

    def to_dict(self) -> dict[str, PropertyValue]:
        """All readable properties."""
        return self._session.schema().list_properties(self._target)

    def __repr__(self) -> str:
        return f"<PropertyBag of {self._target.describe()}>"
