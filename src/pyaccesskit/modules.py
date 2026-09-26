"""VBA modules: the ``db.modules`` collection (creating and reading code needs Microsoft Access)."""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING

from pyaccesskit._ops import design as design_ops
from pyaccesskit.enums import ModuleKind, ObjectKind
from pyaccesskit.errors import ObjectNotFoundError

if TYPE_CHECKING:
    from pyaccesskit._session.session import Session

__all__ = ["Module", "ModuleCollection"]


class Module:
    """A standard or class module (a live, name-based handle)."""

    def __init__(self, session: Session, name: str) -> None:
        self._session = session
        self._name = name

    @property
    def name(self) -> str:
        """The module name."""
        return self._name

    @property
    def kind(self) -> ModuleKind:
        """Standard or class module."""
        return design_ops.read_module(self._session.design(), self._name)[0]

    @property
    def code(self) -> str:
        """The module's VBA source (without Access's class-module header)."""
        return design_ops.read_module(self._session.design(), self._name)[1]

    @code.setter
    def code(self, value: str) -> None:
        self._session.check_writable(f"update module {self._name!r}")
        design_ops.create_module(self._session.design(), self._name, value, self.kind, replace=True)

    def rename(self, new_name: str) -> None:
        """Rename the module."""
        self._session.check_writable(f"rename module {self._name!r}")
        self._session.design().rename_object(ObjectKind.MODULE, self._name, new_name)
        self._name = new_name

    def drop(self) -> None:
        """Delete the module."""
        self._session.check_writable(f"drop module {self._name!r}")
        self._session.design().delete_object(ObjectKind.MODULE, self._name)

    def __repr__(self) -> str:
        return f"<Module {self._name!r}>"


class ModuleCollection:
    """``db.modules``: standard and class modules (form/report modules are managed with their object)."""

    def __init__(self, session: Session) -> None:
        self._session = session

    def names(self) -> list[str]:
        """Module names (read through DAO; does not need Microsoft Access)."""
        return self._session.schema().list_documents(ObjectKind.MODULE)

    def __iter__(self) -> Iterator[Module]:
        return iter([Module(self._session, name) for name in self.names()])

    def __len__(self) -> int:
        return len(self.names())

    def __contains__(self, name: object) -> bool:
        return isinstance(name, str) and any(n.casefold() == name.casefold() for n in self.names())

    def __getitem__(self, name: str) -> Module:
        for actual in self.names():
            if actual.casefold() == name.casefold():
                return Module(self._session, actual)
        raise ObjectNotFoundError(
            f"module {name!r} does not exist", kind=ObjectKind.MODULE, name=name
        )

    def create(
        self, name: str, code: str, *, kind: ModuleKind = ModuleKind.STANDARD, replace: bool = False
    ) -> Module:
        """Create a VBA module from source code.

        ``Option Compare Database`` is added if missing (as Access does); the code is otherwise stored as
        given. VBA source uses the Windows ANSI code page: characters outside it raise ``SpecError``.
        """
        self._session.check_writable(f"create module {name!r}")
        design_ops.create_module(self._session.design(), name, code, kind, replace=replace)
        return Module(self._session, name)

    def drop(self, name: str) -> None:
        """Delete a module."""
        self[name].drop()
