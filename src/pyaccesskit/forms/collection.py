"""Forms: the ``db.forms`` collection and :class:`Form` handles (building forms needs Microsoft Access)."""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from pyaccesskit._backends.protocols import ControlInfo
from pyaccesskit._ops import design as design_ops
from pyaccesskit.enums import ObjectKind
from pyaccesskit.errors import ObjectNotFoundError
from pyaccesskit.forms.builder import FormBuilder
from pyaccesskit.forms.layout import ResolvedForm
from pyaccesskit.forms.spec import FormSpec

if TYPE_CHECKING:
    from pyaccesskit._session.session import Session

__all__ = ["Form", "FormCollection"]


class Form:
    """A saved form (a live, name-based handle)."""

    def __init__(self, session: Session, name: str) -> None:
        self._session = session
        self._name = name

    @property
    def name(self) -> str:
        """The form name."""
        return self._name

    def controls(self) -> list[ControlInfo]:
        """The form's controls as saved (opens the form hidden in Design view)."""
        return self._session.design().form_controls(self._name)

    def check_opens(self) -> None:
        """Open the form in Form view (hidden) and close it; raises if Access reports a problem."""
        self._session.design().check_form_opens(self._name)

    def export_text(self) -> str:
        """The form in Access's ``SaveAsText`` format."""
        return design_ops.export_object(self._session.design(), ObjectKind.FORM, self._name)

    def rename(self, new_name: str) -> None:
        """Rename the form."""
        self._session.check_writable(f"rename form {self._name!r}")
        self._session.design().rename_object(ObjectKind.FORM, self._name, new_name)
        self._name = new_name

    def drop(self) -> None:
        """Delete the form."""
        self._session.check_writable(f"drop form {self._name!r}")
        self._session.design().delete_object(ObjectKind.FORM, self._name)

    def __repr__(self) -> str:
        return f"<Form {self._name!r}>"


class FormCollection:
    """``db.forms``."""

    def __init__(self, session: Session) -> None:
        self._session = session

    def names(self) -> list[str]:
        """Form names (read through DAO; does not need Microsoft Access)."""
        return self._session.schema().list_documents(ObjectKind.FORM)

    def __iter__(self) -> Iterator[Form]:
        return iter([Form(self._session, name) for name in self.names()])

    def __len__(self) -> int:
        return len(self.names())

    def __contains__(self, name: object) -> bool:
        return isinstance(name, str) and any(n.casefold() == name.casefold() for n in self.names())

    def __getitem__(self, name: str) -> Form:
        for actual in self.names():
            if actual.casefold() == name.casefold():
                return Form(self._session, actual)
        raise ObjectNotFoundError(f"form {name!r} does not exist", kind=ObjectKind.FORM, name=name)

    def create(self, name: str, *, replace: bool = False, **form_options: Any) -> FormBuilder[Form]:
        """Start building a form; nothing is created in Access until the builder is saved.

        Example::

            with db.forms.create("frmCustomers", record_source="Customers") as form:
                form.textbox("CustomerName")
            # built atomically here

        Args:
            name: Form name.
            replace: Rebuild the form if it already exists (atomically swapped in).
            **form_options: Any :class:`~pyaccesskit.forms.FormSpec` field (``record_source``, ``caption``,
                ``default_view``...).
        """
        return FormBuilder(
            name, on_save=lambda spec: self.build(spec, replace=replace), **form_options
        )

    def build(self, spec: FormSpec, *, replace: bool = False) -> Form:
        """Build a form from a finished :class:`FormSpec` (atomically)."""
        self._session.check_writable(f"build form {spec.name!r}")
        # design() may switch engines (closing in-process DAO), so resolve the schema backend afterwards.
        design = self._session.design()
        resolved: ResolvedForm = design_ops.build_form(
            self._session.schema(), design, spec, replace=replace
        )
        return Form(self._session, resolved.spec.name)

    def drop(self, name: str) -> None:
        """Delete a form."""
        self[name].drop()
