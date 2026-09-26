# pyright: basic
"""The COM call gateway: error translation, dialog monitoring and exact-argument invocation.

Every group of COM calls made by a backend runs inside :meth:`Com.op`, which

* tells the dialog watchdog that a call is in flight (so a modal dialog is detected, captured and
  dismissed instead of hanging forever);
* translates ``pywintypes.com_error`` into :mod:`pyaccesskit.errors` exceptions with context;
* raises :class:`~pyaccesskit.errors.AccessDialogError` (or warns) if Access showed a dialog meanwhile.

The ``get``/``put``/``call`` helpers invoke ``IDispatch`` by DISPID with *exactly* the given arguments.
pywin32's dynamic wrappers cannot call indexed properties such as ``Form.Section(0)``, raise
``AttributeError`` for members missing from a control's type info, and pad ``Application.Run``'s 30 optional
arguments (spikes S6 and S8); these helpers avoid all three problems.
"""

from __future__ import annotations

import logging
import time
import warnings
from collections.abc import Callable, Iterator, Sequence
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Protocol

import pythoncom
import pywintypes

from pyaccesskit._com import constants as c
from pyaccesskit._com.dispatch import wrap
from pyaccesskit._com.errors import OpContext, translate
from pyaccesskit._com.variants import normalize
from pyaccesskit.enums import DialogPolicy, ObjectKind
from pyaccesskit.errors import (
    AccessDialogError,
    AccessDialogWarning,
    AccessTimeoutError,
    DaoError,
    DialogInfo,
)

__all__ = ["CallMonitor", "Com", "call", "get", "invoke", "put", "retry_busy"]

logger = logging.getLogger("pyaccesskit.com")

_LCID = 0
_BUSY = frozenset({c.RPC_E_CALL_REJECTED, c.RPC_E_SERVERCALL_RETRYLATER})


class CallMonitor(Protocol):
    """Implemented by the dialog watchdog of an Access process."""

    policy: DialogPolicy

    def begin(self, description: str) -> None:
        """An outermost COM operation starts."""
        ...

    def end(self) -> tuple[DialogInfo, ...]:
        """The operation ended; return the dialogs observed while it ran."""
        ...

    @property
    def termination(self) -> tuple[str, str] | None:
        """``(kind, reason)`` if the monitor terminated Access (kind: timeout, dialog, interrupt)."""
        ...


# --------------------------------------------------------------------------------------- invocation
def invoke(obj: Any, name: str, flags: int, *args: Any) -> Any:
    """``IDispatch::Invoke`` ``name`` on ``obj`` with exactly ``args``; wraps/normalizes the result."""
    ole = obj._oleobj_ if hasattr(obj, "_oleobj_") else obj
    dispid = ole.GetIDsOfNames(name)
    result = ole.Invoke(dispid, _LCID, flags, True, *args)
    return normalize(wrap(result))


def get(obj: Any, name: str, *args: Any) -> Any:
    """Read a (possibly indexed) property: ``get(form, "Section", 0)``."""
    return invoke(obj, name, pythoncom.DISPATCH_PROPERTYGET | pythoncom.DISPATCH_METHOD, *args)


def call(obj: Any, name: str, *args: Any) -> Any:
    """Call a method with exactly ``args`` (no padding of optional parameters)."""
    return invoke(obj, name, pythoncom.DISPATCH_METHOD, *args)


def put(obj: Any, name: str, value: Any) -> None:
    """Set a property."""
    ole = obj._oleobj_ if hasattr(obj, "_oleobj_") else obj
    dispid = ole.GetIDsOfNames(name)
    ole.Invoke(dispid, _LCID, pythoncom.DISPATCH_PROPERTYPUT, False, value)


def retry_busy(fn: Callable[..., Any], *args: Any, timeout: float = 30.0) -> Any:
    """Call ``fn`` and retry while the server rejects calls because it is busy (``RPC_E_CALL_REJECTED``).

    A rejected call never started executing, so retrying is always safe.
    """
    deadline = time.monotonic() + timeout
    delay = 0.05
    while True:
        try:
            return fn(*args)
        except pywintypes.com_error as exc:
            if exc.args[0] not in _BUSY or time.monotonic() >= deadline:
                raise
            logger.debug("COM server busy; retrying in %.2fs", delay)
            time.sleep(delay)
            delay = min(delay * 2, 1.0)


# ------------------------------------------------------------------------------------------ gateway
class Com:
    """Per-engine gateway: wraps operations with error translation and dialog monitoring."""

    def __init__(
        self,
        *,
        monitor: CallMonitor | None = None,
        dao_errors: Callable[[], Sequence[DaoError]] | None = None,
    ) -> None:
        self._monitor = monitor
        self._dao_errors = dao_errors
        self._depth = 0

    def _read_dao_errors(self) -> Sequence[DaoError]:
        if self._dao_errors is None:
            return ()
        try:
            return self._dao_errors()
        except Exception:
            return ()

    @contextmanager
    def op(
        self,
        action: str,
        *,
        kind: ObjectKind | None = None,
        name: str | None = None,
        path: Path | None = None,
        sql: str | None = None,
    ) -> Iterator[None]:
        """Run a group of COM calls as one operation (see module docstring)."""
        ctx = OpContext(action, kind, name, path, sql)
        monitor = self._monitor
        outermost = self._depth == 0
        self._depth += 1
        if monitor is not None and outermost:
            monitor.begin(action)
        ended = False

        def finish() -> tuple[DialogInfo, ...]:
            nonlocal ended
            if ended or monitor is None or not outermost:
                return ()
            ended = True
            return monitor.end()

        try:
            logger.debug("COM op: %s", action)
            yield
        except pywintypes.com_error as exc:
            dialogs = finish()
            error = translate(exc, ctx, self._read_dao_errors())
            termination = monitor.termination if monitor is not None else None
            if termination is not None and termination[0] == "timeout":
                raise AccessTimeoutError(
                    f"cannot {action}: {termination[1]}; the owned Access process was terminated",
                    operation=action,
                    details=error.details,
                ) from error
            if dialogs:
                raise AccessDialogError(
                    f"cannot {action}: Microsoft Access showed a dialog",
                    dialogs=dialogs,
                    operation=action,
                ) from error
            raise error from exc
        except BaseException:
            finish()
            raise
        finally:
            self._depth -= 1
        dialogs = finish()
        if dialogs and monitor is not None:
            if monitor.policy is DialogPolicy.FAIL:
                raise AccessDialogError(
                    f"{action}: Microsoft Access showed a dialog that was dismissed automatically",
                    dialogs=dialogs,
                    operation=action,
                )
            if monitor.policy is DialogPolicy.WARN:
                for dialog in dialogs:
                    warnings.warn(
                        f"{action}: {dialog.summary()}", AccessDialogWarning, stacklevel=3
                    )
