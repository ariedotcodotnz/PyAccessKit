# pyright: basic, reportArgumentType=false, reportAttributeAccessIssue=false, reportIndexIssue=false
# (the types-pywin32 stubs are stricter than the COM runtime, e.g. unkOuter=None is valid)
"""Creation of COM objects — always late-bound, never attached to a running instance.

* ``win32com.client.Dispatch(progid)`` first calls ``GetActiveObject`` and would *attach to a user's running
  Access* (ADR 0001, F1). PyAccessKit always uses ``CoCreateInstanceEx`` instead.
* ``Dispatch``/``DispatchEx`` return early- or late-bound wrappers depending on the machine's makepy cache.
  ``win32com.client.dynamic.Dispatch`` keeps the whole object graph late-bound, so behaviour is identical
  on every machine.
"""

from __future__ import annotations

from typing import Any

import pythoncom
import pywintypes
from win32com.client import dynamic

from pyaccesskit._com.runtime import ensure_com_initialized

__all__ = ["create_inproc", "create_local_server", "wrap"]


def wrap(obj: Any) -> Any:
    """Wrap a raw ``PyIDispatch``/``PyIUnknown`` in a late-bound dispatch object."""
    if isinstance(obj, pythoncom.TypeIIDs[pythoncom.IID_IDispatch]):
        return dynamic.Dispatch(obj)
    if isinstance(obj, pythoncom.TypeIIDs[pythoncom.IID_IUnknown]):
        return dynamic.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
    return obj


def create_local_server(progid: str) -> Any:
    """Start a **new** instance of an out-of-process COM server (e.g. ``Access.Application``)."""
    ensure_com_initialized()
    clsid = pywintypes.IID(progid)
    unknown = pythoncom.CoCreateInstanceEx(
        clsid, None, pythoncom.CLSCTX_LOCAL_SERVER, None, (pythoncom.IID_IDispatch,)
    )[0]
    return dynamic.Dispatch(unknown)


def create_inproc(progid: str) -> Any:
    """Load an in-process COM server (e.g. ``DAO.DBEngine.120``) into this Python process."""
    ensure_com_initialized()
    unknown = pythoncom.CoCreateInstance(
        progid, None, pythoncom.CLSCTX_INPROC_SERVER, pythoncom.IID_IDispatch
    )
    return dynamic.Dispatch(unknown)
