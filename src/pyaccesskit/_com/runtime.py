# pyright: basic
"""pywin32 availability and per-thread COM initialisation.

This package (``pyaccesskit._com``) is the only place, together with ``_win``, ``_engines`` and the COM
backends, that imports pywin32. It is imported lazily when a real engine is created, so
``import pyaccesskit`` works on any platform.
"""

from __future__ import annotations

import logging
import threading

import pythoncom

logger = logging.getLogger("pyaccesskit.com")

__all__ = ["ensure_com_initialized"]

_RPC_E_CHANGED_MODE = -2147417850  # 0x80010106
_state = threading.local()


def ensure_com_initialized() -> None:
    """Initialise COM (single-threaded apartment) on the calling thread, once.

    ``import pythoncom`` already initialises the importing thread; worker threads need this call. COM is
    never uninitialised by PyAccessKit (other code on the thread may still hold COM objects). If the thread
    was initialised as MTA by someone else, calls still work through COM marshalling.
    """
    if getattr(_state, "initialized", False):
        return
    try:
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    except pythoncom.com_error as exc:
        if exc.args[0] != _RPC_E_CHANGED_MODE:
            raise
        logger.debug("thread %s is already a multithreaded COM apartment", threading.get_ident())
    _state.initialized = True
