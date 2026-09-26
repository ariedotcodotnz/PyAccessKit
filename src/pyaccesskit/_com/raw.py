# pyright: basic
"""Revocable proxies for the ``db.raw`` escape hatch.

Raw COM objects handed to user code are wrapped so that closing (or upgrading) the session can *revoke*
them: every proxy — including the ones derived from it (``db.raw.dao.TableDefs(0).Fields``...) — drops its
COM reference and raises :class:`~pyaccesskit.errors.SessionClosedError` afterwards. User variables can
therefore never keep an Access process alive or be used against a closed database.
"""

from __future__ import annotations

import threading
import weakref
from collections.abc import Iterator
from typing import Any

from pyaccesskit.errors import SessionClosedError

__all__ = ["ProxyRegistry", "RevocableIterator", "RevocableProxy"]

_SLOTS = ("_pak_target", "_pak_registry", "_pak_label", "__weakref__")


class ProxyRegistry:
    """Tracks every proxy created for a session so they can all be revoked at once."""

    def __init__(self) -> None:
        self._proxies: weakref.WeakSet[RevocableProxy | RevocableIterator] = weakref.WeakSet()
        self._lock = threading.Lock()
        self.revoked = False

    def track(self, proxy: RevocableProxy | RevocableIterator) -> None:
        with self._lock:
            self._proxies.add(proxy)

    def revoke_all(self) -> int:
        """Revoke every live proxy; returns how many were revoked."""
        with self._lock:
            proxies = list(self._proxies)
            self._proxies.clear()
            self.revoked = True
        for proxy in proxies:
            object.__setattr__(
                proxy, "_pak_target", None
            )  # iterators drop their COM enumerator too
        return len(proxies)

    def wrap(self, value: Any, label: str) -> Any:
        """Wrap COM objects and bound COM methods; return plain values unchanged."""
        if value is None or isinstance(value, (str, bytes, int, float, bool, tuple, list, dict)):
            return value
        if hasattr(value, "_oleobj_") or callable(value):
            proxy = RevocableProxy(value, self, label)
            self.track(proxy)
            return proxy
        return value


def _unwrap(value: Any) -> Any:
    if isinstance(value, RevocableProxy):
        return value._pak_require()
    if isinstance(value, (list, tuple)):
        return type(value)(_unwrap(item) for item in value)
    return value


class RevocableProxy:
    """A transparent, revocable stand-in for a COM object (or one of its methods)."""

    __slots__ = _SLOTS

    def __init__(self, target: Any, registry: ProxyRegistry, label: str) -> None:
        object.__setattr__(self, "_pak_target", target)
        object.__setattr__(self, "_pak_registry", registry)
        object.__setattr__(self, "_pak_label", label)

    def _pak_require(self) -> Any:
        target = object.__getattribute__(self, "_pak_target")
        if target is None:
            label = object.__getattribute__(self, "_pak_label")
            raise SessionClosedError(
                f"raw COM object {label} was revoked because its PyAccessKit session was closed or upgraded"
            )
        return target

    def __getattr__(self, name: str) -> Any:
        target = self._pak_require()
        label = object.__getattribute__(self, "_pak_label")
        registry: ProxyRegistry = object.__getattribute__(self, "_pak_registry")
        return registry.wrap(getattr(target, name), f"{label}.{name}")

    def __setattr__(self, name: str, value: Any) -> None:
        setattr(self._pak_require(), name, _unwrap(value))

    def __call__(self, *args: Any, **kwargs: Any) -> Any:
        target = self._pak_require()
        label = object.__getattribute__(self, "_pak_label")
        registry: ProxyRegistry = object.__getattribute__(self, "_pak_registry")
        result = target(*_unwrap(args), **{k: _unwrap(v) for k, v in kwargs.items()})
        return registry.wrap(result, f"{label}(...)")

    def __iter__(self) -> Iterator[Any]:
        target = self._pak_require()
        label = object.__getattribute__(self, "_pak_label")
        registry: ProxyRegistry = object.__getattribute__(self, "_pak_registry")
        iterator = RevocableIterator(iter(target), registry, label)
        registry.track(iterator)
        return iterator

    def __getitem__(self, key: Any) -> Any:
        target = self._pak_require()
        label = object.__getattribute__(self, "_pak_label")
        registry: ProxyRegistry = object.__getattribute__(self, "_pak_registry")
        return registry.wrap(target[_unwrap(key)], f"{label}[{key!r}]")

    def __len__(self) -> int:
        return len(self._pak_require())

    def __bool__(self) -> bool:
        return object.__getattribute__(self, "_pak_target") is not None

    @property
    def revoked(self) -> bool:
        """Whether the underlying COM object has been released."""
        return object.__getattribute__(self, "_pak_target") is None

    def __repr__(self) -> str:
        label = object.__getattribute__(self, "_pak_label")
        state = "revoked" if self.revoked else "live"
        return f"<pyaccesskit raw {label} ({state})>"


class RevocableIterator:
    """Iterates a raw COM collection; revoking it drops the enumerator, so resuming raises."""

    __slots__ = ("_pak_index", *_SLOTS)

    def __init__(self, target: Iterator[Any], registry: ProxyRegistry, label: str) -> None:
        object.__setattr__(self, "_pak_target", target)
        object.__setattr__(self, "_pak_registry", registry)
        object.__setattr__(self, "_pak_label", label)
        object.__setattr__(self, "_pak_index", 0)

    def __iter__(self) -> RevocableIterator:
        return self

    def __next__(self) -> Any:
        target = object.__getattribute__(self, "_pak_target")
        label = object.__getattribute__(self, "_pak_label")
        if target is None:
            raise SessionClosedError(
                f"iteration over raw COM object {label} was revoked because its PyAccessKit session was "
                "closed or upgraded"
            )
        item = next(target)
        index = object.__getattribute__(self, "_pak_index")
        object.__setattr__(self, "_pak_index", index + 1)
        registry: ProxyRegistry = object.__getattribute__(self, "_pak_registry")
        return registry.wrap(item, f"{label}[{index}]")

    @property
    def revoked(self) -> bool:
        """Whether the underlying enumerator has been released."""
        return object.__getattribute__(self, "_pak_target") is None
