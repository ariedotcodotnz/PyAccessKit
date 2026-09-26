"""Backend fixtures for the contract suite.

The same behavioural tests run against every backend. ``fake`` always runs; real backends are added as
integration-marked parameters (see ``tests/integration/backends.py``) so the in-memory fake can never drift
from what Access actually does.
"""

from __future__ import annotations

from collections.abc import Callable, Iterator
from contextlib import AbstractContextManager, contextmanager
from dataclasses import dataclass
from pathlib import Path

import pytest

from pyaccesskit._backends.fake import FakeBackend
from pyaccesskit._backends.protocols import DesignBackend, SchemaBackend


@dataclass
class Backends:
    """A schema backend and (when available) a design backend for one database."""

    name: str
    schema: SchemaBackend
    design: DesignBackend | None


BackendFactory = Callable[[Path], AbstractContextManager[Backends]]


@contextmanager
def _fake(tmp_path: Path) -> Iterator[Backends]:
    backend = FakeBackend(tmp_path / "memory.accdb")
    yield Backends("fake", backend, backend)


FACTORIES: dict[str, BackendFactory] = {"fake": _fake}
PARAMS: list[object] = [pytest.param("fake", id="fake")]

try:  # real backends register themselves when the integration helpers are importable
    from tests.integration.backends import REAL_FACTORIES, REAL_PARAMS
except ImportError:  # pragma: no cover - integration helpers missing
    pass
else:
    FACTORIES.update(REAL_FACTORIES)
    PARAMS.extend(REAL_PARAMS)


@pytest.fixture(params=PARAMS)
def backends(request: pytest.FixtureRequest, tmp_path: Path) -> Iterator[Backends]:
    factory = FACTORIES[request.param]
    with factory(tmp_path) as pair:
        yield pair
