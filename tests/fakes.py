"""Fake engines for testing the session and the public API without Access."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from pyaccesskit._backends.fake import FakeBackend
from pyaccesskit._session.protocols import EnginePlan
from pyaccesskit._session.session import Session
from pyaccesskit.database import AccessDatabase
from pyaccesskit.enums import Engine, Transport
from pyaccesskit.errors import CapabilityError
from pyaccesskit.options import SessionOptions


@dataclass
class FakeProbe:
    inproc: bool = True
    access_installed: bool = True
    reason: str = "fake: DAO missing"

    def inproc_dao(self) -> tuple[bool, str]:
        return self.inproc, "" if self.inproc else self.reason

    def access(self, progid: str) -> bool:
        return self.access_installed


class FakeEngine:
    def __init__(self, plan: EnginePlan, factory: FakeFactory) -> None:
        self.plan = plan
        self.factory = factory
        self.closed = False
        self.terminated = False
        shared = factory.backends.setdefault(plan.path, FakeBackend(plan.path))
        self.backend = shared
        if plan.create:
            plan.path.write_bytes(b"fake accdb")

    @property
    def transport(self) -> Transport:
        return Transport.DAO_INPROC if self.plan.kind == "dao" else Transport.ACCESS_DESIGN

    @property
    def supports_design(self) -> bool:
        return self.plan.kind == "access"

    def schema(self) -> FakeBackend:
        return self.backend

    def design(self) -> FakeBackend:
        if self.plan.kind != "access":
            raise CapabilityError("no design in fake DAO engine")
        return self.backend

    def raw(self, which: str) -> Any:
        return FakeComObject(which)

    def close(self) -> list[BaseException]:
        self.closed = True
        if self.factory.close_interrupt:
            raise KeyboardInterrupt
        return list(self.factory.close_errors)

    def terminate(self) -> None:
        self.terminated = True


class FakeComObject:
    """Looks enough like a COM object for the revocable proxies (has ``_oleobj_``)."""

    _oleobj_ = object()

    def __init__(self, label: str) -> None:
        self.label = label

    def Name(self) -> str:
        return self.label


@dataclass
class FakeFactory:
    engines: list[FakeEngine] = field(default_factory=list)
    backends: dict[Path, FakeBackend] = field(default_factory=dict)
    close_errors: list[BaseException] = field(default_factory=list)
    close_interrupt: bool = False

    def __call__(self, plan: EnginePlan) -> FakeEngine:
        engine = FakeEngine(plan, self)
        self.engines.append(engine)
        return engine


def fake_database(
    tmp_path: Path,
    *,
    create: bool = True,
    engine: Engine = Engine.AUTO,
    readonly: bool = False,
    probe: FakeProbe | None = None,
    factory: FakeFactory | None = None,
    atomic: bool = True,
) -> tuple[AccessDatabase, FakeFactory]:
    """An :class:`AccessDatabase` backed by the in-memory fake engine."""
    factory = factory or FakeFactory()
    target = tmp_path / "fake.accdb"
    working = target.with_name(".fake.pak-test.accdb") if (create and atomic) else target
    if not create and not target.exists():
        target.write_bytes(b"fake accdb")
    session = Session(
        target=target,
        working=working,
        create=create,
        readonly=readonly,
        exclusive=not readonly,
        password=None,
        engine=engine,
        options=SessionOptions(),
        factory=factory,
        probe=probe or FakeProbe(),
    )
    return AccessDatabase(session), factory
