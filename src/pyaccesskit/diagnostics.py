"""Environment diagnostics: which engines PyAccessKit can use on this machine, and why not.

``pyaccesskit doctor`` prints this report; :func:`diagnose` returns it as data (:meth:`Diagnosis.to_dict`
gives a JSON-ready form). Without ``probe=True`` nothing is started: the report comes from the registry, file
headers and loading DAO in-process. With ``probe=True`` a scratch database is built with each usable engine,
in an Access process owned (and closed) by PyAccessKit.
"""

from __future__ import annotations

import dataclasses
import importlib.metadata
import platform
import struct
import sys
import tempfile
import time
import warnings
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

from pyaccesskit._version import __version__
from pyaccesskit.errors import AccessNameWarning, PyAccessKitError

__all__ = [
    "AccessInstallation",
    "Diagnosis",
    "EngineCheck",
    "OwnedProcessReport",
    "ProbeReport",
    "diagnose",
]

EngineName = Literal["dao", "access"]

_INSTALL_ADVICE = (
    "Install Microsoft Access (Microsoft 365, 2016 or later) for full functionality, or the Microsoft 365 "
    "Access Runtime with a Python of the same bitness for DAO-only work."
)
_AC_SYSCMD_RUNTIME = 6


@dataclass(frozen=True)
class EngineCheck:
    """Whether one engine can be used from this Python, with an explanation."""

    engine: EngineName
    available: bool
    detail: str


@dataclass(frozen=True)
class AccessInstallation:
    """Microsoft Access as registered for automation."""

    progid: str
    registered_version: str | None
    executable: str | None
    version: str | None
    bits: int | None
    click_to_run: bool
    products: tuple[str, ...] = ()


@dataclass(frozen=True)
class OwnedProcessReport:
    """An Access process recorded in the ownership ledger, and what ``cleanup`` would do with it."""

    pid: int
    owner_pid: int
    database: str | None
    status: str


@dataclass(frozen=True)
class ProbeReport:
    """The outcome of building a scratch database with one engine."""

    engine: EngineName
    ok: bool
    seconds: float
    detail: str


@dataclass(frozen=True)
class Diagnosis:
    """Everything ``pyaccesskit doctor`` reports.

    Attributes:
        auto_engine: What ``engine="auto"`` selects here (``None``: nothing is usable).
        engines: Availability of in-process DAO and of Microsoft Access.
        ace_oledb: ACE OLEDB providers loadable in-process (used for Decimal columns with in-process DAO).
        owned_processes: Ledger entries of Access processes started by PyAccessKit.
        probes: Results of ``probe=True`` runs.
        problems: Reasons PyAccessKit cannot work (or cannot work fully).
        notes: Useful facts and advice that are not problems.
    """

    pyaccesskit_version: str
    python_version: str
    python_bits: int
    python_executable: str
    operating_system: str
    pywin32_version: str | None
    access: AccessInstallation | None
    engines: tuple[EngineCheck, ...]
    ace_oledb: tuple[str, ...]
    auto_engine: EngineName | None
    owned_processes: tuple[OwnedProcessReport, ...]
    probes: tuple[ProbeReport, ...]
    problems: tuple[str, ...]
    notes: tuple[str, ...]

    @property
    def usable(self) -> bool:
        """At least one engine is available and every probe that ran succeeded."""
        return self.auto_engine is not None and all(probe.ok for probe in self.probes)

    def to_dict(self) -> dict[str, Any]:
        """A JSON-ready dictionary (includes ``usable``)."""
        data = dataclasses.asdict(self)
        data["usable"] = self.usable
        return data


def _package_version(name: str) -> str | None:
    try:
        return importlib.metadata.version(name)
    except importlib.metadata.PackageNotFoundError:
        return None


def diagnose(*, probe: bool = False, progid: str = "Access.Application") -> Diagnosis:
    """Inspect this machine and report which PyAccessKit engines work.

    Args:
        probe: Also build a scratch database with every available engine (starts and closes an owned
            Access process; takes a few seconds).
        progid: The Access ProgID to check (e.g. ``"Access.Application.16"`` on multi-version machines).
    """
    bits = struct.calcsize("P") * 8
    base: dict[str, Any] = {
        "pyaccesskit_version": __version__,
        "python_version": platform.python_version(),
        "python_bits": bits,
        "python_executable": sys.executable,
        "operating_system": platform.platform(),
        "pywin32_version": _package_version("pywin32"),
    }
    problems: list[str] = []
    notes: list[str] = []
    empty: dict[str, Any] = {
        "access": None,
        "engines": (),
        "ace_oledb": (),
        "auto_engine": None,
        "owned_processes": (),
        "probes": (),
    }

    if sys.platform != "win32":
        problems.append(
            "Microsoft Access automation needs Windows; here PyAccessKit can only build and validate specs."
        )
        return Diagnosis(**base, **empty, problems=tuple(problems), notes=())
    gil_check = getattr(sys, "_is_gil_enabled", None)
    if gil_check is not None and not gil_check():
        problems.append(
            "free-threaded Python builds are not supported (COM apartments need the GIL)"
        )
    try:
        from pyaccesskit._engines import probe as facts
    except ImportError as exc:
        problems.append(f"pywin32 is not usable ({exc}); install it with 'pip install pywin32'")
        return Diagnosis(**base, **empty, problems=tuple(problems), notes=())

    access_facts = facts.access_facts(progid)
    c2r = facts.click_to_run()
    access = (
        AccessInstallation(
            progid=progid,
            registered_version=access_facts.current_version,
            executable=access_facts.executable,
            version=access_facts.version,
            bits=access_facts.bits,
            click_to_run=c2r is not None,
            products=c2r.products if c2r else (),
        )
        if access_facts.registered
        else None
    )
    dao = facts.inproc_dao()
    dao_check = EngineCheck(
        "dao",
        dao.available,
        f"DAO.DBEngine.120 loads in this {bits}-bit Python" if dao.available else dao.reason,
    )
    if access is None:
        access_detail = f"{progid} is not registered"
    elif access.executable is None:
        access_detail = f"{progid} is registered but its executable was not found"
    else:
        access_detail = f"Access {access.version or '(unknown version)'}, {access.bits or '?'}-bit"
    access_check = EngineCheck("access", access is not None, access_detail)
    ace = tuple(name for name in facts.ACE_OLEDB_PROGIDS if facts.inproc_registration(name)[bits])

    auto: EngineName | None = "dao" if dao.available else ("access" if access else None)
    if dao.available and access:
        notes.append(
            "engine='auto' uses in-process DAO for schema work and switches to Microsoft Access for forms, "
            "modules and text import/export."
        )
    elif dao.available:
        notes.append(
            "Forms, modules and text import/export need Microsoft Access (not the Runtime)."
        )
        if not ace:
            notes.append(
                "No ACE OLEDB provider is registered for this Python's bitness: Decimal columns cannot be "
                "created in-process."
            )
    elif access:
        other = access.bits if access.bits and access.bits != bits else None
        advice = f" Install a {other}-bit Python to use DAO in-process (faster)." if other else ""
        notes.append(
            "engine='auto' reaches DAO through Microsoft Access, so every session starts an owned, hidden "
            f"MSACCESS.EXE.{advice}"
        )
    else:
        problems.append(f"Neither in-process DAO nor Microsoft Access is usable. {_INSTALL_ADVICE}")
    if access and access.version and not access.version.startswith("16."):
        notes.append(
            f"Access {access.version} is older than the supported floor (16.0, Access 2016); it may work."
        )

    owned = _owned_processes()
    orphans = sum(1 for entry in owned if entry.status == "would-terminate")
    if orphans:
        notes.append(
            f"{orphans} orphaned Access process(es) from crashed PyAccessKit sessions are still running; "
            "run 'pyaccesskit cleanup' to end them."
        )

    probes: list[ProbeReport] = []
    if probe:
        if dao.available:
            probes.append(_probe("dao"))
        if access:
            probes.append(_probe("access"))
        problems.extend(f"{p.engine} probe failed: {p.detail}" for p in probes if not p.ok)

    return Diagnosis(
        **base,
        access=access,
        engines=(dao_check, access_check),
        ace_oledb=ace,
        auto_engine=auto,
        owned_processes=owned,
        probes=tuple(probes),
        problems=tuple(problems),
        notes=tuple(notes),
    )


def _owned_processes() -> tuple[OwnedProcessReport, ...]:
    from pyaccesskit.maintenance import reap_orphans

    return tuple(
        OwnedProcessReport(
            pid=result.entry.pid,
            owner_pid=result.entry.owner_pid,
            database=result.entry.database,
            status=result.action,
        )
        for result in reap_orphans(dry_run=True)
    )


def _probe(engine: EngineName) -> ProbeReport:
    """Build a scratch database end to end with ``engine``."""
    from pyaccesskit.database import AccessDatabase
    from pyaccesskit.schema import Column

    start = time.perf_counter()
    detail = "created a table, wrote and read a row"
    try:
        with (
            warnings.catch_warnings(),
            tempfile.TemporaryDirectory(
                prefix="pyaccesskit-probe-", ignore_cleanup_errors=True
            ) as folder,
            AccessDatabase.create(Path(folder) / "probe.accdb", engine=engine) as db,
        ):
            warnings.simplefilter("ignore", AccessNameWarning)
            db.tables.create(
                "PakProbe",
                columns=[
                    Column.autonumber("ProbeID", primary_key=True),
                    Column.text("ProbeNote", length=20),
                ],
            )
            db.execute("INSERT INTO PakProbe (ProbeNote) VALUES ([note])", {"note": "ok"})
            if db.fetch_all("SELECT ProbeNote FROM PakProbe") != [{"ProbeNote": "ok"}]:
                raise PyAccessKitError("the probe row did not read back correctly")
            if engine == "access":
                db.modules.create(
                    "modPakProbe", "Public Function Ping() As Long\n    Ping = 1\nEnd Function\n"
                )
                app = db.raw.access
                runtime = bool(app.SysCmd(_AC_SYSCMD_RUNTIME))
                detail = (
                    f"Access {app.Version} (build {app.Build}){' Runtime' if runtime else ''}: "
                    "created a table and a VBA module, wrote and read a row"
                )
    except Exception as exc:
        message = str(exc) if isinstance(exc, PyAccessKitError) else f"{type(exc).__name__}: {exc}"
        return ProbeReport(engine, False, round(time.perf_counter() - start, 2), message)
    return ProbeReport(engine, True, round(time.perf_counter() - start, 2), detail)
