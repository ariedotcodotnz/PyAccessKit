"""The ownership ledger: a record of every Access process PyAccessKit started.

Each live Access process owned by PyAccessKit has a small JSON file in
``%LOCALAPPDATA%/PyAccessKit/owned``. It is deleted when the process is shut down cleanly. If Python died
before that (and the kill-on-close job object could not be used), :func:`reap_orphans` finds the leftovers.

A process is only ever terminated when **all** of these hold:

* its owner (the Python process that started it) is gone — PID *and* creation time checked;
* a process with the recorded PID *and* creation time still exists (so the PID was not recycled);
* its image is ``MSACCESS.EXE`` (checked again by the inspector right before terminating).

This module is pure Python; the Windows-specific inspector lives in ``_win.inspector``.
"""

from __future__ import annotations

import json
import logging
import os
import time
from collections.abc import Iterator
from dataclasses import asdict, dataclass
from pathlib import Path, PureWindowsPath
from typing import Literal, Protocol

__all__ = [
    "OwnedProcess",
    "ProcessInspector",
    "ReapResult",
    "ledger_dir",
    "load_all",
    "reap_orphans",
    "record",
    "remove",
]

logger = logging.getLogger("pyaccesskit.process")

LEDGER_ENV = "PYACCESSKIT_LEDGER_DIR"


@dataclass(frozen=True)
class OwnedProcess:
    """One ledger entry."""

    pid: int
    creation_time: float
    image: str
    owner_pid: int
    owner_creation_time: float
    database: str | None
    started_at: float
    version: str

    @property
    def file_name(self) -> str:
        """Ledger file name (unique per process identity)."""
        return f"{self.pid}-{int(self.creation_time * 1000)}.json"


class ProcessInspector(Protocol):
    """Answers questions about processes; implemented with Win32 calls in ``_win.inspector``."""

    def matches(self, pid: int, creation_time: float) -> bool:
        """Whether a running process has exactly this PID and creation time."""
        ...

    def terminate(self, pid: int, creation_time: float, image_name: str) -> bool:
        """Terminate the process if (re-verified) it has this identity and image; return success."""
        ...


@dataclass(frozen=True)
class ReapResult:
    """What :func:`reap_orphans` did with one ledger entry."""

    entry: OwnedProcess
    action: Literal["terminated", "would-terminate", "already-exited", "owner-alive", "failed"]


def ledger_dir() -> Path:
    """The ledger directory (``PYACCESSKIT_LEDGER_DIR`` overrides it, e.g. in tests)."""
    override = os.environ.get(LEDGER_ENV)
    if override:
        return Path(override)
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    return Path(base) / "PyAccessKit" / "owned"


def record(entry: OwnedProcess, directory: Path | None = None) -> Path:
    """Write a ledger entry atomically; returns its path."""
    folder = directory or ledger_dir()
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / entry.file_name
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(asdict(entry), indent=2), encoding="utf-8")
    temporary.replace(path)
    return path


def remove(path: Path) -> None:
    """Delete a ledger entry (missing files are fine)."""
    try:
        path.unlink(missing_ok=True)
    except (
        OSError
    ) as exc:  # pragma: no cover - e.g. a locked file; the entry is simply reaped later
        logger.debug("could not delete ledger entry %s: %s", path, exc)


def load_all(directory: Path | None = None) -> Iterator[tuple[Path, OwnedProcess]]:
    """Yield every readable ledger entry (unreadable files are skipped and logged)."""
    folder = directory or ledger_dir()
    if not folder.is_dir():
        return
    for path in sorted(folder.glob("*.json")):
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            yield path, OwnedProcess(**data)
        except (OSError, ValueError, TypeError) as exc:
            logger.warning("ignoring unreadable ledger entry %s: %s", path, exc)


def reap_orphans(
    inspector: ProcessInspector, *, dry_run: bool = False, directory: Path | None = None
) -> list[ReapResult]:
    """Terminate Access processes whose PyAccessKit owner died without shutting them down.

    Args:
        inspector: Process inspector (Win32 implementation by default in the public API).
        dry_run: Report what would be terminated without terminating anything.
        directory: Ledger directory override.
    """
    results: list[ReapResult] = []
    for path, entry in load_all(directory):
        if inspector.matches(entry.owner_pid, entry.owner_creation_time):
            results.append(ReapResult(entry, "owner-alive"))
            continue
        if not inspector.matches(entry.pid, entry.creation_time):
            if not dry_run:
                remove(path)
            results.append(ReapResult(entry, "already-exited"))
            continue
        if dry_run:
            results.append(ReapResult(entry, "would-terminate"))
            continue
        if inspector.terminate(entry.pid, entry.creation_time, PureWindowsPath(entry.image).name):
            remove(path)
            results.append(ReapResult(entry, "terminated"))
        else:
            results.append(ReapResult(entry, "failed"))
    return results


def now() -> float:
    """Current time (seconds since the epoch); separated for tests."""
    return time.time()
