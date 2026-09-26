"""Maintenance helpers: cleaning up orphaned Access processes."""

from __future__ import annotations

import sys

from pyaccesskit import _ledger
from pyaccesskit._ledger import OwnedProcess, ReapResult

__all__ = ["OwnedProcess", "ReapResult", "owned_processes", "reap_orphans"]


def owned_processes() -> list[OwnedProcess]:
    """Access processes currently recorded in PyAccessKit's ownership ledger (from any Python process)."""
    return [entry for _path, entry in _ledger.load_all()]


def reap_orphans(*, dry_run: bool = False) -> list[ReapResult]:
    """Terminate Access processes that PyAccessKit started but whose Python owner has died.

    Only processes recorded in the ownership ledger are considered, and each one is terminated only if its
    owner is gone *and* its PID, creation time and image (``MSACCESS.EXE``) still match. Access instances
    started by the user or by other tools are never touched.

    Args:
        dry_run: Report what would be terminated without terminating anything.
    """
    if sys.platform != "win32":
        return []  # Access only runs on Windows: nothing can have been started here
    from pyaccesskit._win.inspector import Win32Inspector

    return _ledger.reap_orphans(Win32Inspector(), dry_run=dry_run)
