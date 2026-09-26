"""Test-suite configuration.

Tiers (see docs/contributing.md):

* ``tests/unit``, ``tests/com``, ``tests/contract`` run everywhere by default;
* integration tests (``@pytest.mark.integration``) need Windows with Microsoft Access and/or in-process
  DAO and only run with ``--integration``.

Every integration test is wrapped in a **leak guard**: Access processes that appear during the test must be
gone afterwards. The guard only ever terminates processes recorded in PyAccessKit's ownership ledger — an
Access instance you have open is never touched.
"""

from __future__ import annotations

import sys
import time
from collections.abc import Iterator
from pathlib import Path

import pytest


def pytest_addoption(parser: pytest.Parser) -> None:
    parser.addoption(
        "--integration",
        action="store_true",
        default=False,
        help="run integration tests against real Microsoft Access / DAO (Windows only)",
    )


def pytest_collection_modifyitems(config: pytest.Config, items: list[pytest.Item]) -> None:
    if config.getoption("--integration"):
        return
    skip = pytest.mark.skip(
        reason="integration test: pass --integration to run against real Access/DAO"
    )
    for item in items:
        if "integration" in item.keywords:
            item.add_marker(skip)


@pytest.fixture(autouse=True)
def _isolated_ledger(
    tmp_path_factory: pytest.TempPathFactory, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Keep the test-suite's ownership ledger separate from the user's real one."""
    ledger = tmp_path_factory.getbasetemp() / "ledger"
    monkeypatch.setenv("PYACCESSKIT_LEDGER_DIR", str(ledger))


@pytest.fixture(autouse=True)
def _leak_guard(request: pytest.FixtureRequest) -> Iterator[None]:
    if "integration" not in request.keywords or sys.platform != "win32":
        yield
        return
    from pyaccesskit._win.processes import access_process_ids

    before = access_process_ids()
    yield
    deadline = time.monotonic() + 20
    leaked = access_process_ids() - before
    while leaked and time.monotonic() < deadline:
        time.sleep(0.25)
        leaked = access_process_ids() - before
    if leaked:
        from pyaccesskit._win.inspector import Win32Inspector
        from pyaccesskit.maintenance import owned_processes

        ours = {entry.pid: entry for entry in owned_processes()}
        inspector = Win32Inspector()
        for pid in leaked & ours.keys():
            entry = ours[pid]
            inspector.terminate(entry.pid, entry.creation_time, Path(entry.image).name)
        pytest.fail(f"Access process(es) {sorted(leaked)} were still running after the test")
