"""The documented examples run end to end against real Access (the agent guide embeds example 04)."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest

pytestmark = [pytest.mark.integration, pytest.mark.access]

EXAMPLES = Path(__file__).resolve().parents[2] / "examples"


@pytest.mark.parametrize(
    ("script", "output"),
    [
        ("01_create_crm.py", "crm.accdb"),
        ("03_build_small_app.py", "orders.accdb"),
        ("04_inventory_app.py", "inventory.accdb"),
    ],
)
def test_example_runs(tmp_path: Path, script: str, output: str) -> None:
    target = tmp_path / output
    result = subprocess.run(
        [sys.executable, str(EXAMPLES / script), str(target)],
        capture_output=True,
        text=True,
        timeout=300,
        check=False,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    assert target.exists()
    inspected = subprocess.run(
        [sys.executable, str(EXAMPLES / "02_inspect_database.py"), str(target)],
        capture_output=True,
        text=True,
        timeout=300,
        check=False,
    )
    assert inspected.returncode == 0, inspected.stdout + inspected.stderr
