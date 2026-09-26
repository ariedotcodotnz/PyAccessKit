"""Documentation stays true: generated files are in sync, snippets parse, and their imports exist."""

from __future__ import annotations

import ast
import json
import os
import re
import subprocess
import sys
from pathlib import Path

import pytest
from typer.testing import CliRunner

import pyaccesskit
from pyaccesskit.cli import app as cli_app
from pyaccesskit.cli.agent import guide_path, guide_text

ROOT = Path(__file__).resolve().parents[2]
GUIDE = ROOT / "src" / "pyaccesskit" / "AGENT_GUIDE.md"
PAGES = [*sorted((ROOT / "docs").rglob("*.md")), GUIDE, ROOT / "README.md"]
_PYTHON_BLOCK = re.compile(r"```python\n(.*?)```", re.DOTALL)


def _blocks() -> list[tuple[str, str]]:
    blocks: list[tuple[str, str]] = []
    for page in PAGES:
        for index, match in enumerate(_PYTHON_BLOCK.finditer(page.read_text(encoding="utf-8"))):
            blocks.append((f"{page.relative_to(ROOT).as_posix()}#{index + 1}", match.group(1)))
    return blocks


BLOCKS = _blocks()


def test_generated_docs_are_in_sync() -> None:
    result = subprocess.run(
        [sys.executable, str(ROOT / "scripts" / "sync_docs.py"), "--check"],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 0, result.stdout + result.stderr


def test_guide_contains_the_tested_example() -> None:
    example = (ROOT / "examples" / "04_inventory_app.py").read_text(encoding="utf-8").strip()
    assert example in GUIDE.read_text(encoding="utf-8")


@pytest.mark.parametrize(("where", "code"), BLOCKS, ids=[where for where, _ in BLOCKS])
def test_python_snippets_parse(where: str, code: str) -> None:
    ast.parse(code, filename=where)


@pytest.mark.parametrize(("where", "code"), BLOCKS, ids=[where for where, _ in BLOCKS])
def test_snippet_imports_exist(where: str, code: str) -> None:
    for node in ast.walk(ast.parse(code)):
        if isinstance(node, ast.ImportFrom) and (node.module or "").startswith("pyaccesskit"):
            module = __import__(node.module or "", fromlist=["_"])
            missing = [alias.name for alias in node.names if not hasattr(module, alias.name)]
            assert not missing, f"{where}: {node.module} has no {missing}"


def test_guide_names_are_exported() -> None:
    names = {
        "AccessDatabase", "Column", "Expr", "TableSpec", "IndexSpec", "RelationshipSpec", "QuerySpec",
        "FormView", "LayoutKind", "NumberSize", "Section", "RowSourceType", "ModuleKind", "JoinType",
        "QueryKind", "SessionOptions", "MacroSecurity", "DialogPolicy", "Vba", "cm", "mm", "inch", "pt",
        "twips", "AccessNameWarning", "SpecError", "diagnose",
    }  # fmt: skip
    assert names <= set(pyaccesskit.__all__)


def test_guide_command() -> None:
    runner = CliRunner()
    result = runner.invoke(cli_app.app, ["guide"])
    assert result.exit_code == 0
    assert result.output.startswith("# PyAccessKit")
    assert result.output == guide_text()
    assert Path(guide_path()).is_file()


@pytest.mark.parametrize("kind", ["table", "column", "index", "relationship", "query", "form"])
def test_schema_command(kind: str) -> None:
    result = CliRunner().invoke(cli_app.app, ["schema", kind])
    assert result.exit_code == 0
    assert json.loads(result.output)


def test_schema_all_validates_a_real_spec() -> None:
    data = json.loads(CliRunner().invoke(cli_app.app, ["schema"]).output)
    assert set(data) == {"table", "column", "index", "relationship", "query", "form"}
    assert "columns" in data["table"]["properties"]


def test_guide_can_be_piped_with_an_ansi_code_page() -> None:
    """Agents read the guide through a pipe; on Windows that pipe defaults to cp1252."""
    result = subprocess.run(
        [sys.executable, "-m", "pyaccesskit", "guide"],
        capture_output=True,
        env={**os.environ, "PYTHONIOENCODING": "cp1252"},
        check=False,
    )
    assert result.returncode == 0, result.stderr.decode("utf-8", "replace")
    assert result.stdout.decode("utf-8") == guide_text()
