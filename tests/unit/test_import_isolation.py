"""The COM boundary: importing PyAccessKit must never require pywin32 (Linux CI, spec validation...)."""

from __future__ import annotations

import subprocess
import sys
import textwrap

BLOCKED = (
    "pythoncom",
    "pywintypes",
    "win32com",
    "win32api",
    "win32con",
    "win32event",
    "win32gui",
    "win32job",
    "win32process",
)


def test_public_package_imports_without_pywin32() -> None:
    script = textwrap.dedent(
        f"""
        import importlib.abc, sys

        BLOCKED = {BLOCKED!r}

        class Blocker(importlib.abc.MetaPathFinder):
            def find_spec(self, name, path=None, target=None):
                if name.split(".")[0] in BLOCKED:
                    raise ImportError(f"blocked for test: {{name}}")
                return None

        sys.meta_path.insert(0, Blocker())
        import pyaccesskit
        import pyaccesskit.schema, pyaccesskit.forms, pyaccesskit.units, pyaccesskit.errors, pyaccesskit.enums
        import pyaccesskit._backends.fake, pyaccesskit._ops.schema, pyaccesskit._ops.design, pyaccesskit._text.codec
        leaked = sorted(m for m in sys.modules if m.split(".")[0] in BLOCKED)
        assert not leaked, leaked
        print("ok")
        """
    )
    result = subprocess.run(
        [sys.executable, "-c", script], capture_output=True, text=True, check=False
    )
    assert result.returncode == 0, result.stderr
    assert result.stdout.strip() == "ok"
