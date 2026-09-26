"""Command-line interface (``pyaccesskit``). Heavy imports happen lazily inside :func:`main`."""

from __future__ import annotations


def main() -> None:
    """Entry point for the ``pyaccesskit`` console script."""
    from pyaccesskit.cli.app import run  # noqa: PLC0415 - keep `import pyaccesskit` light

    run()
