"""``pyaccesskit doctor``: which engines work on this machine, and why not."""

from __future__ import annotations

from typing import Annotated

import typer
from rich.markup import escape
from rich.table import Table

from pyaccesskit import diagnostics
from pyaccesskit.cli._output import EXIT_ENVIRONMENT, EXIT_OK, print_json, stdout

_ENGINE_LABELS = {"dao": "in-process DAO", "access": "Microsoft Access"}


def doctor(
    json_output: Annotated[bool, typer.Option("--json", help="Print the report as JSON.")] = False,
    probe: Annotated[
        bool,
        typer.Option(
            "--probe",
            help="Also build a scratch database with each available engine (starts and closes an owned, "
            "hidden Access process).",
        ),
    ] = False,
) -> None:
    """Check which engines PyAccessKit can use here (exit code 3 if none works)."""
    report = diagnostics.diagnose(probe=probe)
    if json_output:
        print_json(report.to_dict())
    else:
        render(report)
    raise typer.Exit(EXIT_OK if report.usable else EXIT_ENVIRONMENT)


def render(report: diagnostics.Diagnosis) -> None:
    """Print a human-readable report."""
    console = stdout()
    console.print(f"[bold]PyAccessKit {report.pyaccesskit_version}[/] environment report\n")
    facts = Table.grid(padding=(0, 2))
    facts.add_column(style="bold")
    facts.add_column()
    facts.add_row(
        "Python",
        escape(f"{report.python_version} ({report.python_bits}-bit)  {report.python_executable}"),
    )
    facts.add_row("Windows", escape(report.operating_system))
    facts.add_row("pywin32", report.pywin32_version or "[red]not installed[/]")
    access = report.access
    if access is None:
        facts.add_row("Access", "[yellow]not registered[/]")
    else:
        edition = (
            f"Click-to-Run: {', '.join(access.products) or 'unknown products'}"
            if access.click_to_run
            else "MSI"
        )
        facts.add_row(
            "Access",
            escape(f"{access.version or 'unknown version'} ({access.bits or '?'}-bit, {edition})"),
        )
        if access.executable:
            facts.add_row("", escape(access.executable))
    facts.add_row("ACE OLEDB", escape(", ".join(report.ace_oledb)) or "none for this bitness")
    console.print(facts)

    console.print("\n[bold]Engines[/]")
    engines = Table.grid(padding=(0, 2))
    for check in report.engines:
        state = "[green]available[/]  " if check.available else "[red]unavailable[/]"
        engines.add_row(f"  {_ENGINE_LABELS[check.engine]}", state, escape(check.detail))
    console.print(engines)
    auto = _ENGINE_LABELS[report.auto_engine] if report.auto_engine else "[red]nothing usable[/]"
    console.print(f"  engine='auto' selects: [bold]{auto}[/]")

    if report.probes:
        console.print("\n[bold]Probes[/]")
        for result in report.probes:
            state = "[green]passed[/]" if result.ok else "[red]failed[/]"
            console.print(
                f"  {_ENGINE_LABELS[result.engine]}: {state} in {result.seconds:.1f}s - "
                f"{escape(result.detail)}"
            )
    if report.owned_processes:
        console.print("\n[bold]Access processes started by PyAccessKit[/]")
        for entry in report.owned_processes:
            console.print(
                f"  PID {entry.pid} (owner PID {entry.owner_pid}): {entry.status}"
                + (f" - {escape(entry.database)}" if entry.database else "")
            )
    if report.notes:
        console.print("\n[bold]Notes[/]")
        for note in report.notes:
            console.print(f"  - {escape(note)}")
    if report.problems:
        console.print("\n[bold red]Problems[/]")
        for problem in report.problems:
            console.print(f"  - {escape(problem)}")
    verdict = "[green]ready[/]" if report.usable else "[red]not usable[/]"
    console.print(f"\nPyAccessKit is {verdict} on this machine.")
