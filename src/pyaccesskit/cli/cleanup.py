"""``pyaccesskit cleanup``: end orphaned Access processes that PyAccessKit started."""

from __future__ import annotations

from typing import Annotated

import typer
from rich.markup import escape

from pyaccesskit.cli._output import EXIT_ERROR, EXIT_OK, handle_errors, print_json, stdout
from pyaccesskit.maintenance import reap_orphans

_DESCRIPTIONS = {
    "terminated": "terminated (its Python owner had exited)",
    "would-terminate": "would be terminated (its Python owner has exited)",
    "already-exited": "already exited; ledger entry removed",
    "owner-alive": "left alone: its Python owner is still running",
    "failed": "could not be terminated",
}


def cleanup(
    dry_run: Annotated[
        bool, typer.Option("--dry-run", help="Only report what would be terminated.")
    ] = False,
    json_output: Annotated[bool, typer.Option("--json", help="Print the result as JSON.")] = False,
) -> None:
    """End Access processes left behind by crashed PyAccessKit sessions.

    Only processes recorded in PyAccessKit's ownership ledger are considered, and only when the Python
    process that started them is gone. Access instances you or other tools started are never touched.
    """
    with handle_errors():
        results = reap_orphans(dry_run=dry_run)
    if json_output:
        print_json(
            [
                {
                    "pid": r.entry.pid,
                    "owner_pid": r.entry.owner_pid,
                    "database": r.entry.database,
                    "action": r.action,
                }
                for r in results
            ]
        )
    else:
        console = stdout()
        if not results:
            console.print("No Access processes started by PyAccessKit are recorded. Nothing to do.")
        for result in results:
            where = f" ({escape(result.entry.database)})" if result.entry.database else ""
            console.print(
                f"PID {result.entry.pid}{where}: {_DESCRIPTIONS.get(result.action, result.action)}"
            )
    raise typer.Exit(EXIT_ERROR if any(r.action == "failed" for r in results) else EXIT_OK)
