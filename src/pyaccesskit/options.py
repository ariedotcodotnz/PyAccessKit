"""Session options (how PyAccessKit starts and supervises Microsoft Access)."""

from __future__ import annotations

from dataclasses import dataclass

from pyaccesskit.enums import DialogPolicy, MacroSecurity

__all__ = ["SessionOptions"]


@dataclass(frozen=True)
class SessionOptions:
    """Advanced settings for a database session. The defaults are right for build scripts.

    Attributes:
        visible: Show the Access window (useful for debugging). Dialog detection then only considers
            standard dialog windows.
        macro_security: What Access may run when it opens a database as its current database. The default,
            ``DISABLE``, never runs AutoExec macros, startup-form code or VBA — whatever the user's Trust
            Center says.
        dialog_policy: What to do when Access shows a modal dialog during an automated call.
        call_timeout: Seconds one operation may take before the owned Access process is terminated
            (``None`` = no limit).
        quit_timeout: Seconds to wait for Access to exit after ``Quit`` before terminating it.
        kill_on_parent_exit: Put Access in a Windows job object so it dies with this Python process.
        access_progid: ProgID used to start Access (e.g. ``"Access.Application.16"`` on multi-version
            machines).
        apply_native_defaults: When DAO creates a database, add the properties Access itself writes to new
            databases (tabbed documents, themed controls...), so the result matches ``engine="access"``.
    """

    visible: bool = False
    macro_security: MacroSecurity = MacroSecurity.DISABLE
    dialog_policy: DialogPolicy = DialogPolicy.FAIL
    call_timeout: float | None = 600.0
    quit_timeout: float = 30.0
    kill_on_parent_exit: bool = True
    access_progid: str = "Access.Application"
    apply_native_defaults: bool = True
