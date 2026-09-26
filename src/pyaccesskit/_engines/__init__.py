# pyright: basic
"""Engine construction (COM). Imported lazily by :mod:`pyaccesskit._session`."""

from __future__ import annotations

from typing import Any

from pyaccesskit._session.protocols import EnginePlan

__all__ = ["open_engine"]


def open_engine(plan: EnginePlan) -> Any:
    """Create the engine described by ``plan`` (see :class:`EnginePlan`)."""
    if plan.kind == "dao":
        from pyaccesskit._engines.inproc import InProcDaoEngine

        return InProcDaoEngine(
            plan.path,
            create=plan.create,
            readonly=plan.readonly,
            exclusive=plan.exclusive,
            password=plan.password,
            native_defaults=plan.options.apply_native_defaults,
        )
    from pyaccesskit._engines.access import AccessEngine
    from pyaccesskit._win.access_process import AccessLaunchOptions

    options = plan.options
    return AccessEngine(
        plan.path,
        create=plan.create,
        readonly=plan.readonly,
        exclusive=plan.exclusive,
        password=plan.password,
        design=plan.design,
        options=AccessLaunchOptions(
            progid=options.access_progid,
            visible=options.visible,
            macro_security=options.macro_security,
            dialog_policy=options.dialog_policy,
            call_timeout=options.call_timeout,
            quit_timeout=options.quit_timeout,
            kill_on_parent_exit=options.kill_on_parent_exit,
        ),
    )
