# Lifecycle & safety

## Sessions

`AccessDatabase.create()` and `AccessDatabase.open()` start a **session**. The session owns everything
that has to be cleaned up: the DAO connection, the Access process if one was started, and the temporary
file of an atomic create. Use the database as a context manager:

```python
with AccessDatabase.open("app.accdb") as db:
    ...
```

`db.close()` does the same explicitly and is idempotent. If you forget both, a finalizer cleans up when the
object is garbage-collected and emits a `ResourceWarning`.

## Atomic creation

```python
with AccessDatabase.create("app.accdb", overwrite=True) as db:
    build_everything(db)
```

The database is built in a hidden sibling file (`.app.pak-1a2b3c4d.accdb`). When the block ends normally,
that file replaces `app.accdb`. When it raises, the file is deleted and `app.accdb`, even an existing one
you asked to overwrite, is untouched. Pass `atomic=False` to build in place.

## Closing, even when things go wrong

Closing runs these steps in order. Each step is guarded, so one failure never skips the rest:

1. Revoke `db.raw` proxies.
2. Close objects PyAccessKit opened in Access, then the database.
3. `Quit` Access without saving (every change PyAccessKit makes is saved explicitly, as it happens).
4. Wait for the process to exit, up to `quit_timeout`. If it is still running, terminate **our** process
   handle.
5. Remove the ledger entry, then commit or discard the atomic temporary file.

If the session is closing because of an exception, cleanup failures are attached to that exception as
notes, so the original error is never masked. On a normal close they are raised together as
`CleanupError`.

## Process ownership

- **A fresh instance, every time.** PyAccessKit starts Access with `CoCreateInstanceEx` as a local server.
  It never uses `Dispatch()` or `GetObject()`, which attach to an Access instance the user already has
  running.
- **Identity, not names.** The process is identified by PID, creation time and image path, and held open
  by handle, so a recycled PID can never be confused with it. PyAccessKit never kills `MSACCESS.EXE` by
  name.
- **Kill on crash.** The process is placed in a Windows job object with *kill-on-close*. If Python dies for
  any reason, including a hard crash or `os._exit`, Windows ends Access too.
- **Ledger.** Each owned process is recorded in `%LOCALAPPDATA%\PyAccessKit\owned`.
  `pyaccesskit cleanup` (or `pyaccesskit.reap_orphans()`) ends only recorded processes whose Python owner
  has exited, and only after re-checking their identity.

## Dialogs, timeouts and Ctrl+C

Office is not designed for unattended automation. A hidden Access can show a modal dialog and wait forever.
PyAccessKit watches for dialogs, but only in windows that belong to its own Access process:

- `dialog_policy="fail"` (default): the dialog is recorded, dismissed (Cancel, then No, then OK) and
  reported as `AccessDialogError`, with its title, text and buttons.
- `dialog_policy="warn"`: dismissed and reported as an `AccessDialogWarning`.
- `dialog_policy="off"`: not watched.

A single operation that runs longer than `call_timeout` (default 600 s) terminates the owned process and
raises `AccessTimeoutError`. Pressing Ctrl+C during a long call terminates the owned process, so
`KeyboardInterrupt` surfaces promptly.

## Threads

COM objects belong to the thread (apartment) that created them. A session can only be used from the thread
that opened it; other threads get `WrongThreadError`. For parallel work, use separate processes, each with
its own session.

## Session options

```python
from pyaccesskit import AccessDatabase, DialogPolicy, MacroSecurity, SessionOptions

options = SessionOptions(visible=True, call_timeout=120, dialog_policy=DialogPolicy.WARN)
with AccessDatabase.open("app.accdb", options=options) as db:
    ...
```

| Option | Default | Meaning |
|---|---|---|
| `visible` | `False` | Show the Access window (useful for debugging) |
| `macro_security` | `MacroSecurity.DISABLE` | What Access may run when it opens the database for design: `DISABLE`, `USE_UI` (the user's Trust Center settings) or `ENABLE` |
| `dialog_policy` | `DialogPolicy.FAIL` | See above |
| `call_timeout` | `600.0` | Seconds per operation; `None` disables the limit |
| `quit_timeout` | `30.0` | Seconds to wait for Access to exit before terminating it |
| `kill_on_parent_exit` | `True` | Use the kill-on-close job object |
| `access_progid` | `"Access.Application"` | For example, `"Access.Application.16"` on machines with several versions |
| `apply_native_defaults` | `True` | Add Access's default database properties when DAO creates a database |
