# Contributing

## Setup

```console
uv sync                      # creates .venv with the dev tools
uv run pytest                # unit, fake-backed contract and CLI tests (no Access needed)
uv run ruff check . && uv run ruff format --check .
uv run pyright
```

## Layout

```
src/pyaccesskit/
  database.py tables.py queries.py …   public API (COM-free)
  schema/ forms/ units.py enums.py     specs, validation, layout (pure Python)
  _ops/                                orchestration: pre-validation, rollback, build-then-swap
  _backends/protocols.py               the SchemaBackend / DesignBackend contracts
  _backends/fake/                      in-memory backend, kept honest by the contract tests
  _backends/dao/ _backends/access/     real backends (COM)
  _session/ _engines/                  session state machine, engine selection
  _win/ _com/                          processes, job objects, watchdog; dispatch, errors, constants
  cli/                                 Typer commands (imported lazily)
```

Only `_com`, `_win`, `_engines` and `_backends/{dao,access}` may import `pywin32`. Ruff's `banned-api` rule
enforces this, and a unit test imports the package with `pywin32` blocked. Everything else must stay pure
Python.

## Tests

| Tier | Folder | Needs | Runs by default |
|---|---|---|---|
| Unit | `tests/unit` | nothing (any OS) | yes |
| Contract | `tests/contract` | nothing for the fake backend; Access/DAO for the real ones | fake only |
| Integration & lifecycle | `tests/integration` | Windows + Access and/or in-process DAO | with `--integration` |

```console
uv run pytest --integration                     # everything this Python can reach
uv run pytest -m "integration and lifecycle" --integration
```

On a machine with 32-bit Office, in-process DAO needs a 32-bit Python. Use a separate environment so the
default `.venv` is untouched:

```console
uv python install cpython-3.12-windows-x86
UV_PROJECT_ENVIRONMENT=.venv-x86 uv run --python cpython-3.12-windows-x86 pytest --integration
```

(In PowerShell: `$env:UV_PROJECT_ENVIRONMENT = ".venv-x86"` before the `uv run` line.)

Integration tests are guarded by a **leak check**. Every test records the `MSACCESS.EXE` processes that
existed before it ran. Afterwards it fails if new ones survive, and it terminates only processes recorded
in the test's own ownership ledger. An Access window you have open while tests run is never touched.
Starting other Access instances *during* a test run can make the leak check fail, though.

## Constants

COM constants live in `src/pyaccesskit/_com/constants.py`, generated from the installed type libraries:

```console
uv run python scripts/gen_constants.py          # regenerate
uv run python scripts/gen_constants.py --check  # fail if the file is out of date
```

An integration test compares the constants with the installed Office.

## Architecture decisions

Significant decisions and findings are recorded in `docs/adr/`. Findings that change behaviour, such as
Access quirks discovered by integration tests, get an entry too.

## Documentation

```console
uv run --group docs mkdocs serve
```
