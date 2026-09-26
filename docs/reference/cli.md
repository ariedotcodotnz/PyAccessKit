# Command line

```console
pyaccesskit [--version] COMMAND [OPTIONS]
python -m pyaccesskit COMMAND [OPTIONS]
```

Exit codes: `0` success, `1` error, `2` usage error, `3` environment unusable. With `--json`, output is
ASCII-safe JSON on stdout; errors go to stderr.

## `doctor`

```console
pyaccesskit doctor [--json] [--probe]
```

Reports Python, Windows, pywin32 and Access versions and bitness, and whether in-process DAO and
Microsoft Access can be used. It also shows what `engine="auto"` will choose and any orphaned Access
processes. `--probe` builds a scratch database with each available engine, starting and closing an owned,
hidden Access. The exit code is 3 when nothing is usable or a probe failed.

JSON fields: `usable`, `auto_engine`, `engines[{engine, available, detail}]`,
`access{version, bits, executable, click_to_run, products}`, `ace_oledb`, `probes[{engine, ok, seconds,
detail}]`, `owned_processes`, `problems`, `notes`, plus version facts.

## `inspect`

```console
pyaccesskit inspect DATABASE [--json] [--counts] [--password PW] [--engine auto|dao|access]
```

Describes tables (columns, types, defaults, indexes), relationships, queries (kind and SQL), and the names
of forms, reports, macros and modules. It opens the database read-only in shared mode, so it works while
the file is open in Access. It never runs startup code. `--counts` adds row counts. The password can also
come from `PYACCESSKIT_PASSWORD`. Passwords inside connection strings are shown as `***`.

JSON: `{path, format_version, transport, tables[{name, linked, rows, spec}], relationships[...],
queries[{kind, name, sql, description, pass_through}], forms, reports, macros, modules, warnings}`. Each
`tables[].spec` is a `TableSpec` in JSON form (`TableSpec.model_validate(entry["spec"])`).

## `cleanup`

```console
pyaccesskit cleanup [--dry-run] [--json]
```

Ends Access processes started by PyAccessKit sessions whose Python process has exited, usually after a
crash. Only processes recorded in PyAccessKit's ownership ledger are considered, and their identity (PID,
creation time, executable) is re-checked first. Access windows you opened yourself are never touched.
The exit code is 1 if a process could not be ended.

## `guide`

```console
pyaccesskit guide [--path]
```

Prints the [guide for AI coding agents](../agents/guide.md) that ships with the installed version, or its
file path with `--path`.

## `schema`

```console
pyaccesskit schema [all|table|column|index|relationship|query|form]
```

Prints the JSON Schema of the spec models. Editors and AI agents can use it to validate specs written as
JSON before touching Access.
