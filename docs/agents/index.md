# Building with AI agents

PyAccessKit is designed so that an AI coding agent can write Microsoft Access software reliably:

- **Everything is text and data.** Tables, relationships, queries and forms are specs: immutable,
  validated Python objects with a JSON form and a JSON Schema (`pyaccesskit schema`). An agent writes specs,
  not COM calls.
- **Mistakes fail early, and with explanations.** Invalid specs raise `SpecError` with every problem
  listed, before Access is touched. Access errors become specific exceptions that say what failed and why.
- **Runs are safe to repeat.** `AccessDatabase.create()` is atomic, and every Access process PyAccessKit
  starts is closed, even when the agent's script crashes. An agent can iterate (write, run, read the error,
  fix) without leaving broken files or stray `MSACCESS.EXE` processes behind.
- **The result can be checked.** `to_spec()` reads back exactly what was built, `pyaccesskit inspect
  --json` describes any database, and `check_opens()` proves each form opens and its VBA compiles.

## Give your agent the guide

The **[agent guide](guide.md)** is a compact, authoritative reference: rules, API cheat sheet, SQL
dialect notes, error table, limits and a complete tested example. It ships with the package, so the
version always matches the installed library:

```console
pyaccesskit guide          # print it
pyaccesskit guide --path   # where the file is, for tools that attach files
```

For tools that read documentation from the web, the site also publishes [`llms.txt`](../llms.txt), an
index, and [`llms-full.txt`](../llms-full.txt), all documentation in one file.

## Instructions for your project

Put something like this in the instructions file your coding agent reads (for example `AGENTS.md` in your
project):

```markdown
## Microsoft Access work

This project builds its Access database with PyAccessKit (Python). Before writing Access code:

1. Run `pyaccesskit guide` and follow it exactly.
2. Run `pyaccesskit doctor --json`; stop and report if `usable` is false.
3. To understand the current database: `pyaccesskit inspect <file> --json`.

Rules: never use win32com/pywin32 directly for Access, never kill MSACCESS.EXE, always use
`with AccessDatabase.create(...)` / `open(...)`, bind SQL values as parameters, and verify by
reopening the file (to_spec, fetch, check_opens) before reporting success.
```

## A good loop

1. **Describe the application** in domain terms: entities, fields, rules, screens, reports.
2. **Let the agent write the specs** (tables, relationships, queries) and a `build()` function, following
   the guide's three-part structure.
3. **Run and verify.** Add forms and VBA only once the schema verifies.
4. **Review the result in Access.** Open the file and check the relationships window, the forms, and the
   data.
5. **Keep the script in source control.** Rebuilding from the script is reproducible: the script is the
   source of your application, the `.accdb` is a build output.

## What agents should not do

- Drive Access with raw `win32com` calls. They attach to the user's own Access, leak processes and give
  unreadable errors. The guide's rules forbid it.
- Change a production database without a backup. Changes to an existing file are applied step by step;
  copy the file first.
- Use features PyAccessKit does not support yet (reports, subforms…) by improvising with `db.raw`. They
  should stop and ask, or leave a clear TODO.
