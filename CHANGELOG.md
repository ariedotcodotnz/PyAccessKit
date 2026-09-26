# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project uses
[Semantic Versioning](https://semver.org/) (0.x releases may contain breaking changes, which are listed
under **Changed**).

## [Unreleased]

## [0.1.0] - 2026-09-26

First release.

### Added

- **Lifecycle.** `AccessDatabase.create()` (atomic by default) and `AccessDatabase.open()` (read-only and
  shared, exclusive, password), used as context managers. Guaranteed cleanup, with cleanup failures
  attached as notes to the original exception. `db.raw` escape hatch with revocable proxies.
- **Engines.** In-process DAO, Access-hosted DAO, and Access design sessions. `engine="auto"` selects one
  and switches to a design session on the first design feature. Bitness-aware diagnostics.
- **Process safety.** Access is always started as a new instance and never attached to. Process identity is
  PID, creation time and image. Kill-on-close job objects, an ownership ledger with `reap_orphans()` /
  `pyaccesskit cleanup`, a dialog watchdog (`AccessDialogError`), per-call timeouts, and Ctrl+C handling.
- **Tables.** Short Text, Long Text, Number (all field sizes), Decimal, Currency, AutoNumber (incrementing
  or Replication ID), Date/Time, Yes/No, Hyperlink and OLE Object columns. Literal and `Expr` defaults,
  validation rules, captions, formats and input masks. Indexes (composite, descending, unique, primary).
  Add, drop and rename columns and indexes, rename and drop tables, property bags. `Table.to_spec()`
  round-trips to the normalized spec.
- **Relationships.** Single and composite keys, referential integrity, cascades and join types, with
  pre-validation of tables, columns, types and unique indexes.
- **Queries and data.** Saved and pass-through queries, query kinds and parameters, and
  `execute()`/`fetch_all()` with parameters bound by name.
- **Forms** (provisional). Labels, text boxes, check boxes, combo boxes and buttons. Stacked and tabular
  layouts, VBA event procedures, form properties, atomic build-then-swap replacement, `check_opens()` and
  `controls()`.
- **VBA modules** (provisional). Standard and class modules; read and replace code.
- **Text import/export** (provisional). `SaveAsText`/`LoadFromText` for forms, reports, macros, queries and
  modules, with the correct encoding per object type; files are UTF-8.
- **Specs.** Immutable, serializable Pydantic models (`TableSpec`, `Column.*`, `IndexSpec`,
  `RelationshipSpec`, `QuerySpec`, `FormSpec`) with JSON Schema. `Length` units (`cm`, `mm`, `inch`, `pt`,
  `twips`).
- **Errors.** A typed exception hierarchy with Access/DAO error numbers and `DBEngine.Errors` details.
- **CLI.** `pyaccesskit doctor [--json] [--probe]`, `pyaccesskit inspect DB [--json] [--counts]` and
  `pyaccesskit cleanup [--dry-run] [--json]`. `python -m pyaccesskit`.
- **Diagnostics.** `pyaccesskit.diagnose()`.
- **For AI agents.** `pyaccesskit guide` prints a version-matched guide (rules, API cheat sheet, SQL dialect
  notes, error table, limits, and a complete tested example); `pyaccesskit schema` prints JSON Schemas of the
  specs; the documentation site publishes `llms.txt` and `llms-full.txt`.
- **Documentation.** Recipes, error / data-type / CLI references, FAQ, and an API reference generated from
  the docstrings.

### Fixed (found while documenting)

- Opening a database for design no longer runs its `StartUpForm` (which could raise dialogs with code
  disabled); the setting is preserved.
- `bytes` query parameters are declared as `LongBinary`, so OLE Object data is no longer corrupted; binary
  values read back as `bytes`.

[Unreleased]: https://github.com/ariedotcodotnz/PyAccessKit/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/ariedotcodotnz/PyAccessKit/releases/tag/v0.1.0
