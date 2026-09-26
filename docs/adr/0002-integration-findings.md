# ADR 0002 — Findings from the first integration runs

- **Status:** accepted
- **Date:** 2026-09-26
- **Context:** the first end-to-end runs of the full stack (smoke script, contract suite on real Access,
  lifecycle and round-trip suites) surfaced behaviours the Phase-0 spikes had not exercised. Each one is now
  covered by a test.

| # | Finding | Decision |
|---|---|---|
| 1 | ``DoCmd.SetWarnings`` raises Access error 2046 ("isn't available now") when no database is open. | Called (best effort) only after a design session has a current database. |
| 2 | Late-bound ``Index.Fields`` of an *appended* index returns a **string** (``"+Col1;-Col2"``), not a collection. | ``_index_fields`` parses the string form (``+`` ascending, ``-`` descending). |
| 3 | pywin32's dynamic dispatch resolves ``Item`` as a *property* on some DAO collections and returns the wrong object. | All DAO collection access goes through the gateway's exact ``IDispatch.Invoke`` (``_item``). |
| 4 | pywin32 treats a naive ``datetime`` as local time and converts it to UTC when building a ``VT_DATE`` (a 13-hour shift was observed). | ``to_variant`` passes wall-clock values as UTC-tagged datetimes; ``normalize`` reads the wall-clock back. |
| 5 | DAO silently ignores ``DefaultValue`` on an auto-increment field, so *random* AutoNumbers (``GenUniqueID()``) cannot be created. | ``AutoNumberColumn(new_values=...)`` was removed rather than shipped half-working. |
| 6 | Releasing a COM proxy after Access has exited makes COM raise-and-handle ``RPC_E_DISCONNECTED``; ``faulthandler`` prints it as a "Windows fatal exception". | ``gc.collect()`` before quitting releases unreachable proxies while Access is alive; the test-suite disables pytest's faulthandler plugin. |
| 7 | ``WaitForSingleObject`` needs ``SYNCHRONIZE``; a query-only handle made live orphans look "already exited". | Identity checks open processes with ``SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION``. |
| 8 | Race: dismissing a dialog unblocks the COM call before the watchdog had recorded the dialog. | Dialogs are recorded *before* dismissal; ``end()`` waits for in-progress handling. |
| 9 | After the watchdog terminated Access (timeout), shutdown's COM steps could only fail. | Shutdown skips COM steps when the owned process is already gone; closing stays quiet. |
| 10 | ``AutomationSecurity`` (``macro_security``) disables VBA and ``AutoExec``, but ``OpenCurrentDatabase`` still opens the database's ``StartUpForm``; with code disabled the form can raise a dialog, so design sessions failed on apps that set one. | Before a design open the engine removes ``StartUpForm`` through DAO (which runs nothing) and restores it as soon as the database is open (or through DAO again if the open fails). |
| 11 | Access-application functions (``Nz``, VBA functions from modules) are unavailable to SQL run through DAO outside the Access UI (DAO error 3085 "Undefined function"). They do work in form control sources. | Documented for users and agents: SQL that Python runs uses engine functions only (``IIf``, ``IsNull``…). |
| 12 | A VBA module that does not compile is only reported when Access next compiles code, e.g. when a form expression uses it (Access error 7960 during the form build). | Documented; ``check_opens()`` on every form is the recommended verification step. A dedicated ``VbaCompileError`` is planned for 0.2. |
| 13 | Implicit QueryDef parameters (a bare ``[name]`` in SQL) are typed as text; binding ``bytes`` to one stores a corrupted, doubled value. ``Parameter.Type`` cannot be changed for implicit parameters. | Ad-hoc SQL declares bytes-valued parameters (``PARAMETERS [x] LongBinary;``) automatically; saved queries with a non-binary parameter reject bytes with ``SpecError`` explaining the declaration. OLE/binary values read back as ``bytes``. |
| 14 | A database created by in-process DAO has no ``Forms``/``Reports``/``Scripts``/``Modules`` containers until Access opens it; reading ``Containers("Forms")`` fails with error 3265. | Listing objects of a kind whose container is missing returns an empty list. |
