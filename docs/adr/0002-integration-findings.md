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
