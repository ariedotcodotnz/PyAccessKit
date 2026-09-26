# ADR 0001 — Phase-0 spike findings

- **Status:** accepted
- **Date:** 2026-09-26
- **Environment:** Microsoft 365 Apps (Current Channel) Access 16.0.20326.20158 **32-bit** Click-to-Run,
  Windows 11, pywin32 312; CPython 3.14 x64 (Access transports) and CPython 3.12 x86 (in-process DAO).
- **Scripts:** `scripts/spikes/s01…s10*.py` (throwaway; reproducible).

Each finding lists the decision it drives. "✅" = assumption from the plan confirmed; "⚠️" = reality
differed and the design was adjusted.

## S1 — process ownership

- ✅ Every `CoCreateInstanceEx(CLSCTX_LOCAL_SERVER)` starts a **new** `MSACCESS.EXE` (~1.1 s).
- ✅ `app.hWndAccessApp()` works on a hidden instance with no database and resolves (via
  `GetWindowThreadProcessId`) to the right process; `QueryFullProcessImageNameW` confirms `MSACCESS.EXE`.
- ✅ A DCOM-launched Access **can** be assigned to our job object; closing the job handle kills it.
- ✅ Defaults: `Visible=False`, `UserControl=False`, `AutomationSecurity=2` (ByUI).
- ✅ `Quit(acQuitSaveNone)` → process exit in ~2.5 s; holding a child COM reference did not block exit.
- `SysCmd(724)` → `'32-bit'`, `SysCmd(720)` → full version string, `SysCmd(6)` → runtime flag. Used by
  `doctor --probe`.

## S2 — `AutomationSecurity` and startup code

| Open path | Opens | AutoExec runs | Dialogs | Design ops |
|---|---|---|---|---|
| Access-hosted DAO (`app.DBEngine.OpenDatabase`) | ✅ | never | 0 | n/a |
| `ByUI` (default) in an untrusted folder | ✅ | no (Trust Center dependent) | 0 | ✅ |
| `ForceDisable` | ✅ | **no** | 0 | ✅ |
| `Low` | ✅ | **yes** | 0 | ✅ |

⚠️ The documentation's warning that ForceDisable "will not open any database" does not apply to modern
Access. **Decision:** default `macro_security = "disable"` (ForceDisable): deterministic, independent of the
user's Trust Center. Code execution (`Application.Run`, 0.2) requires an explicit opt-in.

## S3 — DAO- vs Access-created databases

`NewCurrentDatabase` writes 15 database properties that DAO `CreateDatabase` does not; opening a
DAO-created file in Access only adds nav-pane properties. The user-visible ones include `UseMDIMode=0`
(tabbed documents), `ShowDocumentTabs=True`, `Themed Form Controls=1`.
**Decision:** the in-process DAO transport applies this **native defaults profile** (see
`_backends/dao/profile.py`) so output is identical whichever engine created the file. `dbVersion120` is the
creation format (`dbVersion150/167` still report `Version=12.0`).

## S4 — Decimal(p, s)

- ✅ DAO DDL rejects `DECIMAL` (3292); ADO DDL (`CurrentProject.Connection` or in-process ADODB + ACE
  OLEDB) supports ADD/ALTER/CREATE with precision and scale.
- ⚠️ DAO `CreateField(..., dbDecimal)` silently produces a **BigInt (Type 16)** field, even after setting
  `Precision`/`Scale`. **Never create decimals through DAO.**
- ✅ DAO *reads* precision/scale through the hidden `Field.Properties("Precision"/"Scale")` (only meaningful
  when `Type == 20`) — introspection needs no ADO.
- In-process ADO cannot open a file DAO holds exclusively → the in-process transport closes DAO, runs the
  ADO DDL, and reopens (S10).

## S5 — QueryDefs

- ⚠️ Jet **rewrites** SQL: uppercases keywords, one clause per line, appends `;\r\n` (DDL/UNION are kept
  closer to verbatim). **Decision:** store what Access stores; compare with a whitespace/case-insensitive
  normalizer.
- ✅ No reference validation at save time (missing tables and not-yet-created queries are accepted) →
  `build` needs no topological ordering of queries.
- ✅ Tables and queries share one namespace (3012 / 3010).
- ⚠️ `QueryDef.Type` reads `0` until `QueryDefs.Refresh()`; afterwards append=64, crosstab=16, DDL=96,
  delete=32, make-table=80, pass-through=112, union=128, update=48. **Always refresh before enumerating.**
- Parameters come back **bracketed** (`[pMin]`); a bracketed identifier matching a column is a column, not
  a parameter. Pass-through: `CreateQueryDef(name)` then `Connect` → `SQL` works with no reachable DSN.
- `Recordset.GetRows(n)` returns column-major tuples.

## S6 — forms in a hidden instance

- ⚠️ pywin32 dynamic dispatch cannot call **indexed properties** (`Form.Section(0)` → "Member not found")
  and raises `AttributeError` (not `com_error`) for members missing from a specific control's typeinfo
  (e.g. `Label.ControlSource`). **Decision:** the COM gateway performs its own `IDispatch.Invoke` by DISPID
  with the exact arguments for property reads, indexed properties, and method calls.
- ✅ `CreateForm` → properties → `CreateControl` (+ attached labels via `Parent`) → `HasModule` +
  `Module.AddFromString` → `DoCmd.Close(acSaveYes)` → `DoCmd.Rename` works; geometry round-trips exactly;
  the form opens in Form view (hidden) and binds data.
- ✅ Unsaved forms closed with `acSaveNo` leave nothing behind (free rollback).
- ✅ Header/footer: `DoCmd.RunCommand(acCmdFormHdrFtr)` right after `CreateForm` works hidden;
  `CreateControl` into a missing section raises 2148.
- ✅ Build-then-swap with a hidden `~pak_bak_…` backup name works.
- ⚠️ A modal **"Save As"** dialog appears in the *hidden* instance when `CloseCurrentDatabase` finds unsaved
  new objects. **Decision:** close every open form/report with `acSaveNo` before closing the database, and
  keep the dialog watchdog alive until the process has exited.

## S7 — SaveAsText / LoadFromText

- Forms, queries and macros export as **UTF-16LE with BOM**; they must be re-imported as UTF-16LE (UTF-8
  without BOM loses non-ASCII text).
- ⚠️ Modules are read **and** written in the **ANSI code page** (cp1252 here); UTF-8 is mangled and a UTF-8
  BOM becomes literal `ï»¿` code. **Decision:** module codec = ANSI; reject unrepresentable characters.
- ⚠️ Class modules: Access exports them with four leading `Attribute VB_GlobalNameSpace/Creatable/
  PredeclaredId/Exposed` lines (no `VERSION 1.0 CLASS` header); `LoadFromText` uses those lines to create a
  class module. A VB6-style header yields an empty *standard* module. **Decision:** the codec adds/strips
  exactly Access's header.
- `LoadFromText` silently overwrites existing objects; imported objects persist.
- VBE (`app.VBE.ActiveVBProject.VBComponents`) is accessible without Trust Center changes.

## S8 — dialogs

- ✅ pywin32 releases the GIL during `Invoke`; a watcher thread detects our PID's visible windows while the
  call blocks. `WM_CLOSE` dismisses OK boxes; **Yes/No boxes ignore `WM_CLOSE`** and need `BM_CLICK` on a
  button (preference Cancel → No → OK). Button captions carry mnemonics (`&No`).
- ⚠️ `app.Run("Proc")` through pywin32's typed wrapper fails (`DISP_E_BADPARAMCOUNT`); an exact-argument
  `Invoke` works (`'ok'`, `Add2(2,3) → 5`). Same root cause as S6.

## S9 — errors

- All Access/DAO errors arrive as `DISP_E_EXCEPTION` with `scode = 0x800A0000 | number`; DAO errors carry
  `source='DAO.<Collection>'`, Access errors `source=None` with a filled description.
- Verified numbers: 3265, 3010, 3012, 3191, 3270, 3609 (no unique index for relationship), 3201 (existing
  data violates integrity), 3368, 3125, 3141, 3129, 3075, 3061, 3292, 2102, 2148, 2149, 2220, 2462, 2467,
  2493, 3045, 3024, 3343, 3204, 7865, 7866, 7874.
- ⚠️ `DBEngine.Errors` keeps **stale** entries after Access-level errors → attach it only when its last
  entry matches the current error number.
- ⚠️ **`OpenCurrentDatabase` fails silently** (no exception, no dialog) for locked or non-database files:
  `CurrentProject.FullName == ''`, `CurrentDb() is None`. **Decision:** verify every open and, on failure,
  probe with DAO `OpenDatabase` to raise the precise error (locked / format / missing).

## S10 — in-process DAO (x86 Python, C2R Office)

- ✅ `DAO.DBEngine.120` and ADODB + `Microsoft.ACE.OLEDB.16.0` work in-process from 32-bit Python under
  Click-to-Run (engine ready in ~0.1 s).
- ADO cannot open a file DAO holds exclusively; with DAO closed, ADO decimal DDL works and DAO then reads
  `Precision`/`Scale` correctly.
