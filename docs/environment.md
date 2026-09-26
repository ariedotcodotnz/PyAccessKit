# Environment & troubleshooting

## Supported environments

| | Supported | Notes |
|---|---|---|
| OS | Windows 10 22H2, Windows 11 | Windows Server: in-process DAO only (Microsoft does not support unattended Office automation) |
| Python | CPython 3.11–3.14, 64- and 32-bit | Free-threaded builds are not supported |
| Access | Microsoft 365, 2024, 2021, 2019, 2016 (all version 16.0) | 2013/2010 may work but are untested |
| Access Runtime | In-process DAO (bitness-matched) and Access-hosted DAO | The Runtime cannot design forms or modules |
| File formats | Create `.accdb`; open `.accdb` and `.mdb`; inspect `.accde`/`.mde` | `.adp` is not supported |

The spec models import on any OS. Only opening databases needs Windows and Access or DAO.

## `pyaccesskit doctor`

```console
$ pyaccesskit doctor
PyAccessKit 0.1.0 environment report

Python     3.12.4 (64-bit)  C:\Python312\python.exe
Windows    Windows-11-10.0.26100-SP0
pywin32    311
Access     16.0.19127.20264 (32-bit, Click-to-Run: O365ProPlusRetail)
           C:\Program Files (x86)\Microsoft Office\Root\Office16\MSACCESS.EXE
ACE OLEDB  none for this bitness

Engines
  in-process DAO    unavailable  DAO (Access database engine) is installed for 32-bit programs only, ...
  Microsoft Access  available    Access 16.0.19127.20264, 32-bit
  engine='auto' selects: Microsoft Access
...
```

`doctor` changes nothing: it reads the registry and file headers and tries to load DAO. `--probe`
additionally builds a scratch database with each available engine, in an owned Access instance that is
closed afterwards. `--json` prints the same report for scripts. The exit code is 3 when nothing is usable.

## Common situations

**"DAO is installed for 32-bit programs only, but this Python is 64-bit".** This is normal with Microsoft
365, which is often 32-bit. PyAccessKit then reaches DAO through Access, which works but starts
`MSACCESS.EXE`. For in-process DAO, use a Python of Office's bitness, for example
`uv python install cpython-3.12-windows-x86`.

**`AccessRuntimeOnlyError`.** Only the Access Runtime is installed. Tables, relationships, queries and
data work, but forms, modules and text import/export need the full product.

**`DatabaseLockedError`.** Another process has the database open exclusively, or a stale `.laccdb` lock
file belongs to a crashed process. Close Access, or open read-only (`AccessDatabase.open(path,
readonly=True)` works while the database is open elsewhere).

**`AccessDialogError`.** Access showed a modal dialog in the middle of an operation. The error contains the
dialog's title, text and buttons. Common causes are VBA compile errors in a form's module, broken
references, and startup code of the database (it is disabled by default; see `macro_security`).

**`AccessTimeoutError`.** One operation exceeded `SessionOptions.call_timeout`. The owned Access process
was terminated so your program could continue.

**Leftover `MSACCESS.EXE` after a crash.** Normally Windows ends it with your Python process (job object).
If that was not possible, run `pyaccesskit cleanup --dry-run` to see what PyAccessKit started, and
`pyaccesskit cleanup` to end processes whose Python owner is gone. Access instances you started yourself
are never touched.

**Databases in OneDrive or other synced folders.** Sync clients can lock or copy the file while Access has
it open, which risks corruption. Build into a local folder and copy the finished file.

**Running unattended (CI, services).** Microsoft does not support Office automation from services or
server-side code, because Office may show UI at any time. PyAccessKit's watchdog makes this far more robust,
but run Access-based jobs in an interactive user session (for example a self-hosted runner logged in as a
user). In-process DAO has no such restriction.
