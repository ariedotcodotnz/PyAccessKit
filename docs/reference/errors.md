# Errors

Every exception PyAccessKit raises derives from `PyAccessKitError`. Each carries:

- `str(exc)`: what PyAccessKit was doing, and why it failed, in plain words.
- `exc.operation`: the operation, for example `"create table 'Customers'"`.
- `exc.details`: `ErrorDetails` when Access or DAO reported the error. It holds the error number
  (`details.number`), source, description, HRESULT and the `DBEngine.Errors` entries. The original
  `com_error` is chained as `__cause__`.

Several exceptions also subclass a built-in exception, so generic code keeps working:
`DatabaseNotFoundError` is a `FileNotFoundError`, `DatabaseExistsError` is a `FileExistsError`,
`ObjectNotFoundError` is a `LookupError`, and `SpecError` is a `ValueError`.

## Hierarchy

```text
PyAccessKitError
├── EnvironmentProblem
│   └── EngineUnavailableError              .diagnosis explains what is missing
│       ├── AccessNotInstalledError
│       ├── DaoNotAvailableError            e.g. Python/Office bitness mismatch
│       └── AccessRuntimeOnlyError          design features need full Access
├── SessionError
│   ├── SessionClosedError                  used after close(), or a revoked db.raw object
│   ├── WrongThreadError                    used from another thread
│   ├── CapabilityError                     feature impossible with this engine / read-only session
│   └── ReadOnlyError                       write in a readonly=True session
├── DatabaseError                           .path
│   ├── DatabaseNotFoundError
│   ├── DatabaseExistsError
│   ├── DatabaseLockedError
│   ├── InvalidPasswordError
│   └── UnrecognizedFormatError
├── ObjectError                             .kind, .name
│   ├── ObjectNotFoundError
│   ├── ObjectExistsError
│   └── ObjectInUseError
├── SpecError                               .problems: every validation problem
├── SchemaError
│   ├── RelationshipError
│   └── IntegrityViolationError
├── QueryError                              .sql
│   ├── SqlSyntaxError
│   └── MissingParameterError
├── AccessApplicationError
│   ├── AccessDialogError                   .dialogs: title, text, buttons, action taken
│   ├── AccessTimeoutError
│   └── AccessProcessDiedError
├── CleanupError                            .errors: everything that failed while closing
└── ComError                                an Access/DAO error PyAccessKit has no specific class for
```

Warnings: `AccessNameWarning` for a legal but troublesome name, and `AccessDialogWarning` when
`dialog_policy="warn"` dismisses a dialog.

## Common error numbers

| Number | Meaning | Raised as |
|---|---|---|
| 3010, 3012 | Object already exists | `ObjectExistsError` |
| 3265, 2102, 7874 | Item not found | `ObjectNotFoundError` |
| 3024, 3044 | Database file or path not found | `DatabaseNotFoundError` |
| 3204, 7865 | Database file already exists | `DatabaseExistsError` |
| 3045, 3356, 3734, 7866 | Database in use or locked | `DatabaseLockedError` |
| 3031 | Wrong password | `InvalidPasswordError` |
| 3343 | Not an Access database | `UnrecognizedFormatError` |
| 3075, 3129, 3131, 3134, 3141, 3144 | SQL syntax error | `SqlSyntaxError` |
| 3061 | Too few parameters (unknown name in SQL) | `MissingParameterError` |
| 3022 | Duplicate value in a unique index | `IntegrityViolationError` |
| 3200 | Record cannot be deleted: related records exist | `IntegrityViolationError` |
| 3201 | Related record required | `IntegrityViolationError` |
| 3366, 3368, 3609 | Invalid relationship (field count or types, no unique index) | `RelationshipError` |
| 3085 | Undefined function in expression (`Nz` or a VBA function in DAO SQL) | `ComError` |
| 7960 | VBA did not compile | `ComError` |

## Cleanup errors never hide your error

If an exception is already propagating when the session closes, cleanup problems are added to it as
notes (`exc.__notes__`), and the original exception is the one you see. Only a normal close raises
`CleanupError`.

```python
from pyaccesskit import AccessDatabase, IntegrityViolationError, ObjectExistsError, SpecError

try:
    with AccessDatabase.open("app.accdb") as db:
        db.execute("INSERT INTO Orders (CustomerID) VALUES (999)")
except IntegrityViolationError as exc:
    print("fix the data:", exc)
except (ObjectExistsError, SpecError) as exc:
    print("fix the code:", exc)
```
