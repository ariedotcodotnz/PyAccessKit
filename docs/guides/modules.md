# VBA modules

```python
from pyaccesskit import ModuleKind

db.modules.create(
    "modMath",
    "Public Function Twice(ByVal x As Long) As Long\n    Twice = 2 * x\nEnd Function\n",
)
db.modules.create(
    "clsCounter",
    "Private mCount As Long\n\nPublic Sub Increment()\n    mCount = mCount + 1\nEnd Sub\n",
    kind=ModuleKind.CLASS,
)

module = db.modules["modMath"]
module.kind  # ModuleKind.STANDARD
print(module.code)  # the source, with LF line endings
module.code = module.code.replace("2 * x", "x + x")
module.rename("modArithmetic")
db.modules.drop("modArithmetic")
```

- `Option Compare Database` is added at the top if missing, as Access does. Everything else is stored
  exactly as given.
- `create(..., replace=True)` replaces an existing module. Without it, an existing name raises
  `ObjectExistsError`.
- VBA source is stored in the Windows ANSI code page (for example Windows-1252). Characters that cannot be
  represented there raise `SpecError` instead of turning into `?`.
- Code is not compiled on import. A syntax error surfaces when Access compiles the module, for example when
  a form that uses it opens. Compile checks are planned for 0.2.
- Form and report modules belong to their form or report. Use the form's `on_load`, `module_code` and
  control events instead.

Modules need Microsoft Access (not the Runtime). Listing module names works with any engine:
`db.modules.names()`.
