# Text import/export

Access can write most objects to a text file (`SaveAsText`) and read them back (`LoadFromText`).
PyAccessKit exposes this through `db.objects`, for forms, reports, macros, modules and queries.

```python
text = db.objects.export_text("form", "frmCustomers")  # str, LF line endings
db.objects.import_text("form", "frmCustomersCopy", text)

path = db.objects.save_text("module", "modMath", "src/modMath.bas")  # UTF-8 file
db.objects.load_text("module", "modMath", path, replace=True)

db.objects.names("macro")
db.objects.rename("report", "rptOld", "rptNew")
db.objects.delete("macro", "mcrUnused")
```

## Encodings

Access uses different encodings per object type, and gets confused if you mix them up:

| Kind | What Access reads and writes | What PyAccessKit gives you |
|---|---|---|
| Forms, reports, macros, queries | UTF-16LE with a byte-order mark | `str`, and UTF-8 files with LF line endings |
| Modules | The Windows ANSI code page | `str`, and UTF-8 files with LF line endings |

UTF-16 files are treated as binary by Git. PyAccessKit converts them for you, so exported objects diff
nicely. Class modules carry four `Attribute VB_...` header lines in their text form. Plain VB6-style `.cls`
files are accepted, and their extra header lines are dropped.

## When to use it

Text export is the **full-fidelity** way to move objects between databases, keep designer-made forms in
source control, or look at what Access really stores. It is not the main way to *build* objects: the text
format is undocumented, version-dependent and unvalidated, and `LoadFromText` overwrites without asking.
PyAccessKit therefore refuses to replace an existing object unless you pass `replace=True`.

Stripping volatile content (checksums, printer settings, GUIDs) for cleaner diffs is planned for the
source-control milestone.
