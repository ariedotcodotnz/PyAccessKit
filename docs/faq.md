# FAQ

**Does PyAccessKit need Microsoft Access?**
For tables, relationships, queries and data, you need either Microsoft Access or the Access Runtime with a
Python of the same bitness. For forms, VBA modules and text import/export, you need full Microsoft Access.
`pyaccesskit doctor` tells you what works on your machine.

**Will it close the Access window I'm working in?**
No. PyAccessKit always starts its own hidden Access instance and never attaches to one that is running. It
ends only processes it started, identified by PID, creation time and executable.

**What happens if my script crashes?**
The Access process PyAccessKit started is in a Windows job object that ends it together with Python.
For a new database built with `create()`, the half-built file is discarded. If something is ever left
over, `pyaccesskit cleanup` ends it.

**Why is my table named `Name`/`Date`/`Order` flagged?**
They are Access reserved words or built-in property names. They work only when every SQL statement and
expression writes them in brackets, and they clash with form properties. Use `CustomerName`, `OrderDate`,
`SortOrder`… The warning points at the line of your code that used the name.

**My query works in Access but fails from Python with "Undefined function".**
Queries that Python runs go through DAO outside the Access user interface. There, Access-application
functions such as `Nz()`, and your own VBA functions, do not exist. Use engine functions: `IIf(IsNull(x), 0,
x)` instead of `Nz(x, 0)`. Forms can still use `Nz` and VBA functions.

**Can I build reports, subforms or ribbons?**
Not yet: see the [roadmap](index.md#roadmap). You can import report and macro text exported from Access
(`db.objects.import_text`), or use `db.raw.access` for one-off automation.

**How do I change a column's type?**
Access can't change a column type in place through DAO without risking data loss. Create a new column (or
table), copy the data with an `UPDATE` or `INSERT INTO … SELECT`, then drop the old one. Automated migrations
are planned with `plan`/`apply`.

**Does it work with `.mdb` files?**
You can open, inspect and change them. New databases are always `.accdb`.

**Can I use it from several threads?**
One session per thread. COM objects belong to the thread that created them. For parallel work, use
separate processes, each with its own session and file.

**Can AI coding assistants use PyAccessKit?**
Yes, and the library is designed for it. See [Building with AI agents](agents/index.md).
