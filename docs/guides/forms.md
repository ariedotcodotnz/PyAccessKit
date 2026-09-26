# Forms

!!! note "Provisional"
    The form API is small on purpose and may grow in 0.x releases. Building forms needs Microsoft Access
    (the full product, not the Runtime).

## Building a form

```python
from pyaccesskit import FormView, Vba, cm

with db.forms.create("frmCustomers", record_source="Customers", caption="Customers") as form:
    form.textbox("CustomerName", label="Customer name", width=cm(8))
    form.textbox("Email", width=cm(8))
    form.combobox(
        "CountryID",
        row_source="SELECT CountryID, CountryName FROM Countries ORDER BY CountryName",
        column_count=2,
        column_widths=[cm(0), cm(4)],  # hide the key column, as in Access
    )
    form.checkbox("IsActive")
    form.textbox(
        name="txtCreated", control_source="=Format([CreatedAt], 'Short Date')", locked=True
    )
    form.button("cmdClose", caption="Close", on_click=Vba("DoCmd.Close acForm, Me.Name"))
```

The form is saved when the `with` block ends normally. If the block raises, nothing is saved. Outside a
`with` block, call `form.save()` or `form.discard()`.

### Controls

| Method | Control | Highlights |
|---|---|---|
| `textbox(field, label=, control_source=, format=, enabled=, locked=, after_update=)` | Text box | Bound to `field`, or computed with `control_source="=..."` |
| `checkbox(field, label=, ...)` | Check box | |
| `combobox(field, row_source=, row_source_type=, bound_column=, column_count=, column_widths=, limit_to_list=)` | Combo box | Rows from a table, query, SQL or value list |
| `label(caption)` | Label | |
| `button(name, caption=, on_click=)` | Command button | |

Every control also accepts `name=`, `section=`, `at=(left, top)`, `width=`, `height=`, `visible=` and
`properties={...}`. Text boxes, check boxes and combo boxes get an attached label automatically. Its text
is the field name (or the control name for unbound controls), unless you pass `label="..."` or
`label=False`.

### Layout

Controls are placed for you, and lengths use [units](../concepts/specs.md#units):

- **Stacked** (default for single forms): a label on the left and the control on the right, one row per
  control.
- **Tabular** (default for continuous forms): labels in the form header, and controls side by side in one
  detail row.
- **Explicit**: pass `at=(cm(1), cm(2))` and a size to place a control yourself.

```python
with db.forms.create(
    "frmProducts",
    record_source="Products",
    default_view=FormView.CONTINUOUS,
) as form:
    form.textbox("ProductName", width=cm(6))
    form.textbox("UnitPrice", format="Currency", width=cm(3))
    form.checkbox("Discontinued")
```

Layouts are checked against Access's limits (22 inches per form width and section height) before
anything is built.

### Events and VBA

```python
with db.forms.create("frmOrders", record_source="Orders") as form:
    form.on_load(Vba('Me.Caption = "Orders (" & DCount("*", "Orders") & ")"'))
    form.textbox("Quantity", after_update=Vba("Me.Recalc"))
    form.button("cmdSave", caption="Save", on_click=Vba("DoCmd.RunCommand acCmdSaveRecord"))
    form.module_code("Private Function Helper() As Long\n    Helper = 1\nEnd Function\n")
```

Event procedures (`Form_Load`, `cmdSave_Click`…) are generated in the form's module, which starts with
`Option Compare Database` and `Option Explicit`.

### Form properties

`db.forms.create(name, **options)` accepts: `record_source`, `caption`, `default_view` (single, continuous,
datasheet, split), `layout`, `header`, `width`, `allow_additions`, `allow_edits`, `allow_deletions`,
`data_entry`, `navigation_buttons`, `record_selectors`, `dividing_lines`, `scroll_bars`, `auto_center`,
`pop_up`, `modal`, and `properties={...}` for anything else (set verbatim, like
`properties={"RecordsetType": 2}`).

## Replacing a form safely

`db.forms.create(name, replace=True)` rebuilds an existing form:

1. The new form is built under a temporary name, then saved.
2. The old form is renamed to a backup name, and the new one takes its place.
3. The backup is deleted. If anything fails, the original is restored.

Each build starts from a fresh form, so the 754 *lifetime* controls limit never accumulates.

## Forms as data

A form is described by a `FormSpec`, which can be stored as JSON and built later:

```python
spec = form_builder.to_spec()
text = spec.model_dump_json(indent=2)

from pyaccesskit import FormSpec

db.forms.build(FormSpec.model_validate_json(text), replace=True)
```

## Checking and reading forms

```python
form = db.forms["frmCustomers"]
form.check_opens()  # opens hidden in Form view and closes; raises if Access reports an error
for control in form.controls():
    print(control.name, control.kind, control.control_source, control.left, control.width)
text = form.export_text()  # SaveAsText, as UTF-8 text
form.rename("frmClients")
db.forms.drop("frmClients")
```

Turning arbitrary designer-made forms back into a `FormSpec` is not supported, because that would lose
information. Use [text export](text-io.md) for full fidelity.
