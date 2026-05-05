---
source_id: officeweaver.onlyoffice.9_3.spreadsheet.SheetHelpers
engine: onlyoffice
version: 9.3.1.2
version_family: 9.3
product: spreadsheet
kind: spreadsheet
object: OfficeWeaver Sheet
method: setValue setFormula setFont setFill setBorder setAlignment insertRows freezeRows done raw
keywords:
  - OfficeWeaver
  - Sheet helpers
  - Sheet.setValue
  - Sheet.setFormula
  - Sheet.setFont
  - Sheet.setFill
  - Sheet.freezeRows
  - traced macro
  - outcomes
  - fail fast
---

# OfficeWeaver Sheet Helpers

Use `Sheet.*` helpers for common spreadsheet edits in `excel_run_js_macro`.
They call the editor API internally and return structured `outcomes` with step, action, ok/error, and formula `evaluated_value` where possible.

Prefer helpers over raw API for values, formulas, formatting, rows, columns, sheets, and freeze panes.

## Values And Formulas

```js
Sheet.setValue("A1", "Sales Report");
Sheet.setFormula("F1", "=SUM(E2:E10)");
return Sheet.done("Updated title and total formula");
```

## Formatting

```js
Sheet.setFill("A1:E1", "#1F4E79");
Sheet.setFont("A1:E1", { bold: true, color: "#FFFFFF", size: 12 });
Sheet.setAlignment("A1:E20", { horizontal: "center", vertical: "center", wrap: true });
Sheet.setNumberFormat("E2:E20", "#,##0");
Sheet.setBorder("A1:E20", { color: "#D9E2EC", style: "Thin" });
return Sheet.done("Applied table formatting");
```

## Rows, Columns, And Sizes

```js
Sheet.insertRows(1, 2);
Sheet.setRowHeight(1, 30);
Sheet.setColumnWidth("A", 16);
Sheet.deleteColumns("F", 1);
return Sheet.done("Adjusted report layout");
```

## Sheets

```js
Sheet.addSheet("Summary");
Sheet.setActiveSheet("Summary");
Sheet.setValue("A1", "Summary");
Sheet.setActiveSheet("Sheet1");
return Sheet.done("Created summary sheet");
```

## Freeze Panes

```js
Sheet.freezeRows(2);
Sheet.freezeColumns(1);
Sheet.freezeAt("C3");
return Sheet.done("Updated freeze panes");
```

## Raw Escape Hatch

Use `Sheet.raw` only when no helper exists. Chart creation is one example.

```js
Sheet.raw("customChart", { range: "A1:E10" }, function (Api, Sheet) {
  const ws = Sheet.sheet();
  return ws.AddChart("'Sheet1'!$A$1:$E$10", false, "bar", 2, 4320000, 2520000, 6, 0, 1, 0);
});
return Sheet.done("Added chart");
```
