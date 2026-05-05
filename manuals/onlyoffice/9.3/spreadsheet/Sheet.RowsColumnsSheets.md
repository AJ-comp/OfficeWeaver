---
source_id: officeweaver.onlyoffice.9_3.spreadsheet.SheetRowsColumnsSheets
engine: onlyoffice
version: 9.3.1.2
version_family: 9.3
product: spreadsheet
kind: spreadsheet
object: OfficeWeaver Sheet
method: insertRows deleteRows insertColumns deleteColumns setColumnWidth setRowHeight addSheet renameSheet deleteSheet setActiveSheet
title: Sheet rows, columns, and sheets
keywords:
  - Sheet.insertRows
  - Sheet.deleteRows
  - Sheet.insertColumns
  - Sheet.deleteColumns
  - Sheet.setColumnWidth
  - Sheet.setRowHeight
  - Sheet.addSheet
  - Sheet.setActiveSheet
  - row
  - column
  - sheet
---

# Sheet Rows, Columns, And Sheets

Use `Sheet.*` helpers for layout changes.

```js
Sheet.insertRows(1, 2);
Sheet.setRowHeight(1, 30);
Sheet.setColumnWidth("A", 16);
Sheet.deleteColumns("F", 1);
return Sheet.done("Adjusted report layout");
```

Use sheet helpers for workbook sheet operations.

```js
Sheet.addSheet("Summary");
Sheet.setActiveSheet("Summary");
Sheet.setValue("A1", "Summary");
Sheet.setActiveSheet("Sheet1");
return Sheet.done("Created summary sheet");
```

Use `Sheet.sheet(name?)` and `Sheet.range(address, sheetName?)` only when you need an object for an unwrapped operation inside `Sheet.raw`.
