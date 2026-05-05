---
source_id: officeweaver.onlyoffice.9_3.spreadsheet.SheetFreezePanes
engine: onlyoffice
version: 9.3.1.2
version_family: 9.3
product: spreadsheet
kind: spreadsheet
object: OfficeWeaver Sheet
method: freezeRows freezeColumns freezeAt unfreeze
title: Sheet freeze panes
keywords:
  - Sheet.freezeRows
  - Sheet.freezeColumns
  - Sheet.freezeAt
  - Sheet.unfreeze
  - freeze panes
  - freeze rows
  - freeze columns
  - freeze header
  - fixed header
---

# Sheet Freeze Panes

Use these helpers when the user asks to keep title rows, header rows, or left columns visible while scrolling.

## Freeze Top Rows

```js
Sheet.freezeRows(2);
return Sheet.done("Frozen title and header rows");
```

## Freeze Left Columns

```js
Sheet.freezeColumns(1);
return Sheet.done("Frozen first column");
```

## Freeze At A Cell

`Sheet.freezeAt("C3")` freezes rows above `C3` and columns left of `C3`.

```js
Sheet.freezeAt("C3");
return Sheet.done("Frozen panes at C3");
```

## Unfreeze

```js
Sheet.unfreeze();
return Sheet.done("Removed frozen panes");
```

Do not use raw freeze pane APIs for these operations. Use the helpers so the operation is traced and retriable.
