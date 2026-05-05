---
source_id: spreadsheet.ApiFreezePanes
product: spreadsheet
object: Api ApiWorksheet ApiFreezePanes
method: SetFreezePanesType GetFreezePanes FreezeRows FreezeColumns FreezeAt Unfreeze
title: Freeze panes in ONLYOFFICE spreadsheet macros
source_url: https://api.onlyoffice.com/docs/office-api/usage-api/spreadsheet-api/ApiFreezePanes/
keywords:
  - freeze panes
  - freeze rows
  - freeze columns
  - freeze header
  - freeze title
  - GetFreezePanes
  - FreezeRows
  - FreezeColumns
  - FreezeAt
  - Unfreeze
  - SetFreezePanesType
  - ?� 고정
  - ?�목 고정
  - ?�더 고정
---

# Freeze Panes

ONLYOFFICE spreadsheet macros freeze panes through `ApiWorksheet.GetFreezePanes()`.
Do not call non-existent worksheet methods such as `ws.SetFreezePanes()` or `ws.GetActiveView()`.

## Freeze top rows

Use this when the user asks to keep a title row, header row, or top rows visible while scrolling.

```js
const ws = Sheet.sheet();
Api.SetFreezePanesType("row");
ws.GetFreezePanes().FreezeRows(2);
return { summary: "Frozen title and header rows", changed: true };
```

In OfficeWeaver macros, the helper is preferred:

```js
Sheet.freezeRows(2);
return { summary: "Frozen title and header rows", changed: true };
```

## Freeze left columns

```js
const ws = Sheet.sheet();
Api.SetFreezePanesType("column");
ws.GetFreezePanes().FreezeColumns(1);
return { summary: "Frozen first column", changed: true };
```

Helper:

```js
Sheet.freezeColumns(1);
return { summary: "Frozen first column", changed: true };
```

## Freeze rows and columns at a cell

To freeze rows above and columns left of a cell, call `FreezeAt` with an `ApiRange`.
For example, `C3` freezes rows 1-2 and columns A-B.

```js
const ws = Sheet.sheet();
Api.SetFreezePanesType("cell");
ws.GetFreezePanes().FreezeAt(ws.GetRange("C3"));
return { summary: "Frozen panes at C3", changed: true };
```

Helper:

```js
Sheet.freezeAt("C3");
return { summary: "Frozen panes at C3", changed: true };
```

## Unfreeze

```js
const ws = Sheet.sheet();
ws.GetFreezePanes().Unfreeze();
Api.SetFreezePanesType(null);
return { summary: "Removed frozen panes", changed: true };
```

Helper:

```js
Sheet.unfreeze();
return { summary: "Removed frozen panes", changed: true };
```
