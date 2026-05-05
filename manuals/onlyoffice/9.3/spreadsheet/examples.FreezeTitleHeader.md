---
source_id: spreadsheet.examples.FreezeTitleHeader
engine: onlyoffice
version: 9.3.1.2
version_family: 9.3
product: spreadsheet
kind: spreadsheet
object: ApiWorksheet ApiFreezePanes
method: GetFreezePanes FreezeRows
title: Freeze title and header rows
source_url: local:officeweaver/examples/spreadsheet-freeze-title-header
keywords:
  - freeze title
  - freeze header
  - freeze rows
  - FreezeRows
  - GetFreezePanes
  - fixed header
---

# Freeze Title And Header Rows

When the user asks to keep the title and header visible while scrolling, freeze the top two rows.

## Prefer OfficeWeaver Helper

```js
Sheet.freezeRows(2);
return Sheet.done("Frozen title and header rows");
```

## Raw ONLYOFFICE API

```js
Sheet.raw("freezeRows", { rows: 2 }, function (Api, Sheet) {
  const ws = Sheet.sheet();
  Api.SetFreezePanesType("row");
  ws.GetFreezePanes().FreezeRows(2);
  return { rows: 2 };
});

return Sheet.done("Frozen title and header rows");
```

## Wrong Names

```js
ws.SetFreezePanes(2);       // wrong
ws.GetActiveView();         // wrong
```
