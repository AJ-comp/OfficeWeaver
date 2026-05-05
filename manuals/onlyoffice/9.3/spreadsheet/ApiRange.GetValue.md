---
source_id: spreadsheet.ApiRange.GetValue
engine: onlyoffice
version: 9.3.1.2
version_family: 9.3
product: spreadsheet
kind: spreadsheet
object: ApiRange
method: GetValue
title: ApiRange.GetValue
source_url: https://api.onlyoffice.com/docs/office-api/usage-api/spreadsheet-api/ApiRange/Methods/GetValue/
keywords:
  - value
  - values
  - read cells
  - read range
  - cell value
wrong_patterns:
  - GetValues
---

# ApiRange.GetValue

Use `GetValue()` to read a value from a cell or range.

## Syntax

```js
range.GetValue();
```

## Returns

For a single cell, `GetValue()` returns the cell value. For a multi-cell range, it may return a two-dimensional array depending on the editor version and range shape.

## Correct Example

```js
Sheet.raw("readRange", { range: "A1:E9" }, function (Api, Sheet) {
  const ws = Sheet.sheet();
  return ws.GetRange("A1:E9").GetValue();
});

return Sheet.done("Read range values");
```

## Common Mistake

```js
// Wrong
ws.GetRange("A1:E9").GetValues();

// Correct
ws.GetRange("A1:E9").GetValue();
```
