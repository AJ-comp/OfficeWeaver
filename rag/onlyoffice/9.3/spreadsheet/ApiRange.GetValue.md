---
source_id: spreadsheet.ApiRange.GetValue
product: spreadsheet
object: ApiRange
method: GetValue
title: ApiRange.GetValue
source_url: https://api.onlyoffice.com/docs/office-api/usage-api/spreadsheet-api/ApiRange/Methods/GetValue/
keywords:
  - value
  - values
  - read cells
  - ?€ ê°?  - ê°??½ê¸°
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

For a single cell it returns a string-like scalar. For a multi-cell range it can return a two-dimensional array.

## Correct Example

```js
const ws = Sheet.sheet();
const values = ws.GetRange("A1:E9").GetValue();
return { summary: "Read range values", values };
```

## Common Mistakes

Do not use ExcelJS-style `GetValues()`.

```js
// Wrong
ws.GetRange("A1:E9").GetValues();

// Correct
ws.GetRange("A1:E9").GetValue();
```
