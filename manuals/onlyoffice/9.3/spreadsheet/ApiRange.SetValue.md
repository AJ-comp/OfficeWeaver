---
source_id: spreadsheet.ApiRange.SetValue
engine: onlyoffice
version: 9.3.1.2
version_family: 9.3
product: spreadsheet
kind: spreadsheet
object: ApiRange
method: SetValue
title: ApiRange.SetValue
source_url: https://api.onlyoffice.com/docs/office-api/usage-api/spreadsheet-api/ApiRange/Methods/SetValue/
keywords:
  - set value
  - write cell
  - cell value
  - text value
  - number value
---

# ApiRange.SetValue

Use `SetValue(value)` to write a value to a cell or range.

## Syntax

```js
range.SetValue(value);
```

## Prefer OfficeWeaver Helper

```js
Sheet.setValue("A1", "Sales Report");
Sheet.setValue("E2", 1200000);
return Sheet.done("Updated cell values");
```

## Raw ONLYOFFICE Example

```js
Sheet.raw("setRawValues", { cells: ["A1", "E2"] }, function (Api, Sheet) {
  const ws = Sheet.sheet();
  ws.GetRange("A1").SetValue("Sales Report");
  ws.GetRange("E2").SetValue(1200000);
  return { updated: 2 };
});

return Sheet.done("Updated cell values");
```

## Notes

For formulas, prefer `Sheet.setFormula(...)` so the runtime can normalize the leading equals sign and record an evaluated value when possible.
