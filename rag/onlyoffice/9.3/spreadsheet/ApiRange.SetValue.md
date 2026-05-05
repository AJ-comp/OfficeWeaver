---
source_id: spreadsheet.ApiRange.SetValue
product: spreadsheet
object: ApiRange
method: SetValue
title: ApiRange.SetValue
source_url: https://api.onlyoffice.com/docs/office-api/usage-api/spreadsheet-api/ApiRange/Methods/SetValue/
keywords:
  - set value
  - write cell
  - ?€ ê°??…ë ¥
  - ê°??°ê¸°
---

# ApiRange.SetValue

Use `SetValue(value)` to write a value to a cell or range.

## Syntax

```js
range.SetValue(value);
```

## Correct Example

```js
const ws = Sheet.sheet();
ws.GetRange("A1").SetValue("ë§¤ì¶œ ë³´ê³ ??);
ws.GetRange("E2").SetValue(1200000);
return { summary: "Updated cell values", changed: true };
```

## Notes

For formulas, prefer the dedicated formula APIs if you need formula semantics. For plain text, numbers, labels, and generated helper cells, `SetValue` is the normal choice.
