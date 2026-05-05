---
source_id: spreadsheet.ApiFormatConditions.Add
product: spreadsheet
object: ApiFormatConditions ApiFormatCondition
method: GetFormatConditions Add SetFillColor
title: ApiFormatConditions.Add
source_url: https://api.onlyoffice.com/docs/office-api/usage-api/spreadsheet-api/ApiFormatConditions/Methods/Add/
keywords:
  - conditional formatting
  - format condition
  - highlight values
  - 조건부 ?�식
  - �?강조
---

# ApiFormatConditions.Add

Use `range.GetFormatConditions()` and `formatConditions.Add(...)` to create conditional formatting rules.

## Syntax

```js
const formatConditions = range.GetFormatConditions();
const condition = formatConditions.Add(Type, Operator, Formula1, Formula2);
```

## Correct Example

```js
const ws = Sheet.sheet();
const dataRange = ws.GetRange("E2:E9");
const formatConditions = dataRange.GetFormatConditions();
const condition = formatConditions.Add("xlCellValue", "xlGreater", "2000000");
if (condition) {
  condition.SetFillColor(Sheet.color("#FFE5E5"));
}
return { summary: "Added conditional formatting", changed: true };
```

## Notes

- `Formula1` can be a number-like string.
- Check the returned condition for null before styling it.
