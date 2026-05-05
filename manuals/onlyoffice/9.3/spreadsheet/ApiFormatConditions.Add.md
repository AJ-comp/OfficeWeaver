---
source_id: spreadsheet.ApiFormatConditions.Add
engine: onlyoffice
version: 9.3.1.2
version_family: 9.3
product: spreadsheet
kind: spreadsheet
object: ApiFormatConditions ApiFormatCondition
method: GetFormatConditions Add SetFillColor
keywords:
  - conditional formatting
  - format condition
  - highlight values
---

# Conditional Formatting With Sheet.raw

OfficeWeaver does not yet expose a first-class conditional-formatting helper. Use `Sheet.raw` so the operation is still traced.

```js
Sheet.raw("addConditionalFormatting", { range: "E2:E9" }, function (Api, Sheet) {
  const range = Sheet.range("E2:E9");
  const formatConditions = range.GetFormatConditions();
  const condition = formatConditions.Add("xlCellValue", "xlGreater", "2000000");
  if (condition) {
    condition.SetFillColor(Sheet.color("#FFE5E5"));
  }
  return { condition_created: !!condition };
});

return Sheet.done("Added conditional formatting");
```

Use this only for conditional formatting. For ordinary font, fill, alignment, border, and number formatting, use `Sheet.*` helpers instead.
