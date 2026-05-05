---
source_id: spreadsheet.ApiRange.SetAutoFilter
engine: onlyoffice
version: 9.3.1.2
version_family: 9.3
product: spreadsheet
kind: spreadsheet
object: ApiRange
method: SetAutoFilter
keywords:
  - autofilter
  - filter
  - filtering
---

# AutoFilter With Sheet.raw

OfficeWeaver does not yet expose a first-class AutoFilter helper. Use `Sheet.raw` so the operation is still traced.

## Enable AutoFilter

```js
Sheet.raw("enableAutoFilter", { range: "A1:E9" }, function (Api, Sheet) {
  const table = Sheet.range("A1:E9");
  table.SetAutoFilter();
  return { range: "A1:E9" };
});

return Sheet.done("Enabled AutoFilter for the table");
```

## Filter Values

```js
Sheet.raw("filterValues", { range: "A1:A5" }, function (Api, Sheet) {
  const range = Sheet.range("A1:A5");
  range.SetAutoFilter(1, ["value2", "value3"], "xlFilterValues");
  return { range: "A1:A5", values: ["value2", "value3"] };
});

return Sheet.done("Applied value filter");
```
