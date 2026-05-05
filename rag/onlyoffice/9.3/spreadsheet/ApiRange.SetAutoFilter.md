---
source_id: spreadsheet.ApiRange.SetAutoFilter
product: spreadsheet
object: ApiRange
method: SetAutoFilter
title: ApiRange.SetAutoFilter
source_url: https://api.onlyoffice.com/docs/office-api/usage-api/spreadsheet-api/ApiRange/Methods/SetAutoFilter/
keywords:
  - autofilter
  - filter
  - filtering
  - ?ÑÌÑ∞
  - ?êÎèô?ÑÌÑ∞
---

# ApiRange.SetAutoFilter

Use `SetAutoFilter` to add an AutoFilter to a range.

## Syntax

```js
range.SetAutoFilter(Field, Criteria1, Operator, Criteria2, VisibleDropDown);
```

Most arguments are optional. `Field` is 1-based from the left side of the selected range.

## Correct Example

```js
const ws = Sheet.sheet();
const table = ws.GetRange("A1:E9");
table.SetAutoFilter();
return { summary: "Enabled AutoFilter for the table", changed: true };
```

## Filter Values Example

```js
const ws = Sheet.sheet();
const range = ws.GetRange("A1:A5");
range.SetAutoFilter(1, ["value2", "value3"], "xlFilterValues");
return { summary: "Applied value filter", changed: true };
```
