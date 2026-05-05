---
source_id: spreadsheet.ApiWorksheet.RangeAndSheet
product: spreadsheet
object: Api ApiWorksheet
method: GetActiveSheet GetSheets GetRange
title: Spreadsheet sheet and range access
source_url: https://api.onlyoffice.com/docs/office-api/usage-api/spreadsheet-api/ApiWorksheet/
keywords:
  - active sheet
  - get range
  - worksheet
  - sheet
  - range
  - ?�트
  - 범위
---

# Spreadsheet Sheet And Range Access

In OfficeWeaver macros, prefer the helper functions for sheet and range access:

```js
const ws = Sheet.sheet();             // active sheet
const named = Sheet.sheet("Sheet1");  // exact sheet name
const range = Sheet.range("A1:E9");   // active sheet range
```

Raw ONLYOFFICE API equivalents:

```js
const ws = Api.GetActiveSheet();
const sheets = Api.GetSheets();
const range = ws.GetRange("A1:E9");
```

## Correct Example

```js
const ws = Sheet.sheet();
const used = ws.GetRange("A1:E9");
used.SetWrap(true);
return { summary: "Updated active sheet range", changed: true };
```

## Sheet Name In Chart Data Ranges

For `AddChart`, the range string must include a sheet name:

```js
const chartRange = "'Sheet1'!$A$1:$E$5";
```
