---
source_id: spreadsheet.ApiWorksheet.AddChart
engine: onlyoffice
version: 9.3.1.2
version_family: 9.3
product: spreadsheet
kind: spreadsheet
object: ApiWorksheet
method: AddChart
title: ApiWorksheet.AddChart
source_url: https://api.onlyoffice.com/docs/office-api/usage-api/spreadsheet-api/ApiWorksheet/Methods/AddChart/
keywords:
  - chart
  - graph
  - column chart
  - bar chart
  - line chart
  - pie chart
  - sales chart
wrong_patterns:
  - Api.CreateChart
  - Api.ChartType.COLUMN
  - chart.AddSeries
---

# ApiWorksheet.AddChart

Use `worksheet.AddChart(...)` to create a chart from a spreadsheet range.

Do not use ExcelJS, Office.js, or VBA-style chart APIs. In ONLYOFFICE Spreadsheet API 9.3, common chart creation should start from the worksheet.

## Syntax

```js
worksheet.AddChart(
  sDataRange,
  bInRows,
  sType,
  nStyleIndex,
  nExtX,
  nExtY,
  nFromCol,
  nColOffset,
  nFromRow,
  nRowOffset
);
```

## Important Parameters

- `sDataRange`: range string including the sheet name, for example `"'Sheet1'!$A$1:$E$5"`.
- `bInRows`: `true` reads series from rows. `false` reads series from columns.
- `sType`: chart type string such as `"bar"`, `"bar3D"`, `"lineNormal"`, `"pie"`, or `"doughnut"`.
- `nStyleIndex`: chart style index, often `1` to `48`.
- `nExtX`, `nExtY`: chart size in EMU. A practical approximation is `centimeters * 360000`.
- `nFromCol`, `nFromRow`: zero-based anchor column and row.

## Correct Example

```js
Sheet.raw("addChart", { range: "A1:E5", type: "bar" }, function (Api, Sheet) {
  const ws = Sheet.sheet();
  const chart = ws.AddChart(
    "'Sheet1'!$A$1:$E$5",
    false,
    "bar",
    2,
    12 * 360000,
    7 * 360000,
    6,
    0,
    1,
    0
  );

  if (chart) {
    chart.SetTitle("Sales Status", 13);
  }

  return { chart_created: !!chart };
});

return Sheet.done("Added a sales chart");
```

## Common Mistakes

```js
// Wrong
const chart = Api.CreateChart(Api.ChartType.COLUMN);
chart.AddSeries(...);

// Correct
const chart = ws.AddChart("'Sheet1'!$A$1:$E$5", false, "bar", 2, 12 * 360000, 7 * 360000, 6, 0, 1, 0);
```
