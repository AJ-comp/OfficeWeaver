---
source_id: spreadsheet.ApiWorksheet.AddChart
product: spreadsheet
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
  - 차트
  - 그래??wrong_patterns:
  - Api.CreateChart
  - Api.ChartType.COLUMN
  - chart.AddSeries
---

# ApiWorksheet.AddChart

Use `worksheet.AddChart(...)` to create a chart from a sheet range. In ONLYOFFICE spreadsheet macros, do not use `Api.CreateChart` or `Api.ChartType.COLUMN`.

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

## Parameters That Matter Most

- `sDataRange`: string range including sheet name, e.g. `"'Sheet1'!$A$1:$E$5"`.
- `bInRows`: `true` means series are read from rows. `false` means series are read from columns.
- `sType`: chart type string such as `"bar"`, `"bar3D"`, `"lineNormal"`, `"pie"`, `"doughnut"`.
- `nStyleIndex`: chart style index, commonly `1` to `48`.
- `nExtX`, `nExtY`: size in EMU. A practical conversion is `cm * 360000`.
- `nFromCol`, `nFromRow`: zero-based anchor column and row.

## Correct Column Chart Example

```js
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
  chart.SetTitle("매출 ?�황", 13);
}
return { summary: "Added a sales chart", changed: true };
```

## Common Mistakes

```js
// Wrong
const chart = Api.CreateChart(Api.ChartType.COLUMN);
chart.AddSeries(...);

// Correct
const chart = ws.AddChart("'Sheet1'!$A$1:$E$5", false, "bar", 2, 12 * 360000, 7 * 360000, 6, 0, 1, 0);
```
