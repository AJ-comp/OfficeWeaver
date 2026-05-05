---
source_id: spreadsheet.examples.AddChartFromTable
product: spreadsheet
object: ApiWorksheet ApiChart
method: AddChart SetTitle
title: Example: add chart from an existing table
source_url: local:aide/examples/spreadsheet-add-chart
keywords:
  - add chart
  - graph from data
  - sales chart
  - 차트 추�?
  - 그래??추�?
---

# Example: Add Chart From Existing Table

Use `ApiWorksheet.AddChart` for charts. The data range must include the sheet name.

```js
const ws = Sheet.sheet("Sheet1");

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

return { summary: "Added a chart from the current data", changed: true };
```

## If The Model Starts With The Wrong API

Replace this:

```js
Api.CreateChart(Api.ChartType.COLUMN);
```

with:

```js
ws.AddChart("'Sheet1'!$A$1:$E$5", false, "bar", 2, 12 * 360000, 7 * 360000, 6, 0, 1, 0);
```
