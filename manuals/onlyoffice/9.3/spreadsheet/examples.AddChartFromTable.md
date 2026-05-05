---
source_id: spreadsheet.examples.AddChartFromTable
engine: onlyoffice
version: 9.3.1.2
version_family: 9.3
product: spreadsheet
kind: spreadsheet
object: ApiWorksheet ApiChart
method: AddChart SetTitle
title: Example: add chart from an existing table
source_url: local:officeweaver/examples/spreadsheet-add-chart
keywords:
  - add chart
  - graph from data
  - sales chart
  - bar chart
  - column chart
---

# Example: Add Chart From Existing Table

Use `ApiWorksheet.AddChart` for charts. The data range should include the sheet name.

Because chart support has more version-specific details than basic formatting, wrap chart creation in `Sheet.raw` so OfficeWeaver still records the step.

```js
Sheet.raw("addChart", { range: "A1:E5", type: "bar" }, function (Api, Sheet) {
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
    chart.SetTitle("Sales Status", 13);
  }

  return { chart_created: !!chart };
});

return Sheet.done("Added a chart from the current data");
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
