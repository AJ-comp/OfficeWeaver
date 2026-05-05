---
source_id: spreadsheet.examples.AddChartFromTable
engine: onlyoffice
version: 9.3.1.2
version_family: 9.3
product: spreadsheet
kind: spreadsheet
object: ApiWorksheet ApiChart
method: AddChart SetTitle
keywords:
  - add chart
  - graph from data
  - sales chart
  - bar chart
  - column chart
---

# Example: Add Chart From Existing Table

Charts do not have a first-class OfficeWeaver helper yet. Wrap chart creation in `Sheet.raw` so OfficeWeaver still records the step.

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

If a generated macro tries to use `Api.CreateChart` or `Api.ChartType.COLUMN`, replace that chart attempt with the `Sheet.raw("addChart", ...)` pattern above.
