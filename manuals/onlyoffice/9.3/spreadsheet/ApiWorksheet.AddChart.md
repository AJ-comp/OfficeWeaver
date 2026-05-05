---
source_id: spreadsheet.ApiWorksheet.AddChart
engine: onlyoffice
version: 9.3.1.2
version_family: 9.3
product: spreadsheet
kind: spreadsheet
object: ApiWorksheet
method: AddChart
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

# Charts With Sheet.raw

OfficeWeaver does not yet expose a first-class chart helper. Use `Sheet.raw` for chart creation so the operation still appears as a traced step.

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

Important parameters inside the raw call:

- Data range must include the sheet name, for example `"'Sheet1'!$A$1:$E$5"`.
- `false` reads series from columns. `true` reads series from rows.
- Chart type strings include `"bar"`, `"bar3D"`, `"lineNormal"`, `"pie"`, and `"doughnut"`.
- Chart size uses EMU. A practical approximation is `centimeters * 360000`.

Do not use `Api.CreateChart`, `Api.ChartType.COLUMN`, or `chart.AddSeries` for ONLYOFFICE spreadsheet charts.
