---
source_id: spreadsheet.ApiChart.SeriesAndStyle
engine: onlyoffice
version: 9.3.1.2
version_family: 9.3
product: spreadsheet
kind: spreadsheet
object: ApiChart
method: AddSeria SetSeriesFill SetTitle
keywords:
  - chart series
  - add series
  - chart color
  - chart title
  - chart style
wrong_patterns:
  - AddSeries
  - chart.Series
  - chart.series.push
---

# Chart Series And Styling With Sheet.raw

OfficeWeaver does not yet expose a first-class chart styling helper. Use `Sheet.raw` for chart series and styling so the chart operation still appears in `outcomes`.

```js
Sheet.raw("addStyledChart", { range: "A1:D4", type: "bar3D" }, function (Api, Sheet) {
  const ws = Sheet.sheet("Sheet1");
  const chart = ws.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 12 * 360000, 7 * 360000, 6, 0, 1, 0);

  if (chart) {
    chart.SetTitle("Sales Overview", 13);
    chart.AddSeria("Cost price", "'Sheet1'!$B$4:$D$4");
    chart.SetSeriesFill(Api.CreateSolidFill(Sheet.color("#1F4E79")), 0, false);
    chart.SetSeriesFill(Api.CreateSolidFill(Sheet.color("#F28C28")), 1, false);
  }

  return { chart_created: !!chart };
});

return Sheet.done("Added and styled a chart");
```

Important note: the ONLYOFFICE method for adding an extra series is `AddSeria`, not `AddSeries`.
