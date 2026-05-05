---
source_id: spreadsheet.ApiChart.SeriesAndStyle
product: spreadsheet
object: ApiChart
method: AddSeria SetSeriesFill SetTitle
title: ApiChart series and styling
source_url: https://api.onlyoffice.com/docs/office-api/usage-api/spreadsheet-api/ApiChart/
keywords:
  - chart series
  - add series
  - chart color
  - chart title
  - chart style
  - 차트 계열
  - 차트 ?�상
wrong_patterns:
  - AddSeries
  - chart.Series
  - chart.series.push
---

# ApiChart Series And Styling

ONLYOFFICE spreadsheet chart series APIs use `ApiChart` methods. For adding a series, the method name is `AddSeria`, not `AddSeries`.

## Common Methods

```js
chart.SetTitle("Sales Overview", 13);
chart.AddSeria("Cost price", "'Sheet1'!$B$4:$D$4");
chart.SetSeriesFill(Api.CreateSolidFill(Sheet.color("#1F4E79")), 0, false);
```

## Correct Example

```js
const ws = Sheet.sheet("Sheet1");
const chart = ws.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 12 * 360000, 7 * 360000, 6, 0, 1, 0);

if (chart) {
  chart.SetTitle("Sales Overview", 13);
  chart.AddSeria("Cost price", "'Sheet1'!$B$4:$D$4");
  chart.SetSeriesFill(Api.CreateSolidFill(Sheet.color("#1F4E79")), 0, false);
  chart.SetSeriesFill(Api.CreateSolidFill(Sheet.color("#F28C28")), 1, false);
}

return { summary: "Added and styled a chart", changed: true };
```

## Common Mistakes

```js
// Wrong
chart.AddSeries(...);
chart.series.push(...);

// Correct
chart.AddSeria("Cost price", "'Sheet1'!$B$4:$D$4");
```
