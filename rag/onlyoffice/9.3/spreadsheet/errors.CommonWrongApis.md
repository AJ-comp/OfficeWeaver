---
source_id: spreadsheet.errors.CommonWrongApis
product: spreadsheet
object: ApiRange ApiWorksheet
method: GetValue SetBold AddChart GetFreezePanes FreezeRows
title: Common wrong API names in ONLYOFFICE spreadsheet macros
source_url: local:aide/errors/spreadsheet-common-wrong-apis
keywords:
  - error
  - not a function
  - undefined
  - GetValues
  - SetFontBold
  - Api.ChartType
  - SetFreezePanes
  - GetActiveView
  - ?¤ë¥˜
  - ?¤íŒ¨
---

# Common Wrong API Names

These mistakes usually come from mixing ExcelJS, Office.js, VBA, or other spreadsheet-library patterns with ONLYOFFICE Office JavaScript API.

## Value Reading

```js
// Wrong
range.GetValues();

// Correct
range.GetValue();
```

## Font Bold

```js
// Wrong
range.SetFontBold(true);

// Correct
range.SetBold(true);
```

## Background Fill

```js
// Wrong
range.SetBackgroundColor(Sheet.color("#1F4E79"));

// Correct
range.SetFillColor(Sheet.color("#1F4E79"));
```

## Alignment

```js
// Wrong
range.SetHorizontalAlignment("center");
range.SetVerticalAlignment("center");

// Correct
range.SetAlignHorizontal("center");
range.SetAlignVertical("center");
```

## Chart Creation

```js
// Wrong
const chart = Api.CreateChart(Api.ChartType.COLUMN);

// Correct
const chart = ws.AddChart("'Sheet1'!$A$1:$E$5", false, "bar", 2, 12 * 360000, 7 * 360000, 6, 0, 1, 0);
```

## Freeze Panes

```js
// Wrong
ws.SetFreezePanes(2);
ws.GetActiveView();

// Correct
Api.SetFreezePanesType("row");
ws.GetFreezePanes().FreezeRows(2);

// OfficeWeaver Sheet helper
Sheet.freezeRows(2);
```
