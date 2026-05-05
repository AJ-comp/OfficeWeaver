---
source_id: spreadsheet.errors.CommonWrongApis
engine: onlyoffice
version: 9.3.1.2
version_family: 9.3
product: spreadsheet
kind: spreadsheet
object: ApiRange ApiWorksheet OfficeWeaver
method: GetValue SetBold AddChart GetFreezePanes FreezeRows
title: Common wrong API names in ONLYOFFICE spreadsheet macros
source_url: local:officeweaver/errors/spreadsheet-common-wrong-apis
keywords:
  - error
  - not a function
  - undefined
  - GetValues
  - SetFontBold
  - Api.ChartType
  - SetFreezePanes
  - GetActiveView
---

# Common Wrong API Names

These mistakes usually come from mixing ExcelJS, Office.js, VBA, or other spreadsheet-library patterns with ONLYOFFICE Office JavaScript API.

Prefer `Sheet.*` helpers when they exist. Use raw ONLYOFFICE API only through `Sheet.raw` for operations that are not wrapped yet.

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

// Correct raw API
range.SetBold(true);

// Preferred OfficeWeaver helper
Sheet.setFont("A1:E1", { bold: true });
```

## Background Fill

```js
// Wrong
range.SetBackgroundColor(Sheet.color("#1F4E79"));

// Correct raw API
range.SetFillColor(Sheet.color("#1F4E79"));

// Preferred OfficeWeaver helper
Sheet.setFill("A1:E1", "#1F4E79");
```

## Alignment

```js
// Wrong
range.SetHorizontalAlignment("center");
range.SetVerticalAlignment("center");

// Correct raw API
range.SetAlignHorizontal("center");
range.SetAlignVertical("center");

// Preferred OfficeWeaver helper
Sheet.setAlignment("A1:E1", { horizontal: "center", vertical: "center" });
```

## Chart Creation

```js
// Wrong
const chart = Api.CreateChart(Api.ChartType.COLUMN);

// Correct raw API wrapped in Sheet.raw
Sheet.raw("addChart", { range: "A1:E5" }, function (Api, Sheet) {
  const ws = Sheet.sheet();
  return ws.AddChart("'Sheet1'!$A$1:$E$5", false, "bar", 2, 12 * 360000, 7 * 360000, 6, 0, 1, 0);
});
```

## Freeze Panes

```js
// Wrong
ws.SetFreezePanes(2);
ws.GetActiveView();

// Correct raw API
Api.SetFreezePanesType("row");
ws.GetFreezePanes().FreezeRows(2);

// Preferred OfficeWeaver helper
Sheet.freezeRows(2);
```
