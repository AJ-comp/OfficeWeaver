---
source_id: spreadsheet.ApiRange.SetFontAndFill
engine: onlyoffice
version: 9.3.1.2
version_family: 9.3
product: spreadsheet
kind: spreadsheet
object: ApiRange
method: SetBold SetItalic SetUnderline SetFontColor SetFillColor SetFontSize SetFontName
title: ApiRange font and fill formatting
source_url: https://api.onlyoffice.com/docs/office-api/usage-api/spreadsheet-api/ApiRange/
keywords:
  - bold
  - font color
  - fill color
  - background color
  - header style
  - typography
wrong_patterns:
  - SetFontBold
  - SetBackgroundColor
  - Interior.Color
---

# ApiRange Font And Fill Formatting

Use these `ApiRange` methods for common cell styling:

```js
range.SetBold(true);
range.SetItalic(true);
range.SetUnderline(true);
range.SetFontColor(Sheet.color("#FFFFFF"));
range.SetFillColor(Sheet.color("#1F4E79"));
range.SetFontSize(12);
range.SetFontName("Arial");
```

## Prefer OfficeWeaver Helpers

For common formatting, prefer traced helpers:

```js
Sheet.setFill("A1:E1", "#1F4E79");
Sheet.setFont("A1:E1", { bold: true, color: "#FFFFFF", size: 12, name: "Arial" });
return Sheet.done("Styled the header row");
```

## Raw ONLYOFFICE Example

```js
Sheet.raw("styleHeader", { range: "A1:E1" }, function (Api, Sheet) {
  const ws = Sheet.sheet();
  const header = ws.GetRange("A1:E1");
  header.SetFillColor(Sheet.color("#1F4E79"));
  header.SetFontColor(Sheet.color("#FFFFFF"));
  header.SetBold(true);
  header.SetFontSize(12);
  return { styled: true };
});

return Sheet.done("Styled the header row");
```

## Common Mistakes

```js
// Wrong
header.SetFontBold(true);
header.SetBackgroundColor(Sheet.color("#1F4E79"));
header.Font.Color = "#FFFFFF";
header.Interior.Color = "#1F4E79";

// Correct
header.SetBold(true);
header.SetFillColor(Sheet.color("#1F4E79"));
header.SetFontColor(Sheet.color("#FFFFFF"));
```
