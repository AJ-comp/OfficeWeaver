---
source_id: spreadsheet.ApiRange.SetFontAndFill
product: spreadsheet
object: ApiRange
method: SetBold SetFontColor SetFillColor SetFontSize SetFontName
title: ApiRange font and fill formatting
source_url: https://api.onlyoffice.com/docs/office-api/usage-api/spreadsheet-api/ApiRange/
keywords:
  - bold
  - font color
  - fill color
  - header style
  - 굵게
  - 배경??  - 글?�색
wrong_patterns:
  - SetFontBold
  - SetBackgroundColor
  - Interior.Color
---

# ApiRange Font And Fill Formatting

Use these ApiRange methods for common cell styling:

```js
range.SetBold(true);
range.SetItalic(true);
range.SetUnderline(true);
range.SetFontColor(Sheet.color("#FFFFFF"));
range.SetFillColor(Sheet.color("#1F4E79"));
range.SetFontSize(12);
range.SetFontName("Arial");
```

## Correct Example

```js
const ws = Sheet.sheet();
const header = ws.GetRange("A1:E1");
header.SetFillColor(Sheet.color("#1F4E79"));
header.SetFontColor(Sheet.color("#FFFFFF"));
header.SetBold(true);
header.SetFontSize(12);
return { summary: "Styled the header row", changed: true };
```

## Common Mistakes

Do not invent Excel/VBA/Office.js names:

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
