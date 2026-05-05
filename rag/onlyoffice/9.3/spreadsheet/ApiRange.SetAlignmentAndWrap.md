---
source_id: spreadsheet.ApiRange.SetAlignmentAndWrap
product: spreadsheet
object: ApiRange
method: SetAlignHorizontal SetAlignVertical SetWrap
title: ApiRange alignment and wrap
source_url: https://api.onlyoffice.com/docs/office-api/usage-api/spreadsheet-api/ApiRange/
keywords:
  - align
  - alignment
  - wrap text
  - center
  - ?•ë ¬
  - ì¤„ë°”ê¿?wrong_patterns:
  - SetHorizontalAlignment
  - SetVerticalAlignment
  - SetWrapText
---

# ApiRange Alignment And Wrap

Use `SetAlignHorizontal`, `SetAlignVertical`, and `SetWrap` for cell text layout.

## Syntax

```js
range.SetAlignHorizontal("left" | "center" | "right" | "justify");
range.SetAlignVertical("top" | "center" | "bottom" | "distributed" | "justify");
range.SetWrap(true);
```

## Correct Example

```js
const ws = Sheet.sheet();
const table = ws.GetRange("A1:E9");
table.SetAlignVertical("center");
table.SetWrap(true);
ws.GetRange("A1:E1").SetAlignHorizontal("center");
ws.GetRange("E2:E9").SetAlignHorizontal("right");
return { summary: "Aligned table content", changed: true };
```

## Common Mistakes

```js
// Wrong
range.SetHorizontalAlignment("center");
range.SetVerticalAlignment("center");
range.SetWrapText(true);

// Correct
range.SetAlignHorizontal("center");
range.SetAlignVertical("center");
range.SetWrap(true);
```
