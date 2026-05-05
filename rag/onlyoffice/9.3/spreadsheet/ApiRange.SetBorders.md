---
source_id: spreadsheet.ApiRange.SetBorders
product: spreadsheet
object: ApiRange
method: SetBorders
title: ApiRange.SetBorders
source_url: https://api.onlyoffice.com/docs/office-api/usage-api/spreadsheet-api/ApiRange/Methods/SetBorders/
keywords:
  - border
  - borders
  - table border
  - ?Œë‘ë¦?wrong_patterns:
  - SetBorder
  - Borders.Color
---

# ApiRange.SetBorders

Use `SetBorders(bordersIndex, lineStyle, color)` to set a border on one side or inside edges of a range.

## Syntax

```js
range.SetBorders("Bottom", "Thin", Sheet.color("#D9E2EC"));
```

Common border indexes used in table styling:

```js
"Top"
"Bottom"
"Left"
"Right"
"InsideHorizontal"
"InsideVertical"
```

Common line styles:

```js
"Thin"
"Medium"
"Thick"
"Double"
"Dotted"
"Dashed"
```

## Correct Example

```js
const ws = Sheet.sheet();
const table = ws.GetRange("A1:E9");
["Top", "Bottom", "Left", "Right", "InsideHorizontal", "InsideVertical"].forEach(side => {
  table.SetBorders(side, "Thin", Sheet.color("#D9E2EC"));
});
return { summary: "Applied clean table borders", changed: true };
```

## Common Mistakes

```js
// Wrong
table.SetBorder("all", "#D9E2EC");
table.Borders.Color = "#D9E2EC";

// Correct
table.SetBorders("InsideHorizontal", "Thin", Sheet.color("#D9E2EC"));
```
