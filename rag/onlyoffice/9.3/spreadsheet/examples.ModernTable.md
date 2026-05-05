---
source_id: spreadsheet.examples.ModernTable
product: spreadsheet
object: ApiRange
method: SetFillColor SetFontColor SetBold SetBorders SetAlignHorizontal SetNumberFormat
title: Example: modern Excel table styling
source_url: local:aide/examples/spreadsheet-modern-table
keywords:
  - modern table
  - design refresh
  - report styling
  - zebra
  - ?¸ë ¨?˜ê²Œ
  - ?”ìž??  - ???¤í???---

# Example: Modern Excel Table Styling

Use this pattern when the user asks to make a spreadsheet table cleaner, more modern, more executive, or easier to read.

```js
const ws = Sheet.sheet();
const table = ws.GetRange("A1:E9");
const header = ws.GetRange("A1:E1");
const amount = ws.GetRange("E2:E9");

header.SetFillColor(Sheet.color("#1F4E79"));
header.SetFontColor(Sheet.color("#FFFFFF"));
header.SetBold(true);
header.SetAlignHorizontal("center");

table.SetAlignVertical("center");
table.SetWrap(true);
amount.SetAlignHorizontal("right");
amount.SetNumberFormat("#,##0");

["Top", "Bottom", "Left", "Right", "InsideHorizontal", "InsideVertical"].forEach(side => {
  table.SetBorders(side, "Thin", Sheet.color("#D9E2EC"));
});

for (let row = 2; row <= 9; row++) {
  if (row % 2 === 0) {
    ws.GetRange(`A${row}:E${row}`).SetFillColor(Sheet.color("#F6F8FA"));
  }
}

return { summary: "Applied modern table styling", changed: true };
```

## Notes

- Adjust the range from `excel_get_outline` and `excel_get_range` results.
- If the first row is not the header, target the real header row.
- Avoid changing user data unless the user explicitly asks for cleanup.
