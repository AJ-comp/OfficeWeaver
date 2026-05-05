---
source_id: spreadsheet.examples.ModernTable
engine: onlyoffice
version: 9.3.1.2
version_family: 9.3
product: spreadsheet
kind: spreadsheet
object: OfficeWeaver Sheet
method: setFill setFont setBorder setAlignment setNumberFormat
title: Example: modern spreadsheet table styling
source_url: local:officeweaver/examples/spreadsheet-modern-table
keywords:
  - modern table
  - design refresh
  - report styling
  - zebra
  - dashboard style
---

# Example: Modern Spreadsheet Table Styling

Use this pattern when the user asks to make a spreadsheet table cleaner, more modern, more executive, or easier to read.

Prefer `Sheet.*` helpers so each formatting operation is traced.

```js
Sheet.setFill("A1:E1", "#1F4E79");
Sheet.setFont("A1:E1", { bold: true, color: "#FFFFFF", size: 12, name: "Arial" });
Sheet.setAlignment("A1:E9", { vertical: "center", wrap: true });
Sheet.setAlignment("A1:E1", { horizontal: "center", vertical: "center" });
Sheet.setAlignment("E2:E9", { horizontal: "right" });
Sheet.setNumberFormat("E2:E9", "#,##0");
Sheet.setBorder("A1:E9", { color: "#D9E2EC", style: "Thin" });

for (let row = 2; row <= 9; row++) {
  if (row % 2 === 0) {
    Sheet.setFill(`A${row}:E${row}`, "#F6F8FA");
  }
}

Sheet.setColumnWidth("A", 14);
Sheet.setColumnWidth("B", 18);
Sheet.setColumnWidth("C", 18);
Sheet.setColumnWidth("D", 12);
Sheet.setColumnWidth("E", 16);

return Sheet.done("Applied modern table styling");
```

## Notes

- Adjust the range from outline/range inspection results.
- If the first row is not the header, target the actual header row.
- Do not change user data unless the user explicitly asks for cleanup.
