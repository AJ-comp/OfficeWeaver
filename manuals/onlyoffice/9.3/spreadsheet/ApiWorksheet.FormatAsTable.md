---
source_id: spreadsheet.ApiWorksheet.FormatAsTable
engine: onlyoffice
version: 9.3.1.2
version_family: 9.3
product: spreadsheet
kind: spreadsheet
object: ApiWorksheet
method: FormatAsTable
title: ApiWorksheet.FormatAsTable
source_url: https://api.onlyoffice.com/docs/office-api/usage-api/spreadsheet-api/ApiWorksheet/Methods/FormatAsTable/
keywords:
  - format as table
  - table
  - table style
  - structured table
---

# ApiWorksheet.FormatAsTable

Use `FormatAsTable(range)` to format a cell range as an Excel-like table. The first row of the selected range becomes the table header, so the range should contain at least two rows.

## Syntax

```js
worksheet.FormatAsTable("A1:E10");
```

## Example

```js
Sheet.raw("formatAsTable", { range: "A1:E10" }, function (Api, Sheet) {
  const ws = Sheet.sheet();
  ws.FormatAsTable("A1:E10");
  return { range: "A1:E10" };
});

return Sheet.done("Formatted the range as a table");
```

## Notes

- Use this when the user asks for a real table-style range.
- For custom branding, manual `Sheet.setFill`, `Sheet.setBorder`, and `Sheet.setFont` calls may provide more control and better trace output.
