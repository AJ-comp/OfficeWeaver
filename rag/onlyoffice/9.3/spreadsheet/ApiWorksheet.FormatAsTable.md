---
source_id: spreadsheet.ApiWorksheet.FormatAsTable
product: spreadsheet
object: ApiWorksheet
method: FormatAsTable
title: ApiWorksheet.FormatAsTable
source_url: https://api.onlyoffice.com/docs/office-api/usage-api/spreadsheet-api/ApiWorksheet/Methods/FormatAsTable/
keywords:
  - format as table
  - table
  - ??  - ?åÏù¥Î∏?---

# ApiWorksheet.FormatAsTable

Use `FormatAsTable(range)` to format a cell range as an Excel-like table. The first row of the selected range becomes the table header, so the range must contain at least two rows.

## Syntax

```js
worksheet.FormatAsTable("A1:E10");
```

## Correct Example

```js
const ws = Sheet.sheet();
ws.FormatAsTable("A1:E10");
return { summary: "Formatted the range as a table", changed: true };
```

## Notes

- Use this when the user asks for a real table-style range.
- For custom branding, manual `SetFillColor`, `SetBorders`, and `SetFontColor` calls may provide more control.
