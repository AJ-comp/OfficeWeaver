---
source_id: spreadsheet.ApiRange.SetNumberFormat
product: spreadsheet
object: ApiRange
method: SetNumberFormat
title: ApiRange.SetNumberFormat
source_url: https://api.onlyoffice.com/docs/office-api/usage-api/spreadsheet-api/ApiRange/Methods/SetNumberFormat/
keywords:
  - number format
  - currency
  - percent
  - date format
  - ?«ì ?œì‹
  - ê¸ˆì•¡
  - ?µí™”
---

# ApiRange.SetNumberFormat

Use `SetNumberFormat(format)` to format numbers, currency, dates, percentages, fractions, or text.

## Syntax

```js
range.SetNumberFormat("#,##0");
```

## Useful Formats

```js
"#,##0"       // integer with thousands separator
"#,##0.00"    // decimal
"0.00%"       // percent
"yyyy-mm-dd"  // date-like display
"@"           // text
```

## Correct Example

```js
const ws = Sheet.sheet();
ws.GetRange("E2:E9").SetNumberFormat("#,##0");
ws.GetRange("B2:B9").SetNumberFormat("yyyy-mm-dd");
return { summary: "Applied number formats", changed: true };
```
