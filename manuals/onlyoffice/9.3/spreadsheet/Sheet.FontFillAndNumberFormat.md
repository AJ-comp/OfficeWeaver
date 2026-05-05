---
source_id: officeweaver.onlyoffice.9_3.spreadsheet.SheetFontFillAndNumberFormat
engine: onlyoffice
version: 9.3.1.2
version_family: 9.3
product: spreadsheet
kind: spreadsheet
object: OfficeWeaver Sheet
method: setFont setFill setNumberFormat
title: Sheet font, fill, and number formatting
keywords:
  - Sheet.setFont
  - Sheet.setFill
  - Sheet.setNumberFormat
  - font
  - fill
  - number format
  - currency
  - typography
---

# Sheet Font, Fill, And Number Formatting

Use `Sheet.setFont`, `Sheet.setFill`, and `Sheet.setNumberFormat` for common spreadsheet styling.

```js
Sheet.setFill("A1:E1", "#1F4E79");
Sheet.setFont("A1:E1", {
  bold: true,
  color: "#FFFFFF",
  size: 12,
  name: "Arial"
});
Sheet.setNumberFormat("E2:E20", "#,##0");
return Sheet.done("Styled header and amount column");
```

Useful number formats:

```js
Sheet.setNumberFormat("E2:E20", "#,##0");
Sheet.setNumberFormat("E2:E20", "#,##0.00");
Sheet.setNumberFormat("F2:F20", "0.00%");
Sheet.setNumberFormat("B2:B20", "yyyy-mm-dd");
Sheet.setNumberFormat("A2:A20", "@");
```

Do not use raw font or fill APIs for these operations. Use the helpers so style operations remain traceable.
