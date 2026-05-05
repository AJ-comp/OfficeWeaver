---
source_id: officeweaver.onlyoffice.9_3.spreadsheet.SheetValuesAndFormulas
engine: onlyoffice
version: 9.3.1.2
version_family: 9.3
product: spreadsheet
kind: spreadsheet
object: OfficeWeaver Sheet
method: setValue setFormula
title: Sheet values and formulas
keywords:
  - Sheet.setValue
  - Sheet.setFormula
  - value
  - formula
  - cell value
  - write cell
---

# Sheet Values And Formulas

Use `Sheet.setValue` and `Sheet.setFormula` for common cell writes. These helpers are traced and appear in `outcomes`.

```js
Sheet.setValue("A1", "Sales Report");
Sheet.setValue("E2", 1200000);
Sheet.setFormula("F2", "=SUM(E2:E10)");
return Sheet.done("Updated values and formulas");
```

`Sheet.setFormula` normalizes a missing leading equals sign and tries to record an `evaluated_value` in the outcome.

```js
Sheet.setFormula("F2", "SUM(E2:E10)");
return Sheet.done("Updated total formula");
```

Do not use raw range write APIs for these operations. Use the helpers so the host can see which write succeeded.
