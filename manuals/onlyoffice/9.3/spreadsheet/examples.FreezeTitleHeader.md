---
source_id: spreadsheet.examples.FreezeTitleHeader
engine: onlyoffice
version: 9.3.1.2
version_family: 9.3
product: spreadsheet
kind: spreadsheet
object: OfficeWeaver Sheet
method: freezeRows
keywords:
  - freeze title
  - freeze header
  - freeze rows
  - Sheet.freezeRows
  - fixed header
---

# Freeze Title And Header Rows

When the user asks to keep the title and header visible while scrolling, freeze the top two rows with the OfficeWeaver helper.

```js
Sheet.freezeRows(2);
return Sheet.done("Frozen title and header rows");
```

For a sheet with only a header row and no title row, freeze one row.

```js
Sheet.freezeRows(1);
return Sheet.done("Frozen header row");
```

Do not use raw freeze pane APIs for this task. `Sheet.freezeRows` is the traced helper.
