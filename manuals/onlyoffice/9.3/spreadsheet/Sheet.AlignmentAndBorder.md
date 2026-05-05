---
source_id: officeweaver.onlyoffice.9_3.spreadsheet.SheetAlignmentAndBorder
engine: onlyoffice
version: 9.3.1.2
version_family: 9.3
product: spreadsheet
kind: spreadsheet
object: OfficeWeaver Sheet
method: setAlignment setBorder
title: Sheet alignment and borders
keywords:
  - Sheet.setAlignment
  - Sheet.setBorder
  - align
  - wrap
  - border
  - table border
---

# Sheet Alignment And Borders

Use `Sheet.setAlignment` for horizontal/vertical alignment and wrapping.

```js
Sheet.setAlignment("A1:E20", {
  horizontal: "center",
  vertical: "center",
  wrap: true
});
Sheet.setAlignment("E2:E20", { horizontal: "right" });
return Sheet.done("Aligned table content");
```

Use `Sheet.setBorder` for table borders.

```js
Sheet.setBorder("A1:E20", {
  color: "#D9E2EC",
  style: "Thin",
  sides: ["Top", "Bottom", "Left", "Right", "InsideHorizontal", "InsideVertical"]
});
return Sheet.done("Applied clean table borders");
```

If `sides` is omitted, OfficeWeaver applies the common outside and inside table borders.

Do not use raw alignment or border APIs for these operations. Use the helpers so each formatting step is visible in `outcomes`.
