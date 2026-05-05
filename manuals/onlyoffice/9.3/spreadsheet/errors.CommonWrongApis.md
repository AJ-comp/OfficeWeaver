---
source_id: spreadsheet.errors.CommonWrongApis
engine: onlyoffice
version: 9.3.1.2
version_family: 9.3
product: spreadsheet
kind: spreadsheet
object: OfficeWeaver Sheet
method: setValue setFont setFill setAlignment freezeRows raw
keywords:
  - error
  - not a function
  - undefined
  - wrong api
  - Sheet helpers
---

# Common Spreadsheet Error Recovery

When a spreadsheet macro fails with an unknown method or `not a function` error, first check whether the requested operation has a `Sheet.*` helper.

For helper-covered work, retry with helpers only:

```js
Sheet.setValue("A1", "Sales Report");
Sheet.setFont("A1:E1", { bold: true, color: "#FFFFFF" });
Sheet.setFill("A1:E1", "#1F4E79");
Sheet.setAlignment("A1:E1", { horizontal: "center", vertical: "center", wrap: true });
Sheet.freezeRows(2);
return Sheet.done("Recovered with traced helpers");
```

Use `Sheet.raw` only for operations that do not have a helper yet, such as charts, conditional formatting, AutoFilter, or FormatAsTable.

```js
Sheet.raw("addChart", { range: "A1:E5" }, function (Api, Sheet) {
  const ws = Sheet.sheet();
  return ws.AddChart("'Sheet1'!$A$1:$E$5", false, "bar", 2, 12 * 360000, 7 * 360000, 6, 0, 1, 0);
});
return Sheet.done("Added chart");
```
