---
source_id: officeweaver.onlyoffice.9_3.word.DocTables
engine: onlyoffice
version_family: "9.3"
kind: word
namespace: Doc
method: addTable
---

# Doc Tables

Use `Doc.addTable` for report summary tables, metadata tables, comparison tables, and simple extracted data tables.

```js
Doc.addTable([
  ["Metric", "Value"],
  ["Revenue", "120000"],
  ["Cost", "45000"],
  ["Profit", "75000"]
], {
  widthPercent: 100,
  headerFill: "#1F4E79",
  headerColor: "#FFFFFF",
  borderColor: "#D9E2EC",
  bandRows: true
});
return Doc.done("Added report summary table");
```

Options:

```js
{
  widthPercent: 100,
  headerFill: "#1F4E79",
  headerColor: "#FFFFFF",
  borderColor: "#D9E2EC",
  bandRows: true,
  header: true
}
```

The first row is treated as a header unless `header: false` is provided.
