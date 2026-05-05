---
source_id: officeweaver.onlyoffice.9_3.word.DocLists
engine: onlyoffice
version_family: "9.3"
kind: word
namespace: Doc
method: addList
---

# Doc Lists

Use `Doc.addList` for bulleted and numbered lists.

```js
Doc.addList(["Revenue grew", "Costs remained stable"], {
  type: "bullet",
  spacingAfter: 80
});
return Doc.done("Added summary bullet list");
```

For numbered lists:

```js
Doc.addList(["Collect data", "Review results", "Publish report"], {
  type: "numbered"
});
return Doc.done("Added numbered workflow list");
```

Use `level` only when you need nested list levels.

```js
Doc.addList(["Sub item A", "Sub item B"], {
  type: "bullet",
  level: 1
});
return Doc.done("Added nested list items");
```
