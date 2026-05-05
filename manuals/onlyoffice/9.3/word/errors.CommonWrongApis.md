---
engine: onlyoffice
version_family: "9.3"
kind: word
topic: common-errors
---

# Common Word Error Recovery

When a Word macro fails with an unknown method or `not a function` error, first check whether the requested operation has a `Doc.*` helper.

For helper-covered work, retry with helpers only:

```js
Doc.insertParagraph(0, "Title", { bold: true });
Doc.setParagraphText(0, "Updated title");
Doc.styleParagraph(0, { bold: true, color: "#1F4E79", align: "center" });
Doc.replace("old", "new", { matchCase: false });
Doc.addTable([["A", "B"], ["1", "2"]], {
  headerFill: "#1F4E79",
  borderColor: "#D9E2EC"
});
return Doc.done("Recovered with traced helpers");
```

Use `Doc.raw` only when no helper exists, and keep raw-only code out of normal helper examples.
