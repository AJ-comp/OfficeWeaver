---
source_id: officeweaver.onlyoffice.9_3.word.DocParagraphs
engine: onlyoffice
version_family: "9.3"
kind: word
namespace: Doc
method: setParagraphText addParagraph insertParagraph addHeading styleParagraph
---

# Doc Paragraphs And Headings

Use `Doc.*` helpers for paragraph creation, paragraph replacement, headings, and paragraph-level styling.

```js
Doc.setParagraphText(0, "Quarterly Sales Report");
Doc.styleParagraph(0, {
  bold: true,
  color: "#1F4E79",
  size: 40,
  align: "center",
  spacingAfter: 180
});
return Doc.done("Updated document title");
```

Add new paragraphs at the end:

```js
Doc.addParagraph("Prepared for management", {
  italic: true,
  color: "#666666",
  spacingAfter: 160
});
return Doc.done("Added report metadata");
```

Insert a paragraph at a specific document element index:

```js
Doc.insertParagraph(0, "Confidential", {
  bold: true,
  color: "#B55A00",
  align: "right"
});
return Doc.done("Inserted notice at the top");
```

For headings, prefer `Doc.addHeading` because it applies sensible heading defaults.

```js
Doc.addHeading("Executive Summary", 2, { color: "#2C3E50" });
return Doc.done("Added section heading");
```
