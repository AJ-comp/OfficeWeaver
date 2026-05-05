---
engine: onlyoffice
version_family: "9.3"
kind: word
namespace: Doc
---

# OfficeWeaver Doc Helpers

Use the `Doc` namespace when writing an OfficeWeaver macro for an ONLYOFFICE text document.

Prefer `Doc.*` helpers because each helper is traced as one logical step. If a step fails, OfficeWeaver returns `applied`, `failed`, `stopped_at`, and the per-step `outcomes` array.

## Common Helpers

```js
Doc.clear();
Doc.setParagraphText(0, "Quarterly Sales Report");
Doc.addHeading("Quarterly Sales Report", 1, { color: "#1F4E79" });
Doc.addParagraph("Prepared for management", { italic: true, spacingAfter: 160 });
Doc.styleParagraph(0, { bold: true, size: 40, align: "center" });
Doc.styleRange(0, 24, { bold: true, color: "#FF6F3D" });
Doc.search("draft", false);
Doc.replace("draft", "final", { matchCase: false });
Doc.addList(["Revenue grew", "Costs remained stable"], { type: "bullet" });
Doc.addTable([["Metric", "Value"], ["Revenue", "120000"]], {
  headerFill: "#1F4E79",
  borderColor: "#D9E2EC",
  bandRows: true
});
return Doc.done("Document updated");
```

## Style Options

The most common paragraph and range style options are:

```js
{
  bold: true,
  italic: false,
  underline: false,
  color: "#1F4E79",
  size: 32,
  fontFamily: "Aptos",
  align: "center",
  spacingBefore: 120,
  spacingAfter: 160,
  keepNext: true,
  keepLines: true
}
```

ONLYOFFICE text document font sizes use half-points, so `30` means 15pt. For common headings, `Doc.addHeading` chooses sensible defaults.

## Raw Escape Hatch

Use `Doc.raw(action, details, fn)` only when no `Doc.*` helper exists. Normal Word editing examples should use helpers first.
