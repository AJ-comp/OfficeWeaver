---
source_id: officeweaver.onlyoffice.9_3.word.DocTextStyle
engine: onlyoffice
version_family: "9.3"
kind: word
namespace: Doc
method: styleParagraph styleRange
---

# Doc Text Styling

Use `Doc.styleParagraph` for whole-paragraph styling.

```js
Doc.styleParagraph(0, {
  bold: true,
  color: "#1F4E79",
  size: 32,
  fontFamily: "Aptos",
  align: "center",
  spacingAfter: 160,
  keepNext: true
});
return Doc.done("Styled the first paragraph");
```

Use `Doc.styleRange` only when the host or a previous inspection step already knows the character offsets.

```js
Doc.styleRange(0, 24, {
  bold: true,
  color: "#FF6F3D",
  size: 28,
  fontFamily: "Aptos"
});
return Doc.done("Highlighted the selected text range");
```

Common style options:

```js
{
  bold: true,
  italic: false,
  underline: false,
  color: "#1F4E79",
  size: 32,
  fontFamily: "Aptos",
  align: "left",
  spacingBefore: 120,
  spacingAfter: 160,
  keepNext: true,
  keepLines: true
}
```

ONLYOFFICE text document font sizes use half-points, so `30` means 15pt.
