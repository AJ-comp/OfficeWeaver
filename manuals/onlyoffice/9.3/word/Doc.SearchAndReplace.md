---
source_id: officeweaver.onlyoffice.9_3.word.DocSearchReplace
engine: onlyoffice
version_family: "9.3"
kind: word
namespace: Doc
method: search replace
---

# Doc Search And Replace

Use `Doc.search` when the user asks to find text or when the macro needs to verify whether a phrase exists.

```js
Doc.search("draft", false);
return Doc.done("Searched for draft text");
```

Use `Doc.replace` for simple document-wide replacement.

```js
Doc.replace("draft", "final", { matchCase: false });
return Doc.done("Replaced draft wording");
```

`matchCase` defaults to false-like behavior when omitted.

Do not use browser DOM APIs for an ONLYOFFICE document. Use `Doc.search` and `Doc.replace` for these operations.
