---
source_id: spreadsheet.examples.FreezeTitleHeader
product: spreadsheet
object: ApiWorksheet ApiFreezePanes
method: GetFreezePanes FreezeRows
title: Freeze title and header rows
source_url: local:aide/examples/spreadsheet-freeze-title-header
keywords:
  - freeze title
  - freeze header
  - freeze rows
  - FreezeRows
  - GetFreezePanes
  - ?€ ê³ ì •
  - ?œëª© ê³ ì •
  - ?¤ë” ê³ ì •
---

# Freeze Title And Header Rows

When the user asks "?œëª©?´ë‘ ?¤ë”ë¥?ê³ ì •?´ì¤˜??, freeze the top two rows.

Use a JS macro:

```js
Sheet.freezeRows(2);
return { summary: "?œëª© ?‰ê³¼ ?¤ë” ?‰ì„ ê³ ì •?ˆìŠµ?ˆë‹¤.", changed: true };
```

Raw ONLYOFFICE API:

```js
const ws = Sheet.sheet();
Api.SetFreezePanesType("row");
ws.GetFreezePanes().FreezeRows(2);
return { summary: "?œëª© ?‰ê³¼ ?¤ë” ?‰ì„ ê³ ì •?ˆìŠµ?ˆë‹¤.", changed: true };
```

Wrong names:

```js
ws.SetFreezePanes(2);       // wrong
ws.GetActiveView();         // wrong
```
