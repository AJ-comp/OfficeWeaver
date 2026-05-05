---
engine: onlyoffice
version_family: "9.3"
kind: word
example: modern-report
---

# Example: modern report structure

Use this pattern when the user asks to make a Word document look more professional or report-like.

```js
Doc.clear();

Doc.addHeading("Sales Performance Report", 1, {
  color: "#1F4E79",
  align: "center",
  spacingAfter: 180
});

Doc.addParagraph("Prepared for management", {
  italic: true,
  color: "#666666",
  align: "center",
  spacingAfter: 240
});

Doc.addHeading("Executive Summary", 2, {
  color: "#2C3E50",
  spacingBefore: 120
});

Doc.addParagraph("Revenue increased while operating costs remained stable.", {
  fontFamily: "Aptos",
  size: 24,
  spacingAfter: 160
});

Doc.addTable([
  ["Metric", "Value"],
  ["Revenue", "120000"],
  ["Cost", "45000"],
  ["Profit", "75000"]
], {
  headerFill: "#1F4E79",
  borderColor: "#D9E2EC",
  bandRows: true
});

return Doc.done("Modern report layout created");
```

