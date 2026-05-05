# OfficeWeaver

OfficeWeaver is an AI-friendly automation runtime for office editors.

It is designed for systems where an LLM writes JavaScript to edit a document opened in an editor such as ONLYOFFICE. Raw editor APIs are powerful, but they are not friendly to agents:

- an LLM often invents method names from ExcelJS, Office.js, VBA, or other libraries;
- raw API calls do not provide per-step success/failure logs;
- when a script fails halfway through, it is hard to know what was already applied;
- formula edits often need a computed value echo so the agent can detect errors;
- each host app ends up rebuilding the same tracing and retry machinery.

OfficeWeaver adds a small runtime layer between generated code and the editor API.

```js
Sheet.setValue("A1", "Sales Report");
Sheet.setFormula("F1", "=SUM(E2:E10)");
Sheet.setFont("A1:E1", { bold: true, color: "#FFFFFF" });
Sheet.freezeRows(2);
return Sheet.done("Report header updated");
```

The runtime executes those helpers through the underlying editor API and returns structured outcomes.

```json
{
  "ok": true,
  "engine": "onlyoffice",
  "version": "9.3.1.2",
  "version_family": "9.3",
  "applied": 4,
  "failed": 0,
  "outcomes": [
    { "step": 1, "action": "setValue", "address": "A1", "ok": true },
    { "step": 2, "action": "setFormula", "address": "F1", "ok": true, "evaluated_value": 123000 },
    { "step": 3, "action": "setFont", "address": "A1:E1", "ok": true },
    { "step": 4, "action": "freezeRows", "rows": 2, "ok": true }
  ]
}
```

If a helper fails, OfficeWeaver records the failed step and stops immediately.

```json
{
  "ok": false,
  "applied": 2,
  "failed": 1,
  "stopped_at": 3,
  "outcomes": [
    { "step": 1, "action": "setValue", "address": "A1", "ok": true },
    { "step": 2, "action": "setFont", "address": "A1:E1", "ok": true },
    { "step": 3, "action": "insertRows", "at": 1, "ok": false, "error": "..." }
  ]
}
```

That gives an LLM enough context to retry only the failed part instead of guessing.

## Why Not Just Use Raw APIs?

Raw APIs are still available for advanced cases, but they do not solve agent observability. The editor knows whether a JavaScript command returned or threw, but it does not know that line 1 styled a range, line 2 inserted rows, and line 3 failed. OfficeWeaver makes those steps explicit.

Use raw APIs when a feature is rare or not wrapped yet. Use `Sheet.*` helpers for common spreadsheet edits that need reliable tracing.

```js
Sheet.raw("customChartWork", { range: "A1:E10" }, function (Api, Sheet) {
  const ws = Sheet.sheet();
  return ws.AddChart("'Sheet1'!$A$1:$E$10", false, "bar", 2, 4320000, 2520000, 6, 0, 1, 0);
});
```

## Versioned Adapter Model

Office editor APIs are version-sensitive. OfficeWeaver keeps a stable public helper shape while selecting an engine/version adapter internally.

Planned adapter layout:

```text
src/
  core/
  adapters/
    onlyoffice/
      9.3/
      9.4/
    hancom/
      3.2/
      3.5/
rag/
  onlyoffice/
    9.3/
      spreadsheet/
      document/
      presentation/
    9.4/
  hancom/
    3.5/
      hwp/
```

This lets one application use ONLYOFFICE 9.3 for spreadsheets and Hancom 3.5 for HWP without mixing API references.

## Current Scope

The first implementation targets ONLYOFFICE Spreadsheet API 9.3/9.4 families and exposes:

- `Sheet.setValue`
- `Sheet.setFormula`
- `Sheet.setNumberFormat`
- `Sheet.setFont`
- `Sheet.setFill`
- `Sheet.setBorder`
- `Sheet.setAlignment`
- `Sheet.merge`
- `Sheet.unmerge`
- `Sheet.insertRows`
- `Sheet.deleteRows`
- `Sheet.insertColumns`
- `Sheet.deleteColumns`
- `Sheet.setColumnWidth`
- `Sheet.setRowHeight`
- `Sheet.addSheet`
- `Sheet.renameSheet`
- `Sheet.deleteSheet`
- `Sheet.setActiveSheet`
- `Sheet.freezeRows`
- `Sheet.freezeColumns`
- `Sheet.freezeAt`
- `Sheet.unfreeze`
- `Sheet.raw`
- `Sheet.done`

More helpers should be added based on repeated real-world failures, not by trying to wrap every editor API on day one.

## How A Host App Uses It

OfficeWeaver is not a server. It is a JavaScript runtime loaded inside an editor plugin.

```html
<script src="plugins.js"></script>
<script src="officeweaver.js"></script>
<script src="code.js"></script>
```

The host app still owns:

- LLM tool registration;
- chat UI;
- RAG search;
- document open/draft/save flows;
- retry policy.

OfficeWeaver owns:

- stable helper APIs for generated code;
- editor-version-specific implementation details;
- structured outcomes;
- fail-fast behavior;
- formula evaluated value echo where possible.

The RAG manuals under `rag/` are part of this package. Host apps should read them directly from the installed package, or include them in their build output from the package path. They should not maintain a second hand-edited copy of the same manuals.

## Naming

The project name is `OfficeWeaver`.

The public runtime object is:

```js
OfficeWeaver
```

Generated spreadsheet macros should use:

```js
Sheet.*
```

Future document types should use:

```js
Doc.*
Slide.*
Hwp.*
```
