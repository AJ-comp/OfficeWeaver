# OfficeWeaver

[English](README.md) | [한국어](README.ko.md) | [日本語](README.ja.md)

OfficeWeaver is a small JavaScript runtime for building AI-friendly office document automation.

It sits between an LLM-generated macro and an editor API such as ONLYOFFICE. The goal is simple: let the model edit documents with a stable helper API, while the host application receives reliable per-step execution results.

OfficeWeaver is currently focused on ONLYOFFICE Spreadsheet macros. The package also includes versioned manuals that are ready to be indexed for RAG, so an application can retrieve the right API notes for the editor version it is running.

## The Problem

Office editors expose powerful JavaScript APIs, but raw APIs are not easy for LLM agents to use reliably.

Common problems:

- Models mix up similar APIs from ExcelJS, Office.js, VBA, Google Apps Script, and ONLYOFFICE.
- A model may write `SetFontBold`, `GetValues`, `SetFreezePanes`, or `Api.CreateChart` even when those names do not exist in the current editor.
- A raw JavaScript macro usually fails as one big script. The host can see that it failed, but not which logical edit succeeded before the failure.
- If a formula is written, the agent often needs the evaluated cell value to verify whether the formula worked.
- Each application ends up writing its own helper layer, retry logging, manuals, and editor-version notes.

OfficeWeaver addresses those points by giving the LLM a smaller, traced API to call.

## What OfficeWeaver Solves

Instead of asking the LLM to directly use every raw editor method, let it generate code like this:

```js
Sheet.setValue("A1", "Sales Report");
Sheet.setFormula("F1", "=SUM(E2:E10)");
Sheet.setFont("A1:E1", { bold: true, color: "#FFFFFF", size: 14 });
Sheet.freezeRows(2);
return Sheet.done("Report header updated");
```

OfficeWeaver runs those calls against the actual editor API and returns structured outcomes:

```json
{
  "ok": true,
  "engine": "onlyoffice",
  "version": "9.3.1.2",
  "version_family": "9.3",
  "summary": "Report header updated",
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

If a step fails, OfficeWeaver stops immediately and returns the exact failed step:

```json
{
  "ok": false,
  "summary": "Macro failed at step 3",
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

This makes retries much easier. The agent can see what already happened and what failed.

## Mental Model

```text
User request
  -> host app asks LLM for a macro
  -> LLM writes Sheet.* code
  -> OfficeWeaver executes the helpers
  -> ONLYOFFICE changes the document
  -> OfficeWeaver returns structured outcomes
  -> host app shows success or sends failure details back to the LLM
```

OfficeWeaver is not a chat system, not a vector database, and not a document server. It is the runtime and documentation layer that makes editor automation easier for AI agents.

## Installation

After the package is published:

```bash
npm install @mythosia/officeweaver
```

For local development before publishing:

```bash
npm install ../OfficeWeaver
```

The package contains:

```text
src/officeweaver.js       runtime loaded inside the editor plugin
src/officeweaver.d.ts     TypeScript declarations for host tooling
manuals/                  versioned manuals ready for RAG indexing
README.md                 this guide
README.ko.md              Korean guide
README.ja.md              Japanese guide
```

## Loading In An Editor Plugin

In an ONLYOFFICE plugin, load OfficeWeaver before the code that executes LLM-generated macros.

```html
<script src="plugins.js"></script>
<script src="officeweaver.js"></script>
<script src="code.js"></script>
```

A host app normally copies the runtime from:

```text
node_modules/@mythosia/officeweaver/src/officeweaver.js
```

to the plugin public folder, for example:

```text
wwwroot/onlyoffice-plugin/officeweaver.js
```

OfficeWeaver exposes a global object:

```js
OfficeWeaver
```

The generated spreadsheet macro receives a helper namespace:

```js
Sheet
```

## Running A Spreadsheet Macro

The host app can create an executable command from LLM-generated code:

```js
const command = OfficeWeaver.buildSpreadsheetMacroCommand(`
Sheet.setValue("A1", "Sales Report");
Sheet.setFont("A1:E1", { bold: true, color: "#FFFFFF", size: 14 });
Sheet.setFill("A1:E1", "#1F4E79");
return Sheet.done("Styled report title");
`, {
  engine: "onlyoffice",
  version: "9.3.1.2"
});

const result = command();
```

Inside ONLYOFFICE, `command()` expects the editor `Api` object to be available. OfficeWeaver uses that object internally.

## What The LLM Should Use

For common spreadsheet work, ask the LLM to use `Sheet.*` helpers first.

### Values And Formulas

```js
Sheet.setValue("A1", "Sales Report");
Sheet.setFormula("F1", "=SUM(E2:E10)");
return Sheet.done("Updated title and total formula");
```

### Formatting

```js
Sheet.setFill("A1:E1", "#1F4E79");
Sheet.setFont("A1:E1", { bold: true, color: "#FFFFFF", size: 12, name: "Arial" });
Sheet.setAlignment("A1:E20", { horizontal: "center", vertical: "center", wrap: true });
Sheet.setNumberFormat("E2:E20", "#,##0");
return Sheet.done("Applied table formatting");
```

### Rows, Columns, And Layout

```js
Sheet.insertRows(1, 2);
Sheet.setRowHeight(1, 30);
Sheet.setColumnWidth("A", 16);
Sheet.deleteColumns("F", 1);
return Sheet.done("Adjusted report layout");
```

### Freeze Panes

```js
Sheet.freezeRows(2);
return Sheet.done("Frozen title and header rows");
```

### Raw API Escape Hatch

Not every editor API is wrapped. For rare operations, use `Sheet.raw` so the operation is still traced.

```js
Sheet.raw("customChart", { range: "A1:E10" }, function (Api, Sheet) {
  const ws = Sheet.sheet();
  return ws.AddChart("'Sheet1'!$A$1:$E$10", false, "bar", 2, 12 * 360000, 7 * 360000, 6, 0, 1, 0);
});

return Sheet.done("Added chart");
```

Prefer this over untracked raw code because `Sheet.raw` still creates an outcome entry.

## Why Not Just Give The LLM Raw ONLYOFFICE API Docs?

You still need API docs. OfficeWeaver includes them under `manuals/`.

However, raw docs alone do not solve execution observability:

- Raw calls do not produce a per-step `outcomes` array.
- A failure does not automatically say which logical operation failed.
- Formula verification must be implemented by every macro or host app.
- Common wrong method names need repeated correction.

OfficeWeaver combines two ideas:

- RAG manuals tell the model what exists.
- `Sheet.*` helpers make common actions traceable and safer to retry.

## Manuals For RAG

The package includes versioned manuals. They are plain Markdown files, but they are written and tagged so host applications can index them into a vector database for RAG.

The files under `manuals/` are source documents, not embeddings. A host application should split these Markdown files into chunks, create embeddings for each chunk, store those embeddings in a vector database, and retrieve the most relevant chunks when asking the LLM to generate OfficeWeaver macro code.

```text
manuals/
  manifest.json
  onlyoffice/
    9.3/
      spreadsheet/
        OfficeWeaver.SheetHelpers.md
        ApiWorksheet.AddChart.md
        examples.ModernTable.md
        errors.CommonWrongApis.md
```

Host apps should read manuals directly from the installed package. Do not maintain a second hand-edited copy in the app project.

Recommended metadata filters:

- `engine`: `onlyoffice`
- `version_family`: `9.3`
- `kind`: `spreadsheet`

When a new editor version is supported, add a new version family folder such as:

```text
manuals/onlyoffice/9.4/spreadsheet/
```

If your application only wants latest docs, delete older embedded chunks and index the latest manifest version again.

## Versioned Adapter Model

Office APIs are version-sensitive. A helper that works for ONLYOFFICE 9.3 may need a different implementation for ONLYOFFICE 9.4 or another office suite.

OfficeWeaver keeps this shape:

```text
src/
  officeweaver.js
  adapters/
    onlyoffice/
      9.3/
      9.4/
    hancom/
      3.2/
      3.5/
manuals/
  onlyoffice/
    9.3/
    9.4/
  hancom/
    3.5/
```

The first version is intentionally small and currently implements the runtime in one file. The public design leaves room for version-specific adapters as the wrapper grows.

This lets a future host app choose combinations like:

```text
ONLYOFFICE Spreadsheet: 9.3.1.2
Hancom HWP: 3.5
```

and retrieve only the manuals for that exact engine/version family.

## Current Spreadsheet Helpers

Current helper list:

- `Sheet.setValue(address, value, sheetName?)`
- `Sheet.setFormula(address, formula, sheetName?)`
- `Sheet.setNumberFormat(address, format, sheetName?)`
- `Sheet.setFont(address, options, sheetName?)`
- `Sheet.setFill(address, color, sheetName?)`
- `Sheet.setBorder(address, options, sheetName?)`
- `Sheet.setAlignment(address, options, sheetName?)`
- `Sheet.merge(address, sheetName?, across?)`
- `Sheet.unmerge(address, sheetName?)`
- `Sheet.insertRows(at, count?, sheetName?)`
- `Sheet.deleteRows(at, count?, sheetName?)`
- `Sheet.insertColumns(at, count?, sheetName?)`
- `Sheet.deleteColumns(at, count?, sheetName?)`
- `Sheet.setColumnWidth(column, width, sheetName?)`
- `Sheet.setRowHeight(row, height, sheetName?)`
- `Sheet.addSheet(name)`
- `Sheet.renameSheet(from, to)`
- `Sheet.deleteSheet(name)`
- `Sheet.setActiveSheet(name)`
- `Sheet.freezeRows(rows, sheetName?)`
- `Sheet.freezeColumns(columns, sheetName?)`
- `Sheet.freezeAt(address, sheetName?)`
- `Sheet.unfreeze(sheetName?)`
- `Sheet.raw(action, details, fn)`
- `Sheet.outcomes()`
- `Sheet.done(summary, data?)`

More helpers should be added based on repeated real-world failures, not by trying to wrap every editor API on day one.

## How A Host App Integrates It

A typical app integration has four parts:

1. Install the package.
2. Copy `src/officeweaver.js` into the editor plugin build output.
3. Index `manuals/` into the app's vector database.
4. Register one macro execution tool that runs generated `Sheet.*` code through OfficeWeaver.

For example:

```text
Application startup
  -> read node_modules/@mythosia/officeweaver/manuals/manifest.json
  -> if embedded version differs, delete old OfficeWeaver chunks
  -> index manuals/onlyoffice/9.3/spreadsheet/*.md

Chat request
  -> retrieve relevant RAG chunks
  -> ask LLM to write Sheet.* macro code
  -> execute with OfficeWeaver
  -> if result.ok is false, send result.outcomes back for retry
```

## Development

Install dependencies if needed:

```bash
npm install
```

Run tests:

```bash
npm test
```

Run syntax checks:

```bash
npm run check
```

Preview the npm package contents:

```bash
npm pack --dry-run
```

Publish after the package is ready:

```bash
npm publish --access public
```

## Current Limitations

- The runtime currently targets ONLYOFFICE Spreadsheet API 9.3/9.4 families.
- Document, presentation, and HWP helpers are planned but not implemented yet.
- Complex features such as advanced charts may still need `Sheet.raw` and RAG references.
- OfficeWeaver does not replace the host app's draft, save, permission, or undo policy.

## Naming

Project name:

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
