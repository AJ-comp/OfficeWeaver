const assert = require('node:assert/strict');
const OfficeWeaver = require('../src/officeweaver');

class MockRange {
  constructor(sheet, address) {
    this.sheet = sheet;
    this.address = address;
  }

  SetValue(value) {
    this.sheet.cells[this.address] = this.sheet.cells[this.address] || {};
    this.sheet.cells[this.address].value = value;
  }

  GetValue() {
    const cell = this.sheet.cells[this.address];
    return cell ? cell.value : null;
  }

  GetCells() {
    return this;
  }

  SetNumberFormat(format) {
    this._style('numberFormat', format);
  }

  SetBold(value) {
    this._style('bold', value);
  }

  SetItalic(value) {
    this._style('italic', value);
  }

  SetUnderline(value) {
    this._style('underline', value);
  }

  SetFontColor(value) {
    this._style('fontColor', value);
  }

  SetFillColor(value) {
    this._style('fillColor', value);
  }

  SetFontSize(value) {
    this._style('fontSize', value);
  }

  SetFontName(value) {
    this._style('fontName', value);
  }

  SetBorders(side, style, color) {
    this._style('borders', this._styleValue('borders', []).concat([{ side, style, color }]));
  }

  SetAlignHorizontal(value) {
    this._style('horizontal', value);
  }

  SetAlignVertical(value) {
    this._style('vertical', value);
  }

  SetWrap(value) {
    this._style('wrap', value);
  }

  Merge(across) {
    this._style('merged', across === true);
  }

  UnMerge() {
    this._style('merged', false);
  }

  Insert(direction) {
    this.sheet.operations.push({ action: 'insert', address: this.address, direction });
  }

  Delete(direction) {
    this.sheet.operations.push({ action: 'delete', address: this.address, direction });
  }

  _style(name, value) {
    const cell = this.sheet.cells[this.address] = this.sheet.cells[this.address] || {};
    cell.styles = cell.styles || {};
    cell.styles[name] = value;
  }

  _styleValue(name, fallback) {
    const cell = this.sheet.cells[this.address];
    return cell && cell.styles && cell.styles[name] ? cell.styles[name] : fallback;
  }
}

class MockFreezePanes {
  constructor(sheet) {
    this.sheet = sheet;
  }

  FreezeRows(rows) {
    this.sheet.freeze = { type: 'row', rows };
  }

  FreezeColumns(columns) {
    this.sheet.freeze = { type: 'column', columns };
  }

  FreezeAt(range) {
    this.sheet.freeze = { type: 'cell', address: range.address };
  }

  Unfreeze() {
    this.sheet.freeze = null;
  }
}

class MockSheet {
  constructor(name) {
    this.name = name;
    this.cells = {};
    this.operations = [];
    this.freeze = null;
    this.active = false;
  }

  GetName() {
    return this.name;
  }

  SetName(name) {
    this.name = name;
  }

  SetActive() {
    this.active = true;
  }

  Delete() {
    this.deleted = true;
  }

  GetRange(address) {
    return new MockRange(this, address);
  }

  SetColumnWidth(index, width) {
    this.operations.push({ action: 'setColumnWidth', index, width });
  }

  SetRowHeight(index, height) {
    this.operations.push({ action: 'setRowHeight', index, height });
  }

  GetFreezePanes() {
    return new MockFreezePanes(this);
  }
}

function createApi() {
  const sheet = new MockSheet('Sheet1');
  const sheets = [sheet];

  return {
    sheet,
    sheets,
    freezeType: null,
    CreateColorFromRGB(r, g, b) {
      return { r, g, b };
    },
    GetActiveSheet() {
      return sheet;
    },
    GetSheets() {
      return sheets;
    },
    AddSheet(name) {
      const next = new MockSheet(name);
      sheets.push(next);
      return next;
    },
    SetFreezePanesType(type) {
      this.freezeType = type;
    }
  };
}

function runMacro(code) {
  const api = createApi();
  global.Api = api;

  try {
    const command = OfficeWeaver.buildSpreadsheetMacroCommand(code, {
      engine: 'onlyoffice',
      version: '9.3.1.2'
    });
    return { api, result: command() };
  } finally {
    delete global.Api;
  }
}

class MockParagraph {
  constructor(text = '') {
    this.text = text;
    this.styles = {};
    this.numbering = null;
  }

  AddText(text) {
    this.text += String(text);
    return { text: String(text) };
  }

  RemoveAllElements() {
    this.text = '';
    return true;
  }

  GetText() {
    return this.text;
  }

  SetBold(value) {
    this.styles.bold = value;
    return this;
  }

  SetItalic(value) {
    this.styles.italic = value;
    return this;
  }

  SetUnderline(value) {
    this.styles.underline = value;
    return this;
  }

  SetColor(value) {
    this.styles.color = value;
    return this;
  }

  SetFontSize(value) {
    this.styles.fontSize = value;
    return this;
  }

  SetFontFamily(value) {
    this.styles.fontFamily = value;
    return this;
  }

  SetSpacingAfter(value) {
    this.styles.spacingAfter = value;
    return true;
  }

  SetSpacingBefore(value) {
    this.styles.spacingBefore = value;
    return true;
  }

  SetJc(value) {
    this.styles.align = value;
    return true;
  }

  SetKeepNext(value) {
    this.styles.keepNext = value;
    return true;
  }

  SetKeepLines(value) {
    this.styles.keepLines = value;
    return true;
  }

  SetStyle(value) {
    this.styles.style = value;
    return true;
  }

  SetNumbering(value) {
    this.numbering = value;
    return true;
  }

  SetContextualSpacing(value) {
    this.styles.contextualSpacing = value;
    return true;
  }
}

class MockTextRange {
  constructor(doc, start, end) {
    this.doc = doc;
    this.start = start;
    this.end = end;
    this.styles = {};
  }

  SetBold(value) {
    this.styles.bold = value;
    return this;
  }

  SetItalic(value) {
    this.styles.italic = value;
    return this;
  }

  SetUnderline(value) {
    this.styles.underline = value;
    return this;
  }

  SetColor(value) {
    this.styles.color = value;
    return this;
  }

  SetFontSize(value) {
    this.styles.fontSize = value;
    return this;
  }

  SetFontFamily(value) {
    this.styles.fontFamily = value;
    return this;
  }
}

class MockStyle {
  constructor(name, type) {
    this.name = name;
    this.type = type;
    this.tablePr = {
      SetTableBorderAll: (...args) => {
        this.border = args;
        return true;
      }
    };
  }

  SetBasedOn(style) {
    this.basedOn = style;
  }

  GetTablePr() {
    return this.tablePr;
  }
}

class MockTableCell {
  constructor() {
    this.paragraphs = [];
    this.shading = null;
  }

  SetShd(type, r, g, b, isAuto) {
    this.shading = { type, r, g, b, isAuto };
    return true;
  }
}

class MockTable {
  constructor(rows, columns) {
    this.rows = rows;
    this.columns = columns;
    this.cells = Array.from({ length: rows }, () => Array.from({ length: columns }, () => new MockTableCell()));
    this.width = null;
    this.style = null;
  }

  SetWidth(type, value) {
    this.width = { type, value };
  }

  SetStyle(style) {
    this.style = style;
  }

  SetTableLook(...args) {
    this.look = args;
  }

  GetCell(row, column) {
    return this.cells[row] && this.cells[row][column] ? this.cells[row][column] : null;
  }

  AddElement(cell, index, paragraph) {
    cell.paragraphs.splice(index, 0, paragraph);
    return true;
  }
}

class MockDocument {
  constructor() {
    this.elements = [new MockParagraph('Hello world')];
    this.ranges = [];
    this.styles = {};
    this.replacements = [];
  }

  GetElement(index) {
    return this.elements[index] || null;
  }

  Push(element) {
    this.elements.push(element);
    return true;
  }

  AddElement(index, element) {
    this.elements.splice(index, 0, element);
    return true;
  }

  RemoveAllElements() {
    this.elements = [new MockParagraph()];
    return true;
  }

  GetRange(start, end) {
    const range = new MockTextRange(this, start, end);
    this.ranges.push(range);
    return range;
  }

  Search(text, matchCase) {
    const needle = matchCase ? String(text) : String(text).toLowerCase();
    return this.elements.filter((element) => {
      const value = matchCase ? (element.text || '') : (element.text || '').toLowerCase();
      return value.includes(needle);
    });
  }

  SearchAndReplace(props) {
    this.replacements.push(props);
    this.elements.forEach((element) => {
      if (typeof element.text === 'string') {
        element.text = element.text.split(props.searchString).join(props.replaceString);
      }
    });
    return true;
  }

  CreateStyle(name, type = 'paragraph') {
    const style = new MockStyle(name, type);
    this.styles[name] = style;
    return style;
  }

  GetStyle(name) {
    return this.styles[name] || new MockStyle(name, 'paragraph');
  }

  CreateNumbering(type = 'bullet') {
    return {
      type,
      GetLevel(level) {
        return { type, level };
      }
    };
  }

  GetAllParagraphs() {
    return this.elements.filter((element) => element instanceof MockParagraph);
  }
}

function createTextApi() {
  const doc = new MockDocument();
  return {
    doc,
    GetDocument() {
      return doc;
    },
    CreateParagraph() {
      return new MockParagraph();
    },
    CreateTable(rows, columns) {
      return new MockTable(rows, columns);
    },
    HexColor(hex) {
      return { hex };
    },
    RGB(r, g, b) {
      return { r, g, b };
    }
  };
}

function runTextMacro(code) {
  const api = createTextApi();
  global.Api = api;

  try {
    const command = OfficeWeaver.buildTextDocumentMacroCommand(code, {
      engine: 'onlyoffice',
      version: '9.3.1.2'
    });
    return { api, result: command() };
  } finally {
    delete global.Api;
  }
}

assert.equal(OfficeWeaver.resolveVersionFamily('onlyoffice', '9.3.1.2'), '9.3');
assert.equal(OfficeWeaver.resolveVersionFamily('onlyoffice', '10.0.0'), '10.0.0');
assert.equal(
  OfficeWeaver.normalizeSpreadsheetMacroCode('range.SetFontBold(true); range.SetBackgroundColor(color);'),
  'range.SetBold(true); range.SetFillColor(color);'
);

{
  const { api, result } = runMacro(`
Sheet.setValue("A1", "Sales Report");
Sheet.setFormula("F1", "SUM(E2:E10)");
Sheet.setFont("A1:E1", { bold: true, color: "#fff", size: 12 });
Sheet.freezeRows(2);
return Sheet.done("Report header updated", false);
`);

  assert.equal(result.ok, true);
  assert.equal(result.applied, 4);
  assert.equal(result.failed, 0);
  assert.equal(result.summary, 'Report header updated');
  assert.equal(result.result.data, false);
  assert.equal(api.sheet.cells.A1.value, 'Sales Report');
  assert.equal(api.sheet.cells.F1.value, '=SUM(E2:E10)');
  assert.equal(api.sheet.cells['A1:E1'].styles.bold, true);
  assert.deepEqual(api.sheet.cells['A1:E1'].styles.fontColor, { r: 255, g: 255, b: 255 });
  assert.deepEqual(api.sheet.freeze, { type: 'row', rows: 2 });
  assert.equal(result.outcomes[1].evaluated_value, '=SUM(E2:E10)');
}

{
  const { api, result } = runMacro(`
Sheet.setValue("A1", "ok");
Sheet.setFill("A1", "not-a-color");
Sheet.setValue("A2", "should not run");
`);

  assert.equal(result.ok, false);
  assert.equal(result.applied, 1);
  assert.equal(result.failed, 1);
  assert.equal(result.stopped_at, 2);
  assert.match(result.outcomes[1].error, /invalid_color/);
  assert.equal(api.sheet.cells.A1.value, 'ok');
  assert.equal(api.sheet.cells.A2, undefined);
}

{
  const { result } = runMacro(`
Sheet.raw("custom", { target: "A1" }, function (Api, Sheet) {
  return { ok: true, name: Sheet.sheet().GetName() };
});
return Sheet.done("raw done");
`);

  assert.equal(result.ok, true);
  assert.equal(result.applied, 1);
  assert.equal(result.summary, 'raw done');
  assert.deepEqual(result.outcomes[0].result, { ok: true, name: 'Sheet1' });
}

{
  const { api, result } = runTextMacro(`
Doc.setParagraphText(0, "Quarterly report");
Doc.styleParagraph(0, { bold: true, color: "#1F4E79", size: 32, align: "center" });
Doc.addParagraph("Prepared by OfficeWeaver", { italic: true, spacingAfter: 160 });
Doc.replace("OfficeWeaver", "Mythosia", { matchCase: true });
return Doc.done("Document title updated");
`);

  assert.equal(result.ok, true);
  assert.equal(result.kind, 'word');
  assert.equal(result.applied, 4);
  assert.equal(result.summary, 'Document title updated');
  assert.equal(api.doc.elements[0].text, 'Quarterly report');
  assert.equal(api.doc.elements[0].styles.bold, true);
  assert.deepEqual(api.doc.elements[0].styles.color, { hex: '#1F4E79' });
  assert.equal(api.doc.elements[1].text, 'Prepared by Mythosia');
  assert.equal(result.outcomes[3].result, true);
}

{
  const { api, result } = runTextMacro(`
Doc.addHeading("Sales Summary", 1, { color: "#2C3E50" });
Doc.addList(["Revenue grew", "Costs remained stable"], { type: "bullet" });
Doc.addTable([
  ["Metric", "Value"],
  ["Revenue", "120000"],
  ["Cost", "45000"]
], { headerFill: "#1F4E79", borderColor: "#D9E2EC", bandRows: true });
return Doc.done("Report structure created");
`);

  assert.equal(result.ok, true);
  assert.equal(result.applied, 3);
  assert.equal(api.doc.elements[1].text, 'Sales Summary');
  assert.equal(api.doc.elements[1].styles.bold, true);
  assert.equal(api.doc.elements[2].numbering.type, 'bullet');
  const table = api.doc.elements[4];
  assert.equal(table.rows, 3);
  assert.equal(table.columns, 2);
  assert.equal(table.cells[0][0].paragraphs[0].text, 'Metric');
  assert.deepEqual(table.cells[0][0].shading, { type: 'clear', r: 31, g: 78, b: 121, isAuto: false });
}

{
  const { api, result } = runTextMacro(`
Doc.addParagraph("ok");
Doc.styleParagraph(0, { background: "not-a-color" });
Doc.addParagraph("should not run");
`);

  assert.equal(result.ok, false);
  assert.equal(result.applied, 1);
  assert.equal(result.stopped_at, 2);
  assert.match(result.outcomes[1].error, /invalid_color/);
  assert.equal(api.doc.elements.length, 2);
}

console.log('OfficeWeaver tests passed');
