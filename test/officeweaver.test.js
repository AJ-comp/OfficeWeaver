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

assert.equal(OfficeWeaver.resolveVersionFamily('onlyoffice', '9.3.1.2'), '9.3');
assert.equal(OfficeWeaver.resolveVersionFamily('onlyoffice', '9.4.0'), '9.4');
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

console.log('OfficeWeaver tests passed');
