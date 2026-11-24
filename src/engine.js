'use strict';

const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const SOURCE_PATH = path.join(__dirname, 'shopee-v10-base.js');
const source = fs.readFileSync(SOURCE_PATH, 'utf8');

class FakeRange {
  constructor(sheet, row, col, numRows = 1, numCols = 1) {
    Object.assign(this, { sheet, row, col, numRows, numCols });
  }
  getValues() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) {
      const line = [];
      for (let c = 0; c < this.numCols; c++) line.push(this.sheet.values[this.row - 1 + r]?.[this.col - 1 + c] ?? '');
      out.push(line);
    }
    return out;
  }
  setValues(values) {
    for (let r = 0; r < values.length; r++) {
      const targetRow = this.row - 1 + r;
      if (!this.sheet.values[targetRow]) this.sheet.values[targetRow] = [];
      for (let c = 0; c < values[r].length; c++) this.sheet.values[targetRow][this.col - 1 + c] = values[r][c];
    }
    return this;
  }
  setFontWeight() { return this; }
  setFontColor() { return this; }
  setBackground() { return this; }
  setFontFamily() { return this; }
  setFontSize() { return this; }
}

class FakeSheet {
  constructor(name, values = []) { this.name = name; this.values = values.map(row => row.slice()); }
  getDataRange() {
    const rows = Math.max(1, this.getLastRow());
    const cols = Math.max(1, this.getLastColumn());
    return new FakeRange(this, 1, 1, rows, cols);
  }
  getRange(row, col, numRows = 1, numCols = 1) { return new FakeRange(this, row, col, numRows, numCols); }
  getLastRow() {
    let last = this.values.length;
    while (last > 0 && this.values[last - 1].every(v => v === '' || v === null || v === undefined)) last--;
    return last;
  }
  getLastColumn() { return this.values.reduce((max, row) => Math.max(max, row.length), 0); }
  setColumnWidth() { return this; }
  setFrozenRows() { return this; }
  hideSheet() { return this; }
}

class FakeSpreadsheet {
  constructor(sheets) { this.sheets = sheets; }
  getSheetByName(name) { return this.sheets.find(s => s.name === name) || null; }
  insertSheet(name) { const sheet = new FakeSheet(name); this.sheets.push(sheet); return sheet; }
  deleteSheet(sheet) { this.sheets = this.sheets.filter(s => s !== sheet); }
  getSheets() { return this.sheets.slice(); }
  setActiveSheet() { return this; }
  moveActiveSheet() { return this; }
  toast() { return this; }
}

function runConsolidation({ orderRows, inventoryRows }) {
  const spreadsheet = new FakeSpreadsheet([
    new FakeSheet('Raw', orderRows),
    new FakeSheet('PRODUCT_DB', inventoryRows)
  ]);
  const ui = { alert(message) { throw new Error(String(message)); }, createMenu() { return { addItem() { return this; }, addToUi() { return this; } }; } };
  const context = {
    console,
    SpreadsheetApp: {
      getActiveSpreadsheet: () => spreadsheet,
      getUi: () => ui
    },
    Math, Date, JSON, RegExp, String, Number, Boolean, Array, Object, parseInt, parseFloat, isNaN
  };
  vm.createContext(context);
  vm.runInContext(source, context, { filename: 'shopee-v10-base.js', timeout: 15000 });
  context.consolidateOrders();
  const output = spreadsheet.getSheetByName('Consolidated');
  const summary = spreadsheet.getSheetByName('Summary');
  const rows = output ? output.getDataRange().getValues() : [];
  const nonEmptyRows = rows;
  const lineRows = rows.filter(row => row.some(v => v !== '' && v !== null && v !== undefined));
  const matched = lineRows.filter(row => String(row[6] || '').trim()).length;
  const categories = summary ? summary.getDataRange().getValues().slice(1).filter(row => row[0] && row[0] !== 'TOTAL' && row[0] !== 'Total Orders (unique)').length : 0;
  return {
    rows: nonEmptyRows,
    summary: summary ? summary.getDataRange().getValues() : [],
    stats: { lineItems: lineRows.length, matched, unmatched: Math.max(0, lineRows.length - matched), categories }
  };
}

module.exports = { runConsolidation };
