const fs = require('fs');
const path = require('path');
const { seedDefaultSpreadsheet } = require('./seedData');

function isCellEmpty(val) {
  return val === null || val === undefined || val === '';
}

function parseA1(a1) {
  const match = a1.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
  if (!match) return null;
  const colToNum = (col) => {
    let num = 0;
    for (let i = 0; i < col.length; i++) {
      num = num * 26 + (col.toUpperCase().charCodeAt(i) - 64);
    }
    return num;
  };
  const startCol = colToNum(match[1]);
  const startRow = parseInt(match[2], 10);
  const endCol = match[3] ? colToNum(match[3]) : startCol;
  const endRow = match[4] ? parseInt(match[4], 10) : startRow;
  return {
    row: Math.min(startRow, endRow),
    col: Math.min(startCol, endCol),
    numRows: Math.abs(endRow - startRow) + 1,
    numCols: Math.abs(endCol - startCol) + 1
  };
}

class Range {
  constructor(sheet, row, col, numRows = 1, numCols = 1) {
    this._sheet = sheet;
    this._row = Math.max(1, row);
    this._col = Math.max(1, col);
    this._numRows = Math.max(0, numRows);
    this._numCols = Math.max(0, numCols);
  }

  getRow() { return this._row; }
  getColumn() { return this._col; }
  getLastRow() { return this._row + this._numRows - 1; }
  getLastColumn() { return this._col + this._numCols - 1; }
  getNumRows() { return this._numRows; }
  getNumColumns() { return this._numCols; }

  getValue() {
    return this._sheet._getCellValue(this._row, this._col);
  }

  setValue(val) {
    this._sheet._setCellValue(this._row, this._col, val);
    return this;
  }

  getValues() {
    const result = [];
    for (let r = 0; r < this._numRows; r++) {
      const rowArr = [];
      for (let c = 0; c < this._numCols; c++) {
        rowArr.push(this._sheet._getCellValue(this._row + r, this._col + c));
      }
      result.push(rowArr);
    }
    return result;
  }

  setValues(values2D) {
    if (!Array.isArray(values2D)) return this;
    for (let r = 0; r < values2D.length; r++) {
      const rowArr = values2D[r];
      if (!Array.isArray(rowArr)) continue;
      for (let c = 0; c < rowArr.length; c++) {
        this._sheet._setCellValue(this._row + r, this._col + c, rowArr[c]);
      }
    }
    return this;
  }

  clearContent() {
    for (let r = 0; r < this._numRows; r++) {
      for (let c = 0; c < this._numCols; c++) {
        this._sheet._setCellValue(this._row + r, this._col + c, '');
      }
    }
    return this;
  }

  clear() {
    return this.clearContent();
  }

  setBackground(color) { return this; }
  setBackgrounds(colors) { return this; }
  setFontColor(color) { return this; }
  setFontColors(colors) { return this; }
  setFontWeight(weight) { return this; }
  setNumberFormat(format) { return this; }
  setNumberFormats(formats) { return this; }
  setWrap(wrap) { return this; }
  setHorizontalAlignment(alignment) { return this; }
  setVerticalAlignment(alignment) { return this; }

  offset(rowOffset, columnOffset, optNumRows, optNumColumns) {
    return new Range(
      this._sheet,
      this._row + rowOffset,
      this._col + columnOffset,
      optNumRows !== undefined ? optNumRows : this._numRows,
      optNumColumns !== undefined ? optNumColumns : this._numCols
    );
  }
}

class Sheet {
  constructor(spreadsheet, name) {
    this._spreadsheet = spreadsheet;
    this._name = name;
    this._grid = [];
    this._hidden = false;
  }

  getName() { return this._name; }
  setName(name) { this._name = name; return this; }

  hideSheet() { this._hidden = true; return this; }
  showSheet() { this._hidden = false; return this; }
  isSheetHidden() { return this._hidden; }

  _ensureGrid(row, col) {
    while (this._grid.length < row) {
      this._grid.push([]);
    }
    const rowArr = this._grid[row - 1];
    while (rowArr.length < col) {
      rowArr.push('');
    }
  }

  _getCellValue(row, col) {
    if (row < 1 || col < 1 || row > this._grid.length) return '';
    const rowArr = this._grid[row - 1];
    if (!rowArr || col > rowArr.length) return '';
    const val = rowArr[col - 1];
    return val !== undefined ? val : '';
  }

  _setCellValue(row, col, val) {
    if (row < 1 || col < 1) return;
    this._ensureGrid(row, col);
    this._grid[row - 1][col - 1] = val;
    this._spreadsheet._markDirty();
  }

  getLastRow() {
    for (let r = this._grid.length - 1; r >= 0; r--) {
      const rowArr = this._grid[r];
      if (rowArr) {
        for (let c = 0; c < rowArr.length; c++) {
          if (!isCellEmpty(rowArr[c])) {
            return r + 1;
          }
        }
      }
    }
    return 0;
  }

  getLastColumn() {
    let maxCol = 0;
    for (let r = 0; r < this._grid.length; r++) {
      const rowArr = this._grid[r];
      if (rowArr) {
        for (let c = rowArr.length - 1; c >= 0; c--) {
          if (!isCellEmpty(rowArr[c])) {
            if (c + 1 > maxCol) maxCol = c + 1;
            break;
          }
        }
      }
    }
    return maxCol;
  }

  getRange(a, b, c, d) {
    if (typeof a === 'string') {
      const parsed = parseA1(a);
      if (parsed) {
        return new Range(this, parsed.row, parsed.col, parsed.numRows, parsed.numCols);
      }
      return new Range(this, 1, 1, 1, 1);
    }
    return new Range(this, a, b, c !== undefined ? c : 1, d !== undefined ? d : 1);
  }

  getDataRange() {
    const lastRow = Math.max(1, this.getLastRow());
    const lastCol = Math.max(1, this.getLastColumn());
    return new Range(this, 1, 1, lastRow, lastCol);
  }

  appendRow(rowArray) {
    const nextRow = this.getLastRow() + 1;
    this._ensureGrid(nextRow, rowArray.length);
    for (let c = 0; c < rowArray.length; c++) {
      this._grid[nextRow - 1][c] = rowArray[c];
    }
    this._spreadsheet._markDirty();
    return this;
  }

  insertRowAfter(afterPosition) {
    const pos = Math.max(0, afterPosition);
    this._grid.splice(pos, 0, []);
    this._spreadsheet._markDirty();
    return this;
  }

  insertRowBefore(beforePosition) {
    const pos = Math.max(0, beforePosition - 1);
    this._grid.splice(pos, 0, []);
    this._spreadsheet._markDirty();
    return this;
  }

  deleteRow(rowPosition) {
    if (rowPosition >= 1 && rowPosition <= this._grid.length) {
      this._grid.splice(rowPosition - 1, 1);
      this._spreadsheet._markDirty();
    }
    return this;
  }

  deleteRows(rowPosition, howMany = 1) {
    if (rowPosition >= 1 && rowPosition <= this._grid.length) {
      this._grid.splice(rowPosition - 1, howMany);
      this._spreadsheet._markDirty();
    }
    return this;
  }

  clear() {
    this._grid = [];
    this._spreadsheet._markDirty();
    return this;
  }

  clearContents() {
    return this.clear();
  }

  getMaxRows() {
    return Math.max(1000, this._grid.length);
  }

  getMaxColumns() {
    let max = 26;
    for (const r of this._grid) {
      if (r && r.length > max) max = r.length;
    }
    return max;
  }

  setFrozenRows(n) { return this; }
  setFrozenColumns(n) { return this; }
  setTabColor(color) { return this; }
}

class Spreadsheet {
  constructor(id, storageDir) {
    this._id = id;
    this._storageDir = storageDir;
    this._filePath = path.join(storageDir, `${id}.json`);
    this._sheets = new Map();
    this._dirty = false;
    this._load();
  }

  getId() { return this._id; }
  getName() { return 'Spreadsheet_' + this._id; }

  _markDirty() {
    this._dirty = true;
  }

  _load() {
    try {
      if (!fs.existsSync(this._filePath)) {
        seedDefaultSpreadsheet(this._storageDir, this._id);
      }
      if (fs.existsSync(this._filePath)) {
        const raw = fs.readFileSync(this._filePath, 'utf8');
        const json = JSON.parse(raw);
        for (const [name, rows] of Object.entries(json.sheets || {})) {
          const sheet = new Sheet(this, name);
          sheet._grid = (rows || []).map(row => 
            (row || []).map(cell => {
              if (typeof cell === 'string' && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(cell)) {
                const d = new Date(cell);
                return isNaN(d.getTime()) ? cell : d;
              }
              return cell;
            })
          );
          this._sheets.set(name, sheet);
        }
      }
    } catch (e) {
      // ignore read error
    }
  }

  flush() {
    if (!this._dirty) return;
    try {
      if (!fs.existsSync(this._storageDir)) {
        fs.mkdirSync(this._storageDir, { recursive: true });
      }
      const data = { sheets: {} };
      for (const [name, sheet] of this._sheets.entries()) {
        data.sheets[name] = sheet._grid.map(row =>
          row.map(cell => (cell instanceof Date ? cell.toISOString() : cell))
        );
      }
      fs.writeFileSync(this._filePath, JSON.stringify(data, null, 2), 'utf8');
      this._dirty = false;
    } catch (e) {
      // ignore write error
    }
  }

  getSheetByName(name) {
    return this._sheets.get(name) || null;
  }

  getSheets() {
    return Array.from(this._sheets.values());
  }

  getActiveSheet() {
    const first = this.getSheets()[0];
    return first || null;
  }

  insertSheet(name) {
    let sheet = this._sheets.get(name);
    if (!sheet) {
      sheet = new Sheet(this, name);
      this._sheets.set(name, sheet);
      this._markDirty();
    }
    return sheet;
  }

  deleteSheet(sheet) {
    if (typeof sheet === 'string') {
      this._sheets.delete(sheet);
    } else if (sheet && sheet.getName) {
      this._sheets.delete(sheet.getName());
    }
    this._markDirty();
  }
}

const storageDir = path.resolve(__dirname, '../../data/sheets');
const spreadsheetCache = new Map();

class SpreadsheetApp {
  static openById(id) {
    const key = String(id || 'default');
    if (!spreadsheetCache.has(key)) {
      spreadsheetCache.set(key, new Spreadsheet(key, storageDir));
    }
    return spreadsheetCache.get(key);
  }

  static getActiveSpreadsheet() {
    const defaultId = process.env.SPREADSHEET_ID || '1qZnZov1l9DfvhTzruqlLubGhydcOOFKLf2kG8xplNLM';
    return SpreadsheetApp.openById(defaultId);
  }

  static flush() {
    for (const ss of spreadsheetCache.values()) {
      ss.flush();
    }
  }
}

module.exports = { SpreadsheetApp, Spreadsheet, Sheet, Range };
