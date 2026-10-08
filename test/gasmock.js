// テスト用の最小限の SpreadsheetApp / PropertiesService 等のモック（DBシートの読み書きのみ）
class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const row = this.sh.data[this.r - 1 + i] || [];
      out.push(Array.from({ length: this.nc }, (_, j) => (row[this.c - 1 + j] === undefined ? '' : row[this.c - 1 + j])));
    }
    return out;
  }
  setValues(v) {
    if (this.r + this.nr - 1 > this.sh.maxRows) throw new Error('範囲がシート外です');
    v.forEach((row, i) => {
      const t = (this.sh.data[this.r - 1 + i] = this.sh.data[this.r - 1 + i] || []);
      row.forEach((x, j) => { t[this.c - 1 + j] = x; });
    });
    return this;
  }
  getRow() { return this.r; }
  setFontWeight() { return this; }
  setBackground() { return this; }
  setNumberFormat() { return this; }
  createTextFinder(text) {
    const self = this;
    return {
      matchEntireCell() { return this; },
      findAll() {
        const res = [];
        self.getValues().forEach((row, i) => row.forEach((x, j) => {
          if (String(x) === String(text)) res.push(new Range(self.sh, self.r + i, self.c + j, 1, 1));
        }));
        return res;
      }
    };
  }
}
class Sheet {
  constructor(name) { this.name = name; this.data = []; this.maxRows = 1000; }
  getName() { return this.name; }
  getRange(r, c, nr, nc) { return new Range(this, r, c, nr || 1, nc || 1); }
  getLastRow() { let n = this.data.length; while (n > 0 && !(this.data[n - 1] || []).some((x) => x !== '' && x !== undefined)) n--; return n; }
  getMaxRows() { return this.maxRows; }
  insertRowsAfter(_, n) { this.maxRows += n; }
  deleteRows(start, n) { this.data.splice(start - 1, n); }
  setFrozenRows() {}
}
class Spreadsheet {
  constructor() { this.sheets = []; }
  getSheetByName(n) { return this.sheets.find((s) => s.name === n) || null; }
  insertSheet(n) { const s = new Sheet(n); this.sheets.push(s); return s; }
  getSheets() { return this.sheets; }
}
module.exports = function install(ctx) {
  const db = new Spreadsheet();
  const props = {};
  Object.assign(ctx, {
    SpreadsheetApp: { openById: () => db, flush() {} },
    PropertiesService: { getScriptProperties: () => ({
      getProperties: () => ({ ...props }), getProperty: (k) => props[k] || null, setProperty: (k, v) => { props[k] = v; }
    }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    Session: { getActiveUser: () => ({ getEmail: () => 'sv@example.com' }) },
    Utilities: { formatDate: (d) => d.toISOString().slice(0, 16).replace('T', ' ') },
    Logger: { log() {} }
  });
  return { db, props };
};
