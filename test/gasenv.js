// GAS 実行環境のモック（SpreadsheetApp / DriveApp / Drive / PropertiesService など）。
// 未実装のメソッドを呼ぶと例外にして、GAS に存在しない呼び出しや想定外の使い方を検出する。
// 実際の GAS と同じく、範囲外・0行の getRange、サイズ不一致の setValues、最後のシートの削除などは例外にする。
const fs = require('fs');
const path = require('path');
const vm = require('vm');

function strict(obj, name) {
  return new Proxy(obj, {
    get(t, k) {
      if (k in t || typeof k === 'symbol' || k === 'then' || k === 'toJSON') return t[k];
      throw new Error(`モック未実装: ${name}.${String(k)}`);
    }
  });
}

function colNum(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

let idSeq = 0;
const newId = (p) => `${p}${++idSeq}`;

function makeRange(sh, r, c, nr, nc) {
  if (!(r >= 1 && c >= 1)) throw new Error(`範囲の開始位置が不正です (${r},${c})`);
  if (!(nr >= 1 && nc >= 1)) throw new Error('The number of rows in the range must be at least 1.');
  if (r + nr - 1 > sh.maxRows || c + nc - 1 > sh.maxCols) {
    throw new Error(`The coordinates of the range are outside the dimensions of the sheet. (${sh.name} R${r}C${c} ${nr}x${nc})`);
  }
  const cellGet = (i, j) => {
    const row = sh.data[r - 1 + i];
    const v = row ? row[c - 1 + j] : undefined;
    return v === undefined || v === null ? '' : v;
  };
  const cellSet = (i, j, v) => {
    const t = (sh.data[r - 1 + i] = sh.data[r - 1 + i] || []);
    if (typeof v === 'string' && v.startsWith('=')) {
      sh.formulas[`${r + i},${c + j}`] = v;
      t[c - 1 + j] = '#FORMULA';
    } else {
      delete sh.formulas[`${r + i},${c + j}`];
      t[c - 1 + j] = v;
    }
  };
  const self = {
    getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => cellGet(i, j))),
    getFormulas: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => sh.formulas[`${r + i},${c + j}`] || '')),
    setValues: (v) => {
      if (!Array.isArray(v) || v.length !== nr || v.some((row) => !Array.isArray(row) || row.length !== nc)) {
        throw new Error(`The number of rows/columns in the data does not match the range (${nr}x${nc})`);
      }
      v.forEach((row, i) => row.forEach((x, j) => {
        if (x === undefined || (typeof x === 'object' && x !== null && !(x instanceof Date))) {
          throw new Error(`setValues に不正な値: ${JSON.stringify(x)}`);
        }
        cellSet(i, j, x);
      }));
      return range;
    },
    setValue: (x) => {
      if (nr !== 1 || nc !== 1) throw new Error('setValue は単一セルで使う想定');
      if (x === undefined) throw new Error('setValue に undefined');
      cellSet(0, 0, x);
      return range;
    },
    getRow: () => r,
    setFontWeight: () => range,
    setBackground: () => range,
    setNumberFormat: () => range,
    createTextFinder: (text) => {
      const finder = strict({
        matchEntireCell: () => finder,
        findAll: () => {
          const res = [];
          for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) {
            if (String(cellGet(i, j)) === String(text)) res.push(makeRange(sh, r + i, c + j, 1, 1));
          }
          return res;
        }
      }, 'TextFinder');
      return finder;
    }
  };
  const range = strict(self, 'Range');
  return range;
}

function makeSheet(ss, name, data) {
  const sh = { name, data: data || [], formulas: {}, maxRows: 1000, maxCols: 26, ss };
  sh.api = strict({
    getName: () => sh.name,
    getRange: (a, b, c, d) => {
      if (typeof a === 'string') {
        const m = /^([A-Z]+)(\d+)$/.exec(a);
        if (!m) throw new Error('A1表記が不正: ' + a);
        return makeRange(sh, Number(m[2]), colNum(m[1]), 1, 1);
      }
      return makeRange(sh, a, b, c === undefined ? 1 : c, d === undefined ? 1 : d);
    },
    getDataRange: () => makeRange(sh, 1, 1, Math.max(1, sh.api.getLastRow()), Math.max(1, sh.data.reduce((m, row) => Math.max(m, (row || []).length), 1))),
    getLastRow: () => {
      let n = sh.data.length;
      while (n > 0 && !(sh.data[n - 1] || []).some((x) => x !== '' && x !== undefined && x !== null)) n--;
      return n;
    },
    getMaxRows: () => sh.maxRows,
    insertRowsAfter: (after, n) => { if (after > sh.maxRows) throw new Error('行位置が不正'); sh.maxRows += n; },
    deleteRows: (start, n) => {
      if (start < 1 || start + n - 1 > sh.maxRows) throw new Error('deleteRows 範囲外');
      sh.data.splice(start - 1, n);
      sh.maxRows -= n;
    },
    setFrozenRows: () => {},
    setColumnWidth: () => {},
    clear: () => { sh.data = []; sh.formulas = {}; }
  }, 'Sheet');
  return sh;
}

function createEnv(opts) {
  const props = {};
  const files = {}; // id -> {id, name, mime, parents:[folderId], trashed, ss?, blob?}
  const folders = {}; // id -> {id, name, parent}
  const mails = [];
  const spreadsheets = {}; // id -> ss

  function makeSpreadsheet(id) {
    const ss = { id, sheets: [] };
    ss.api = strict({
      getId: () => id,
      getUrl: () => `https://docs.google.com/spreadsheets/d/${id}`,
      getSheetByName: (n) => { const s = ss.sheets.find((x) => x.name === n); return s ? s.api : null; },
      getSheets: () => ss.sheets.map((s) => s.api),
      insertSheet: (n) => {
        if (ss.sheets.some((x) => x.name === n)) throw new Error(`シート名「${n}」は既に存在します`);
        const s = makeSheet(ss, n);
        ss.sheets.push(s);
        return s.api;
      },
      deleteSheet: (shApi) => {
        if (ss.sheets.length <= 1) throw new Error('最後のシートは削除できません');
        ss.sheets = ss.sheets.filter((s) => s.api !== shApi);
      }
    }, 'Spreadsheet');
    spreadsheets[id] = ss;
    return ss;
  }
  function cloneSpreadsheet(src, id) {
    const ss = makeSpreadsheet(id);
    src.sheets.forEach((s) => {
      const c = makeSheet(ss, s.name, JSON.parse(JSON.stringify(s.data)));
      c.formulas = { ...s.formulas };
      c.maxRows = s.maxRows;
      ss.sheets.push(c);
    });
    return ss;
  }

  function folderApi(id) {
    const f = folders[id];
    if (!f) throw new Error('フォルダが見つかりません: ' + id);
    const iter = (list) => { let i = 0; return strict({ hasNext: () => i < list.length, next: () => list[i++] }, 'Iterator'); };
    return strict({
      getId: () => id,
      getName: () => f.name,
      getFoldersByName: (n) => iter(Object.values(folders).filter((x) => x.parent === id && x.name === n).map((x) => folderApi(x.id))),
      getFolders: () => iter(Object.values(folders).filter((x) => x.parent === id).map((x) => folderApi(x.id))),
      getFiles: () => iter(Object.values(files).filter((x) => x.parents.includes(id) && !x.trashed).map((x) => fileApi(x.id))),
      createFolder: (n) => { const nid = newId('folder'); folders[nid] = { id: nid, name: n, parent: id }; return folderApi(nid); },
      createFile: (blob) => {
        const nid = newId('file');
        files[nid] = { id: nid, name: blob.getName(), mime: blob.getContentType(), parents: [id], trashed: false, bytes: blob.bytes };
        return fileApi(nid);
      }
    }, 'Folder');
  }
  function fileApi(id) {
    const f = files[id];
    if (!f) throw new Error('ファイルが見つかりません: ' + id);
    return strict({
      getId: () => id,
      getName: () => f.name,
      setName: (n) => { f.name = n; },
      getMimeType: () => f.mime,
      getUrl: () => `https://drive.google.com/file/d/${id}`,
      isTrashed: () => f.trashed,
      setTrashed: (v) => { f.trashed = v; },
      moveTo: (folder) => { f.parents = [folder.getId()]; },
      makeCopy: (name, folder) => {
        const nid = newId('ss');
        files[nid] = { id: nid, name, mime: f.mime, parents: [folder.getId()], trashed: false };
        if (spreadsheets[id]) cloneSpreadsheet(spreadsheets[id], nid);
        return fileApi(nid);
      }
    }, 'File');
  }

  const MimeType = { GOOGLE_SHEETS: 'application/vnd.google-apps.spreadsheet' };
  const g = {
    MimeType,
    PropertiesService: strict({
      getScriptProperties: () => strict({
        getProperties: () => ({ ...props }),
        getProperty: (k) => (k in props ? props[k] : null),
        setProperty: (k, v) => { props[k] = String(v); }
      }, 'Properties')
    }, 'PropertiesService'),
    SpreadsheetApp: strict({
      openById: (id) => { if (!spreadsheets[id] || files[id].trashed) throw new Error('スプレッドシートが見つかりません: ' + id); return spreadsheets[id].api; },
      create: (name) => {
        const id = newId('ss');
        files[id] = { id, name, mime: MimeType.GOOGLE_SHEETS, parents: ['root'], trashed: false };
        const ss = makeSpreadsheet(id);
        ss.sheets.push(makeSheet(ss, 'シート1'));
        return ss.api;
      },
      flush: () => {}
    }, 'SpreadsheetApp'),
    DriveApp: strict({ getFolderById: folderApi, getFileById: fileApi }, 'DriveApp'),
    Drive: strict({
      Files: strict({
        copy: (resource, fileId, options) => {
          const src = files[fileId];
          if (!src) throw new Error('ファイルが見つかりません: ' + fileId);
          if (!resource || !resource.name || resource.mimeType !== MimeType.GOOGLE_SHEETS || !Array.isArray(resource.parents)) {
            throw new Error('Drive.Files.copy の引数が不正: ' + JSON.stringify(resource));
          }
          if (!options || options.supportsAllDrives !== true) throw new Error('supportsAllDrives 未指定');
          const id = newId('ss');
          files[id] = { id, name: resource.name, mime: MimeType.GOOGLE_SHEETS, parents: resource.parents, trashed: false };
          cloneSpreadsheet(src.excelContent, id);
          return { id };
        }
      }, 'Drive.Files')
    }, 'Drive'),
    Utilities: strict({
      formatDate: (d, tz, pattern) => {
        if (!(d instanceof Date)) throw new Error('formatDate に Date 以外');
        const p = (n) => String(n).padStart(2, '0');
        return pattern.replace('yyyy', d.getFullYear()).replace('MM', p(d.getMonth() + 1)).replace('dd', p(d.getDate()))
          .replace('HH', p(d.getHours())).replace('mm', p(d.getMinutes())).replace('ss', p(d.getSeconds()));
      },
      base64Decode: (s) => Array.from(Buffer.from(s, 'base64')),
      newBlob: (bytes, type, name) => strict({ getName: () => name, getContentType: () => type, bytes }, 'Blob')
    }, 'Utilities'),
    MailApp: strict({ sendEmail: (m) => { if (!m.to || !m.subject) throw new Error('メール引数不正'); mails.push(m); } }, 'MailApp'),
    LockService: strict({ getScriptLock: () => strict({ tryLock: () => true, releaseLock: () => {} }, 'Lock') }, 'LockService'),
    Session: strict({ getActiveUser: () => strict({ getEmail: () => 'sv@example.com' }, 'User') }, 'Session'),
    Logger: strict({ log: () => {} }, 'Logger')
  };

  // ルートフォルダと、アップロード済みのテンプレートExcel
  folders.root = { id: 'root', name: 'マイドライブ', parent: null };
  folders[opts.rootFolderId] = { id: opts.rootFolderId, name: 'SVレポート', parent: 'root' };
  const excel = { sheets: [] };
  const check = makeSheet(excel, 'チェック項目', JSON.parse(JSON.stringify(opts.templateValues)));
  Object.assign(check.formulas, opts.templateFormulas || {});
  excel.sheets.push(check);
  excel.sheets.push(makeSheet(excel, '校舎一覧', JSON.parse(JSON.stringify(opts.schoolList))));
  files.xls = { id: 'xls', name: 'SVレポート_2026.10ver.xls', mime: 'application/vnd.ms-excel', parents: ['root'], trashed: false, excelContent: excel };

  return { g, props, files, folders, spreadsheets, mails };
}

/** 貼り付け用の1ファイル（dist）または src/*.gs を、モック環境で読み込む */
function loadApp(env, source) {
  // Date はテスト側と同じものを共有する（別レルムだと instanceof Date が成り立たないため）
  const ctx = vm.createContext({ ...env.g, Date });
  vm.runInContext(source, ctx, { filename: 'bundle.gs' });
  return ctx;
}

module.exports = { createEnv, loadApp };
