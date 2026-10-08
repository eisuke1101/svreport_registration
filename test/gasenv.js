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
    getNumRows: () => nr,
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
  const sh = { name, data: data || [], formulas: {}, maxRows: 1000, maxCols: 26, ss, gid: ++idSeq };
  sh.api = strict({
    getName: () => sh.name,
    getSheetId: () => sh.gid,
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
      getActiveSheet: () => (ss.active ? ss.sheets.find((s) => s.name === ss.active.sheet).api : ss.sheets[0].api),
      getActiveRangeList: () => {
        if (!ss.active) return null;
        const sh = ss.sheets.find((s) => s.name === ss.active.sheet);
        return strict({ getRanges: () => ss.active.ranges.map(([r, n]) => makeRange(sh, r, 1, n, 3)) }, 'RangeList');
      },
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
      getFilesByName: (n) => iter(Object.values(files).filter((x) => x.parents.includes(id) && !x.trashed && x.name === n).map((x) => fileApi(x.id))),
      getFoldersByName: (n) => iter(Object.values(folders).filter((x) => x.parent === id && x.name === n).map((x) => folderApi(x.id))),
      getFolders: () => iter(Object.values(folders).filter((x) => x.parent === id).map((x) => folderApi(x.id))),
      getFiles: () => iter(Object.values(files).filter((x) => x.parents.includes(id) && !x.trashed).map((x) => fileApi(x.id))),
      createFolder: (n) => { const nid = newId('folder'); folders[nid] = { id: nid, name: n, parent: id }; return folderApi(nid); },
      createFile: (blob) => {
        const nid = newId('file');
        files[nid] = { id: nid, name: blob.getName(), mime: blob.getContentType(), parents: [id], trashed: false, bytes: blob.bytes, source: blob.source };
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
  const triggers = [];
  const fetches = [];
  const sleeps = [];
  const ui = { alerts: [], prompts: [], answer: 'OK', promptText: '' };
  const fetchPlan = []; // テストで設定する応答コードの予定（空なら 200）
  function makeBlob(bytes, type, name, source) {
    const b = { bytes, type, name, source };
    const api = strict({
      getName: () => b.name, getContentType: () => b.type, bytes: b.bytes, source: b.source,
      setName: (n) => { b.name = n; return api; }
    }, 'Blob');
    return api;
  }
  const Button = { OK: 'OK', CANCEL: 'CANCEL' };
  const uiApi = strict({
    ButtonSet: { OK: 'OK', OK_CANCEL: 'OK_CANCEL' },
    Button,
    alert: (title, msg, buttons) => { ui.alerts.push({ title, msg, buttons }); return buttons === 'OK_CANCEL' ? ui.answer : 'OK'; },
    prompt: (title, msg) => {
      ui.prompts.push({ title, msg });
      return strict({ getSelectedButton: () => ui.answer, getResponseText: () => ui.promptText }, 'PromptResponse');
    },
    createMenu: (name) => {
      const menu = { name, items: [] };
      ui.menu = menu;
      const m = strict({
        addItem: (label, fn) => { menu.items.push({ label, fn }); return m; },
        addSeparator: () => m,
        addToUi: () => {}
      }, 'Menu');
      return m;
    }
  }, 'Ui');
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
      flush: () => {},
      getUi: () => uiApi,
      getActiveSpreadsheet: () => (ui.activeSpreadsheetId ? spreadsheets[ui.activeSpreadsheetId].api : null)
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
      newBlob: (bytes, type, name) => makeBlob(bytes, type, name),
      sleep: (ms) => { sleeps.push(ms); }
    }, 'Utilities'),
    MailApp: strict({ sendEmail: (m) => { if (!m.to || !m.subject) throw new Error('メール引数不正'); mails.push(m); } }, 'MailApp'),
    ScriptApp: strict({
      getProjectTriggers: () => triggers.map((t) => strict({ getHandlerFunction: () => t.fn, getTriggerSourceId: () => t.source }, 'Trigger')),
      newTrigger: (fn) => {
        const t = { fn };
        const b = strict({
          forSpreadsheet: (idOrSs) => { t.source = typeof idOrSs === 'string' ? idOrSs : idOrSs.getId(); return b; },
          onOpen: () => { t.type = 'open'; return b; },
          create: () => { if (!t.source || !t.type) throw new Error('トリガー設定不足'); triggers.push(t); return {}; }
        }, 'TriggerBuilder');
        return b;
      },
      getOAuthToken: () => 'token'
    }, 'ScriptApp'),
    UrlFetchApp: strict({
      fetch: (url, opts) => {
        const m = /^https:\/\/docs\.google\.com\/spreadsheets\/d\/([^/]+)\/export\?(.*)$/.exec(url);
        if (!m) throw new Error('想定外のURL: ' + url);
        if (!opts || !opts.headers || opts.headers.Authorization !== 'Bearer token') throw new Error('認証ヘッダなし');
        const q = Object.fromEntries(new URLSearchParams(m[2]));
        fetches.push({ id: m[1], q });
        const code = fetchPlan.length ? fetchPlan.shift() : 200;
        const ss = spreadsheets[m[1]];
        if (code === 200 && (!ss || files[m[1]].trashed)) throw new Error('存在しないファイルのエクスポート');
        if (code === 200 && !ss.sheets.some((s) => String(s.gid) === q.gid)) throw new Error('gid 不一致');
        return strict({
          getResponseCode: () => code,
          getBlob: () => makeBlob([37, 80, 68, 70], 'application/pdf', 'export.pdf', m[1])
        }, 'HTTPResponse');
      }
    }, 'UrlFetchApp'),
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

  return { g, props, files, folders, spreadsheets, mails, triggers, fetches, fetchPlan, sleeps, ui };
}

/** 貼り付け用の1ファイル（dist）または src/*.gs を、モック環境で読み込む */
function loadApp(env, source) {
  // Date はテスト側と同じものを共有する（別レルムだと instanceof Date が成り立たないため）
  const ctx = vm.createContext({ ...env.g, Date });
  vm.runInContext(source, ctx, { filename: 'bundle.gs' });
  return ctx;
}

module.exports = { createEnv, loadApp };
