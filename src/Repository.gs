/**
 * 管理用スプレッドシート（DB）の読み書き。
 */

/** シートごとの文字列として扱う列（数値・日付への自動変換を防ぐ） */
var TEXT_COLUMNS = {
  schools: ['校舎ID'],
  items: ['版', '項目キー', 'No', '対象外業態'],
  reports: ['レポートID', '校舎ID', '実施日', 'テンプレート版', 'ファイルID', '送信ID'],
  lines: ['レポートID', '校舎ID', '項目キー'],
  photos: ['レポートID', '項目キー', 'ファイルID'],
  importLog: ['ファイルID', 'レポートID']
};

var db_cache_ = null;

function db_() {
  return db_cache_ || (db_cache_ = SpreadsheetApp.openById(getConfig_().dbId));
}

function sheetOf_(key) {
  var sh = db_().getSheetByName(SHEETS[key]);
  if (!sh) throw new Error('シート「' + SHEETS[key] + '」がありません。setup() を実行してください。');
  return sh;
}

/** シートが無ければ作成し、見出しと文字列列の書式を設定する */
function ensureSheet_(ss, key) {
  var sh = ss.getSheetByName(SHEETS[key]) || ss.insertSheet(SHEETS[key]);
  var headers = HEADERS[key];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e8eaed');
  sh.setFrozenRows(1);
  (TEXT_COLUMNS[key] || []).forEach(function (h) {
    sh.getRange(1, headers.indexOf(h) + 1, sh.getMaxRows(), 1).setNumberFormat('@');
  });
  return sh;
}

function readObjects_(key) {
  var sh = sheetOf_(key);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var headers = HEADERS[key];
  return sh.getRange(2, 1, last - 1, headers.length).getValues().map(function (row, i) {
    var o = { _row: i + 2 };
    headers.forEach(function (h, j) { o[h] = row[j]; });
    return o;
  });
}

function toRow_(key, obj) {
  return HEADERS[key].map(function (h) { return obj[h] === undefined || obj[h] === null ? '' : obj[h]; });
}

function appendObjects_(key, objects) {
  if (!objects.length) return;
  var sh = sheetOf_(key);
  var rows = objects.map(function (o) { return toRow_(key, o); });
  var start = sh.getLastRow() + 1;
  var shortage = start + rows.length - 1 - sh.getMaxRows();
  if (shortage > 0) sh.insertRowsAfter(sh.getMaxRows(), shortage + 500); // 追加行は直前行の書式（文字列列）を引き継ぐ
  sh.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
}

/** 1列目が value に一致する行を削除する */
function deleteRowsByFirstColumn_(key, value, predicate) {
  var sh = sheetOf_(key);
  var last = sh.getLastRow();
  if (last < 2) return;
  var found = sh.getRange(2, 1, last - 1, 1).createTextFinder(String(value)).matchEntireCell(true).findAll();
  var rows = found.map(function (r) { return r.getRow(); });
  if (predicate) {
    var width = HEADERS[key].length;
    rows = rows.filter(function (r) { return predicate(sh.getRange(r, 1, 1, width).getValues()[0]); });
  }
  rows.sort(function (a, b) { return b - a; });
  // 連続する行はまとめて削除
  var i = 0;
  while (i < rows.length) {
    var end = rows[i];
    var start = end;
    while (i + 1 < rows.length && rows[i + 1] === start - 1) { i++; start--; }
    sh.deleteRows(start, end - start + 1);
    i++;
  }
}

function isTrue_(v) {
  return v === true || String(v).toUpperCase() === 'TRUE';
}

/* ---------- マスタ ---------- */

function getSchools_() {
  return readObjects_('schools')
    .filter(function (r) { return r['校舎ID'] !== '' && isTrue_(r['有効']); })
    .map(function (r) {
      return { id: String(r['校舎ID']), business: String(r['業態']), region: String(r['地域']), name: String(r['校舎名']) };
    });
}

function getSchool_(schoolId) {
  var hit = getSchools_().filter(function (s) { return s.id === String(schoolId); })[0];
  if (!hit) throw new Error('校舎ID ' + schoolId + ' が校舎マスタにありません');
  return hit;
}

function toItem_(r) {
  var num = function (v) { return v === '' || v === null || isNaN(Number(v)) ? null : Number(v); };
  return {
    key: String(r['項目キー']),
    no: r['No'] === '' ? '' : (isNaN(Number(r['No'])) ? String(r['No']) : Number(r['No'])),
    type: String(r['種別']) || '通常',
    category: String(r['区分']),
    text: String(r['項目']),
    penalties: { L: num(r['減点L']), M: num(r['減点M']), H: num(r['減点H']) },
    row: Number(r['出力行']),
    excludedBiz: String(r['対象外業態'] || '').split(/[,、，\s]+/).filter(String)
  };
}

/** 現行版の項目マスタ */
function getItems_() {
  var version = props_().getProperty('CURRENT_ITEM_VERSION');
  var rows = readObjects_('items').filter(function (r) { return isTrue_(r['有効']); });
  if (!version) throw new Error('CURRENT_ITEM_VERSION が未設定です。setup() を実行してください。');
  var items = rows.filter(function (r) { return String(r['版']) === version; }).map(toItem_);
  if (!items.length) throw new Error('項目マスタに版 ' + version + ' の項目がありません');
  return { version: version, items: items };
}

/* ---------- 登録履歴・明細 ---------- */

function toReport_(r) {
  return {
    reportId: String(r['レポートID']), year: Number(r['年度']), round: Number(r['実施回']),
    schoolId: String(r['校舎ID']), schoolName: String(r['校舎名']), business: String(r['業態']),
    date: formatDate_(r['実施日']), am: String(r['AM']), sv: String(r['SV']),
    cash: r['現金'], bank: r['通帳残'], tms: r['TMS金額'],
    total: Number(r['減点合計']) || 0, score: Number(r['点数']),
    version: String(r['テンプレート版']), fileId: String(r['ファイルID']), url: String(r['ファイルURL']),
    revision: Number(r['版数']) || 1, registeredBy: String(r['登録者']),
    createdAt: formatDate_(r['登録日時'], 'yyyy-MM-dd HH:mm'), updatedAt: formatDate_(r['更新日時'], 'yyyy-MM-dd HH:mm'),
    kind: String(r['区分']), submissionId: String(r['送信ID']), _row: r._row
  };
}

function findReport_(reportId) {
  var hit = readObjects_('reports').filter(function (r) { return String(r['レポートID']) === reportId; })[0];
  return hit ? toReport_(hit) : null;
}

/** 校舎のレポートを新しい順に返す */
function reportsOfSchool_(schoolId) {
  return readObjects_('reports')
    .filter(function (r) { return String(r['校舎ID']) === String(schoolId); })
    .map(toReport_)
    .sort(function (a, b) { return ordinalOf(b.year, b.round) - ordinalOf(a.year, a.round); });
}

/** 指定の実施回より前のレポート（新しい順、最大 n 件）。未実施の回は飛ばして直近を参照する */
function previousReports_(schoolId, year, round, n) {
  var current = ordinalOf(year, round);
  return reportsOfSchool_(schoolId)
    .filter(function (r) { return ordinalOf(r.year, r.round) < current; })
    .slice(0, n);
}

/** 明細 {項目キー: {level, base, count, applied, na, note}} */
function getLines_(reportId) {
  var sh = sheetOf_('lines');
  var last = sh.getLastRow();
  var result = {};
  if (last < 2) return result;
  var found = sh.getRange(2, 1, last - 1, 1).createTextFinder(reportId).matchEntireCell(true).findAll();
  if (!found.length) return result;
  var rows = found.map(function (r) { return r.getRow(); });
  var min = Math.min.apply(null, rows);
  var max = Math.max.apply(null, rows);
  sh.getRange(min, 1, max - min + 1, HEADERS.lines.length).getValues().forEach(function (row) {
    if (String(row[0]) !== reportId) return;
    result[String(row[4])] = {
      level: String(row[5]), base: Number(row[6]) || 0, count: Number(row[7]) || 0,
      applied: Number(row[8]) || 0, na: isTrue_(row[9]), note: String(row[10])
    };
  });
  return result;
}

/** 登録履歴（1行）と明細を保存する。同じレポートIDは置き換える */
function saveReport_(header, lines) {
  var existing = findReport_(header['レポートID']);
  var sh = sheetOf_('reports');
  if (existing) {
    sh.getRange(existing._row, 1, 1, HEADERS.reports.length).setValues([toRow_('reports', header)]);
  } else {
    appendObjects_('reports', [header]);
  }
  deleteRowsByFirstColumn_('lines', header['レポートID']);
  appendObjects_('lines', lines);
}

/* ---------- 写真 ---------- */

function getPhotos_(reportId) {
  return readObjects_('photos')
    .filter(function (r) { return String(r['レポートID']) === reportId; })
    .map(function (r) {
      return { itemKey: String(r['項目キー']), fileId: String(r['ファイルID']), name: String(r['ファイル名']), url: String(r['URL']) };
    });
}

function removePhotoRecords_(reportId, fileIds) {
  if (!fileIds.length) return;
  deleteRowsByFirstColumn_('photos', reportId, function (row) { return fileIds.indexOf(String(row[2])) >= 0; });
}

/* ---------- 取込ログ ---------- */

function logImport_(fileId, fileName, result, reportId, message) {
  appendObjects_('importLog', [{
    '日時': new Date(), 'ファイルID': fileId, 'ファイル名': fileName, '結果': result, 'レポートID': reportId || '', 'メッセージ': message || ''
  }]);
}

function processedImportFileIds_() {
  var done = {};
  readObjects_('importLog').forEach(function (r) {
    if (r['結果'] === 'OK' || r['結果'] === 'スキップ') done[String(r['ファイルID'])] = true;
  });
  return done;
}
