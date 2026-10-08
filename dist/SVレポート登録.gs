// 自動生成ファイル（npm run bundle）。src/*.gs を編集すること。
// 初期設定: 上部のプルダウンで「setup」を選んで ▷実行
// ===== Setup.gs =====
/**
 * 初期設定（スクリプトエディタから手動実行）。
 *
 * 事前準備:
 *   1. テンプレートExcel（SVレポート_2026.10ver.xls）を Drive にアップロードし、ファイルIDを
 *      スクリプトプロパティ TEMPLATE_XLS_FILE_ID に設定する。
 *   2. （任意）ROOT_FOLDER_ID / NOTIFY_TO を設定する。ROOT_FOLDER_ID 未設定時は既定のフォルダ。
 *
 * setup() が行うこと（何度実行しても安全）:
 *   - 保存先フォルダに「_システム」フォルダを作成
 *   - テンプレートを Google スプレッドシートに変換（TEMPLATE_SPREADSHEET_ID）
 *   - 管理用スプレッドシートを作成し各シートを用意（DB_SPREADSHEET_ID）
 *   - テンプレートから校舎マスタ・項目マスタを作成（未作成の場合のみ）
 */
function setup() {
  var p = props_();
  var rootId = p.getProperty('ROOT_FOLDER_ID') || DEFAULT_ROOT_FOLDER_ID;
  p.setProperty('ROOT_FOLDER_ID', rootId);
  var sys = systemFolder_();

  var templateId = p.getProperty('TEMPLATE_SPREADSHEET_ID');
  if (!templateId) {
    var xlsId = p.getProperty('TEMPLATE_XLS_FILE_ID');
    if (!xlsId) {
      throw new Error('スクリプトプロパティ TEMPLATE_XLS_FILE_ID に、Drive にアップロードしたテンプレート（.xls）のファイルIDを設定してください。');
    }
    templateId = convertToSpreadsheet_(xlsId, 'SVレポート_テンプレート', sys.getId());
    p.setProperty('TEMPLATE_SPREADSHEET_ID', templateId);
  }

  var dbId = p.getProperty('DB_SPREADSHEET_ID');
  var ss;
  if (dbId) {
    ss = SpreadsheetApp.openById(dbId);
  } else {
    ss = SpreadsheetApp.create('SVレポート_管理DB');
    DriveApp.getFileById(ss.getId()).moveTo(sys);
    p.setProperty('DB_SPREADSHEET_ID', ss.getId());
  }
  Object.keys(SHEETS).forEach(function (key) { ensureSheet_(ss, key); });
  ss.getSheets().forEach(function (sh) {
    var known = Object.keys(SHEETS).some(function (k) { return SHEETS[k] === sh.getName(); });
    if (!known && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
  db_cache_ = ss;

  importMastersFromTemplate_(templateId);
  Logger.log('セットアップ完了\n管理DB: %s\nテンプレート: %s', ss.getUrl(), SpreadsheetApp.openById(templateId).getUrl());
}

/**
 * テンプレートの改版時に実行する。
 * TEMPLATE_XLS_FILE_ID を新しいExcelのIDに変更してから実行すると、テンプレートを差し替え、
 * 新しい版の項目を項目マスタに追加して現行版を切り替える。
 * 注意: 項目キーは No から自動採番する。改版で No がずれた場合は、前回・前々回と正しく
 * 紐付くよう、項目マスタの新しい版の「項目キー」を旧版に合わせて修正すること。
 */
function updateTemplate() {
  var p = props_();
  var xlsId = p.getProperty('TEMPLATE_XLS_FILE_ID');
  if (!xlsId) throw new Error('TEMPLATE_XLS_FILE_ID を設定してください');
  var templateId = convertToSpreadsheet_(xlsId, 'SVレポート_テンプレート_' + formatDate_(new Date(), 'yyyyMMdd'), systemFolder_().getId());
  p.setProperty('TEMPLATE_SPREADSHEET_ID', templateId);
  importMastersFromTemplate_(templateId);
}

function systemFolder_() {
  var root = DriveApp.getFolderById(props_().getProperty('ROOT_FOLDER_ID') || DEFAULT_ROOT_FOLDER_ID);
  return getOrCreateFolder_(root, '_システム');
}

/** Excel を Google スプレッドシートに変換（既にスプレッドシートならコピー）して新しいファイルIDを返す */
function convertToSpreadsheet_(fileId, name, parentId) {
  var src = DriveApp.getFileById(fileId);
  if (src.getMimeType() === MimeType.GOOGLE_SHEETS) {
    return src.makeCopy(name, DriveApp.getFolderById(parentId)).getId();
  }
  var copied = Drive.Files.copy(
    { name: name, mimeType: MimeType.GOOGLE_SHEETS, parents: [parentId] },
    fileId,
    { supportsAllDrives: true }
  );
  return copied.id;
}

function importMastersFromTemplate_(templateId) {
  var tpl = SpreadsheetApp.openById(templateId);
  var sh = tpl.getSheetByName(REPORT_SHEET_NAME) || tpl.getSheets()[0];
  var values = sh.getRange(1, 1, sh.getLastRow(), 15).getValues();
  var version = parseVersion(values);
  if (!version) throw new Error('テンプレートのO1セルから版を読み取れません');

  // 項目マスタ（その版が未登録の場合のみ追加）
  var exists = readObjects_('items').some(function (r) { return String(r['版']) === version; });
  if (!exists) {
    appendObjects_('items', parseTemplateItems(values).map(function (i) {
      return {
        '版': version, '項目キー': i.key, 'No': String(i.no), '種別': i.type, '区分': i.category, '項目': i.text,
        '減点L': i.penalties.L === null ? '' : i.penalties.L,
        '減点M': i.penalties.M === null ? '' : i.penalties.M,
        '減点H': i.penalties.H === null ? '' : i.penalties.H,
        '出力行': i.row, '対象外業態': guessExcludedBusinesses(i.text).join(','), '有効': true
      };
    }));
  }
  props_().setProperty('CURRENT_ITEM_VERSION', version);

  // 校舎マスタ（空の場合のみテンプレートの校舎一覧から作成）
  var schoolSheet = tpl.getSheetByName(SCHOOL_LIST_SHEET_NAME);
  if (schoolSheet && readObjects_('schools').length === 0) {
    var list = parseSchoolList(schoolSheet.getDataRange().getValues());
    appendObjects_('schools', list.map(function (s) {
      return {
        '校舎ID': s.id, '業態': s.business, '地域': s.region, '校舎名': s.name, '経営': s.management,
        '会社': s.company, '開校日': s.openDate, '備考': s.remarks, '有効': true
      };
    }));
  }
}

// ===== Import.gs =====
/**
 * 過去データの取込（スクリプトエディタから手動実行）。
 *
 * スクリプトプロパティ IMPORT_FOLDER_ID のフォルダ（サブフォルダ含む）にある過去のSVレポート
 * （.xls / .xlsx / Google スプレッドシート）を読み取り、登録履歴・明細に「取込」として登録する。
 * - 校舎ID（E2）と実施日（G2）から年度・実施回を判定。同じ実施回が登録済みならスキップ
 * - 項目は項目文で現行の項目マスタと照合（版が違っても文言が同じなら紐付く）
 * - L列（今回の適用減点）と K列（連続指摘回数）をそのまま取り込む
 * - 実行時間の上限（6分）に近づいたら中断する。未処理が残った場合は再実行すると続きから処理する
 * 結果は管理DBの「取込ログ」シートに記録される。
 */
function importPastReports() {
  var folderId = props_().getProperty('IMPORT_FOLDER_ID');
  if (!folderId) throw new Error('スクリプトプロパティ IMPORT_FOLDER_ID に取込元フォルダのIDを設定してください');
  var started = Date.now();
  var done = processedImportFileIds_();
  var master = getItems_();
  var schools = {};
  readObjects_('schools').forEach(function (r) {
    schools[String(r['校舎ID'])] = { id: String(r['校舎ID']), name: String(r['校舎名']), business: String(r['業態']) };
  });
  var workFolder = getOrCreateFolder_(systemFolder_(), '_取込作業');
  var count = { ok: 0, skip: 0, error: 0, remaining: 0 };

  listSpreadsheetFiles_(DriveApp.getFolderById(folderId)).forEach(function (file) {
    if (done[file.getId()]) return;
    if (Date.now() - started > 4.5 * 60 * 1000) {
      count.remaining++;
      return;
    }
    try {
      var res = importOne_(file, master, schools, workFolder);
      logImport_(file.getId(), file.getName(), res.status, res.reportId, res.message);
      count[res.status === 'OK' ? 'ok' : 'skip']++;
    } catch (e) {
      logImport_(file.getId(), file.getName(), 'エラー', '', e.message);
      count.error++;
    }
  });
  Logger.log('取込: 成功 %s / スキップ %s / エラー %s / 未処理 %s', count.ok, count.skip, count.error, count.remaining);
  if (count.remaining) Logger.log('未処理のファイルがあります。importPastReports() を再実行してください。');
}

var IMPORTABLE_MIME_TYPES = [
  'application/vnd.google-apps.spreadsheet',
  'application/vnd.ms-excel',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
];

function listSpreadsheetFiles_(folder) {
  var result = [];
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    if (IMPORTABLE_MIME_TYPES.indexOf(f.getMimeType()) >= 0) result.push(f);
  }
  var subs = folder.getFolders();
  while (subs.hasNext()) result = result.concat(listSpreadsheetFiles_(subs.next()));
  return result;
}

function toDateString_(v) {
  if (v instanceof Date) return formatDate_(v);
  var m = /(\d{4})[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})/.exec(String(v || ''));
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : '';
}

function importOne_(file, master, schools, workFolder) {
  var tempId = null;
  var ssId = file.getId();
  if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) {
    tempId = convertToSpreadsheet_(file.getId(), '取込_' + file.getName(), workFolder.getId());
    ssId = tempId;
  }
  try {
    var ss = SpreadsheetApp.openById(ssId);
    var sh = ss.getSheetByName(REPORT_SHEET_NAME) || ss.getSheets()[0];
    var parsed = parseReportSheet(sh.getRange(1, 1, sh.getLastRow(), 15).getValues());
    var h = parsed.header;
    if (!h.schoolId) throw new Error('校舎ID（E2）が空です');
    var dateStr = toDateString_(h.date);
    if (!dateStr) throw new Error('実施日（G2）が読み取れません');

    var period = periodOf(dateStr);
    var school = schools[h.schoolId] || { id: h.schoolId, name: h.schoolName, business: h.business };
    var reportId = reportIdOf(period.year, period.round, school.id);
    if (findReport_(reportId)) return { status: 'スキップ', reportId: reportId, message: '登録済み' };

    // 項目の照合（全文一致 → 先頭一致）
    var byText = {};
    master.items.forEach(function (i) { byText[normalizeItemText(i.text)] = i; });
    var mapped = {};
    var unmatched = [];
    parsed.items.forEach(function (pi) {
      var norm = normalizeItemText(pi.text);
      var mi = byText[norm] || master.items.filter(function (i) {
        var n = normalizeItemText(i.text);
        return n.indexOf(norm.slice(0, 12)) === 0 || norm.indexOf(n.slice(0, 12)) === 0;
      })[0];
      if (!mi || mapped[mi.key]) {
        if (pi.applied < 0) unmatched.push('No.' + pi.no + '(' + pi.applied + ')');
        return;
      }
      mapped[mi.key] = {
        level: inferLevel(pi, pi.applied, pi.count),
        base: pi.applied < 0 ? pi.applied / Math.min(pi.count || 1, MAX_MULTIPLIER) : 0,
        count: pi.count, applied: pi.applied, na: /^対象外/.test(pi.note), note: pi.note
      };
    });

    var total = parsed.items.reduce(function (s, pi) { return s + (pi.applied < 0 ? pi.applied : 0); }, 0);
    var now = new Date();
    saveReport_({
      'レポートID': reportId, '年度': period.year, '実施回': period.round, '校舎ID': school.id,
      '校舎名': school.name, '業態': school.business, '実施日': dateStr, 'AM': h.am, 'SV': h.sv,
      '現金': h.cash, '通帳残': h.bank, 'TMS金額': h.tms, '減点合計': total, '点数': 100 + total,
      'テンプレート版': h.version, 'ファイルID': file.getId(), 'ファイルURL': file.getUrl(), '版数': 1,
      '登録者': currentUserEmail_(), '登録日時': now, '更新日時': now, '区分': '取込', '送信ID': ''
    }, master.items.map(function (i) {
      var l = mapped[i.key] || { level: '', base: 0, count: 0, applied: 0, na: false, note: '' };
      return {
        'レポートID': reportId, '校舎ID': school.id, '年度': period.year, '実施回': period.round,
        '項目キー': i.key, 'レベル': l.level, '基本減点': l.base, '連続回数': l.count,
        '適用減点': l.applied, '対象外': l.na, '備考': l.note
      };
    }));
    return {
      status: 'OK', reportId: reportId,
      message: unmatched.length ? '項目マスタと照合できない減点あり（合計点には含む）: ' + unmatched.join(' ') : ''
    };
  } finally {
    if (tempId) DriveApp.getFileById(tempId).setTrashed(true);
  }
}

// ===== Config.gs =====
/**
 * 設定。値はスクリプトプロパティに保存する（setup() が自動設定）。
 *   ROOT_FOLDER_ID           保存先フォルダ（この下に「{年度}年度/第{回}回」を作成）
 *   TEMPLATE_XLS_FILE_ID     元のExcelテンプレート（setup() の入力。Driveにアップロードしたもの）
 *   TEMPLATE_SPREADSHEET_ID  変換後のテンプレート（setup() が作成）
 *   DB_SPREADSHEET_ID        管理用スプレッドシート（setup() が作成）
 *   NOTIFY_TO                登録通知の送信先メールアドレス
 */
var DEFAULT_ROOT_FOLDER_ID = '1fCaEEFJMTKgGPPINg7G5TuBNuohbwymo';
var TIME_ZONE = 'Asia/Tokyo';
var REPORT_SHEET_NAME = 'チェック項目';
var SCHOOL_LIST_SHEET_NAME = '校舎一覧';
var PHOTO_SHEET_NAME = '写真';

var SHEETS = {
  schools: '校舎マスタ',
  items: '項目マスタ',
  reports: '登録履歴',
  lines: '明細',
  photos: '写真',
  importLog: '取込ログ'
};

var HEADERS = {
  schools: ['校舎ID', '業態', '地域', '校舎名', '経営', '会社', '開校日', '備考', '有効'],
  items: ['版', '項目キー', 'No', '種別', '区分', '項目', '減点L', '減点M', '減点H', '出力行', '対象外業態', '有効'],
  reports: ['レポートID', '年度', '実施回', '校舎ID', '校舎名', '業態', '実施日', 'AM', 'SV', '現金', '通帳残', 'TMS金額',
    '減点合計', '点数', 'テンプレート版', 'ファイルID', 'ファイルURL', '版数', '登録者', '登録日時', '更新日時', '区分', '送信ID'],
  lines: ['レポートID', '校舎ID', '年度', '実施回', '項目キー', 'レベル', '基本減点', '連続回数', '適用減点', '対象外', '備考'],
  photos: ['レポートID', '項目キー', 'ファイルID', 'ファイル名', 'URL', '登録日時'],
  importLog: ['日時', 'ファイルID', 'ファイル名', '結果', 'レポートID', 'メッセージ']
};

function props_() {
  return PropertiesService.getScriptProperties();
}

function getConfig_() {
  var p = props_().getProperties();
  var config = {
    rootFolderId: p.ROOT_FOLDER_ID || DEFAULT_ROOT_FOLDER_ID,
    templateId: p.TEMPLATE_SPREADSHEET_ID,
    dbId: p.DB_SPREADSHEET_ID,
    notifyTo: p.NOTIFY_TO || ''
  };
  if (!config.templateId || !config.dbId) {
    throw new Error('初期設定が未完了です。スクリプトエディタで setup() を実行してください。');
  }
  return config;
}

function formatDate_(d, pattern) {
  if (!d) return '';
  if (!(d instanceof Date)) return String(d);
  return Utilities.formatDate(d, TIME_ZONE, pattern || 'yyyy-MM-dd');
}

function parseDate_(dateStr) {
  var m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(String(dateStr));
  if (!m) throw new Error('日付の形式が不正です: ' + dateStr);
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

// ===== Domain.gs =====
/**
 * 業務ルール（純粋関数）。
 * サーバ側の確定計算と、画面側のリアルタイム計算（sharedSource() で注入）の両方で使う。
 * GAS のサービス（SpreadsheetApp 等）をここで呼ばないこと。
 */

/** 連続指摘の倍率上限（5回目以降は4回目と同じ） */
var MAX_MULTIPLIER = 4;

/** 重要事項の「該当」を表すレベル値 */
var LEVEL_IMPORTANT = 'X';

/**
 * 実施日（yyyy-mm-dd）から年度・実施回を求める。
 * 第1回: 1〜3月、第2回: 4〜6月、第3回: 7〜9月、第4回: 10〜12月（年度＝暦年）。
 */
function periodOf(dateStr) {
  var m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(String(dateStr || ''));
  if (!m) throw new Error('実施日の形式が不正です: ' + dateStr);
  var year = Number(m[1]);
  var month = Number(m[2]);
  if (month < 1 || month > 12) throw new Error('実施日の月が不正です: ' + dateStr);
  return { year: year, round: Math.floor((month - 1) / 3) + 1 };
}

/** レポートID = {年度}-{実施回}-{校舎ID} */
function reportIdOf(year, round, schoolId) {
  return year + '-' + round + '-' + String(schoolId).trim();
}

/** 実施回の前後比較用の序数（年度をまたいで比較できる） */
function ordinalOf(year, round) {
  return Number(year) * 10 + Number(round);
}

/** 項目がその業態で対象外か */
function isExcludedFor(item, business) {
  return !!business && (item.excludedBiz || []).indexOf(String(business)) >= 0;
}

/** 項目で選択できるレベルの一覧 [{level, penalty}] */
function levelsOf(item) {
  if (item.type === '重要') {
    return item.penalties.L != null ? [{ level: LEVEL_IMPORTANT, penalty: item.penalties.L }] : [];
  }
  return ['L', 'M', 'H']
    .filter(function (lv) { return item.penalties[lv] != null; })
    .map(function (lv) { return { level: lv, penalty: item.penalties[lv] }; });
}

/** レベルに対応する基本減点。選択不可なら null */
function basePenaltyOf(item, level) {
  var hit = levelsOf(item).filter(function (x) { return x.level === level; })[0];
  return hit ? hit.penalty : null;
}

/**
 * 1項目の今回結果を計算する。
 * - 今回減点あり かつ 前回（直近の実施済みレポート）でも減点あり → 連続回数 = 前回の連続回数 + 1
 *   （レベル違いでも連続。間の回が未実施でも直近の実施済みレポートを参照するためリセットされない）
 * - 適用減点 = 基本減点 × min(連続回数, 4)
 * @param {Object} item 項目マスタ
 * @param {Object} input {level, na, note}
 * @param {Object} prevLine 前回の明細 {applied, count}（無ければ null）
 * @param {string} business 校舎の業態
 */
function computeLine(item, input, prevLine, business) {
  input = input || {};
  var excluded = isExcludedFor(item, business);
  var line = {
    key: item.key, level: '', base: 0, count: 0, applied: 0,
    na: excluded || !!input.na, forcedNa: excluded,
    note: input.note ? String(input.note) : ''
  };
  if (line.na || !input.level) return line;

  var base = basePenaltyOf(item, input.level);
  if (base == null) {
    throw new Error('項目 ' + item.no + ' ではレベル「' + input.level + '」を選択できません');
  }
  if (!(base < 0)) return line;

  var prevApplied = prevLine ? Number(prevLine.applied) || 0 : 0;
  var prevCount = prevApplied < 0 ? (Number(prevLine.count) || 1) : 0;
  line.level = input.level;
  line.base = base;
  line.count = prevCount + 1;
  line.applied = base * Math.min(line.count, MAX_MULTIPLIER);
  return line;
}

/**
 * レポート全体を計算する。
 * @param {Object[]} items 項目マスタ（並び順どおり）
 * @param {Object} inputs {項目キー: {level, na, note}}
 * @param {Object} prevLines {項目キー: {applied, count}}（前回レポートの明細）
 * @param {string} business 校舎の業態
 */
function computeReport(items, inputs, prevLines, business) {
  inputs = inputs || {};
  prevLines = prevLines || {};
  var lines = items.map(function (item) {
    return computeLine(item, inputs[item.key], prevLines[item.key], business);
  });
  var total = lines.reduce(function (sum, l) { return sum + l.applied; }, 0);
  return { lines: lines, total: total, score: 100 + total };
}

// ===== Parser.gs =====
/**
 * SVレポート（チェック項目シート）の解析（純粋関数）。
 * values は sheet.getRange('A1:O{最終行}').getValues() の2次元配列。
 */

var SHEET_LAYOUT = {
  firstItemRow: 6,     // 重要事項の先頭行
  col: { no: 1, category: 2, text: 3, penaltyL: 8, penaltyM: 9, penaltyH: 10, note: 6, count: 11, current: 12, prev: 14, prev2: 15 },
  cell: {
    version: 'O1', business: 'B2', schoolName: 'C2', schoolId: 'E2', date: 'G2',
    am: 'M3', sv: 'M5', cash: 'E8', bank: 'E9', tms: 'G10'
  }
};

function cellAt_(values, a1) {
  var m = /^([A-Z]+)(\d+)$/.exec(a1);
  var col = 0;
  for (var i = 0; i < m[1].length; i++) col = col * 26 + (m[1].charCodeAt(i) - 64);
  var row = values[Number(m[2]) - 1];
  return row ? row[col - 1] : '';
}

function penaltyValue_(v) {
  if (v === '' || v === null || v === undefined || v === '*') return null;
  var n = Number(v);
  return isNaN(n) ? null : n;
}

/** テンプレート版（例: '2026.10ver.' → '2026.10'） */
function parseVersion(values) {
  var v = String(cellAt_(values, SHEET_LAYOUT.cell.version) || '').trim();
  return v.replace(/ver\.?$/i, '').trim();
}

/** 項目文の比較用正規化（過去ファイル取込時の照合に使う） */
function normalizeItemText(s) {
  return String(s || '').replace(/[\s　■※・、。，．,.（）()]/g, '').toLowerCase();
}

/**
 * チェック項目を抽出する。
 * - A列に番号がある行 → 通常項目（No, 区分, 項目, L/M/H減点）
 * - 番号が無く C列が「■」で始まり H列に減点がある行（6〜8行）→ 重要事項
 * - C列が空白・全角空白で始まる行 → 直前の項目文の続き
 * H列が「減点total」の行で終了。
 * @return {Object[]} [{key, no, type, category, text, penalties:{L,M,H}, row}]
 */
function parseTemplateItems(values) {
  var c = SHEET_LAYOUT.col;
  var items = [];
  var category = '';
  var importantNo = 0;
  for (var r = SHEET_LAYOUT.firstItemRow; r <= values.length; r++) {
    var row = values[r - 1];
    if (String(row[c.penaltyL - 1]).indexOf('減点total') >= 0) break;
    if (row[c.category - 1]) category = String(row[c.category - 1]).trim();
    var text = String(row[c.text - 1] || '');
    var no = row[c.no - 1];

    if (no !== '' && no !== null && !isNaN(Number(no))) {
      items.push({
        key: 'I' + ('0' + Number(no)).slice(-2), no: Number(no), type: '通常', category: category,
        text: text.trim(), row: r,
        penalties: { L: penaltyValue_(row[c.penaltyL - 1]), M: penaltyValue_(row[c.penaltyM - 1]), H: penaltyValue_(row[c.penaltyH - 1]) }
      });
    } else if (/^■/.test(text) && penaltyValue_(row[c.penaltyL - 1]) != null) {
      importantNo++;
      items.push({
        key: 'J' + ('0' + importantNo).slice(-2), no: '重' + importantNo, type: '重要', category: category || '重要事項',
        text: text.trim(), row: r,
        penalties: { L: penaltyValue_(row[c.penaltyL - 1]), M: null, H: null }
      });
    } else if (/^[\s　]/.test(text) && text.trim() && items.length) {
      items[items.length - 1].text += '\n' + text.trim();
    }
  }
  return items;
}

/** 項目文の注記から対象外業態を推定する（初期値。マスタで修正可能） */
function guessExcludedBusinesses(text) {
  var result = [];
  var re = /([^\s（(、・]+)は対象外/g;
  var m;
  while ((m = re.exec(text))) result.push(m[1]);
  return result;
}

/**
 * 登録済みのSVレポート（過去ファイル）を読み取る。
 * @return {{header:Object, items:Object[]}} items は parseTemplateItems の結果に count/applied を付けたもの
 */
function parseReportSheet(values) {
  var cell = SHEET_LAYOUT.cell;
  var c = SHEET_LAYOUT.col;
  var header = {
    version: parseVersion(values),
    business: String(cellAt_(values, cell.business) || '').trim(),
    schoolName: String(cellAt_(values, cell.schoolName) || '').trim(),
    schoolId: String(cellAt_(values, cell.schoolId) || '').trim().replace(/\.0$/, ''),
    date: cellAt_(values, cell.date),
    am: String(cellAt_(values, cell.am) || '').trim(),
    sv: String(cellAt_(values, cell.sv) || '').trim(),
    cash: cellAt_(values, cell.cash),
    bank: cellAt_(values, cell.bank),
    tms: cellAt_(values, cell.tms)
  };
  var items = parseTemplateItems(values).map(function (item) {
    var row = values[item.row - 1];
    var applied = Number(row[c.current - 1]) || 0;
    var count = Number(row[c.count - 1]) || (applied < 0 ? 1 : 0);
    item.applied = applied;
    item.count = applied < 0 ? count : 0;
    item.note = String(row[c.note - 1] || '').trim();
    return item;
  });
  return { header: header, items: items };
}

/** 適用減点と連続回数から元のレベルを推定する（一致しなければ ''） */
function inferLevel(item, applied, count) {
  if (!(applied < 0)) return '';
  var base = applied / Math.min(count || 1, MAX_MULTIPLIER);
  var hit = levelsOf(item).filter(function (x) { return x.penalty === base; })[0];
  return hit ? hit.level : '';
}

/**
 * 校舎一覧シートを読み取る（3行目以降、A列=ID）。
 * @return {Object[]} [{id, business, region, name, management, company, openDate, remarks}]
 */
function parseSchoolList(values) {
  var list = [];
  for (var r = 2; r < values.length; r++) {
    var row = values[r];
    if (row[0] === '' || row[0] === null || isNaN(Number(row[0]))) continue;
    list.push({
      id: String(Number(row[0])), business: String(row[1] || '').trim(), region: String(row[3] || '').trim(),
      name: String(row[4] || '').trim(), management: String(row[6] || '').trim(), company: String(row[7] || '').trim(),
      openDate: row[8], remarks: String(row[9] || '').trim()
    });
  }
  return list;
}

// ===== Repository.gs =====
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

// ===== ReportWriter.gs =====
/**
 * テンプレートを複製してレポートのスプレッドシートを作成・更新する。写真の保存も担当。
 * 保存先: {ROOT}/{年度}年度/第{回}回/
 */

function getOrCreateFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function reportFolder_(year, round) {
  var root = DriveApp.getFolderById(getConfig_().rootFolderId);
  return getOrCreateFolder_(getOrCreateFolder_(root, year + '年度'), '第' + round + '回');
}

function sanitizeFileName_(s) {
  return String(s).replace(/[\\/:*?"<>|]/g, '_');
}

/** 例: SVレポート_2026年度第4回_2351_山口宇部校_20261008 */
function reportFileName_(r) {
  return sanitizeFileName_('SVレポート_' + r.year + '年度第' + r.round + '回_' + r.schoolId + '_' + r.schoolName + '_' + String(r.date).replace(/-/g, ''));
}

/** テンプレートの値（数式は数式のまま）を取得する */
function baselineValues_(range) {
  var values = range.getValues();
  var formulas = range.getFormulas();
  return values.map(function (row, i) {
    return row.map(function (v, j) { return formulas[i][j] || v; });
  });
}

/**
 * @param {Object} p
 *   report       {reportId, year, round, schoolId, schoolName, business, date, am, sv, cash, bank, tms}
 *   items        項目マスタ
 *   lines        computeReport().lines
 *   prevLines    前回の明細 {項目キー: {applied}}
 *   prev2Lines   前々回の明細
 *   photos       [{itemKey, name, url}]
 *   existingFileId 再登録時の既存ファイルID（同じファイルを更新してURLを維持する）
 * @return {{fileId:string, url:string}}
 */
function writeReportFile_(p) {
  var config = getConfig_();
  var c = SHEET_LAYOUT.col;
  var cell = SHEET_LAYOUT.cell;
  var folder = reportFolder_(p.report.year, p.report.round);
  var name = reportFileName_(p.report);

  var file = null;
  if (p.existingFileId) {
    try {
      file = DriveApp.getFileById(p.existingFileId);
      if (file.isTrashed()) file = null;
    } catch (e) {
      file = null;
    }
  }
  if (file) {
    file.setName(name);
    file.moveTo(folder);
  } else {
    file = DriveApp.getFileById(config.templateId).makeCopy(name, folder);
  }

  var ss = SpreadsheetApp.openById(file.getId());
  var sh = ss.getSheetByName(REPORT_SHEET_NAME) || ss.getSheets()[0];
  var tpl = SpreadsheetApp.openById(config.templateId);
  var tsh = tpl.getSheetByName(REPORT_SHEET_NAME) || tpl.getSheets()[0];

  // ヘッダ（校舎名・業態はVLOOKUP式を値で置き換える）
  sh.getRange(cell.business).setValue(p.report.business);
  sh.getRange(cell.schoolName).setValue(p.report.schoolName);
  sh.getRange(cell.schoolId).setValue(Number(p.report.schoolId) || p.report.schoolId);
  sh.getRange(cell.date).setValue(parseDate_(p.report.date));
  sh.getRange(cell.am).setValue(p.report.am);
  sh.getRange(cell.sv).setValue(p.report.sv);
  sh.getRange(cell.cash).setValue(p.report.cash === '' ? '' : Number(p.report.cash));
  sh.getRange(cell.bank).setValue(p.report.bank === '' ? '' : Number(p.report.bank));
  sh.getRange(cell.tms).setValue(p.report.tms === '' ? '' : Number(p.report.tms));

  // 項目行（F列=備考、K列=連続回数、L列=今回、N列=前回、O列=前々回）。テンプレートの値を土台に毎回作り直す
  var rows = p.items.map(function (i) { return i.row; });
  var r1 = Math.min.apply(null, rows);
  var n = Math.max.apply(null, rows) - r1 + 1;
  var ko = baselineValues_(tsh.getRange(r1, c.count, n, c.prev2 - c.count + 1));
  var notes = baselineValues_(tsh.getRange(r1, c.note, n, 1));
  var lineByKey = {};
  p.lines.forEach(function (l) { lineByKey[l.key] = l; });
  var penalty = function (l) { return l && Number(l.applied) < 0 ? Number(l.applied) : ''; };

  p.items.forEach(function (item) {
    var i = item.row - r1;
    var l = lineByKey[item.key];
    ko[i][0] = l.applied < 0 ? l.count : '';
    ko[i][c.current - c.count] = penalty(l);
    ko[i][c.prev - c.count] = penalty(p.prevLines[item.key]);
    ko[i][c.prev2 - c.count] = penalty(p.prev2Lines[item.key]);
    notes[i][0] = l.na ? '対象外' + (l.note ? '：' + l.note : '') : l.note;
  });
  sh.getRange(r1, c.count, n, ko[0].length).setValues(ko);
  sh.getRange(r1, c.note, n, 1).setValues(notes);

  // 出力ファイルには校舎一覧（非表示シート）は不要
  var schoolSheet = ss.getSheetByName(SCHOOL_LIST_SHEET_NAME);
  if (schoolSheet) ss.deleteSheet(schoolSheet);

  writePhotoSheet_(ss, p.items, p.photos);
  SpreadsheetApp.flush();
  return { fileId: file.getId(), url: ss.getUrl() };
}

function writePhotoSheet_(ss, items, photos) {
  var ps = ss.getSheetByName(PHOTO_SHEET_NAME);
  if (!photos.length) {
    if (ps) ss.deleteSheet(ps);
    return;
  }
  if (!ps) ps = ss.insertSheet(PHOTO_SHEET_NAME);
  ps.clear();
  var itemByKey = {};
  items.forEach(function (i) { itemByKey[i.key] = i; });
  var rows = [['No', '項目', 'ファイル名', 'リンク']].concat(photos.map(function (ph) {
    var item = itemByKey[ph.itemKey];
    return [item ? item.no : '全体', item ? item.text : '', ph.name, '=HYPERLINK("' + ph.url + '","開く")'];
  }));
  ps.getRange(1, 1, rows.length, 4).setValues(rows);
  ps.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e8eaed');
  ps.setColumnWidth(2, 420);
  ps.setColumnWidth(3, 260);
}

/* ---------- 写真 ---------- */

function photoFolder_(report) {
  var parent = getOrCreateFolder_(reportFolder_(report.year, report.round), '写真');
  return getOrCreateFolder_(parent, sanitizeFileName_(report.reportId + '_' + report.schoolName));
}

/**
 * 画面から送られた写真（dataURL）をDriveに保存し、写真シートの行データを返す。
 * @param {Object[]} photos [{itemKey, itemNo, name, dataUrl}]
 */
function savePhotos_(report, photos) {
  if (!photos || !photos.length) return [];
  var folder = photoFolder_(report);
  var stamp = Utilities.formatDate(new Date(), TIME_ZONE, 'yyyyMMdd_HHmmss');
  return photos.map(function (ph, idx) {
    var m = /^data:([^;]+);base64,(.*)$/.exec(ph.dataUrl || '');
    if (!m) throw new Error('写真データが不正です: ' + ph.name);
    var name = sanitizeFileName_((ph.itemNo || '全体') + '_' + stamp + '_' + (idx + 1) + '_' + (ph.name || 'photo.jpg'));
    var file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name));
    return {
      'レポートID': report.reportId, '項目キー': ph.itemKey || '', 'ファイルID': file.getId(),
      'ファイル名': name, 'URL': file.getUrl(), '登録日時': new Date()
    };
  });
}

/** 指定レポートに属する写真だけをゴミ箱へ移す */
function trashPhotos_(reportId, fileIds) {
  if (!fileIds || !fileIds.length) return;
  var own = getPhotos_(reportId).map(function (p) { return p.fileId; });
  var targets = fileIds.filter(function (id) { return own.indexOf(String(id)) >= 0; });
  targets.forEach(function (id) {
    try {
      DriveApp.getFileById(id).setTrashed(true);
    } catch (e) {
      // 既に削除済みなら記録だけ消す
    }
  });
  removePhotoRecords_(reportId, targets);
}

// ===== Notify.gs =====
/**
 * 登録通知（固定の送信先1件。スクリプトプロパティ NOTIFY_TO）
 */
function notifyRegistered_(report, items, lines, isUpdate) {
  var to = getConfig_().notifyTo;
  if (!to) return false;
  var itemByKey = {};
  items.forEach(function (i) { itemByKey[i.key] = i; });
  var hits = lines.filter(function (l) { return l.applied < 0; });

  var subject = '【SVレポート' + (isUpdate ? '更新' : '登録') + '】' + report.year + '年度第' + report.round + '回 '
    + report.schoolName + '（' + report.schoolId + '） ' + report.score + '点';
  var body = [
    'SVレポートが' + (isUpdate ? '更新' : '登録') + 'されました。',
    '',
    '校舎　　: ' + report.schoolName + '（' + report.schoolId + ' / ' + report.business + '）',
    '実施回　: ' + report.year + '年度 第' + report.round + '回',
    '実施日　: ' + report.date,
    'AM / SV : ' + report.am + ' / ' + report.sv,
    '点数　　: ' + report.score + '点（減点 ' + report.total + '）',
    '登録者　: ' + report.registeredBy,
    '',
    '■指摘項目（' + hits.length + '件）'
  ].concat(hits.map(function (l) {
    var item = itemByKey[l.key];
    return '・No.' + item.no + ' ' + item.text.split('\n')[0] + ' ' + l.applied + (l.count >= 2 ? '（連続' + l.count + '回目）' : '');
  })).concat(['', 'レポート: ' + report.url]).join('\n');

  MailApp.sendEmail({ to: to, subject: subject, body: body, name: 'SVレポート登録' });
  return true;
}

// ===== Code.gs =====
/**
 * ウェブアプリのエントリポイントと画面から呼ばれるAPI（api* 関数）。
 * google.script.run は Date を返せないため、戻り値は文字列・数値のみで構成する。
 */

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('SVレポート登録')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/** 画面でもサーバと同じ計算をするため、Domain.gs の関数ソースを埋め込む */
function sharedSource() {
  return [
    'var MAX_MULTIPLIER = ' + MAX_MULTIPLIER + ';',
    'var LEVEL_IMPORTANT = ' + JSON.stringify(LEVEL_IMPORTANT) + ';',
    periodOf, reportIdOf, ordinalOf, isExcludedFor, levelsOf, basePenaltyOf, computeLine, computeReport
  ].map(String).join('\n\n');
}

function currentUserEmail_() {
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch (e) {
    return '';
  }
}

/** 起動時に必要なマスタ */
function apiGetBootstrap() {
  var master = getItems_();
  var reports = readObjects_('reports');
  var names = function (col) {
    var seen = {};
    reports.forEach(function (r) { if (r[col]) seen[String(r[col])] = true; });
    return Object.keys(seen).sort();
  };
  return {
    schools: getSchools_(),
    version: master.version,
    items: master.items,
    amNames: names('AM'),
    svNames: names('SV'),
    userEmail: currentUserEmail_(),
    notifyEnabled: !!getConfig_().notifyTo
  };
}

function reportSummary_(r) {
  if (!r) return null;
  return {
    reportId: r.reportId, year: r.year, round: r.round, date: r.date, am: r.am, sv: r.sv,
    cash: r.cash, bank: r.bank, tms: r.tms, score: r.score, total: r.total, url: r.url,
    revision: r.revision, updatedAt: r.updatedAt, registeredBy: r.registeredBy, kind: r.kind
  };
}

/**
 * 校舎・実施日を選んだときに、既存登録と前回・前々回を返す。
 */
function apiGetContext(schoolId, dateStr) {
  var school = getSchool_(schoolId);
  var period = periodOf(dateStr);
  var reportId = reportIdOf(period.year, period.round, school.id);
  var existing = findReport_(reportId);
  var prevs = previousReports_(school.id, period.year, period.round, 2);
  var withLines = function (r) {
    if (!r) return null;
    var s = reportSummary_(r);
    s.lines = getLines_(r.reportId);
    return s;
  };
  var later = reportsOfSchool_(school.id).filter(function (r) {
    return ordinalOf(r.year, r.round) > ordinalOf(period.year, period.round);
  });
  return {
    reportId: reportId,
    year: period.year,
    round: period.round,
    school: school,
    existing: existing ? Object.assign(withLines(existing), { photos: getPhotos_(reportId) }) : null,
    prev: withLines(prevs[0]),
    prev2: withLines(prevs[1]),
    laterReports: later.map(function (r) { return r.year + '年度第' + r.round + '回'; })
  };
}

/**
 * 登録。
 * @param {Object} payload {submissionId, schoolId, date, am, sv, cash, bank, tms,
 *   inputs: {項目キー: {level, na, note}}, photos: [{itemKey, itemNo, name, dataUrl}],
 *   removedPhotoIds: [], overwrite: boolean}
 */
function apiSubmitReport(payload) {
  validatePayload_(payload);
  var school = getSchool_(payload.schoolId);
  var master = getItems_();
  var period = periodOf(payload.date);
  var reportId = reportIdOf(period.year, period.round, school.id);

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('他の登録処理中です。しばらくしてから再度登録してください。');
  var report, computed, isUpdate;
  try {
    var existing = findReport_(reportId);
    if (existing && payload.submissionId && existing.submissionId === payload.submissionId) {
      // 通信切れ等で同じ送信が再送された場合は二重登録しない
      return { status: 'ok', reportId: reportId, url: existing.url, score: existing.score, duplicate: true };
    }
    if (existing && !payload.overwrite) {
      return { status: 'exists', reportId: reportId, existing: reportSummary_(existing) };
    }
    isUpdate = !!existing;

    var prevs = previousReports_(school.id, period.year, period.round, 2);
    var prevLines = prevs[0] ? getLines_(prevs[0].reportId) : {};
    var prev2Lines = prevs[1] ? getLines_(prevs[1].reportId) : {};
    computed = computeReport(master.items, payload.inputs, prevLines, school.business);

    report = {
      reportId: reportId, year: period.year, round: period.round,
      schoolId: school.id, schoolName: school.name, business: school.business,
      date: payload.date, am: String(payload.am).trim(), sv: String(payload.sv).trim(),
      cash: payload.cash, bank: payload.bank, tms: payload.tms,
      total: computed.total, score: computed.score, registeredBy: currentUserEmail_()
    };

    // 写真
    trashPhotos_(reportId, payload.removedPhotoIds || []);
    appendObjects_('photos', savePhotos_(report, payload.photos || []));
    var photos = getPhotos_(reportId);

    var out = writeReportFile_({
      report: report, items: master.items, lines: computed.lines,
      prevLines: prevLines, prev2Lines: prev2Lines, photos: photos,
      existingFileId: existing ? existing.fileId : ''
    });
    report.url = out.url;

    var now = new Date();
    saveReport_({
      'レポートID': reportId, '年度': period.year, '実施回': period.round, '校舎ID': school.id,
      '校舎名': school.name, '業態': school.business, '実施日': payload.date, 'AM': report.am, 'SV': report.sv,
      '現金': payload.cash, '通帳残': payload.bank, 'TMS金額': payload.tms,
      '減点合計': computed.total, '点数': computed.score, 'テンプレート版': master.version,
      'ファイルID': out.fileId, 'ファイルURL': out.url, '版数': existing ? existing.revision + 1 : 1,
      '登録者': report.registeredBy, '登録日時': existing ? existing.createdAt : now, '更新日時': now,
      '区分': '登録', '送信ID': payload.submissionId || ''
    }, computed.lines.map(function (l) {
      return {
        'レポートID': reportId, '校舎ID': school.id, '年度': period.year, '実施回': period.round,
        '項目キー': l.key, 'レベル': l.level, '基本減点': l.base, '連続回数': l.count,
        '適用減点': l.applied, '対象外': l.na, '備考': l.note
      };
    }));
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  var notified = false;
  var warning = '';
  try {
    notified = notifyRegistered_(report, master.items, computed.lines, isUpdate);
  } catch (e) {
    warning = '通知メールの送信に失敗しました: ' + e.message;
  }
  return {
    status: 'ok', reportId: reportId, url: report.url, score: computed.score, total: computed.total,
    updated: isUpdate, notified: notified, warning: warning
  };
}

function validatePayload_(p) {
  var errors = [];
  if (!p) throw new Error('登録データがありません');
  if (!p.schoolId) errors.push('校舎');
  if (!p.date) errors.push('実施日');
  if (!String(p.am || '').trim()) errors.push('AM');
  if (!String(p.sv || '').trim()) errors.push('SV');
  ['cash', 'bank', 'tms'].forEach(function (k, i) {
    if (p[k] === '' || p[k] === null || p[k] === undefined || isNaN(Number(p[k]))) {
      errors.push(['現金', '通帳残', 'TMS金額'][i]);
    }
  });
  if (errors.length) throw new Error('未入力または不正な項目があります: ' + errors.join('、'));
}

// ===== ファイル終端（全1217行）: この行まで貼り付けられていればOK =====
