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
