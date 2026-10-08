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
