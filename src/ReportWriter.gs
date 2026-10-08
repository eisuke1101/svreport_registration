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
