/**
 * レポート（スプレッドシート）の PDF 出力。管理DBのメニューから実行する。
 *
 * - 出力先: {保存先}/{年度}年度/第{回}回/PDF/  （年度・実施回は管理DB「登録履歴」の値で判定）
 * - ファイル名: レポートと同じ名前 + .pdf。同名のPDFがあれば古い方をゴミ箱へ移して置き換える
 * - 出力後、「登録履歴」の「PDF出力」列に「済」を記入
 * - 一括出力: 「PDF出力」が空欄の行のみ / 個別出力: 選択した行のみ（「済」でも出し直す）
 * - 再登録（上書き）された行は「PDF出力」が空欄に戻るので、次の一括出力で出し直される
 * - 実行時間の上限（6分）に近づいたら中断。残りは再実行で続きから処理される
 */

var PDF_COLUMN = 'PDF出力';
var PDF_DONE = '済';
var PDF_FOLDER_NAME = 'PDF';
var PDF_TIME_LIMIT_MS = 5 * 60 * 1000;

/* ---------- メニュー ---------- */

/** 管理DBを開いたときにメニューを追加する（setup() がインストール型トリガーとして登録） */
function onDbOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SVレポート')
    .addItem('PDF出力（未出力をすべて）', 'menuExportPendingPdfs')
    .addItem('PDF出力（選択した行）', 'menuExportSelectedPdfs')
    .addToUi();
}

/** 管理DBのメニュー用トリガーを登録する（重複して登録しない） */
function installDbMenuTrigger_(dbId) {
  var exists = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'onDbOpen' && t.getTriggerSourceId() === dbId;
  });
  if (!exists) ScriptApp.newTrigger('onDbOpen').forSpreadsheet(dbId).onOpen().create();
}

function menuExportPendingPdfs() {
  var ui = SpreadsheetApp.getUi();
  var targets = readObjects_('reports').filter(function (r) {
    return r['レポートID'] !== '' && String(r[PDF_COLUMN]).trim() === '';
  });
  if (!targets.length) {
    ui.alert('PDF出力', '「PDF出力」が空欄の行はありません。', ui.ButtonSet.OK);
    return;
  }
  if (ui.alert('PDF出力', '「PDF出力」が空欄の ' + targets.length + ' 件をPDFに出力します。よろしいですか？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  ui.alert('PDF出力', pdfResultMessage_(exportPdfs_(targets)), ui.ButtonSet.OK);
}

function menuExportSelectedPdfs() {
  var ui = SpreadsheetApp.getUi();
  var ids = selectedReportIds_();
  if (ids === null) {
    // 選択範囲を取得できない場合は入力してもらう
    var res = ui.prompt('PDF出力（個別）',
      '出力する行の行番号（例: 5,8,12）またはレポートID（例: 2026-4-2351）を入力してください。', ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) return;
    ids = reportIdsFromInput_(res.getResponseText());
  }
  if (!ids.length) {
    ui.alert('PDF出力', '「登録履歴」シートで出力したい行を選択してから実行してください（複数行・複数範囲の選択可）。', ui.ButtonSet.OK);
    return;
  }
  var all = readObjects_('reports');
  var targets = ids.map(function (id) {
    return all.filter(function (r) { return String(r['レポートID']) === id; })[0];
  }).filter(Boolean);
  if (!targets.length) {
    ui.alert('PDF出力', '対象のレポートが見つかりません。', ui.ButtonSet.OK);
    return;
  }
  var names = targets.slice(0, 10).map(function (r) {
    return '・' + r['年度'] + '年度第' + r['実施回'] + '回 ' + r['校舎名'] + (String(r[PDF_COLUMN]) === PDF_DONE ? '（出力済→出し直し）' : '');
  }).join('\n') + (targets.length > 10 ? '\n…ほか ' + (targets.length - 10) + ' 件' : '');
  if (ui.alert('PDF出力（個別）', targets.length + ' 件をPDFに出力します。\n\n' + names, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  ui.alert('PDF出力', pdfResultMessage_(exportPdfs_(targets)), ui.ButtonSet.OK);
}

/**
 * 管理DBで選択中の「登録履歴」の行のレポートID。
 * 選択範囲を取得できない（管理DB以外から実行された等）ときは null。
 */
function selectedReportIds_() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active || active.getId() !== getConfig_().dbId) return null;
  var sheet = active.getActiveSheet();
  if (!sheet || sheet.getName() !== SHEETS.reports) return [];
  var list = active.getActiveRangeList();
  if (!list) return [];
  var rows = {};
  list.getRanges().forEach(function (rg) {
    for (var r = rg.getRow(); r < rg.getRow() + rg.getNumRows(); r++) if (r >= 2) rows[r] = true;
  });
  var rowNums = Object.keys(rows).map(Number).sort(function (a, b) { return a - b; });
  if (!rowNums.length) return [];
  var last = sheet.getLastRow();
  return rowNums.filter(function (r) { return r <= last; }).map(function (r) {
    return String(sheet.getRange(r, 1).getValues()[0][0]);
  }).filter(String);
}

/** 入力（行番号 または レポートID のカンマ区切り）をレポートIDの配列にする */
function reportIdsFromInput_(text) {
  var all = readObjects_('reports');
  var ids = [];
  String(text || '').split(/[,、，\s]+/).filter(String).forEach(function (token) {
    if (/^\d+$/.test(token)) {
      var hit = all.filter(function (r) { return r._row === Number(token); })[0];
      if (hit) ids.push(String(hit['レポートID']));
    } else {
      ids.push(token);
    }
  });
  return ids.filter(function (id, i) { return id && ids.indexOf(id) === i; });
}

/* ---------- 出力処理 ---------- */

/** 出力先フォルダ: {年度}年度/第{回}回/PDF */
function pdfFolder_(year, round) {
  return getOrCreateFolder_(reportFolder_(year, round), PDF_FOLDER_NAME);
}

/**
 * @param {Object[]} targets readObjects_('reports') の行
 * @return {{ok:string[], errors:string[], remaining:number}}
 */
function exportPdfs_(targets) {
  ensurePdfColumn_();
  var started = Date.now();
  var result = { ok: [], errors: [], remaining: 0 };
  targets.forEach(function (r) {
    var label = r['年度'] + '年度第' + r['実施回'] + '回 ' + r['校舎名'] + '（' + r['レポートID'] + '）';
    if (Date.now() - started > PDF_TIME_LIMIT_MS) {
      result.remaining++;
      return;
    }
    try {
      exportOnePdf_(r);
      markPdfDone_(String(r['レポートID']));
      result.ok.push(label);
    } catch (e) {
      result.errors.push(label + ': ' + e.message);
    }
  });
  return result;
}

function exportOnePdf_(r) {
  var year = Number(r['年度']);
  var round = Number(r['実施回']);
  if (!year || !(round >= 1 && round <= 4)) throw new Error('年度・実施回が不正です');
  var fileId = String(r['ファイルID'] || '');
  if (!fileId) throw new Error('ファイルIDが空です');
  var file;
  try {
    file = DriveApp.getFileById(fileId);
  } catch (e) {
    throw new Error('レポートファイルが見つかりません（削除された可能性があります）');
  }
  if (file.isTrashed()) throw new Error('レポートファイルがゴミ箱にあります');

  // 過去データ取込のExcelは一時的にスプレッドシートへ変換して出力する
  var tempId = null;
  var ssId = fileId;
  if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) {
    tempId = convertToSpreadsheet_(fileId, 'PDF作業_' + file.getName(), getOrCreateFolder_(systemFolder_(), '_取込作業').getId());
    ssId = tempId;
  }
  try {
    var name = file.getName().replace(/\.(xlsx?|xls)$/i, '') + '.pdf';
    var blob = exportSheetAsPdf_(ssId).setName(name);
    var folder = pdfFolder_(year, round);
    var old = folder.getFilesByName(name);
    while (old.hasNext()) old.next().setTrashed(true);
    folder.createFile(blob);
  } finally {
    if (tempId) DriveApp.getFileById(tempId).setTrashed(true);
  }
}

/** 「チェック項目」シートの印刷範囲（A1:O最終行）を A4縦・1ページに収めてPDF化 */
function exportSheetAsPdf_(spreadsheetId) {
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sh = ss.getSheetByName(REPORT_SHEET_NAME) || ss.getSheets()[0];
  var params = {
    format: 'pdf', gid: sh.getSheetId(), range: 'A1:O' + sh.getLastRow(),
    size: 'A4', portrait: 'true', scale: '4', // 4 = ページに合わせる
    gridlines: 'false', printtitle: 'false', sheetnames: 'false', pagenum: 'UNDEFINED',
    top_margin: '0.4', bottom_margin: '0.4', left_margin: '0.4', right_margin: '0.4',
    horizontal_alignment: 'CENTER'
  };
  var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?' + Object.keys(params).map(function (k) {
    return k + '=' + encodeURIComponent(params[k]);
  }).join('&');
  var options = { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
  for (var attempt = 1; ; attempt++) {
    var res = UrlFetchApp.fetch(url, options);
    var code = res.getResponseCode();
    if (code === 200) return res.getBlob();
    // 短時間に連続で出力すると 429（回数制限）になることがあるので待って再試行
    if (code === 429 && attempt < 4) {
      Utilities.sleep(5000 * attempt);
      continue;
    }
    throw new Error('PDF変換に失敗しました（HTTP ' + code + '）');
  }
}

/** 「登録履歴」に「PDF出力」列の見出しが無ければ追加する（setup 未再実行の環境向け） */
function ensurePdfColumn_() {
  var sh = sheetOf_('reports');
  var col = HEADERS.reports.indexOf(PDF_COLUMN) + 1;
  if (sh.getRange(1, col).getValues()[0][0] !== PDF_COLUMN) sh.getRange(1, col).setValue(PDF_COLUMN);
}

/** 出力済みを記入（行は登録処理で更新されることがあるので、書く直前にレポートIDで探す） */
function markPdfDone_(reportId) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('他の処理中のため「済」を記入できませんでした');
  try {
    var hit = findReport_(reportId);
    if (!hit) throw new Error('登録履歴に行が見つかりません');
    sheetOf_('reports').getRange(hit._row, HEADERS.reports.indexOf(PDF_COLUMN) + 1).setValue(PDF_DONE);
  } finally {
    lock.releaseLock();
  }
}

function pdfResultMessage_(res) {
  var lines = ['出力: ' + res.ok.length + ' 件'];
  if (res.errors.length) {
    lines.push('エラー: ' + res.errors.length + ' 件（「PDF出力」は空欄のまま）');
    lines = lines.concat(res.errors.slice(0, 10).map(function (e) { return '・' + e; }));
    if (res.errors.length > 10) lines.push('…ほか ' + (res.errors.length - 10) + ' 件');
  }
  if (res.remaining) lines.push('時間切れで未処理: ' + res.remaining + ' 件（もう一度実行してください）');
  return lines.join('\n');
}
