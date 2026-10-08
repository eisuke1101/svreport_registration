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
 *   - 管理DBにメニュー「SVレポート」（PDF出力）を追加するトリガーを登録
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
  installDbMenuTrigger_(ss.getId());
  Logger.log('セットアップ完了\n管理DB: %s\nテンプレート: %s\n管理DBを開き直すとメニュー「SVレポート」が表示されます。',
    ss.getUrl(), SpreadsheetApp.openById(templateId).getUrl());
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
  requireDriveService_();
  var copied = Drive.Files.copy(
    { name: name, mimeType: MimeType.GOOGLE_SHEETS, parents: [parentId] },
    fileId,
    { supportsAllDrives: true }
  );
  return copied.id;
}

/** Excel の変換に使う Drive API（拡張サービス）が有効か確認する */
function requireDriveService_() {
  if (typeof Drive === 'undefined') {
    throw new Error('Drive API（拡張サービス）が有効になっていません。'
      + 'エディタ左の「サービス」の＋ →「Drive API」を選び、バージョン v3・ID「Drive」のまま「追加」してから再実行してください。');
  }
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
