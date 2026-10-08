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
