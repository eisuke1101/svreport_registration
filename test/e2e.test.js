// 貼り付け用の1ファイル（dist/SVレポート登録.gs）を GAS モック環境で動かし、
// setup → 画面初期表示 → 登録 → 出力シート → 上書き → 過去データ取込 を通しで検証する。
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { createEnv, loadApp } = require('./gasenv');

const ROOT = path.join(__dirname, '..');
const BUNDLE = path.join(ROOT, 'dist', 'SVレポート登録.gs');
const templateValues = require('./fixtures/template_values.json');
const ROOT_FOLDER = '1fCaEEFJMTKgGPPINg7G5TuBNuohbwymo';

const templateFormulas = {
  '2,2': '=VLOOKUP(E2,校舎一覧!A:B,2,FALSE())', '2,3': '=VLOOKUP(E2,校舎一覧!A:E,5,FALSE())', '10,5': '=E8+E9',
  '107,12': '=SUM(L6:M106)', '107,14': '=SUM(N6:N106)', '107,15': '=SUM(O6:O106)',
  '108,12': '=100+L107', '108,14': '=100+N107', '108,15': '=100+O107'
};
const schoolList = [
  [new Date(2026, 9, 2), '', '', '', '', '', 3],
  ['ID', '業態', '', '地域', '校舎正式名', '', '経営', '会社', '開校日', '備考'],
  [2351, 'ITTO', 35, '山口', '山口宇部校', 1, '直営', '㈱ITTO', new Date(2023, 10, 1), '2026/9譲渡'],
  [9001, 'アスモ', 13, '東京', '足立北千住校', 1, '直営', '㈱ITTO', new Date(2025, 9, 1), ''],
  [9002, 'みやび', 34, '広島', '広島/福山校', 1, '直営', '㈱ITTO', new Date(2025, 1, 1), '']
];

function boot() {
  assert.ok(fs.existsSync(BUNDLE), 'dist/SVレポート登録.gs がありません（npm run bundle）');
  const env = createEnv({ rootFolderId: ROOT_FOLDER, templateValues, templateFormulas, schoolList });
  env.props.TEMPLATE_XLS_FILE_ID = 'xls';
  env.props.NOTIFY_TO = 'boss@example.com';
  const app = loadApp(env, fs.readFileSync(BUNDLE, 'utf8'));
  app.setup();
  return { env, app };
}

/** google.script.run で返せない値（Date・関数・undefined）が含まれていないか */
function assertTransferable(v, where) {
  if (v instanceof Date) assert.fail(`${where} に Date が含まれています（画面に返せません）`);
  if (typeof v === 'function' || v === undefined) assert.fail(`${where} に返せない値: ${typeof v}`);
  if (v && typeof v === 'object') Object.keys(v).forEach((k) => assertTransferable(v[k], `${where}.${k}`));
}

function folderByPath(env, names) {
  let id = ROOT_FOLDER;
  for (const n of names) {
    const f = Object.values(env.folders).find((x) => x.parent === id && x.name === n);
    assert.ok(f, `フォルダ ${names.join('/')} がありません`);
    id = f.id;
  }
  return id;
}
function outputSheet(env, fileId) {
  return env.spreadsheets[fileId].sheets.find((s) => s.name === 'チェック項目');
}
function cell(sh, a1) {
  const v = sh.api.getRange(a1).getValues()[0][0];
  return v;
}
function rowOf(app, key) {
  return app.getItems_().items.find((i) => i.key === key).row;
}

const base = { schoolId: '2351', am: '山田AM', sv: '佐藤SV', cash: 10000, bank: 25000, tms: 35000 };

test('setup: テンプレート変換・管理DB・マスタ作成、2回実行しても重複しない', () => {
  const { env, app } = boot();
  assert.ok(env.props.TEMPLATE_SPREADSHEET_ID);
  assert.ok(env.props.DB_SPREADSHEET_ID);
  assert.equal(env.props.CURRENT_ITEM_VERSION, '2026.10');
  const sys = folderByPath(env, ['_システム']);
  assert.ok(env.files[env.props.DB_SPREADSHEET_ID].parents.includes(sys));
  const db = env.spreadsheets[env.props.DB_SPREADSHEET_ID];
  assert.deepEqual(db.sheets.map((s) => s.name).sort(), ['写真', '取込ログ', '明細', '校舎マスタ', '登録履歴', '項目マスタ'].sort());
  assert.equal(app.getItems_().items.length, 92);
  assert.equal(app.getSchools_().length, 3);
  app.setup();
  assert.equal(app.readObjects_('items').length, 92);
  assert.equal(app.readObjects_('schools').length, 3);
});

test('setup: TEMPLATE_XLS_FILE_ID 未設定なら分かるエラー', () => {
  const env = createEnv({ rootFolderId: ROOT_FOLDER, templateValues, templateFormulas, schoolList });
  const app = loadApp(env, fs.readFileSync(BUNDLE, 'utf8'));
  assert.throws(() => app.setup(), /TEMPLATE_XLS_FILE_ID/);
});

test('画面: 初期データ・前回前々回の取得結果が google.script.run で返せる形', () => {
  const { app } = boot();
  const b = app.apiGetBootstrap();
  assertTransferable(b, 'apiGetBootstrap');
  assert.equal(b.items.length, 92);
  app.apiSubmitReport({ ...base, submissionId: 's1', date: '2026-08-01', inputs: { I01: { level: 'M' } } });
  const ctx = app.apiGetContext('2351', '2026-10-08');
  assertTransferable(ctx, 'apiGetContext');
  assert.equal(ctx.prev.reportId, '2026-3-2351');
  assertTransferable(app.apiGetContext('2351', '2026-08-01'), 'apiGetContext(既存)');
});

test('登録: 指定フォルダにExcel形式のシートを作成し、セルに正しく書き込む', () => {
  const { env, app } = boot();
  const r1 = app.apiSubmitReport({ ...base, submissionId: 's1', date: '2026-08-01', inputs: { I01: { level: 'M' }, J01: { level: 'X' } } });
  const res = app.apiSubmitReport({
    ...base, submissionId: 's2', date: '2026-10-08',
    inputs: { I01: { level: 'L', note: '剥がれ' }, J01: { level: 'X' }, I05: { level: 'H', na: true, note: '確認不可' } },
    photos: [{ itemKey: 'I01', itemNo: '1', name: 'a.jpg', dataUrl: 'data:image/jpeg;base64,/9j/AA==' }]
  });
  assertTransferable(res, 'apiSubmitReport');
  assert.equal(r1.status, 'ok');
  assert.equal(res.status, 'ok');
  assert.equal(res.total, -2 * 2 - 200);

  const folder = folderByPath(env, ['2026年度', '第4回']);
  const out = Object.values(env.files).find((f) => f.parents.includes(folder) && f.mime === 'application/vnd.google-apps.spreadsheet');
  assert.equal(out.name, 'SVレポート_2026年度第4回_2351_山口宇部校_20261008');
  const sh = outputSheet(env, out.id);
  assert.equal(cell(sh, 'B2'), 'ITTO');
  assert.equal(cell(sh, 'C2'), '山口宇部校');
  assert.equal(cell(sh, 'E2'), 2351);
  assert.ok(cell(sh, 'G2') instanceof Date);
  assert.equal(cell(sh, 'M3'), '山田AM');
  assert.equal(cell(sh, 'M5'), '佐藤SV');
  assert.equal(cell(sh, 'E8'), 10000);
  assert.equal(cell(sh, 'E9'), 25000);
  assert.equal(cell(sh, 'G10'), 35000);
  assert.equal(sh.formulas['10,5'], '=E8+E9');
  assert.equal(cell(sh, 'F10'), '[TMS金額]'); // ラベルを消していない
  const r01 = rowOf(app, 'I01');
  assert.equal(cell(sh, 'K' + r01), 2);
  assert.equal(cell(sh, 'L' + r01), -4);
  assert.equal(cell(sh, 'N' + r01), -5);
  assert.equal(cell(sh, 'O' + r01), '');
  assert.equal(cell(sh, 'F' + r01), '剥がれ');
  assert.equal(cell(sh, 'L6'), -200);
  assert.equal(cell(sh, 'N6'), -100);
  const r05 = rowOf(app, 'I05');
  assert.equal(cell(sh, 'L' + r05), '');
  assert.equal(cell(sh, 'F' + r05), '対象外：確認不可');
  assert.equal(sh.formulas['107,12'], '=SUM(L6:M106)');
  const ss = env.spreadsheets[out.id];
  assert.ok(!ss.sheets.some((s) => s.name === '校舎一覧'));
  const photo = ss.sheets.find((s) => s.name === '写真');
  assert.match(photo.formulas['2,4'], /^=HYPERLINK\("https:\/\/drive.google.com\/file\/d\/file\d+","開く"\)$/);
  assert.ok(folderByPath(env, ['2026年度', '第4回', '写真', '2026-4-2351_山口宇部校']));
  assert.equal(env.mails.length, 2);
  assert.match(env.mails[1].body, /連続2回目/);
});

test('登録: 上書きで同じファイルを更新し、写真の削除・明細の置換ができる', () => {
  const { env, app } = boot();
  const photo = { itemKey: 'I02', itemNo: '2', name: 'b.jpg', dataUrl: 'data:image/jpeg;base64,/9j/AA==' };
  const first = app.apiSubmitReport({ ...base, submissionId: 'a', date: '2026-10-01', inputs: { I01: { level: 'H' } }, photos: [photo] });
  const exists = app.apiSubmitReport({ ...base, submissionId: 'b', date: '2026-11-01', inputs: {} });
  assert.equal(exists.status, 'exists');
  const ctx = app.apiGetContext('2351', '2026-11-01');
  const photoId = ctx.existing.photos[0].fileId;
  const second = app.apiSubmitReport({ ...base, submissionId: 'b', date: '2026-11-01', inputs: { I02: { level: 'L' } }, removedPhotoIds: [photoId], overwrite: true });
  assert.equal(second.url, first.url);
  assert.equal(env.files[photoId].trashed, true);
  assert.equal(app.getPhotos_('2026-4-2351').length, 0);
  const sh = outputSheet(env, second.url.split('/').pop());
  assert.equal(cell(sh, 'L' + rowOf(app, 'I01')), '');
  assert.equal(cell(sh, 'L' + rowOf(app, 'I02')), -2);
  assert.ok(!env.spreadsheets[second.url.split('/').pop()].sheets.some((s) => s.name === '写真'));
  assert.equal(app.readObjects_('lines').length, 92);
  // ファイル名の禁止文字は置換
  app.apiSubmitReport({ ...base, schoolId: '9002', submissionId: 'c', date: '2026-10-01', inputs: {} });
  assert.ok(Object.values(env.files).some((f) => f.name === 'SVレポート_2026年度第4回_9002_広島_福山校_20261001'));
});

test('過去データ取込: Excelを変換して取り込み、前回として参照できる。再実行はスキップ', () => {
  const { env, app } = boot();
  env.folders.imp = { id: 'imp', name: '過去分', parent: 'root' };
  const past = (fileId, date, edits) => {
    const v = JSON.parse(JSON.stringify(templateValues));
    v[1][4] = 2351; v[1][6] = date;
    edits.forEach(([a1, val]) => {
      const m = /^([A-Z])(\d+)$/.exec(a1);
      v[Number(m[2]) - 1][m[1].charCodeAt(0) - 65] = val;
    });
    const sheet = { name: 'チェック項目', data: v, formulas: {}, maxRows: 1000, maxCols: 26 };
    env.files[fileId] = { id: fileId, name: fileId + '.xls', mime: 'application/vnd.ms-excel', parents: ['imp'], trashed: false, excelContent: { sheets: [sheet] } };
  };
  past('p1', new Date(2026, 4, 20), [['L12', -10], ['K12', 2], ['L6', -100]]);
  past('p2', '2026/8/5', [['L12', -15], ['K12', 3]]);
  env.props.IMPORT_FOLDER_ID = 'imp';
  app.importPastReports();
  const log = app.readObjects_('importLog');
  assert.deepEqual(log.map((r) => r['結果']), ['OK', 'OK'], JSON.stringify(log.map((r) => r['メッセージ'])));
  const ctx = app.apiGetContext('2351', '2026-10-08');
  assert.equal(ctx.prev.reportId, '2026-3-2351');
  assert.equal(ctx.prev.lines.I01.applied, -15);
  assert.equal(ctx.prev.lines.I01.level, 'M');
  assert.equal(ctx.prev2.lines.J01.applied, -100);
  // 取込データから連続4回目
  const res = app.apiSubmitReport({ ...base, submissionId: 'x', date: '2026-10-08', inputs: { I01: { level: 'L' } } });
  assert.equal(res.total, -2 * 4);
  // 一時変換ファイルはゴミ箱へ、再実行はスキップ
  assert.ok(Object.values(env.files).filter((f) => f.name.startsWith('取込_')).every((f) => f.trashed));
  app.importPastReports();
  assert.equal(app.readObjects_('importLog').length, 2);
});

test('画面HTML: テンプレートを展開した結果のスクリプトが構文エラーにならない', () => {
  const env = createEnv({ rootFolderId: ROOT_FOLDER, templateValues, templateFormulas, schoolList });
  const app = loadApp(env, fs.readFileSync(BUNDLE, 'utf8'));
  app.HtmlService = { createHtmlOutputFromFile: (n) => ({ getContent: () => fs.readFileSync(path.join(ROOT, 'src', n + '.html'), 'utf8') }) };
  let html = fs.readFileSync(path.join(ROOT, 'src', 'index.html'), 'utf8');
  html = html.replace(/<\?!=\s*([\s\S]*?)\s*\?>/g, (_, expr) => vm.runInContext(expr, app));
  assert.doesNotMatch(html, /<\?/);
  const scripts = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map((m) => m[1]);
  assert.equal(scripts.length, 2);
  scripts.forEach((s, i) => assert.doesNotThrow(() => new vm.Script(s, { filename: `script${i}` })));
  const ids = [...html.matchAll(/\$\('([A-Za-z0-9]+)'\)/g)].map((m) => m[1]);
  ids.forEach((id) => assert.match(html, new RegExp(`id="${id}"`), `画面に id="${id}" がありません`));
});

test('setup / 取込: Drive API（拡張サービス）が無効なら追加方法を案内するエラー', () => {
  const env = createEnv({ rootFolderId: ROOT_FOLDER, templateValues, templateFormulas, schoolList });
  env.props.TEMPLATE_XLS_FILE_ID = 'xls';
  const driveService = env.g.Drive;
  delete env.g.Drive;
  const app = loadApp(env, fs.readFileSync(BUNDLE, 'utf8'));
  assert.throws(() => app.setup(), /Drive API（拡張サービス）が有効になっていません.*サービス/);
  // 有効化して再実行すれば完了する（途中で作られた _システム フォルダがあっても重複しない）
  app.Drive = driveService;
  app.setup();
  assert.equal(app.getItems_().items.length, 92);
  assert.equal(Object.values(env.folders).filter((f) => f.name === '_システム').length, 1);
  const { env: env2, app: app2 } = boot();
  delete app2.Drive;
  env2.props.IMPORT_FOLDER_ID = 'root';
  assert.throws(() => app2.importPastReports(), /Drive API（拡張サービス）/);
  assert.equal(app2.readObjects_('importLog').length, 0);
});

/* ---------- PDF出力（管理DBのメニュー） ---------- */

function pdfFiles(env, year, round) {
  const folder = folderByPath(env, [`${year}年度`, `第${round}回`, 'PDF']);
  return Object.values(env.files).filter((f) => f.parents.includes(folder) && !f.trashed);
}
function pdfColumn(app) {
  return Object.fromEntries(app.readObjects_('reports').map((r) => [r['レポートID'], r['PDF出力']]));
}
function bootWithReports() {
  const ctx = boot();
  const { app } = ctx;
  app.apiSubmitReport({ ...base, submissionId: 'r3', date: '2026-08-01', inputs: { I01: { level: 'M' } } });
  app.apiSubmitReport({ ...base, submissionId: 'r4', date: '2026-10-08', inputs: {} });
  app.apiSubmitReport({ ...base, schoolId: '9001', submissionId: 'a4', date: '2026-11-20', inputs: {} });
  return ctx;
}

test('PDF: setup で管理DBにメニュー用トリガーを1つだけ登録し、メニューの関数が存在する', () => {
  const { env, app } = boot();
  app.setup();
  assert.equal(env.triggers.length, 1);
  assert.deepEqual({ ...env.triggers[0] }, { fn: 'onDbOpen', source: env.props.DB_SPREADSHEET_ID, type: 'open' });
  app.onDbOpen();
  assert.equal(env.ui.menu.name, 'SVレポート');
  assert.equal(env.ui.menu.items.length, 2);
  env.ui.menu.items.forEach((i) => assert.equal(typeof app[i.fn], 'function', i.fn));
});

test('PDF一括: 空欄の行だけを 年度/回/PDF に出力し「済」を記入、2回目は対象なし', () => {
  const { env, app } = bootWithReports();
  app.menuExportPendingPdfs();
  assert.match(env.ui.alerts[0].msg, /空欄の 3 件/);
  assert.match(env.ui.alerts[1].msg, /出力: 3 件/);
  const p4 = pdfFiles(env, 2026, 4).map((f) => f.name).sort();
  assert.deepEqual(p4, ['SVレポート_2026年度第4回_2351_山口宇部校_20261008.pdf', 'SVレポート_2026年度第4回_9001_足立北千住校_20261120.pdf']);
  assert.deepEqual(pdfFiles(env, 2026, 3).map((f) => f.name), ['SVレポート_2026年度第3回_2351_山口宇部校_20260801.pdf']);
  assert.ok(pdfFiles(env, 2026, 4).every((f) => f.mime === 'application/pdf'));
  // チェック項目シートの印刷範囲を A4縦・1ページで
  const f = env.fetches[0];
  const report = env.spreadsheets[f.id];
  assert.equal(f.q.gid, String(report.sheets.find((s) => s.name === 'チェック項目').gid));
  assert.equal(f.q.range, 'A1:O108');
  assert.equal(f.q.format, 'pdf');
  assert.equal(f.q.size, 'A4');
  assert.equal(f.q.portrait, 'true');
  assert.equal(f.q.scale, '4');
  assert.deepEqual(Object.values(pdfColumn(app)), ['済', '済', '済']);
  assert.equal(env.spreadsheets[env.props.DB_SPREADSHEET_ID].sheets.find((s) => s.name === '登録履歴').data[0][23], 'PDF出力');
  // 2回目
  app.menuExportPendingPdfs();
  assert.match(env.ui.alerts[2].msg, /空欄の行はありません/);
  assert.equal(env.fetches.length, 3);
});

test('PDF一括: キャンセルしたら何もしない', () => {
  const { env, app } = bootWithReports();
  env.ui.answer = 'CANCEL';
  app.menuExportPendingPdfs();
  assert.equal(env.fetches.length, 0);
  assert.deepEqual(Object.values(pdfColumn(app)), ['', '', '']);
});

test('PDF: 再登録（上書き）で「PDF出力」が空欄に戻り、次の一括で出し直して古いPDFを置き換える', () => {
  const { env, app } = bootWithReports();
  app.menuExportPendingPdfs();
  app.apiSubmitReport({ ...base, submissionId: 'r4b', date: '2026-10-08', inputs: { I02: { level: 'L' } }, overwrite: true });
  assert.equal(pdfColumn(app)['2026-4-2351'], '');
  assert.equal(pdfColumn(app)['2026-3-2351'], '済');
  app.menuExportPendingPdfs();
  assert.match(env.ui.alerts[2].msg, /空欄の 1 件/);
  assert.equal(pdfFiles(env, 2026, 4).length, 2); // 同名は1つだけ（古いものはゴミ箱）
  assert.equal(Object.values(env.files).filter((x) => x.name === 'SVレポート_2026年度第4回_2351_山口宇部校_20261008.pdf' && x.trashed).length, 1);
  assert.equal(pdfColumn(app)['2026-4-2351'], '済');
});

test('PDF個別: 選択した行だけを出力（出力済みでも出し直す）', () => {
  const { env, app } = bootWithReports();
  const dbId = env.props.DB_SPREADSHEET_ID;
  app.menuExportPendingPdfs();
  const before = env.fetches.length;
  env.ui.activeSpreadsheetId = dbId;
  env.spreadsheets[dbId].active = { sheet: '登録履歴', ranges: [[3, 1]] }; // 3行目 = 2026-4-2351
  app.menuExportSelectedPdfs();
  assert.match(env.ui.alerts.at(-2).msg, /1 件をPDFに出力します[\s\S]*2026年度第4回 山口宇部校（出力済→出し直し）/);
  assert.equal(env.fetches.length, before + 1);
  assert.equal(pdfFiles(env, 2026, 4).length, 2);
});

test('PDF個別: 複数範囲・見出し行を含む選択、別シートでの実行', () => {
  const { env, app } = bootWithReports();
  const dbId = env.props.DB_SPREADSHEET_ID;
  env.ui.activeSpreadsheetId = dbId;
  env.spreadsheets[dbId].active = { sheet: '登録履歴', ranges: [[1, 2], [4, 1]] }; // 見出し+2行目、4行目
  app.menuExportSelectedPdfs();
  assert.deepEqual(pdfColumn(app), { '2026-3-2351': '済', '2026-4-2351': '', '2026-4-9001': '済' });
  env.spreadsheets[dbId].active = { sheet: '明細', ranges: [[2, 1]] };
  app.menuExportSelectedPdfs();
  assert.match(env.ui.alerts.at(-1).msg, /「登録履歴」シートで出力したい行を選択/);
});

test('PDF個別: 選択範囲を取得できないときは行番号・レポートIDの入力で指定できる', () => {
  const { env, app } = bootWithReports();
  env.ui.promptText = '2, 2026-4-9001';
  app.menuExportSelectedPdfs();
  assert.equal(env.ui.prompts.length, 1);
  assert.deepEqual(pdfColumn(app), { '2026-3-2351': '済', '2026-4-2351': '', '2026-4-9001': '済' });
});

test('PDF: エラー時は「PDF出力」を空欄のまま、429 は待って再試行', () => {
  const { env, app } = bootWithReports();
  const fileOf = (id) => app.findReport_(id).fileId;
  env.files[fileOf('2026-3-2351')].trashed = true;
  env.fetchPlan.push(429, 200); // 1件目（第4回）: 1回目429→再試行で成功
  env.fetchPlan.push(429, 429, 429, 429); // 2件目: 4回とも429で失敗
  app.menuExportPendingPdfs();
  const msg = env.ui.alerts.at(-1).msg;
  assert.match(msg, /出力: 1 件/);
  assert.match(msg, /エラー: 2 件/);
  assert.match(msg, /ゴミ箱/);
  assert.match(msg, /HTTP 429/);
  assert.deepEqual(env.sleeps, [5000, 5000, 10000, 15000]);
  assert.deepEqual(pdfColumn(app), { '2026-3-2351': '', '2026-4-2351': '済', '2026-4-9001': '' });
});

test('PDF: 時間切れで中断した分は空欄のまま残り、再実行で処理される', () => {
  const { env, app } = bootWithReports();
  app.PDF_TIME_LIMIT_MS = -1;
  app.menuExportPendingPdfs();
  assert.match(env.ui.alerts.at(-1).msg, /時間切れで未処理: 3 件/);
  app.PDF_TIME_LIMIT_MS = 5 * 60 * 1000;
  app.menuExportPendingPdfs();
  assert.deepEqual(Object.values(pdfColumn(app)), ['済', '済', '済']);
});

test('PDF: 過去データ取込（Excel）の行も一時変換して出力し、一時ファイルを削除', () => {
  const { env, app } = boot();
  env.folders.imp = { id: 'imp', name: '過去分', parent: 'root' };
  const v = JSON.parse(JSON.stringify(templateValues));
  v[1][4] = 2351; v[1][6] = new Date(2026, 4, 20);
  env.files.p1 = { id: 'p1', name: 'SV山口宇部_2026春.xls', mime: 'application/vnd.ms-excel', parents: ['imp'], trashed: false,
    excelContent: { sheets: [{ name: 'チェック項目', data: v, formulas: {}, maxRows: 1000, maxCols: 26, gid: 777 }] } };
  env.props.IMPORT_FOLDER_ID = 'imp';
  app.importPastReports();
  app.menuExportPendingPdfs();
  assert.match(env.ui.alerts.at(-1).msg, /出力: 1 件/);
  assert.deepEqual(pdfFiles(env, 2026, 2).map((f) => f.name), ['SV山口宇部_2026春.pdf']);
  assert.ok(Object.values(env.files).filter((f) => f.name.startsWith('PDF作業_')).every((f) => f.trashed));
});

test('PDF: setup を再実行していない古い管理DB（見出しなし）でも見出しを補って出力できる', () => {
  const { env, app } = bootWithReports();
  const sh = env.spreadsheets[env.props.DB_SPREADSHEET_ID].sheets.find((s) => s.name === '登録履歴');
  sh.data[0][23] = '';
  app.menuExportPendingPdfs();
  assert.equal(sh.data[0][23], 'PDF出力');
  assert.deepEqual(Object.values(pdfColumn(app)), ['済', '済', '済']);
});
