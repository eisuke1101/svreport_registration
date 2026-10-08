const test = require('node:test');
const assert = require('node:assert/strict');
const load = require('./load');
const install = require('./gasmock');
const values = require('./fixtures/template_values.json');

function setupEnv() {
  const g = load('Domain.gs', 'Parser.gs', 'Config.gs', 'Repository.gs', 'Notify.gs', 'Code.gs');
  const { db, props } = install(g);
  Object.assign(props, { DB_SPREADSHEET_ID: 'db', TEMPLATE_SPREADSHEET_ID: 'tpl', CURRENT_ITEM_VERSION: '2026.10', NOTIFY_TO: 'boss@example.com' });
  Object.keys(g.SHEETS).forEach((k) => g.ensureSheet_(db, k));
  g.appendObjects_('schools', [{ '校舎ID': '2351', '業態': 'ITTO', '地域': '山口', '校舎名': '山口宇部校', '有効': true }]);
  g.appendObjects_('items', g.parseTemplateItems(values).map((i) => ({
    '版': '2026.10', '項目キー': i.key, 'No': String(i.no), '種別': i.type, '区分': i.category, '項目': i.text,
    '減点L': i.penalties.L ?? '', '減点M': i.penalties.M ?? '', '減点H': i.penalties.H ?? '', '出力行': i.row,
    '対象外業態': g.guessExcludedBusinesses(i.text).join(','), '有効': true
  })));
  // Drive・メール関連はスタブ
  g.written = [];
  g.mails = [];
  g.writeReportFile_ = (p) => { g.written.push(p); return { fileId: p.existingFileId || 'file-' + p.report.reportId, url: 'https://docs/' + p.report.reportId }; };
  g.savePhotos_ = (r, photos) => photos.map((ph, i) => ({ 'レポートID': r.reportId, '項目キー': ph.itemKey, 'ファイルID': 'ph' + i, 'ファイル名': ph.name, 'URL': 'u', '登録日時': '' }));
  g.trashPhotos_ = () => {};
  g.MailApp = { sendEmail: (m) => g.mails.push(m) };
  return { g, db };
}

const base = { schoolId: '2351', am: '山田', sv: '佐藤', cash: 1000, bank: 2000, tms: 3000 };

test('登録 → 前回・前々回の取得 → 連続指摘（未実施の回は飛ばす）', () => {
  const { g, db } = setupEnv();
  const r1 = g.apiSubmitReport({ ...base, submissionId: 'a', date: '2026-02-01', inputs: { I01: { level: 'M' } } });
  assert.equal(r1.status, 'ok');
  assert.equal(r1.score, 95);
  assert.equal(r1.warning, '');
  // 第2回は未実施。第3回で I01 を L 指摘 → 連続2回目
  const r3 = g.apiSubmitReport({ ...base, submissionId: 'b', date: '2026-08-01', inputs: { I01: { level: 'L' }, J01: { level: 'X' } } });
  assert.equal(r3.total, -2 * 2 - 100);
  const ctx = g.apiGetContext('2351', '2026-11-01');
  assert.equal(ctx.reportId, '2026-4-2351');
  assert.equal(ctx.prev.reportId, '2026-3-2351');
  assert.equal(ctx.prev2.reportId, '2026-1-2351');
  assert.equal(ctx.prev.lines.I01.applied, -4);
  assert.equal(ctx.prev.lines.I01.count, 2);
  // 出力に渡した前回・前々回
  const out = g.written[1];
  assert.equal(out.prevLines.I01.applied, -5);
  assert.deepEqual(Object.keys(out.prev2Lines), []);
  // 年度をまたいで連続: 2027年第1回
  g.apiSubmitReport({ ...base, submissionId: 'c', date: '2026-11-01', inputs: { I01: { level: 'H' } } });
  const r5 = g.apiSubmitReport({ ...base, submissionId: 'd', date: '2027-01-15', inputs: { I01: { level: 'H' } } });
  assert.equal(r5.total, -10 * 4);
  const r6 = g.apiSubmitReport({ ...base, submissionId: 'e', date: '2027-04-15', inputs: { I01: { level: 'L' } } });
  assert.equal(r6.total, -2 * 4); // 5回目以降も×4
  assert.equal(db.getSheetByName('明細').getLastRow() - 1, 92 * 5);
  assert.equal(g.mails.length, 5);
  assert.match(g.mails[0].subject, /2026年度第1回 山口宇部校（2351） 95点/);
});

test('同じ実施回の再登録は上書き確認、上書きで明細を置換', () => {
  const { g, db } = setupEnv();
  g.apiSubmitReport({ ...base, submissionId: 'a', date: '2026-10-01', inputs: { I01: { level: 'M' } } });
  const again = g.apiSubmitReport({ ...base, submissionId: 'b', date: '2026-12-20', inputs: { I02: { level: 'H' } } });
  assert.equal(again.status, 'exists');
  const ow = g.apiSubmitReport({ ...base, submissionId: 'b', date: '2026-12-20', inputs: { I02: { level: 'H' } }, overwrite: true });
  assert.equal(ow.status, 'ok');
  assert.equal(ow.updated, true);
  assert.equal(g.written[1].existingFileId, 'file-2026-4-2351'); // 同じファイルを更新
  const reports = db.getSheetByName('登録履歴');
  assert.equal(reports.getLastRow(), 2);
  assert.equal(reports.getRange(2, 18).getValues()[0][0], 2); // 版数
  const lines = g.getLines_('2026-4-2351');
  assert.equal(lines.I01.applied, 0);
  assert.equal(lines.I02.applied, -10);
  assert.equal(db.getSheetByName('明細').getLastRow() - 1, 92);
});

test('同じ送信IDの再送は二重登録しない', () => {
  const { g } = setupEnv();
  g.apiSubmitReport({ ...base, submissionId: 'same', date: '2026-10-01', inputs: {} });
  const dup = g.apiSubmitReport({ ...base, submissionId: 'same', date: '2026-10-01', inputs: {} });
  assert.equal(dup.duplicate, true);
  assert.equal(g.written.length, 1);
});

test('必須項目（現金・通帳残・TMS金額 等）の検証', () => {
  const { g } = setupEnv();
  assert.throws(() => g.apiSubmitReport({ ...base, tms: '', date: '2026-10-01', inputs: {} }), /TMS金額/);
  assert.throws(() => g.apiSubmitReport({ ...base, sv: ' ', date: '2026-10-01', inputs: {} }), /SV/);
});

test('写真の記録と出力', () => {
  const { g } = setupEnv();
  g.apiSubmitReport({ ...base, submissionId: 'p', date: '2026-10-01', inputs: {}, photos: [{ itemKey: 'I05', itemNo: '5', name: 'a.jpg', dataUrl: 'data:image/jpeg;base64,AA==' }] });
  assert.equal(g.written[0].photos.length, 1);
  assert.equal(g.getPhotos_('2026-4-2351')[0].itemKey, 'I05');
});
