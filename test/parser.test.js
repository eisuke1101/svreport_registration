const test = require('node:test');
const assert = require('node:assert/strict');
const load = require('./load');
const values = require('./fixtures/template_values.json');

const g = load('Domain.gs', 'Parser.gs');

test('テンプレートから重要事項3件＋通常89件を抽出', () => {
  const items = g.parseTemplateItems(values);
  assert.equal(items.filter(i => i.type === '重要').length, 3);
  assert.equal(items.filter(i => i.type === '通常').length, 89);
  assert.equal(items[0].key, 'J01');
  assert.equal(items[0].row, 6);
  assert.equal(items[0].penalties.L, -100);
  const i1 = items.find(i => i.key === 'I01');
  assert.equal(i1.row, 12);
  assert.equal(i1.category, '看板');
  assert.deepEqual({ ...i1.penalties }, { L: -2, M: -5, H: -10 });
  const last = items[items.length - 1];
  assert.equal(last.key, 'I89');
  assert.equal(last.row, 106);
});

test('「*」は選択不可(null)', () => {
  const i6 = g.parseTemplateItems(values).find(i => i.no === 6);
  assert.deepEqual({ ...i6.penalties }, { L: null, M: -3, H: null });
});

test('複数行にまたがる項目文を結合', () => {
  const i74 = g.parseTemplateItems(values).find(i => i.no === 74);
  assert.match(i74.text, /整合と\n適正な授業日程組み/);
  assert.match(i74.text, /※3名程度/);
  const j3 = g.parseTemplateItems(values).find(i => i.key === 'J03');
  assert.doesNotMatch(j3.text, /TMS現金出納帳/);
});

test('版と対象外業態の推定', () => {
  assert.equal(g.parseVersion(values), '2026.10');
  const welcome = g.parseTemplateItems(values).find(i => /Welcome/.test(i.text));
  assert.deepEqual([...g.guessExcludedBusinesses(welcome.text)], ['アスモ']);
});

test('過去レポートの読み取りとレベル推定', () => {
  const v = JSON.parse(JSON.stringify(values));
  v[1][4] = 2351; v[1][6] = '2026-07-15';
  v[11][10] = 2; v[11][11] = -10;   // No.1: 連続2回目、-5(M)×2
  v[12][11] = -2;                   // No.2: 初回 L
  const r = g.parseReportSheet(v);
  assert.equal(r.header.schoolId, '2351');
  const i1 = r.items.find(i => i.key === 'I01');
  assert.equal(i1.applied, -10);
  assert.equal(i1.count, 2);
  assert.equal(g.inferLevel(i1, i1.applied, i1.count), 'M');
  const i2 = r.items.find(i => i.key === 'I02');
  assert.equal(i2.count, 1);
  assert.equal(g.inferLevel(i2, i2.applied, i2.count), 'L');
});
