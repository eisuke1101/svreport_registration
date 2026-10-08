const test = require('node:test');
const assert = require('node:assert/strict');
const load = require('./load');

const g = load('Domain.gs');
const item = { key: 'I12', no: 12, type: '通常', penalties: { L: -1, M: -3, H: -5 }, excludedBiz: ['アスモ'] };
const imp = { key: 'J01', no: '重1', type: '重要', penalties: { L: -100, M: null, H: null } };
const starM = { key: 'I06', no: 6, type: '通常', penalties: { L: null, M: -3, H: null } };

test('実施回: 1〜3月=第1回、以降3か月単位', () => {
  assert.deepEqual({ ...g.periodOf('2026-01-01') }, { year: 2026, round: 1 });
  assert.deepEqual({ ...g.periodOf('2026-03-31') }, { year: 2026, round: 1 });
  assert.deepEqual({ ...g.periodOf('2026-04-01') }, { year: 2026, round: 2 });
  assert.deepEqual({ ...g.periodOf('2026-09-30') }, { year: 2026, round: 3 });
  assert.deepEqual({ ...g.periodOf('2026-10-08') }, { year: 2026, round: 4 });
  assert.throws(() => g.periodOf('2026/10/08'));
});

test('年度をまたいで序数で比較できる', () => {
  assert.ok(g.ordinalOf(2025, 4) < g.ordinalOf(2026, 1));
  assert.equal(g.reportIdOf(2026, 4, 2351), '2026-4-2351');
});

test('初回指摘は倍率1', () => {
  const l = g.computeLine(item, { level: 'M' }, null, 'ITTO');
  assert.equal(l.count, 1);
  assert.equal(l.applied, -3);
});

test('レベルが違っても連続扱い、適用減点は今回の基本減点×回数', () => {
  const l = g.computeLine(item, { level: 'H' }, { applied: -3, count: 1 }, 'ITTO');
  assert.equal(l.count, 2);
  assert.equal(l.applied, -10);
});

test('5回目以降は×4', () => {
  assert.equal(g.computeLine(item, { level: 'L' }, { applied: -3, count: 3 }, 'ITTO').applied, -4);
  const l5 = g.computeLine(item, { level: 'L' }, { applied: -4, count: 4 }, 'ITTO');
  assert.equal(l5.count, 5);
  assert.equal(l5.applied, -4);
});

test('前回が減点なしならリセット', () => {
  assert.equal(g.computeLine(item, { level: 'L' }, { applied: 0, count: 0 }, 'ITTO').count, 1);
});

test('取込データで連続回数が空でも、減点ありなら1回扱い', () => {
  assert.equal(g.computeLine(item, { level: 'L' }, { applied: -3, count: '' }, 'ITTO').count, 2);
});

test('重要事項も倍率の対象', () => {
  const l = g.computeLine(imp, { level: 'X' }, { applied: -100, count: 1 }, 'ITTO');
  assert.equal(l.applied, -200);
});

test('選択不可レベル（*）はエラー', () => {
  assert.throws(() => g.computeLine(starM, { level: 'L' }, null, 'ITTO'));
  assert.equal(g.computeLine(starM, { level: 'M' }, null, 'ITTO').applied, -3);
});

test('業態で対象外の項目は減点されない', () => {
  const l = g.computeLine(item, { level: 'H' }, null, 'アスモ');
  assert.equal(l.na, true);
  assert.equal(l.forcedNa, true);
  assert.equal(l.applied, 0);
});

test('手動の対象外', () => {
  const l = g.computeLine(item, { level: 'H', na: true, note: '併設なし' }, null, 'ITTO');
  assert.equal(l.applied, 0);
  assert.equal(l.note, '併設なし');
});

test('レポート合計と点数', () => {
  const r = g.computeReport([item, imp], { I12: { level: 'M' }, J01: { level: 'X' } }, { J01: { applied: -100, count: 1 } }, 'ITTO');
  assert.equal(r.total, -203);
  assert.equal(r.score, -103);
});
