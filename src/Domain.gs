/**
 * 業務ルール（純粋関数）。
 * サーバ側の確定計算と、画面側のリアルタイム計算（sharedSource() で注入）の両方で使う。
 * GAS のサービス（SpreadsheetApp 等）をここで呼ばないこと。
 */

/** 連続指摘の倍率上限（5回目以降は4回目と同じ） */
var MAX_MULTIPLIER = 4;

/** 重要事項の「該当」を表すレベル値 */
var LEVEL_IMPORTANT = 'X';

/**
 * 実施日（yyyy-mm-dd）から年度・実施回を求める。
 * 第1回: 1〜3月、第2回: 4〜6月、第3回: 7〜9月、第4回: 10〜12月（年度＝暦年）。
 */
function periodOf(dateStr) {
  var m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(String(dateStr || ''));
  if (!m) throw new Error('実施日の形式が不正です: ' + dateStr);
  var year = Number(m[1]);
  var month = Number(m[2]);
  if (month < 1 || month > 12) throw new Error('実施日の月が不正です: ' + dateStr);
  return { year: year, round: Math.floor((month - 1) / 3) + 1 };
}

/** レポートID = {年度}-{実施回}-{校舎ID} */
function reportIdOf(year, round, schoolId) {
  return year + '-' + round + '-' + String(schoolId).trim();
}

/** 実施回の前後比較用の序数（年度をまたいで比較できる） */
function ordinalOf(year, round) {
  return Number(year) * 10 + Number(round);
}

/** 項目がその業態で対象外か */
function isExcludedFor(item, business) {
  return !!business && (item.excludedBiz || []).indexOf(String(business)) >= 0;
}

/** 項目で選択できるレベルの一覧 [{level, penalty}] */
function levelsOf(item) {
  if (item.type === '重要') {
    return item.penalties.L != null ? [{ level: LEVEL_IMPORTANT, penalty: item.penalties.L }] : [];
  }
  return ['L', 'M', 'H']
    .filter(function (lv) { return item.penalties[lv] != null; })
    .map(function (lv) { return { level: lv, penalty: item.penalties[lv] }; });
}

/** レベルに対応する基本減点。選択不可なら null */
function basePenaltyOf(item, level) {
  var hit = levelsOf(item).filter(function (x) { return x.level === level; })[0];
  return hit ? hit.penalty : null;
}

/**
 * 1項目の今回結果を計算する。
 * - 今回減点あり かつ 前回（直近の実施済みレポート）でも減点あり → 連続回数 = 前回の連続回数 + 1
 *   （レベル違いでも連続。間の回が未実施でも直近の実施済みレポートを参照するためリセットされない）
 * - 適用減点 = 基本減点 × min(連続回数, 4)
 * @param {Object} item 項目マスタ
 * @param {Object} input {level, na, note}
 * @param {Object} prevLine 前回の明細 {applied, count}（無ければ null）
 * @param {string} business 校舎の業態
 */
function computeLine(item, input, prevLine, business) {
  input = input || {};
  var excluded = isExcludedFor(item, business);
  var line = {
    key: item.key, level: '', base: 0, count: 0, applied: 0,
    na: excluded || !!input.na, forcedNa: excluded,
    note: input.note ? String(input.note) : ''
  };
  if (line.na || !input.level) return line;

  var base = basePenaltyOf(item, input.level);
  if (base == null) {
    throw new Error('項目 ' + item.no + ' ではレベル「' + input.level + '」を選択できません');
  }
  if (!(base < 0)) return line;

  var prevApplied = prevLine ? Number(prevLine.applied) || 0 : 0;
  var prevCount = prevApplied < 0 ? (Number(prevLine.count) || 1) : 0;
  line.level = input.level;
  line.base = base;
  line.count = prevCount + 1;
  line.applied = base * Math.min(line.count, MAX_MULTIPLIER);
  return line;
}

/**
 * レポート全体を計算する。
 * @param {Object[]} items 項目マスタ（並び順どおり）
 * @param {Object} inputs {項目キー: {level, na, note}}
 * @param {Object} prevLines {項目キー: {applied, count}}（前回レポートの明細）
 * @param {string} business 校舎の業態
 */
function computeReport(items, inputs, prevLines, business) {
  inputs = inputs || {};
  prevLines = prevLines || {};
  var lines = items.map(function (item) {
    return computeLine(item, inputs[item.key], prevLines[item.key], business);
  });
  var total = lines.reduce(function (sum, l) { return sum + l.applied; }, 0);
  return { lines: lines, total: total, score: 100 + total };
}
