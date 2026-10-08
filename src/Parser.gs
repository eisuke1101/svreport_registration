/**
 * SVレポート（チェック項目シート）の解析（純粋関数）。
 * values は sheet.getRange('A1:O{最終行}').getValues() の2次元配列。
 */

var SHEET_LAYOUT = {
  firstItemRow: 6,     // 重要事項の先頭行
  col: { no: 1, category: 2, text: 3, penaltyL: 8, penaltyM: 9, penaltyH: 10, note: 6, count: 11, current: 12, prev: 14, prev2: 15 },
  cell: {
    version: 'O1', business: 'B2', schoolName: 'C2', schoolId: 'E2', date: 'G2',
    am: 'M3', sv: 'M5', cash: 'E8', bank: 'E9', tms: 'G10'
  }
};

function cellAt_(values, a1) {
  var m = /^([A-Z]+)(\d+)$/.exec(a1);
  var col = 0;
  for (var i = 0; i < m[1].length; i++) col = col * 26 + (m[1].charCodeAt(i) - 64);
  var row = values[Number(m[2]) - 1];
  return row ? row[col - 1] : '';
}

function penaltyValue_(v) {
  if (v === '' || v === null || v === undefined || v === '*') return null;
  var n = Number(v);
  return isNaN(n) ? null : n;
}

/** テンプレート版（例: '2026.10ver.' → '2026.10'） */
function parseVersion(values) {
  var v = String(cellAt_(values, SHEET_LAYOUT.cell.version) || '').trim();
  return v.replace(/ver\.?$/i, '').trim();
}

/** 項目文の比較用正規化（過去ファイル取込時の照合に使う） */
function normalizeItemText(s) {
  return String(s || '').replace(/[\s　■※・、。，．,.（）()]/g, '').toLowerCase();
}

/**
 * チェック項目を抽出する。
 * - A列に番号がある行 → 通常項目（No, 区分, 項目, L/M/H減点）
 * - 番号が無く C列が「■」で始まり H列に減点がある行（6〜8行）→ 重要事項
 * - C列が空白・全角空白で始まる行 → 直前の項目文の続き
 * H列が「減点total」の行で終了。
 * @return {Object[]} [{key, no, type, category, text, penalties:{L,M,H}, row}]
 */
function parseTemplateItems(values) {
  var c = SHEET_LAYOUT.col;
  var items = [];
  var category = '';
  var importantNo = 0;
  for (var r = SHEET_LAYOUT.firstItemRow; r <= values.length; r++) {
    var row = values[r - 1];
    if (String(row[c.penaltyL - 1]).indexOf('減点total') >= 0) break;
    if (row[c.category - 1]) category = String(row[c.category - 1]).trim();
    var text = String(row[c.text - 1] || '');
    var no = row[c.no - 1];

    if (no !== '' && no !== null && !isNaN(Number(no))) {
      items.push({
        key: 'I' + ('0' + Number(no)).slice(-2), no: Number(no), type: '通常', category: category,
        text: text.trim(), row: r,
        penalties: { L: penaltyValue_(row[c.penaltyL - 1]), M: penaltyValue_(row[c.penaltyM - 1]), H: penaltyValue_(row[c.penaltyH - 1]) }
      });
    } else if (/^■/.test(text) && penaltyValue_(row[c.penaltyL - 1]) != null) {
      importantNo++;
      items.push({
        key: 'J' + ('0' + importantNo).slice(-2), no: '重' + importantNo, type: '重要', category: category || '重要事項',
        text: text.trim(), row: r,
        penalties: { L: penaltyValue_(row[c.penaltyL - 1]), M: null, H: null }
      });
    } else if (/^[\s　]/.test(text) && text.trim() && items.length) {
      items[items.length - 1].text += '\n' + text.trim();
    }
  }
  return items;
}

/** 項目文の注記から対象外業態を推定する（初期値。マスタで修正可能） */
function guessExcludedBusinesses(text) {
  var result = [];
  var re = /([^\s（(、・]+)は対象外/g;
  var m;
  while ((m = re.exec(text))) result.push(m[1]);
  return result;
}

/**
 * 登録済みのSVレポート（過去ファイル）を読み取る。
 * @return {{header:Object, items:Object[]}} items は parseTemplateItems の結果に count/applied を付けたもの
 */
function parseReportSheet(values) {
  var cell = SHEET_LAYOUT.cell;
  var c = SHEET_LAYOUT.col;
  var header = {
    version: parseVersion(values),
    business: String(cellAt_(values, cell.business) || '').trim(),
    schoolName: String(cellAt_(values, cell.schoolName) || '').trim(),
    schoolId: String(cellAt_(values, cell.schoolId) || '').trim().replace(/\.0$/, ''),
    date: cellAt_(values, cell.date),
    am: String(cellAt_(values, cell.am) || '').trim(),
    sv: String(cellAt_(values, cell.sv) || '').trim(),
    cash: cellAt_(values, cell.cash),
    bank: cellAt_(values, cell.bank),
    tms: cellAt_(values, cell.tms)
  };
  var items = parseTemplateItems(values).map(function (item) {
    var row = values[item.row - 1];
    var applied = Number(row[c.current - 1]) || 0;
    var count = Number(row[c.count - 1]) || (applied < 0 ? 1 : 0);
    item.applied = applied;
    item.count = applied < 0 ? count : 0;
    item.note = String(row[c.note - 1] || '').trim();
    return item;
  });
  return { header: header, items: items };
}

/** 適用減点と連続回数から元のレベルを推定する（一致しなければ ''） */
function inferLevel(item, applied, count) {
  if (!(applied < 0)) return '';
  var base = applied / Math.min(count || 1, MAX_MULTIPLIER);
  var hit = levelsOf(item).filter(function (x) { return x.penalty === base; })[0];
  return hit ? hit.level : '';
}

/**
 * 校舎一覧シートを読み取る（3行目以降、A列=ID）。
 * @return {Object[]} [{id, business, region, name, management, company, openDate, remarks}]
 */
function parseSchoolList(values) {
  var list = [];
  for (var r = 2; r < values.length; r++) {
    var row = values[r];
    if (row[0] === '' || row[0] === null || isNaN(Number(row[0]))) continue;
    list.push({
      id: String(Number(row[0])), business: String(row[1] || '').trim(), region: String(row[3] || '').trim(),
      name: String(row[4] || '').trim(), management: String(row[6] || '').trim(), company: String(row[7] || '').trim(),
      openDate: row[8], remarks: String(row[9] || '').trim()
    });
  }
  return list;
}
