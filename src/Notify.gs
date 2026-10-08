/**
 * 登録通知（固定の送信先1件。スクリプトプロパティ NOTIFY_TO）
 */
function notifyRegistered_(report, items, lines, isUpdate) {
  var to = getConfig_().notifyTo;
  if (!to) return false;
  var itemByKey = {};
  items.forEach(function (i) { itemByKey[i.key] = i; });
  var hits = lines.filter(function (l) { return l.applied < 0; });

  var subject = '【SVレポート' + (isUpdate ? '更新' : '登録') + '】' + report.year + '年度第' + report.round + '回 '
    + report.schoolName + '（' + report.schoolId + '） ' + report.score + '点';
  var body = [
    'SVレポートが' + (isUpdate ? '更新' : '登録') + 'されました。',
    '',
    '校舎　　: ' + report.schoolName + '（' + report.schoolId + ' / ' + report.business + '）',
    '実施回　: ' + report.year + '年度 第' + report.round + '回',
    '実施日　: ' + report.date,
    'AM / SV : ' + report.am + ' / ' + report.sv,
    '点数　　: ' + report.score + '点（減点 ' + report.total + '）',
    '登録者　: ' + report.registeredBy,
    '',
    '■指摘項目（' + hits.length + '件）'
  ].concat(hits.map(function (l) {
    var item = itemByKey[l.key];
    return '・No.' + item.no + ' ' + item.text.split('\n')[0] + ' ' + l.applied + (l.count >= 2 ? '（連続' + l.count + '回目）' : '');
  })).concat(['', 'レポート: ' + report.url]).join('\n');

  MailApp.sendEmail({ to: to, subject: subject, body: body, name: 'SVレポート登録' });
  return true;
}
