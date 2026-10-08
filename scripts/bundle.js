// src/*.gs を1つのファイル dist/SVレポート登録.gs にまとめる（GASエディタへの貼り付け用）
const fs = require('fs');
const path = require('path');
// Setup を先頭に置く（GASエディタの実行プルダウンは先頭の関数 setup が初期選択になる）
const order = ['Setup', 'Import', 'Config', 'Domain', 'Parser', 'Repository', 'ReportWriter', 'Notify', 'Code'];
const src = path.join(__dirname, '..', 'src');
const out = order.map((n) => `// ===== ${n}.gs =====\n` + fs.readFileSync(path.join(src, n + '.gs'), 'utf8')).join('\n');
fs.mkdirSync(path.join(__dirname, '..', 'dist'), { recursive: true });
const body = '// 自動生成ファイル（npm run bundle）。src/*.gs を編集すること。\n// 初期設定: 上部のプルダウンで「setup」を選んで ▷実行\n' + out;
const lines = body.split('\n').length + 1;
// 最終行の目印。エディタでこの行が見えていれば最後まで貼り付けられている
fs.writeFileSync(path.join(__dirname, '..', 'dist', 'SVレポート登録.gs'), body + `\n// ===== ファイル終端（全${lines}行）: この行まで貼り付けられていればOK =====\n`);
