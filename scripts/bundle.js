// src/*.gs を1つのファイル dist/SVレポート登録.gs にまとめる（GASエディタへの貼り付け用）
const fs = require('fs');
const path = require('path');
const order = ['Config', 'Domain', 'Parser', 'Repository', 'ReportWriter', 'Notify', 'Code', 'Setup', 'Import'];
const src = path.join(__dirname, '..', 'src');
const out = order.map((n) => `// ===== ${n}.gs =====\n` + fs.readFileSync(path.join(src, n + '.gs'), 'utf8')).join('\n');
fs.mkdirSync(path.join(__dirname, '..', 'dist'), { recursive: true });
fs.writeFileSync(path.join(__dirname, '..', 'dist', 'SVレポート登録.gs'), '// 自動生成ファイル（npm run bundle）。src/*.gs を編集すること。\n' + out);
