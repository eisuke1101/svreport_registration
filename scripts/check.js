// 実装後のエラーチェック一式: npm run check
//  1. dist/SVレポート登録.gs を作り直す
//  2. 各 .gs と dist、画面の <script> の構文チェック
//  3. トップレベルの関数・変数名の重複チェック（GAS は全ファイルが同じ空間に読み込まれるため）
//  4. 必須のエントリポイント（setup / doGet / api* など）の存在チェック
//  5. 全テスト（単体・登録処理・GASモック環境での通し試験）
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { execSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const SRC = path.join(ROOT, 'src');
let failed = false;
const ok = (msg) => console.log('  OK  ' + msg);
const ng = (msg) => { console.log('  NG  ' + msg); failed = true; };

console.log('[1] 1ファイル版の生成');
execSync('node scripts/bundle.js', { cwd: ROOT, stdio: 'inherit' });
const bundle = fs.readFileSync(path.join(ROOT, 'dist', 'SVレポート登録.gs'), 'utf8');
const m = /全(\d+)行/.exec(bundle.trim().split('\n').pop());
m && Number(m[1]) === bundle.trim().split('\n').length ? ok(`dist/SVレポート登録.gs（${m[1]}行、終端の目印あり）`) : ng('終端の目印の行数が一致しません');

console.log('[2] 構文チェック');
const gsFiles = fs.readdirSync(SRC).filter((f) => f.endsWith('.gs'));
for (const f of gsFiles) {
  try { new vm.Script(fs.readFileSync(path.join(SRC, f), 'utf8'), { filename: f }); ok(f); } catch (e) { ng(`${f}: ${e.message}`); }
}
try { new vm.Script(bundle, { filename: 'bundle' }); ok('dist/SVレポート登録.gs'); } catch (e) { ng(`dist: ${e.message}`); }
const js = fs.readFileSync(path.join(SRC, 'js.html'), 'utf8').match(/<script>([\s\S]*)<\/script>/)[1];
try { new vm.Script(js, { filename: 'js.html' }); ok('js.html'); } catch (e) { ng(`js.html: ${e.message}`); }
try { JSON.parse(fs.readFileSync(path.join(SRC, 'appsscript.json'), 'utf8')); ok('appsscript.json'); } catch (e) { ng(`appsscript.json: ${e.message}`); }

console.log('[3] 名前の重複');
const seen = {};
for (const f of gsFiles) {
  const text = fs.readFileSync(path.join(SRC, f), 'utf8');
  for (const m of text.matchAll(/^(?:function\s+([A-Za-z0-9_$]+)|var\s+([A-Za-z0-9_$]+))/gm)) {
    const name = m[1] || m[2];
    if (seen[name]) ng(`${name} が ${seen[name]} と ${f} の両方にあります`); else seen[name] = f;
  }
}
if (!failed) ok(`${Object.keys(seen).length} 個の名前に重複なし`);

console.log('[4] エントリポイント');
const ctx = vm.createContext({});
vm.runInContext(bundle, ctx);
['setup', 'updateTemplate', 'importPastReports', 'onDbOpen', 'menuExportPendingPdfs', 'menuExportSelectedPdfs', 'doGet', 'include', 'sharedSource',
  'apiGetBootstrap', 'apiGetContext', 'apiSubmitReport'].forEach((n) => (typeof ctx[n] === 'function' ? ok(n) : ng(`${n} がありません`)));
const fnLine = bundle.split('\n').findIndex((l) => /^function\s/.test(l));
/^function setup\(/.test(bundle.split('\n')[fnLine]) ? ok('先頭の関数は setup（エディタの実行プルダウン初期値）') : ng('先頭の関数が setup ではありません');
['index', 'css', 'js'].forEach((n) => (fs.existsSync(path.join(SRC, n + '.html')) ? ok(n + '.html') : ng(n + '.html がありません')));

console.log('[5] テスト');
try {
  const out = execSync('node --test test/*.test.js', { cwd: ROOT, encoding: 'utf8', shell: '/bin/bash' });
  const pass = /# pass (\d+)/.exec(out)[1];
  ok(`全 ${pass} 件成功`);
} catch (e) {
  console.log(e.stdout);
  ng('テスト失敗');
}

console.log(failed ? '\n結果: NG' : '\n結果: すべてOK');
process.exit(failed ? 1 : 0);
