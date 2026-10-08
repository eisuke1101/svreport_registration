// GAS の .gs ファイル（純粋関数部分）を Node で読み込むためのヘルパー
const fs = require('fs');
const path = require('path');
const vm = require('vm');

module.exports = function load(...files) {
  const ctx = vm.createContext({});
  for (const f of files) {
    vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'src', f), 'utf8'), ctx, { filename: f });
  }
  return ctx;
};
