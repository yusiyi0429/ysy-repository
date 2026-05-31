/**
 * 将 Luckysheet / jQuery 复制到 frontend/vendor，供内网/容器离线部署。
 * 用法：在 projects 目录执行  node scripts/copy-frontend-vendor.js
 */
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const FRONTEND = path.join(ROOT, 'frontend');
const NM = path.join(FRONTEND, 'node_modules');

function rmrf(dir) {
  if (!fs.existsSync(dir)) return;
  fs.rmSync(dir, { recursive: true, force: true });
}

function copyDir(src, dest) {
  fs.mkdirSync(dest, { recursive: true });
  for (const name of fs.readdirSync(src)) {
    const s = path.join(src, name);
    const d = path.join(dest, name);
    if (fs.statSync(s).isDirectory()) copyDir(s, d);
    else fs.copyFileSync(s, d);
  }
}

/** 删除 Luckysheet 包内应用不需要的演示/扩展文件 */
function pruneLuckysheetVendor(vendorRoot) {
  const removePaths = [
    'demoData',
    'expendPlugins',
    'index.html',
    'luckysheet.esm.js',
    'luckysheet.esm.js.map',
    'luckysheet.umd.js.map',
    path.join('assets', 'iconfont', 'demo_index.html'),
    path.join('assets', 'iconfont', 'demo.css'),
  ];
  for (const rel of removePaths) {
    const p = path.join(vendorRoot, rel);
    rmrf(p);
  }
}

const luckysheetDist = path.join(NM, 'luckysheet', 'dist');
const jquerySrc = path.join(NM, 'jquery', 'dist', 'jquery.min.js');

if (!fs.existsSync(luckysheetDist)) {
  console.error('缺少 node_modules/luckysheet，请先在 frontend 目录执行：');
  console.error('  npm install luckysheet@2.1.13 jquery@3.6.4 --no-save');
  process.exit(1);
}
if (!fs.existsSync(jquerySrc)) {
  console.error('缺少 jquery，请先 npm install jquery@3.6.4');
  process.exit(1);
}

const vendorLuckysheet = path.join(FRONTEND, 'vendor', 'luckysheet');
const vendorJqueryDir = path.join(FRONTEND, 'vendor', 'jquery');

rmrf(vendorLuckysheet);
copyDir(luckysheetDist, vendorLuckysheet);
pruneLuckysheetVendor(vendorLuckysheet);

fs.mkdirSync(vendorJqueryDir, { recursive: true });
fs.copyFileSync(jquerySrc, path.join(vendorJqueryDir, 'jquery.min.js'));

console.log('OK: frontend/vendor/luckysheet (pruned)');
console.log('OK: frontend/vendor/jquery/jquery.min.js');
