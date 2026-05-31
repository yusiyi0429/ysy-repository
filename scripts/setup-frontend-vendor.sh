#!/bin/bash
# 生成 frontend/vendor（Luckysheet + jQuery），供内网/容器离线部署
set -e
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT/frontend"
npm install luckysheet@2.1.13 jquery@3.6.4 --no-save
cd "$ROOT"
node scripts/copy-frontend-vendor.js
echo "Done. Vendor files are under frontend/vendor/"
