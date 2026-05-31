# 生成 frontend/vendor（Luckysheet + jQuery），供内网/容器离线部署
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root
Set-Location frontend
if (-not (Test-Path package.json)) { npm init -y | Out-Null }
npm install luckysheet@2.1.13 jquery@3.6.4 --no-save
Set-Location $Root
node scripts/copy-frontend-vendor.js
Write-Host "Done. Vendor files are under frontend/vendor/"
