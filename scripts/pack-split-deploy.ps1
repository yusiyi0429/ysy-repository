# 打包内网部署文件并切分为 <=50KB 分片
# 用法: .\scripts\pack-split-deploy.ps1 [-ChunkSizeKB 50] [-OutputDir "$env:USERPROFILE\Desktop\split0531"]

param(
    [int]$ChunkSizeKB = 50,
    [string]$OutputDir = ""
)

$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $PSScriptRoot
if (-not $OutputDir) {
    $OutputDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "split0531"
}

$ChunkSize = $ChunkSizeKB * 1024
$Staging = Join-Path $env:TEMP "tacit-deploy-staging-$(Get-Date -Format 'yyyyMMddHHmmss')"
$BundleName = "tacit-knowledge-deploy-2.0.0-arm64"
$ZipPath = Join-Path $env:TEMP "$BundleName.zip"

if (Test-Path $OutputDir) {
    Remove-Item $OutputDir -Recurse -Force
}
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
New-Item -ItemType Directory -Path $Staging -Force | Out-Null

Write-Host "=== 打包内网部署包 ===" -ForegroundColor Cyan
Write-Host "源项目: $ProjectDir"
Write-Host "输出目录: $OutputDir"
Write-Host "分片大小: ${ChunkSizeKB} KB"
Write-Host ""

# 复制部署所需内容（排除大体积/无关目录）
$excludeDirs = @('.git', 'node_modules', '.cursor', '__pycache__', 'workspace', 'logs', 'split0531', 'split0525', 'split0526', 'split0527')
$excludeFiles = @('*.tar', '*.pyc', '*.log', 'debug-*.log')

robocopy $ProjectDir $Staging /E /XD $excludeDirs /XF $excludeFiles /NFL /NDL /NJH /NJS /nc /ns /np | Out-Null

# 若存在镜像 tar，一并打入包（可能很大，分片会很多）
foreach ($tar in Get-ChildItem $ProjectDir -Filter "tacit-knowledge-externalization*.tar" -ErrorAction SilentlyContinue) {
    Copy-Item $tar.FullName (Join-Path $Staging $tar.Name) -Force
    Write-Host "已纳入镜像: $($tar.Name) ($([math]::Round($tar.Length/1MB,1)) MB)"
}

# 写入分片说明
$readme = @"
# 隐性知识显性化 — 内网传输分片包

## 文件说明
- ${BundleName}.zip.part0000 ... : 压缩包分片（每片 <= ${ChunkSizeKB} KB）
- ${BundleName}.zip.manifest.json : 分片清单与校验
- REASSEMBLE.sh : Linux 内网合并脚本
- REASSEMBLE.ps1 : Windows 合并脚本

## Linux 内网合并
cd /data/隐性知识显性化
chmod +x REASSEMBLE.sh
./REASSEMBLE.sh

## 合并后部署
unzip ${BundleName}.zip -d deploy
cd deploy
docker load -i tacit-knowledge-externalization-*.tar   # 若有镜像文件
docker compose -p tacit-knowledge up -d
curl http://127.0.0.1:5000/api/version
"@
Set-Content -Path (Join-Path $Staging "SPLIT_README.md") -Value $readme -Encoding UTF8

if (Test-Path $ZipPath) { Remove-Item $ZipPath -Force }
Compress-Archive -Path (Join-Path $Staging '*') -DestinationPath $ZipPath -CompressionLevel Optimal
$zipSize = (Get-Item $ZipPath).Length
Write-Host "压缩包: $ZipPath ($([math]::Round($zipSize/1MB,2)) MB)"

# 切分
$partNum = 0
$fs = [System.IO.File]::OpenRead($ZipPath)
$buffer = New-Object byte[] $ChunkSize
try {
    while ($true) {
        $read = $fs.Read($buffer, 0, $ChunkSize)
        if ($read -le 0) { break }
        $partName = "${BundleName}.zip.part{0:D4}" -f $partNum
        $partPath = Join-Path $OutputDir $partName
        if ($read -eq $ChunkSize) {
            [System.IO.File]::WriteAllBytes($partPath, $buffer)
        } else {
            $chunk = $buffer[0..($read - 1)]
            [System.IO.File]::WriteAllBytes($partPath, $chunk)
        }
        $partNum++
    }
} finally {
    $fs.Close()
}

$parts = Get-ChildItem $OutputDir -Filter "${BundleName}.zip.part*"
$manifest = @{
    bundle_name   = "$BundleName.zip"
    app_version   = (Get-Content (Join-Path $ProjectDir "VERSION") -ErrorAction SilentlyContinue | Select-Object -First 1)
    chunk_size_kb = $ChunkSizeKB
    chunk_count   = $parts.Count
    total_bytes   = $zipSize
    parts         = @($parts | ForEach-Object {
        @{
            file = $_.Name
            bytes = $_.Length
            sha256 = (Get-FileHash $_.FullName -Algorithm SHA256).Hash
        }
    })
}
$manifestPath = Join-Path $OutputDir "${BundleName}.zip.manifest.json"
$manifest | ConvertTo-Json -Depth 5 | Set-Content $manifestPath -Encoding UTF8

# 合并脚本
$reassembleSh = @'
#!/bin/bash
set -euo pipefail
BUNDLE="tacit-knowledge-deploy-2.0.0-arm64.zip"
OUT="${BUNDLE}"
PARTS=( $(ls -1 ${BUNDLE}.part* 2>/dev/null | sort) )
if [ ${#PARTS[@]} -eq 0 ]; then
  echo "未找到分片 ${BUNDLE}.part*"
  exit 1
fi
echo "合并 ${#PARTS[@]} 个分片 -> $OUT"
cat "${PARTS[@]}" > "$OUT"
unzip -o "$OUT" -d deploy_extracted
echo "完成。部署目录: $(pwd)/deploy_extracted"
echo "请进入 deploy_extracted 执行 docker load / docker compose"
'@
Set-Content (Join-Path $OutputDir "REASSEMBLE.sh") -Value $reassembleSh -Encoding UTF8NoBOM

$reassemblePs = @'
$Bundle = "tacit-knowledge-deploy-2.0.0-arm64.zip"
$parts = Get-ChildItem "$Bundle.part*" | Sort-Object Name
if (-not $parts) { throw "未找到分片" }
$out = Join-Path $PSScriptRoot $Bundle
$fs = [IO.File]::Create($out)
foreach ($p in $parts) {
  $bytes = [IO.File]::ReadAllBytes($p.FullName)
  $fs.Write($bytes, 0, $bytes.Length)
}
$fs.Close()
Write-Host "已合并: $out"
Expand-Archive -Path $out -DestinationPath (Join-Path $PSScriptRoot "deploy_extracted") -Force
'@
Set-Content (Join-Path $OutputDir "REASSEMBLE.ps1") -Value $reassemblePs -Encoding UTF8

Copy-Item $manifestPath (Join-Path $OutputDir "MANIFEST.json") -Force
Copy-Item (Join-Path $Staging "SPLIT_README.md") (Join-Path $OutputDir "README.md") -Force

Remove-Item $Staging -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $ZipPath -Force -ErrorAction SilentlyContinue

$maxPart = ($parts | Measure-Object -Property Length -Maximum).Maximum
Write-Host ""
Write-Host "完成" -ForegroundColor Green
Write-Host "  分片数: $($parts.Count)"
Write-Host "  最大分片: $([math]::Round($maxPart/1KB,1)) KB (限制 ${ChunkSizeKB} KB)"
Write-Host "  目录: $OutputDir"
