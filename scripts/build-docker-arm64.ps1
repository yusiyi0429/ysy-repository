# 构建 linux/arm64 Docker 镜像并导出离线 tar 包（内网部署）
# 用法: .\scripts\build-docker-arm64.ps1
# 可选: .\scripts\build-docker-arm64.ps1 -AppVersion 2.0.1

param(
    [string]$AppVersion = "",
    [string]$OutputDir = ""
)

$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectDir

$ImageName = "tacit-knowledge-externalization"
$Platform = "linux/arm64"
$PlatformSlug = "arm64"
$DateTag = Get-Date -Format "yyyyMMdd-HHmm"
$UtcNow = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

$VersionFile = Join-Path $ProjectDir "VERSION"
if (-not $AppVersion) {
    if (-not (Test-Path $VersionFile)) {
        Write-Error "缺少 VERSION 文件，请创建或使用 -AppVersion 参数"
    }
    $AppVersion = (Get-Content $VersionFile -TotalCount 1).Trim()
}
if (-not $AppVersion) {
    Write-Error "APP 版本号为空"
}

$VersionTag = "${AppVersion}-${PlatformSlug}-${DateTag}"
$TarBaseName = "${ImageName}-${VersionTag}"
if (-not $OutputDir) {
    $OutputDir = $ProjectDir
}
$TarFile = Join-Path $OutputDir "${TarBaseName}.tar"
$TarLatest = Join-Path $ProjectDir "tacit-knowledge-externalization-arm64.tar"
$ManifestFile = Join-Path $OutputDir "${TarBaseName}.manifest.json"

Write-Host "=== 隐性知识显性化 ARM64 镜像构建 ===" -ForegroundColor Cyan
Write-Host "APP 版本:     $AppVersion"
Write-Host "镜像标签:     ${ImageName}:${VersionTag}"
Write-Host "目标平台:     $Platform"
Write-Host ""

if (-not (Get-Command docker -ErrorAction SilentlyContinue)) {
    Write-Error "未找到 docker 命令，请先安装 Docker Desktop 并启用 buildx"
}

docker info | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Error "Docker 未运行，请启动 Docker Desktop"
}

# 使用 desktop-linux（本机守护进程），可复用已拉取的 python 基础镜像，避免 container builder 重复访问 docker.io
docker buildx use desktop-linux 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    docker buildx use default 2>&1 | Out-Null
}

# 中文路径易导致 buildx 失败：复制到英文临时目录构建
$BuildDir = "C:\tacit-knowledge-docker-build"
if (Test-Path $BuildDir) { Remove-Item $BuildDir -Recurse -Force }
New-Item -ItemType Directory -Path $BuildDir | Out-Null
robocopy $ProjectDir $BuildDir /E /XD .git node_modules .cursor __pycache__ split0525 split0526 split0527 /XF *.tar *.pyc /NFL /NDL /NJH /NJS /nc /ns /np | Out-Null
Set-Location $BuildDir

Write-Host "开始构建镜像 ${ImageName}:${VersionTag} ..." -ForegroundColor Yellow
$buildArgs = @(
    "buildx", "build",
    "--platform", $Platform,
    "--build-arg", "APP_VERSION=$AppVersion",
    "--build-arg", "IMAGE_VERSION=$VersionTag",
    "--build-arg", "BUILD_DATE=$UtcNow",
    "--build-arg", "APP_PLATFORM=$Platform",
    "--tag", "${ImageName}:${VersionTag}",
    "--tag", "${ImageName}:${AppVersion}-arm64",
    "--file", "docker/Dockerfile",
    "--pull=false",
    "--load",
    "--progress=plain",
    "."
)
docker @buildArgs
if ($LASTEXITCODE -ne 0) {
    Write-Error "镜像构建失败"
}

Write-Host "导出离线包: $TarFile" -ForegroundColor Yellow
if (Test-Path $TarFile) { Remove-Item $TarFile -Force }
docker save "${ImageName}:${VersionTag}" -o $TarFile
Copy-Item $TarFile $TarLatest -Force

$manifest = @{
    app_name          = $ImageName
    app_version       = $AppVersion
    image             = "${ImageName}:${VersionTag}"
    image_version     = $VersionTag
    platform          = $Platform
    built_at_utc      = $UtcNow
    tar_file          = (Split-Path -Leaf $TarFile)
    tar_latest_alias  = (Split-Path -Leaf $TarLatest)
    health_check      = "GET /api/version 或 GET /api/health"
    compose_image_line = "image: ${ImageName}:${VersionTag}"
}
$manifest | ConvertTo-Json -Depth 4 | Set-Content -Path $ManifestFile -Encoding UTF8

$sizeMb = [math]::Round((Get-Item $TarFile).Length / 1MB, 1)
Write-Host ""
Write-Host "构建完成" -ForegroundColor Green
Write-Host "  APP 版本:       $AppVersion"
Write-Host "  镜像标签:       ${ImageName}:${VersionTag}"
Write-Host "  架构:           $Platform"
Write-Host "  离线包:         $TarFile ($sizeMb MB)"
Write-Host "  最新别名副本:   $TarLatest"
Write-Host "  版本清单:       $ManifestFile"
Write-Host ""
Write-Host "内网 ARM64 部署:" -ForegroundColor Cyan
Write-Host "  docker load -i $(Split-Path -Leaf $TarFile)"
Write-Host "  mkdir -p workspace logs"
Write-Host "  # docker-compose.yml 中设置: image: ${ImageName}:${VersionTag}"
Write-Host "  docker compose up -d"
Write-Host "  curl http://localhost:5000/api/version"
