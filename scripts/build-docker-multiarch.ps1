
# 构建 linux/arm64 和 linux/amd64 Docker 镜像并导出离线 tar 包（内网部署）
# 用法: .\scripts\build-docker-multiarch.ps1
# 可选: .\scripts\build-docker-multiarch.ps1 -AppVersion 2.0.1

param(
    [string]$AppVersion = "",
    [string]$OutputDir = "",
    [switch]$Arm64Only = $false,
    [switch]$Amd64Only = $false
)

$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectDir

$ImageName = "tacit-knowledge-externalization"
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

if (-not $OutputDir) {
    $OutputDir = $ProjectDir
}

if (-not (Get-Command docker -ErrorAction SilentlyContinue)) {
    Write-Error "未找到 docker 命令，请先安装 Docker Desktop 并启用 buildx"
}

docker info 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Error "Docker 未运行，请启动 Docker Desktop"
}

docker buildx use desktop-linux 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    docker buildx use default 2>&1 | Out-Null
}

# 中文路径易导致 buildx 失败：复制到英文临时目录构建
$BuildDir = "C:\tacit-knowledge-docker-build"
Write-Host "=== 复制项目到临时构建目录: $BuildDir ===" -ForegroundColor Cyan
if (Test-Path $BuildDir) { Remove-Item $BuildDir -Recurse -Force }
New-Item -ItemType Directory -Path $BuildDir | Out-Null
robocopy $ProjectDir $BuildDir /E /XD .git node_modules .cursor __pycache__ /XF *.tar *.pyc *.ps1 /NFL /NDL /NJH /NJS /nc /ns /np | Out-Null
Set-Location $BuildDir

# 确定要构建的架构
$platforms = @()
if ($Amd64Only) {
    $platforms = @("linux/amd64")
} elseif ($Arm64Only) {
    $platforms = @("linux/arm64")
} else {
    $platforms = @("linux/amd64", "linux/arm64")
}

Write-Host ""
Write-Host "=== 隐性知识显性化 Docker 镜像构建 ===" -ForegroundColor Cyan
Write-Host "APP 版本:     $AppVersion"
Write-Host "构建架构:     $($platforms -join ', ')"
Write-Host ""

$results = @()

foreach ($Platform in $platforms) {
    $PlatformSlug = if ($Platform -like "*arm64*") { "arm64" } else { "amd64" }
    $VersionTag = "${AppVersion}-${PlatformSlug}-${DateTag}"
    $TarBaseName = "${ImageName}-${VersionTag}"
    $TarFile = Join-Path $OutputDir "${TarBaseName}.tar"
    $ManifestFile = Join-Path $OutputDir "${TarBaseName}.manifest.json"

    Write-Host ">>> 构建 ${Platform} 镜像: ${ImageName}:${VersionTag}" -ForegroundColor Yellow

    $buildArgs = @(
        "buildx", "build",
        "--platform", $Platform,
        "--build-arg", "APP_VERSION=$AppVersion",
        "--build-arg", "IMAGE_VERSION=$VersionTag",
        "--build-arg", "BUILD_DATE=$UtcNow",
        "--build-arg", "APP_PLATFORM=$Platform",
        "--tag", "${ImageName}:${VersionTag}",
        "--tag", "${ImageName}:${AppVersion}-${PlatformSlug}",
        "--file", "docker/Dockerfile",
        "--pull=false",
        "--load",
        "--progress=plain",
        "."
    )
    docker @buildArgs
    if ($LASTEXITCODE -ne 0) {
        Write-Error "镜像构建失败 (${Platform})"
    }

    Write-Host ">>> 导出离线包: $TarFile" -ForegroundColor Yellow
    if (Test-Path $TarFile) { Remove-Item $TarFile -Force }
    docker save "${ImageName}:${VersionTag}" -o $TarFile

    $sizeMb = [math]::Round((Get-Item $TarFile).Length / 1MB, 1)

    $manifest = @{
        app_name          = $ImageName
        app_version       = $AppVersion
        image             = "${ImageName}:${VersionTag}"
        image_version     = $VersionTag
        platform          = $Platform
        platform_slug     = $PlatformSlug
        built_at_utc      = $UtcNow
        tar_file          = (Split-Path -Leaf $TarFile)
        size_mb           = $sizeMb
        health_check      = "GET /api/version"
        compose_image_line = "image: ${ImageName}:${VersionTag}"
    }
    $manifest | ConvertTo-Json -Depth 4 | Set-Content -Path $ManifestFile -Encoding UTF8

    $results += [PSCustomObject]@{
        Platform     = $PlatformSlug
        Tag          = $VersionTag
        TarFile      = $TarFile
        SizeMB       = $sizeMb
        ManifestFile = $ManifestFile
    }

    Write-Host "  完成: $TarBaseName.tar ($sizeMb MB)" -ForegroundColor Green
    Write-Host ""
}

Write-Host ""
Write-Host "=== 全部构建完成 ===" -ForegroundColor Green
foreach ($r in $results) {
    Write-Host "  [$($r.Platform)] $($r.Tag)  ($($r.SizeMB) MB)"
    Write-Host "     离线包: $($r.TarFile)"
    Write-Host "     清单:   $($r.ManifestFile)"
}

Write-Host ""
Write-Host "部署命令:" -ForegroundColor Cyan
foreach ($r in $results) {
    Write-Host "  # --- $($r.Platform) ---"
    Write-Host "  docker load -i $((Split-Path -Leaf $r.TarFile))"
    Write-Host "  docker run -d -p 5000:5000 --name tacit-knowledge ${ImageName}:$($r.Tag)"
}
