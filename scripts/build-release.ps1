param(
    [ValidateSet("windows-x64", "linux-x64", "all")]
    [string]$Platform = "all",
    [string]$Version = "",
    [string]$OutDir = "target/release",
    [switch]$SkipTests,
    [switch]$SkipSidecarBuild
)

$ErrorActionPreference = "Stop"
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$targetRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot "target"))
$outputRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot $OutDir))
if (-not $outputRoot.StartsWith($targetRoot + [IO.Path]::DirectorySeparatorChar,
        [StringComparison]::OrdinalIgnoreCase)) {
    throw "输出目录必须位于仓库 target 目录下：$outputRoot"
}

[xml]$pom = Get-Content -LiteralPath (Join-Path $repoRoot "pom.xml")
$projectVersion = [string]$pom.project.version
if ([string]::IsNullOrWhiteSpace($Version)) {
    $Version = $projectVersion
}
if ([string]::IsNullOrWhiteSpace($projectVersion) -or $projectVersion.Contains('$')) {
    throw "无法从 pom.xml 解析出明确的项目版本号"
}
if ([string]::IsNullOrWhiteSpace($Version) -or $Version.Contains('$')) {
    throw "无法从 pom.xml 解析出明确的发布版本号"
}

$trackedStatus = & git -C $repoRoot status --porcelain --untracked-files=no
if ($LASTEXITCODE -ne 0) { throw "读取 Git 状态失败" }
if ($trackedStatus) {
    throw "发布构建要求已跟踪文件没有未提交改动，请先提交或还原这些改动"
}
$commit = (& git -C $repoRoot rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0) { throw "无法读取当前 Git 提交号" }

$maven = Join-Path $repoRoot ".mvn\apache-maven-3.9.12\bin\mvn.cmd"
if (-not (Test-Path -LiteralPath $maven)) { throw "未找到仓库内置 Maven：$maven" }
$mavenArgs = @(if ($SkipSidecarBuild) { "package" } else { "clean"; "package" })
if ($SkipTests) { $mavenArgs += "-DskipTests" }
Write-Host "执行 Maven：$($mavenArgs -join ' ')"
& $maven @mavenArgs
if ($LASTEXITCODE -ne 0) { throw "Maven 发布构建失败" }

$jar = Join-Path $repoRoot "target\paper-to-word-$projectVersion.jar"
if (-not (Test-Path -LiteralPath $jar)) { throw "未找到发布 JAR：$jar" }

$sidecarRoot = Join-Path $repoRoot "target\vector-sidecar-dist"
if (-not $SkipSidecarBuild) {
    & (Join-Path $PSScriptRoot "build-vector-sidecar.ps1") -Platform $Platform `
        -OutDir "target/vector-sidecar-dist"
    if ($LASTEXITCODE -ne 0) { throw "真矢量离线 sidecar 构建失败" }
}

$platforms = if ($Platform -eq "all") { @("windows-x64", "linux-x64") } else { @($Platform) }
$versionRoot = Join-Path $outputRoot $Version
$stageRoot = Join-Path $versionRoot "stage"
if (Test-Path -LiteralPath $versionRoot) {
    Remove-Item -LiteralPath $versionRoot -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $stageRoot | Out-Null

function Write-ReleaseFiles([string]$bundle, [string]$platformName) {
    New-Item -ItemType Directory -Force -Path $bundle | Out-Null
    Copy-Item -LiteralPath $jar -Destination (Join-Path $bundle "paper-to-word.jar")
    Copy-Item -LiteralPath (Join-Path $repoRoot "exam-template.json") -Destination $bundle
    Copy-Item -LiteralPath (Join-Path $repoRoot "README.md") -Destination $bundle
    Copy-Item -LiteralPath (Join-Path $repoRoot "README.zh-CN.md") -Destination $bundle
    New-Item -ItemType Directory -Force -Path (Join-Path $bundle "docs") | Out-Null
    Copy-Item -LiteralPath (Join-Path $repoRoot "docs\linux-runtime.md") -Destination (Join-Path $bundle "docs")
    Copy-Item -LiteralPath (Join-Path $repoRoot "docs\vector-runtime-licenses.md") -Destination (Join-Path $bundle "docs")

    $releaseInfo = [ordered]@{
        product = "latextomathtype"
        version = $Version
        platform = $platformName
        commit = $commit
        java = "21 或更高版本"
        node = "v24.9.0（随包提供）"
        mathJax = "3.2.2"
        saxonJs = "2.7.0"
        preview = "MathJax SVG -> Batik -> 纯 POLYPOLYGON WMF"
    }
    $releaseInfo | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $bundle "release-info.json") -Encoding utf8

    @"
latextomathtype $Version（$platformName）

运行要求：Java 21 或更高版本。Node、MathJax 和 Saxon-JS 已包含在 vector-sidecar 中。
运行时完全离线，不会执行 npm，也不会访问网络。

健康检查：http://127.0.0.1:8081/api/export/health
Word 导出接口：POST http://127.0.0.1:8081/api/export/word

Windows：在 PowerShell 中运行 .\run.ps1
Linux：运行 chmod +x run.sh && ./run.sh
"@ | Set-Content -LiteralPath (Join-Path $bundle "RELEASE-README.txt") -Encoding utf8
}

$artifacts = @()
foreach ($name in $platforms) {
    $packageName = "latextomathtype-$Version-$name"
    $platformStage = Join-Path $stageRoot $name
    $bundle = Join-Path $platformStage $packageName
    Write-ReleaseFiles $bundle $name

    if ($name -eq "windows-x64") {
        $sidecar = Join-Path $sidecarRoot "vector-sidecar-windows-x64.zip"
        if (-not (Test-Path -LiteralPath $sidecar)) { throw "缺少 Windows sidecar：$sidecar" }
        Expand-Archive -LiteralPath $sidecar -DestinationPath $bundle
        @'
param([int]$Port = 8081)
$ErrorActionPreference = "Stop"
$cache = Join-Path $PSScriptRoot "cache"
New-Item -ItemType Directory -Force -Path $cache | Out-Null
$env:SERVER_PORT = "$Port"
Push-Location $PSScriptRoot
try {
    & java "-Dpaperword.render.cache.enabled=true" "-Dpaperword.render.cache.dir=$cache" -jar "paper-to-word.jar"
    exit $LASTEXITCODE
} finally {
    Pop-Location
}
'@ | Set-Content -LiteralPath (Join-Path $bundle "run.ps1") -Encoding utf8
        $artifact = Join-Path $versionRoot "$packageName.zip"
        Compress-Archive -Path $bundle -DestinationPath $artifact -CompressionLevel Optimal
    } else {
        $sidecar = Join-Path $sidecarRoot "vector-sidecar-linux-x64.tar.gz"
        if (-not (Test-Path -LiteralPath $sidecar)) { throw "缺少 Linux sidecar：$sidecar" }
        & tar -xzf $sidecar -C $bundle
        if ($LASTEXITCODE -ne 0) { throw "无法解压 Linux sidecar" }
        $runSh = @'
#!/bin/sh
set -eu
ROOT=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
CACHE="$ROOT/cache"
mkdir -p "$CACHE"
cd "$ROOT"
exec java \
  -Dpaperword.render.cache.enabled=true \
  -Dpaperword.render.cache.dir="$CACHE" \
  -jar paper-to-word.jar
'@
        [IO.File]::WriteAllText(
            (Join-Path $bundle "run.sh"),
            $runSh,
            [Text.UTF8Encoding]::new($false))
        $artifact = Join-Path $versionRoot "$packageName.tar.gz"
        & tar -czf $artifact -C $platformStage $packageName
        if ($LASTEXITCODE -ne 0) { throw "无法创建 Linux 发布压缩包" }
    }
    $artifacts += Get-Item -LiteralPath $artifact
}

$checksums = foreach ($artifact in $artifacts) {
    $hash = (Get-FileHash -LiteralPath $artifact.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
    "$hash  $($artifact.Name)"
}
$checksums | Set-Content -LiteralPath (Join-Path $versionRoot "SHA256SUMS.txt") -Encoding ascii
Remove-Item -LiteralPath $stageRoot -Recurse -Force

[pscustomobject]@{
    version = $Version
    commit = $commit
    output = $versionRoot
    artifacts = @($artifacts.Name)
} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $versionRoot "release-manifest.json") -Encoding utf8

Write-Host "发布版 $Version 已构建，来源提交：$commit"
$artifacts | ForEach-Object { Write-Host $_.FullName }
Write-Host "校验文件：$(Join-Path $versionRoot "SHA256SUMS.txt")"
