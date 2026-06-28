param(
    [int]$IntervalSeconds = 60,
    [int]$Samples = 0,
    [string]$OutDir = "J:\latextomathtype\analysis\codex-health"
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$script = Join-Path $PSScriptRoot "watch_codex_health.ps1"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$launchDir = Join-Path $OutDir "launch-$stamp"
New-Item -ItemType Directory -Force -Path $launchDir | Out-Null

$out = Join-Path $launchDir "watch.out.log"
$err = Join-Path $launchDir "watch.err.log"
$pidPath = Join-Path $launchDir "pid.txt"

$args = @(
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", $script,
    "-IntervalSeconds", "$IntervalSeconds",
    "-Samples", "$Samples",
    "-OutDir", $OutDir
)

$process = Start-Process -FilePath "powershell.exe" -ArgumentList $args -WorkingDirectory $repoRoot -WindowStyle Hidden -PassThru -RedirectStandardOutput $out -RedirectStandardError $err
$process.Id | Set-Content -LiteralPath $pidPath -Encoding ASCII

Write-Host "pid=$($process.Id)"
Write-Host "launchDir=$launchDir"
Write-Host "stdout=$out"
Write-Host "stderr=$err"
