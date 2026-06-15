param(
    [int]$Start = 1,
    [int]$End = 0,
    [int]$ChunkSize = 25,
    [string]$AnalysisDir = "J:\latextomathtype\analysis",
    [string]$LatexRoot = "J:\latextomathtype\analysis\xsc-latex-fixed",
    [string]$Manifest = "J:\latextomathtype\analysis\xsc-latex-fixed\manifest.json",
    [string]$RequestDir = "",
    [string]$RunRoot = "J:\latextomathtype\analysis\unattended-runs",
    [int]$CommandTimeoutMinutes = 45
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$script = Join-Path $PSScriptRoot "run_xsc_unattended_acceptance.ps1"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$launchDir = Join-Path $RunRoot "launch-$stamp"
New-Item -ItemType Directory -Force -Path $launchDir | Out-Null

$out = Join-Path $launchDir "launcher.out.log"
$err = Join-Path $launchDir "launcher.err.log"
$pidPath = Join-Path $launchDir "pid.txt"

$args = @(
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", $script,
    "-Start", "$Start",
    "-ChunkSize", "$ChunkSize",
    "-AnalysisDir", $AnalysisDir,
    "-LatexRoot", $LatexRoot,
    "-Manifest", $Manifest,
    "-RequestDir", $RequestDir,
    "-RunRoot", $RunRoot,
    "-CommandTimeoutMinutes", "$CommandTimeoutMinutes"
)
if ($End -gt 0) {
    $args += @("-End", "$End")
}

$process = Start-Process -FilePath "powershell.exe" -ArgumentList $args -WorkingDirectory $repoRoot -WindowStyle Hidden -PassThru -RedirectStandardOutput $out -RedirectStandardError $err
$process.Id | Set-Content -LiteralPath $pidPath -Encoding ASCII

Write-Host "pid=$($process.Id)"
Write-Host "launchDir=$launchDir"
Write-Host "stdout=$out"
Write-Host "stderr=$err"
