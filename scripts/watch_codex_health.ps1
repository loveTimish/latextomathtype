param(
    [int]$IntervalSeconds = 60,
    [int]$Samples = 0,
    [string]$OutDir = "J:\latextomathtype\analysis\codex-health"
)

$ErrorActionPreference = "Stop"

$packageRoot = "C:\Users\11703\AppData\Local\Packages\OpenAI.Codex_2p2nqsd0c76g0"
$logRoot = Join-Path $packageRoot "LocalCache\Local\Codex\Logs"
$crashRoot = Join-Path $packageRoot "LocalCache\Roaming\Codex\web\Codex\Crashpad\reports"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$outPath = Join-Path $OutDir "codex-health-$stamp.jsonl"
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$count = 0
while ($true) {
    $processes = @(Get-Process | Where-Object { $_.ProcessName -ieq "Codex" -or $_.ProcessName -ieq "codex" } |
        Select-Object Id, ProcessName, Path, StartTime, CPU, WorkingSet64, PrivateMemorySize64)

    $latestLogs = @()
    if (Test-Path -LiteralPath $logRoot) {
        $latestLogs = @(Get-ChildItem -LiteralPath $logRoot -Recurse -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 5 FullName, LastWriteTime, Length)
    }

    $latestCrashes = @()
    if (Test-Path -LiteralPath $crashRoot) {
        $latestCrashes = @(Get-ChildItem -LiteralPath $crashRoot -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 5 FullName, LastWriteTime, Length)
    }

    $record = [ordered]@{
        time = (Get-Date).ToString("o")
        processCount = $processes.Count
        totalWorkingSetMb = [Math]::Round((($processes | Measure-Object -Property WorkingSet64 -Sum).Sum / 1MB), 1)
        maxWorkingSetMb = [Math]::Round((($processes | Measure-Object -Property WorkingSet64 -Maximum).Maximum / 1MB), 1)
        processes = $processes
        latestLogs = $latestLogs
        latestCrashes = $latestCrashes
    }
    ($record | ConvertTo-Json -Compress -Depth 8) | Add-Content -LiteralPath $outPath -Encoding UTF8
    Write-Host "sample=$count processes=$($record.processCount) totalMb=$($record.totalWorkingSetMb) maxMb=$($record.maxWorkingSetMb)"

    $count++
    if ($Samples -gt 0 -and $count -ge $Samples) {
        break
    }
    Start-Sleep -Seconds $IntervalSeconds
}

Write-Host "healthLog=$outPath"
