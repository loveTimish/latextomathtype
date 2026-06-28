param(
    [int]$Start = 1,
    [int]$End = 0,
    [int]$ChunkSize = 25,
    [string]$AnalysisDir = "J:\latextomathtype\analysis",
    [string]$LatexRoot = "J:\latextomathtype\analysis\xsc-latex-fixed",
    [string]$Manifest = "J:\latextomathtype\analysis\xsc-latex-fixed\manifest.json",
    [string]$RequestDir = "",
    [string]$RunRoot = "J:\latextomathtype\analysis\unattended-runs",
    [int]$CommandTimeoutMinutes = 45,
    [switch]$SkipRequestBuild,
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$mvn = Join-Path $repoRoot ".mvn\apache-maven-3.9.12\bin\mvn.cmd"
if (-not (Test-Path -LiteralPath $mvn)) {
    throw "Missing Maven runtime: $mvn"
}
if (-not (Test-Path -LiteralPath $Manifest)) {
    throw "Missing manifest: $Manifest"
}

$manifestItems = @(Get-Content -LiteralPath $Manifest -Raw | ConvertFrom-Json | ForEach-Object { $_ })
$manifestByIndex = @{}
foreach ($manifestItem in $manifestItems) {
    $manifestByIndex[[int]$manifestItem.index] = $manifestItem
}
if ($End -le 0) {
    $End = ($manifestItems | Measure-Object -Property index -Maximum).Maximum
}
if ($ChunkSize -lt 1) {
    throw "ChunkSize must be >= 1"
}

$runStamp = Get-Date -Format "yyyyMMdd-HHmmss"
$runDir = Join-Path $RunRoot $runStamp
$logDir = Join-Path $runDir "logs"
$sizeDir = Join-Path $runDir "size-compare"
$docxRoot = Join-Path $runDir "docx"
$requestRoot = if ($RequestDir) { $RequestDir } else { Join-Path $runDir "requests" }
$summaryPath = Join-Path $runDir "summary.json"
$checkpointPath = Join-Path $runDir "checkpoint.jsonl"

New-Item -ItemType Directory -Force -Path $logDir, $sizeDir, $docxRoot, $requestRoot | Out-Null

function Write-Checkpoint {
    param([hashtable]$Record)
    $Record.time = (Get-Date).ToString("o")
    ($Record | ConvertTo-Json -Compress -Depth 8) | Add-Content -LiteralPath $checkpointPath -Encoding UTF8
}

function ConvertTo-CommandLineArgument {
    param([string]$Argument)
    if ($null -eq $Argument) {
        return '""'
    }
    if ($Argument -notmatch '[\s"]') {
        return $Argument
    }
    $escaped = $Argument -replace '\\(?=\\*")', '$&$&'
    $escaped = $escaped -replace '"', '\"'
    $escaped = $escaped -replace '(\\+)$', '$1$1'
    return '"' + $escaped + '"'
}

function Invoke-Logged {
    param(
        [string]$Name,
        [string]$FilePath,
        [string[]]$Arguments,
        [string]$WorkingDirectory
    )
    $out = Join-Path $logDir "$Name.out.log"
    $err = Join-Path $logDir "$Name.err.log"
    Write-Host "RUN $Name"
    Write-Host "$FilePath $($Arguments -join ' ')"
    if ($DryRun) {
        return
    }
    $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $FilePath
    $startInfo.WorkingDirectory = $WorkingDirectory
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.Arguments = (($Arguments | ForEach-Object { ConvertTo-CommandLineArgument $_ }) -join " ")
    $process = [System.Diagnostics.Process]::new()
    $process.StartInfo = $startInfo
    [void]$process.Start()
    $stdoutTask = $process.StandardOutput.ReadToEndAsync()
    $stderrTask = $process.StandardError.ReadToEndAsync()
    $timeoutMs = [Math]::Max(1, $CommandTimeoutMinutes) * 60 * 1000
    if (-not $process.WaitForExit($timeoutMs)) {
        $children = @(Get-CimInstance Win32_Process | Where-Object { $_.ParentProcessId -eq $process.Id })
        foreach ($child in $children) {
            Stop-Process -Id $child.ProcessId -Force -ErrorAction SilentlyContinue
        }
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
        Write-Checkpoint @{ event = "timeout"; name = $Name; timeoutMinutes = $CommandTimeoutMinutes; stdout = $out; stderr = $err }
        throw "$Name timed out after $CommandTimeoutMinutes minutes. See $err"
    }
    $stdoutTask.Wait()
    $stderrTask.Wait()
    [System.IO.File]::WriteAllText($out, $stdoutTask.Result, [System.Text.Encoding]::UTF8)
    [System.IO.File]::WriteAllText($err, $stderrTask.Result, [System.Text.Encoding]::UTF8)
    $process.Refresh()
    $exitCode = $process.ExitCode
    if ($exitCode -ne 0) {
        Write-Checkpoint @{ event = "failed"; name = $Name; exitCode = $exitCode; stdout = $out; stderr = $err }
        throw "$Name failed with exit code $exitCode. See $err"
    }
}

Write-Host "runDir=$runDir"
Write-Host "range=$Start..$End chunkSize=$ChunkSize"
Write-Checkpoint @{ event = "start"; start = $Start; end = $End; chunkSize = $ChunkSize; runDir = $runDir }

for ($chunkStart = $Start; $chunkStart -le $End; $chunkStart += $ChunkSize) {
    $chunkEnd = [Math]::Min($End, $chunkStart + $ChunkSize - 1)
    $chunkName = "$chunkStart-$chunkEnd"
    $chunkRequestDir = if ($SkipRequestBuild -and $RequestDir) { $requestRoot } else { Join-Path $requestRoot $chunkName }
    $chunkDocxDir = Join-Path $docxRoot $chunkName
    New-Item -ItemType Directory -Force -Path $chunkRequestDir, $chunkDocxDir | Out-Null
    Write-Checkpoint @{ event = "chunk-start"; start = $chunkStart; end = $chunkEnd }

    if (-not $SkipRequestBuild) {
        Invoke-Logged `
            -Name "requests-$chunkName" `
            -FilePath "python" `
            -Arguments @(
                "scripts\make_full_batch10_requests.py",
                "--start", "$chunkStart",
                "--end", "$chunkEnd",
                "--latex-root", $LatexRoot,
                "--fallback-latex-root", $LatexRoot,
                "--out-dir", $chunkRequestDir,
                "--manifest", $Manifest
            ) `
            -WorkingDirectory $repoRoot
    }

    Invoke-Logged `
        -Name "mvn-$chunkName" `
        -FilePath $mvn `
        -Arguments @(
            "-q",
            "-Dtest=XscFullBatch10DocxTest",
            "-Dxsc.analysis.dir=$AnalysisDir",
            "-Dxsc.full.request.dir=$chunkRequestDir",
            "-Dxsc.full.run.output.dir=$chunkDocxDir",
            "-Dxsc.full.start=$chunkStart",
            "-Dxsc.full.end=$chunkEnd",
            "test"
        ) `
        -WorkingDirectory $repoRoot

    if ($DryRun) {
        continue
    }

    for ($i = $chunkStart; $i -le $chunkEnd; $i++) {
        $item = $manifestByIndex[$i]
        if ($null -eq $item) {
            throw "Missing manifest item for index $i"
        }
        $pad = "{0:D2}" -f $i
        $generatedItem = Get-ChildItem -LiteralPath $chunkDocxDir -File -Filter "*.docx" |
            Where-Object { $_.Name -like "*_$pad.docx" } |
            Select-Object -First 1
        if ($null -eq $generatedItem) {
            throw "Missing generated docx for index $pad in $chunkDocxDir"
        }
        $generated = $generatedItem.FullName
        $request = Join-Path $chunkRequestDir "full-$pad.request.json"
        if (-not (Test-Path -LiteralPath $request)) {
            throw "Missing request json: $request"
        }
        $docSizeDir = Join-Path $sizeDir $pad
        New-Item -ItemType Directory -Force -Path $docSizeDir | Out-Null
        Invoke-Logged `
            -Name "compare-$pad" `
            -FilePath "python" `
            -Arguments @(
                "scripts\compare_docx_pair_metrics.py",
                $item.sourcePath,
                $generated,
                $docSizeDir,
                "--generated-request",
                $request
            ) `
            -WorkingDirectory $repoRoot
    }

    Invoke-Logged `
        -Name "summary-$chunkName" `
        -FilePath "python" `
        -Arguments @(
            "scripts\summarize_xsc_size_acceptance.py",
            $sizeDir,
            "--out",
            $summaryPath,
            "--limit",
            "20"
        ) `
        -WorkingDirectory $repoRoot

    Write-Checkpoint @{ event = "chunk-complete"; start = $chunkStart; end = $chunkEnd; docxDir = $chunkDocxDir; requestDir = $chunkRequestDir; summary = $summaryPath }
}

if (-not $DryRun) {
    Invoke-Logged `
        -Name "summary-final" `
        -FilePath "python" `
        -Arguments @(
            "scripts\summarize_xsc_size_acceptance.py",
            $sizeDir,
            "--out",
            $summaryPath,
            "--limit",
            "50"
        ) `
        -WorkingDirectory $repoRoot
}

Write-Checkpoint @{ event = "complete"; summary = $summaryPath; sizeDir = $sizeDir }
Write-Host "summary=$summaryPath"
