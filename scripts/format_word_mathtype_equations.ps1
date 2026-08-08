param(
    [Parameter(Mandatory = $true)]
    [string]$InputDocx,
    [Parameter(Mandatory = $true)]
    [string]$OutputDocx,
    [string]$OutReport = "",
    [ValidateRange(1, 1000)]
    [int]$ChunkSize = 12,
    [ValidateRange(0, 1000000)]
    [int]$ExpectedObjectCount = 0,
    [int[]]$PreserveObjectIndices = @(),
    [int[]]$FormatObjectIndices = @(),
    [ValidateSet("NewEquation", "Clipboard")]
    [string]$PreferenceSource = "NewEquation",
    [string]$PreferenceDocx = "",
    [ValidateRange(1, 1000000)]
    [int]$PreferenceObjectIndex = 1,
    [ValidateRange(5, 600)]
    [int]$DialogTimeoutSeconds = 120,
    [switch]$RepeatFirstChunk,
    [switch]$Visible
)

$ErrorActionPreference = "Stop"
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class WordWindowNative {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}
'@

function Resolve-RequiredPath {
    param([string]$Path)
    $resolved = if ([IO.Path]::IsPathRooted($Path)) {
        [IO.Path]::GetFullPath($Path)
    } else {
        [IO.Path]::GetFullPath((Join-Path (Get-Location) $Path))
    }
    if (-not (Test-Path -LiteralPath $resolved)) { throw "file not found: $resolved" }
    return $resolved
}

function Resolve-OutputPath {
    param([string]$Path)
    if ([IO.Path]::IsPathRooted($Path)) { return [IO.Path]::GetFullPath($Path) }
    return [IO.Path]::GetFullPath((Join-Path (Get-Location) $Path))
}

function Get-MathTypeObjects {
    param($Document)
    $objects = @()
    foreach ($shape in $Document.InlineShapes) {
        $progId = ""
        try { $progId = [string]$shape.OLEFormat.ProgID } catch { $progId = "" }
        if ($progId -eq "Equation.DSMT4") {
            $objects += [pscustomobject]@{
                rangeStart = [int]$shape.Range.Start
                rangeEnd = [int]$shape.Range.End
            }
        }
    }
    return @($objects)
}

$inputPath = Resolve-RequiredPath $InputDocx
$preferencePath = if ($PreferenceSource -eq "Clipboard") {
    if ([string]::IsNullOrWhiteSpace($PreferenceDocx)) {
        throw "PreferenceDocx is required when PreferenceSource is Clipboard"
    }
    Resolve-RequiredPath $PreferenceDocx
} else { $null }
$outputPath = Resolve-OutputPath $OutputDocx
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $outputPath) | Out-Null
if ([string]::IsNullOrWhiteSpace($OutReport)) {
    $OutReport = [IO.Path]::ChangeExtension($outputPath, ".format-report.json")
}
$reportPath = Resolve-OutputPath $OutReport

$watcherScript = Join-Path $PSScriptRoot "watch_mathtype_format_dialog.ps1"
if (-not (Test-Path -LiteralPath $watcherScript)) { throw "dialog watcher not found: $watcherScript" }
$watcherResult = Join-Path ([IO.Path]::GetTempPath()) ("mathtype-format-" + [guid]::NewGuid() + ".json")

if (-not (Get-Process MathType -ErrorAction SilentlyContinue)) {
    $mathTypeServer = New-Object -ComObject Equation.DSMT4
    [Runtime.InteropServices.Marshal]::ReleaseComObject($mathTypeServer) | Out-Null
    Start-Sleep -Milliseconds 500
}

$word = New-Object -ComObject Word.Application
$word.Visible = $true
$word.DisplayAlerts = 0
$word.ScreenUpdating = $true
$doc = $null
$watcherProcess = $null
try {
    if ($preferencePath) {
        $preferenceDocument = $word.Documents.Open([string]$preferencePath, $false, $true)
        try {
            $preferenceObjects = @(Get-MathTypeObjects $preferenceDocument)
            if ($preferenceObjects.Count -lt $PreferenceObjectIndex) {
                throw "preference object index $PreferenceObjectIndex is outside 1..$($preferenceObjects.Count): $preferencePath"
            }
            $preferenceObject = $preferenceObjects[$PreferenceObjectIndex - 1]
            $preferenceDocument.Range(
                $preferenceObject.rangeStart,
                $preferenceObject.rangeEnd).Copy()
        } finally {
            $preferenceDocument.Close($false) | Out-Null
            [Runtime.InteropServices.Marshal]::ReleaseComObject($preferenceDocument) | Out-Null
        }
    }
    $doc = $word.Documents.Open([string]$inputPath, $false, $false)
    $doc.Activate()
    $beforeObjects = @(Get-MathTypeObjects $doc)
    $beforeCount = $beforeObjects.Count
    $beforeInlineShapeCount = [int]$doc.InlineShapes.Count
    if ($ExpectedObjectCount -gt 0 -and $beforeCount -ne $ExpectedObjectCount) {
        throw "input contains $beforeCount MathType objects; expected $ExpectedObjectCount"
    }

    $wordProcess = Get-Process WINWORD | Where-Object { $_.MainWindowHandle -eq $word.ActiveWindow.Hwnd } | Select-Object -First 1
    if (-not $wordProcess) {
        $wordProcess = Get-Process WINWORD | Sort-Object StartTime -Descending | Select-Object -First 1
    }
    if (-not $wordProcess) { throw "could not identify the Word process" }
    [void][WordWindowNative]::SetForegroundWindow($wordProcess.MainWindowHandle)
    Start-Sleep -Milliseconds 300

    $started = Get-Date
    $formattedCount = 0
    $statistics = @()
    $checkpointSaved = $false
    $repeatedFirstChunk = $false
    $preserveSet = @{}
    foreach ($objectIndex in $PreserveObjectIndices) {
        if ($objectIndex -lt 1 -or $objectIndex -gt $beforeCount) {
            throw "preserved object index $objectIndex is outside 1..$beforeCount"
        }
        $preserveSet[$objectIndex] = $true
    }
    if ($FormatObjectIndices.Count -gt 0) {
        $formatIndices = @($FormatObjectIndices | Sort-Object -Unique)
        foreach ($objectIndex in $formatIndices) {
            if ($objectIndex -lt 1 -or $objectIndex -gt $beforeCount) {
                throw "format object index $objectIndex is outside 1..$beforeCount"
            }
            if ($preserveSet.ContainsKey($objectIndex)) {
                throw "object index $objectIndex cannot be both formatted and preserved"
            }
        }
    } else {
        $formatIndices = @(1..$beforeCount | Where-Object { -not $preserveSet.ContainsKey($_) })
    }
    $contiguousFormatIndices = $true
    for ($index = 1; $index -lt $formatIndices.Count; $index++) {
        if ($formatIndices[$index] -ne $formatIndices[$index - 1] + 1) {
            $contiguousFormatIndices = $false
            break
        }
    }
    $effectiveChunkSize = if ($preserveSet.Count -gt 0 -or -not $contiguousFormatIndices) { 1 } else { $ChunkSize }
    for ($cursor = 0; $cursor -lt $formatIndices.Count; $cursor += $effectiveChunkSize) {
        $first = [int]$formatIndices[$cursor]
        $lastCursor = [Math]::Min($cursor + $effectiveChunkSize - 1, $formatIndices.Count - 1)
        $last = [int]$formatIndices[$lastCursor]
        $expectedChunkCount = $last - $first + 1
        $formatScope = if ($preserveSet.Count -gt 0 -or $beforeCount -gt $ChunkSize -or
                $beforeInlineShapeCount -ne $beforeCount) {
            "Selection"
        } else {
            "WholeDocument"
        }
        if ($formatScope -eq "Selection") {
            $currentObjects = @(Get-MathTypeObjects $doc)
            if ($currentObjects.Count -ne $beforeCount) {
                throw "MathType object count changed before chunk $first-$last`: before=$beforeCount current=$($currentObjects.Count)"
            }
            $rangeStart = $currentObjects[$first - 1].rangeStart
            $rangeEnd = $currentObjects[$last - 1].rangeEnd
            $doc.Range($rangeStart, $rangeEnd).Select()
        }

        $watcherArguments = @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", ('"' + $watcherScript + '"'),
            "-WordProcessId", [string]$wordProcess.Id,
            "-ResultPath", ('"' + $watcherResult + '"'),
            "-FormatScope", $formatScope,
            "-PreferenceSource", $PreferenceSource,
            "-TimeoutSeconds", [string]$DialogTimeoutSeconds
        )
        $watcherProcess = Start-Process -FilePath "powershell.exe" -ArgumentList $watcherArguments `
            -WindowStyle Hidden -PassThru
        $word.Run("MTCommand_FormatEqns")
        $watcherProcess.WaitForExit()
        $watcherExitCode = $watcherProcess.ExitCode
        $watcherProcess.Dispose()
        $watcherProcess = $null
        if ($watcherExitCode -ne 0 -or -not (Test-Path -LiteralPath $watcherResult)) {
            throw "MathType Format Equations dialog automation failed with exit code $watcherExitCode"
        }
        $dialogReport = Get-Content -Raw -LiteralPath $watcherResult | ConvertFrom-Json
        $chunkFormatted = [int]$dialogReport.formattedObjectCount
        if ($chunkFormatted -ne $expectedChunkCount) {
            throw "MathType formatted only $chunkFormatted of $expectedChunkCount objects in chunk $first-$last"
        }
        $formattedCount += $chunkFormatted
        $statistics += [string]$dialogReport.statisticsText
        Remove-Item -LiteralPath $watcherResult -Force
        if (-not $checkpointSaved) {
            $doc.SaveAs([ref]$outputPath, [ref]16)
            $checkpointSaved = $true
        } else {
            $doc.Save()
        }
        Write-Host "formattedChunk=$first-$last formattedTotal=$formattedCount checkpoint=$outputPath"
        if ($RepeatFirstChunk -and $cursor -eq 0 -and -not $repeatedFirstChunk) {
            $repeatedFirstChunk = $true
            $cursor -= $effectiveChunkSize
            Write-Host "repeatFirstChunk=$first-$last"
        }
    }

    $afterCount = @(Get-MathTypeObjects $doc).Count
    if ($afterCount -ne $beforeCount) {
        throw "MathType object count changed during formatting: before=$beforeCount after=$afterCount"
    }
    if (-not $checkpointSaved) {
        $doc.SaveAs([ref]$outputPath, [ref]16)
    } else {
        $doc.Save()
    }
    [ordered]@{
        schemaVersion = 1
        inputDocx = $inputPath
        outputDocx = $outputPath
        preferenceSource = if ($PreferenceSource -eq "Clipboard") {
            "clipboard Equation.DSMT4 preferences: $preferencePath object=$PreferenceObjectIndex"
        } else { "MathType new equation preferences" }
        objectCountBefore = $beforeCount
        objectCountAfter = $afterCount
        inlineShapeCount = $beforeInlineShapeCount
        nonMathTypeInlineShapeCount = $beforeInlineShapeCount - $beforeCount
        formattedObjectCount = $formattedCount
        preservedObjectCount = $beforeCount - $formatIndices.Count
        chunkSize = $ChunkSize
        repeatedFirstChunk = $repeatedFirstChunk
        statisticsText = [string]::Join("`r`n", $statistics)
        elapsedMs = [int]((Get-Date) - $started).TotalMilliseconds
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $reportPath -Encoding UTF8

    [pscustomobject]@{
        OutputDocx = $outputPath
        Report = $reportPath
        ObjectCount = $afterCount
    }
} finally {
    if ($watcherProcess) {
        if (-not $watcherProcess.HasExited) { $watcherProcess.Kill() }
        $watcherProcess.Dispose()
    }
    if ($doc) {
        $doc.Close($false) | Out-Null
        [Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
    }
    if ($word) {
        $word.Quit() | Out-Null
        [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    if (Test-Path -LiteralPath $watcherResult) {
        Remove-Item -LiteralPath $watcherResult -Force
    }
}
