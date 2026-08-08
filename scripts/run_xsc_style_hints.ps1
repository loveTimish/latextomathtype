param(
    [int]$Start = 1,
    [int]$End = 155,
    [string]$DatasetDir = "F:\资料\xsc资料\word_files",
    [string]$AnalysisDir = "J:\latextomathtype\analysis",
    [string]$DocxToLatexDir = "J:\docxtolatex\docxtolatex-main",
    [int]$MaxParallel = 16
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$outDir = Join-Path $AnalysisDir "mtef-style-hints"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
if (-not (Test-Path -LiteralPath $DocxToLatexDir)) {
    throw "missing docxtolatex dir: $DocxToLatexDir"
}

Push-Location $DocxToLatexDir
try {
    $extractorSource = Join-Path $repoRoot "scripts\extract_mtef_style_hints.go"
    $toolDir = Join-Path $repoRoot "target\tools"
    $extractorExe = Join-Path $toolDir "extract_mtef_style_hints.exe"
    New-Item -ItemType Directory -Force -Path $toolDir | Out-Null
    if (-not (Test-Path -LiteralPath $extractorExe) -or
            (Get-Item -LiteralPath $extractorSource).LastWriteTimeUtc -gt
                (Get-Item -LiteralPath $extractorExe).LastWriteTimeUtc) {
        go build -o $extractorExe $extractorSource
        if ($LASTEXITCODE -ne 0) { throw "failed to build $extractorSource" }
    }
    $pending = New-Object System.Collections.Generic.Queue[object]
    for ($i = $Start; $i -le $End; $i++) {
        $sourceDocx = Join-Path $DatasetDir "$i.docx"
        $outPath = Join-Path $outDir "$i.style-hints.json"
        if (-not (Test-Path -LiteralPath $sourceDocx)) {
            throw "missing source docx: $sourceDocx"
        }
        if (Test-Path -LiteralPath $outPath) {
            try {
                Get-Content -LiteralPath $outPath -Raw -Encoding UTF8 | ConvertFrom-Json | Out-Null
                continue
            } catch {
                Remove-Item -LiteralPath $outPath -Force
            }
        }
        $pending.Enqueue([pscustomobject]@{ index = $i; source = $sourceDocx; output = $outPath })
    }
    $running = New-Object System.Collections.Generic.List[object]
    while ($pending.Count -gt 0 -or $running.Count -gt 0) {
        while ($pending.Count -gt 0 -and $running.Count -lt $MaxParallel) {
            $item = $pending.Dequeue()
            $process = Start-Process -FilePath $extractorExe `
                -ArgumentList @($item.source, $item.output) -PassThru -NoNewWindow
            $running.Add([pscustomobject]@{ process = $process; item = $item })
        }
        for ($position = $running.Count - 1; $position -ge 0; $position--) {
            $active = $running[$position]
            if (-not $active.process.HasExited) { continue }
            $active.process.WaitForExit()
            $validOutput = $false
            if (Test-Path -LiteralPath $active.item.output) {
                try {
                    Get-Content -LiteralPath $active.item.output -Raw -Encoding UTF8 |
                        ConvertFrom-Json | Out-Null
                    $validOutput = $true
                } catch {
                    $validOutput = $false
                }
            }
            if (-not $validOutput) {
                throw "extract_mtef_style_hints failed for $($active.item.source) (exit=$($active.process.ExitCode))"
            }
            $running.RemoveAt($position)
        }
        if ($running.Count -gt 0) { Start-Sleep -Milliseconds 50 }
    }
} finally {
    Pop-Location
}

Write-Host "styleHints=$outDir"
