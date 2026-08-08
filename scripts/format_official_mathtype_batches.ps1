param(
    [string]$Manifest = "target\mathtype-standard\manifest.json",
    [switch]$Resume
)

$ErrorActionPreference = "Stop"
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$manifestPath = if ([IO.Path]::IsPathRooted($Manifest)) {
    [IO.Path]::GetFullPath($Manifest)
} else {
    [IO.Path]::GetFullPath((Join-Path $repoRoot $Manifest))
}
$data = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$formatter = Join-Path $PSScriptRoot "format_word_mathtype_equations.ps1"
$preferenceDocx = if ($data.formatPreferenceDocx -and
        (Test-Path -LiteralPath ([string]$data.formatPreferenceDocx))) {
    [string]$data.formatPreferenceDocx
} else {
    @($data.batches | ForEach-Object { [string]$_.standardDocx } |
        Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1)[0]
}
if ([string]::IsNullOrWhiteSpace($preferenceDocx)) {
    throw "manifest does not contain a standard DOCX for MathType clipboard preferences"
}
$expectedPreferenceSource = "clipboard Equation.DSMT4 preferences: $preferenceDocx"

function Invoke-IsolatedFormatter {
    param(
        [string]$InputDocx,
        [string]$OutputDocx,
        [string]$OutReport,
        [int]$ExpectedObjectCount,
        [int]$ChunkSize = 12,
        [int[]]$PreserveObjectIndices = @()
    )
    $powerShellExe = Join-Path $PSHOME "pwsh.exe"
    if (-not (Test-Path -LiteralPath $powerShellExe)) {
        $powerShellExe = Join-Path $PSHOME "powershell.exe"
    }
    $formatterArguments = @(
        "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", $formatter,
        "-InputDocx", $InputDocx, "-OutputDocx", $OutputDocx, "-OutReport", $OutReport,
        "-ExpectedObjectCount", "$ExpectedObjectCount", "-ChunkSize", "$ChunkSize",
        "-PreferenceSource", "Clipboard", "-PreferenceDocx", $preferenceDocx, "-Visible"
    )
    if ($PreserveObjectIndices.Count -gt 0) {
        $formatterArguments += "-PreserveObjectIndices"
        $formatterArguments += @($PreserveObjectIndices | ForEach-Object { "$_" })
    }
    & $powerShellExe @formatterArguments | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "isolated MathType formatter failed with exit code $LASTEXITCODE for $InputDocx"
    }
}

$limitationPath = Join-Path $repoRoot "docs\reference\mathtype\word-format-limitations.json"
$limitations = Get-Content -LiteralPath $limitationPath -Raw -Encoding UTF8 | ConvertFrom-Json
$selectionIndices = @{}
$shardIndices = @{}
$preserveIndices = @{}
foreach ($limitation in $limitations.limitations) {
    if ([string]$limitation.mode -eq "selection") {
        $selectionIndices[[int]$limitation.globalIndex] = $true
    } elseif ([string]$limitation.mode -eq "preserve") {
        $preserveIndices[[int]$limitation.globalIndex] = $true
    }
}
foreach ($group in $limitations.selectionGroups) {
    foreach ($globalIndex in $group.globalIndices) {
        $selectionIndices[[int]$globalIndex] = $true
    }
}
foreach ($group in $limitations.shardGroups) {
    foreach ($globalIndex in $group.globalIndices) {
        $shardIndices[[int]$globalIndex] = $true
    }
}
$reports = @()

foreach ($batch in $data.batches) {
    $batchDir = Split-Path -Parent ([string]$batch.standardDocx)
    $generated = Join-Path $batchDir "generated.docx"
    $formatted = Join-Path $batchDir "generated-formatted.docx"
    $report = Join-Path $batchDir "generated-formatted.report.json"
    if (-not (Test-Path -LiteralPath $generated)) {
        throw "missing generated batch DOCX: $generated"
    }
    if ($null -ne $batch.requiresWordFormat -and -not [bool]$batch.requiresWordFormat) {
        Copy-Item -LiteralPath $generated -Destination $formatted -Force
        $formatResult = [pscustomobject][ordered]@{
            schemaVersion = 1
            inputDocx = $generated
            outputDocx = $formatted
            preferenceSource = "preserved: pinned MathType formatter limitation"
            objectCountBefore = [int]$batch.formulaCount
            objectCountAfter = [int]$batch.formulaCount
            formattedObjectCount = 0
            preservedObjectCount = [int]$batch.formulaCount
            elapsedMs = 0
        }
        $formatResult | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $report -Encoding UTF8
        if ([bool]$batch.requiresFormattedGeneratedReference) {
            Copy-Item -LiteralPath $formatted -Destination ([string]$batch.standardDocx) -Force
        }
        $reports += $formatResult
        Write-Host "Preserve: $($batch.name) retained $($batch.formulaCount) editable objects (formatter limitation)"
        continue
    }
    $formulaRows = @(Get-Content -LiteralPath ([string]$batch.formulas) -Raw -Encoding UTF8 | ConvertFrom-Json)
    $requiresSelection = @($formulaRows | Where-Object {
        $selectionIndices.ContainsKey([int]$_.globalIndex)
    }).Count -gt 0
    $preserveCount = @($formulaRows | Where-Object {
        $preserveIndices.ContainsKey([int]$_.globalIndex)
    }).Count
    $preserveObjectIndices = @()
    for ($formulaIndex = 0; $formulaIndex -lt $formulaRows.Count; $formulaIndex++) {
        if ($preserveIndices.ContainsKey([int]$formulaRows[$formulaIndex].globalIndex)) {
            $preserveObjectIndices += $formulaIndex + 1
        }
    }
    if ($preserveCount -eq $formulaRows.Count) {
        Copy-Item -LiteralPath $generated -Destination $formatted -Force
        $formatResult = [pscustomobject][ordered]@{
            schemaVersion = 1
            inputDocx = $generated
            outputDocx = $formatted
            preferenceSource = "preserved: MathType formatter corrupts verified special-symbol encoding"
            objectCountBefore = [int]$batch.formulaCount
            objectCountAfter = [int]$batch.formulaCount
            formattedObjectCount = 0
            preservedObjectCount = [int]$batch.formulaCount
            elapsedMs = 0
        }
        $formatResult | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $report -Encoding UTF8
        $reports += $formatResult
        Write-Host "Preserve: $($batch.name) retained $($batch.formulaCount) verified special-symbol objects"
        continue
    }
    $requiresShards = @($formulaRows | Where-Object {
        $shardIndices.ContainsKey([int]$_.globalIndex)
    }).Count -gt 0
    if ($Resume -and (Test-Path -LiteralPath $formatted) -and (Test-Path -LiteralPath $report)) {
        $existing = Get-Content -LiteralPath $report -Raw -Encoding UTF8 | ConvertFrom-Json
        $isCurrent = (Get-Item -LiteralPath $formatted).LastWriteTimeUtc -ge
            (Get-Item -LiteralPath $generated).LastWriteTimeUtc
        $preferenceCurrent = if ($requiresShards) {
            [string]$existing.preferenceSource -eq "$expectedPreferenceSource; one-object shards"
        } else {
            [string]$existing.preferenceSource -eq $expectedPreferenceSource
        }
        $shardsCurrent = $true
        if ($requiresShards) {
            $shardDir = Join-Path $batchDir "format-shards"
            for ($index = 0; $index -lt $formulaRows.Count; $index++) {
                $inputShard = Join-Path $shardDir ("input-{0:D3}.docx" -f $index)
                if (-not (Test-Path -LiteralPath $inputShard) -or
                        (Get-Item -LiteralPath $inputShard).LastWriteTimeUtc -lt
                            (Get-Item -LiteralPath $generated).LastWriteTimeUtc) {
                    $shardsCurrent = $false
                    break
                }
            }
        }
        $preserveCurrent = [int]$existing.preservedObjectCount -eq $preserveCount
        if ($isCurrent -and $preferenceCurrent -and $shardsCurrent -and $preserveCurrent -and
                ([int]$existing.formattedObjectCount + [int]$existing.preservedObjectCount) -eq
                    [int]$batch.formulaCount) {
            if ([bool]$batch.requiresFormattedGeneratedReference) {
                Copy-Item -LiteralPath $formatted -Destination ([string]$batch.standardDocx) -Force
            }
            $reports += $existing
            Write-Host "Resume: $($batch.name) already formatted ($($batch.formulaCount) objects)"
            continue
        }
    }
    $chunkSize = if ($requiresSelection) { 1 } else { 12 }
    if ($requiresShards) {
        $classpathPath = Join-Path $repoRoot "target\classpath.txt"
        if (-not (Test-Path -LiteralPath $classpathPath)) {
            throw "missing Java classpath generated by Maven: $classpathPath"
        }
        $javaClasspath = (Join-Path $repoRoot "target\classes") + ";" +
            (Get-Content -LiteralPath $classpathPath -Raw).Trim()
        $shardDir = Join-Path $batchDir "format-shards"
        New-Item -ItemType Directory -Force -Path $shardDir | Out-Null
        $formattedShards = @()
        $formattedShardCount = 0
        $preservedShardCount = 0
        for ($index = 0; $index -lt $formulaRows.Count; $index++) {
            $inputShard = Join-Path $shardDir ("input-{0:D3}.docx" -f $index)
            $formattedShard = Join-Path $shardDir ("formatted-{0:D3}.docx" -f $index)
            $shardReport = Join-Path $shardDir ("formatted-{0:D3}.report.json" -f $index)
            $preserveShard = $preserveIndices.ContainsKey([int]$formulaRows[$index].globalIndex)
            $expectedShardPreference = if ($preserveShard) {
                "preserved: MathType formatter corrupts verified special-symbol encoding"
            } else {
                $expectedPreferenceSource
            }
            if ($Resume -and (Test-Path -LiteralPath $inputShard) -and
                    (Test-Path -LiteralPath $formattedShard) -and (Test-Path -LiteralPath $shardReport)) {
                $existingShard = Get-Content -LiteralPath $shardReport -Raw -Encoding UTF8 | ConvertFrom-Json
                $inputCurrent = (Get-Item -LiteralPath $inputShard).LastWriteTimeUtc -ge
                    (Get-Item -LiteralPath $generated).LastWriteTimeUtc
                $shardCurrent = (Get-Item -LiteralPath $formattedShard).LastWriteTimeUtc -ge
                    (Get-Item -LiteralPath $inputShard).LastWriteTimeUtc
                $shardPreferenceCurrent = [string]$existingShard.preferenceSource -eq
                    $expectedShardPreference
                if ($inputCurrent -and $shardCurrent -and $shardPreferenceCurrent -and
                        ([int]$existingShard.formattedObjectCount + [int]$existingShard.preservedObjectCount) -eq 1) {
                    $formattedShards += $formattedShard
                    $formattedShardCount += [int]$existingShard.formattedObjectCount
                    $preservedShardCount += [int]$existingShard.preservedObjectCount
                    continue
                }
            }
            & java -cp $javaClasspath com.lz.paperword.core.mtef.MathTypeFormulaDocxCli `
                $inputShard ([string]$formulaRows[$index].formula) | Out-Null
            if ($LASTEXITCODE -ne 0) { throw "failed to generate shard $index for $($batch.name)" }
            if ($preserveShard) {
                Copy-Item -LiteralPath $inputShard -Destination $formattedShard -Force
                $shardResult = [pscustomobject][ordered]@{
                    schemaVersion = 1
                    inputDocx = $inputShard
                    outputDocx = $formattedShard
                    preferenceSource = $expectedShardPreference
                    objectCountBefore = 1
                    objectCountAfter = 1
                    formattedObjectCount = 0
                    preservedObjectCount = 1
                    elapsedMs = 0
                }
                $shardResult | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $shardReport -Encoding UTF8
            } else {
                Invoke-IsolatedFormatter -InputDocx $inputShard -OutputDocx $formattedShard `
                    -OutReport $shardReport -ExpectedObjectCount 1
                $shardResult = Get-Content -LiteralPath $shardReport -Raw -Encoding UTF8 | ConvertFrom-Json
            }
            $formattedShards += $formattedShard
            $formattedShardCount += [int]$shardResult.formattedObjectCount
            $preservedShardCount += [int]$shardResult.preservedObjectCount
            # The MathType OLE server retains formatter state after Word exits.
            # A fresh server per shard prevents the next one-object document
            # from inheriting a blocked radical/template conversion state.
            $taskMathType = @(Get-Process MathType -ErrorAction SilentlyContinue |
                Where-Object { [string]::IsNullOrEmpty($_.MainWindowTitle) })
            foreach ($process in $taskMathType) {
                Stop-Process -Id $process.Id -Force
            }
            Start-Sleep -Milliseconds 500
        }
        & (Join-Path $PSScriptRoot "merge_word_mathtype_shards.ps1") `
            -InputDocx $formattedShards -OutputDocx $formatted `
            -ExpectedObjectCount ([int]$batch.formulaCount) | Out-Null
        $formatResult = [pscustomobject][ordered]@{
            schemaVersion = 1
            inputDocx = $generated
            outputDocx = $formatted
            preferenceSource = "$expectedPreferenceSource; one-object shards"
            objectCountBefore = [int]$batch.formulaCount
            objectCountAfter = [int]$batch.formulaCount
            formattedObjectCount = $formattedShardCount
            preservedObjectCount = $preservedShardCount
            chunkSize = 1
            shardCount = [int]$batch.formulaCount
        }
        $formatResult | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $report -Encoding UTF8
    } else {
        Invoke-IsolatedFormatter -InputDocx $generated -OutputDocx $formatted -OutReport $report `
            -ExpectedObjectCount ([int]$batch.formulaCount) -ChunkSize $chunkSize `
            -PreserveObjectIndices $preserveObjectIndices
        $formatResult = Get-Content -LiteralPath $report -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    if (([int]$formatResult.formattedObjectCount + [int]$formatResult.preservedObjectCount) -ne [int]$batch.formulaCount) {
        throw "MathType formatted $($formatResult.formattedObjectCount) of $($batch.formulaCount) objects in $($batch.name)"
    }
    if ([bool]$batch.requiresFormattedGeneratedReference) {
        Copy-Item -LiteralPath $formatted -Destination ([string]$batch.standardDocx) -Force
    }
    $reports += $formatResult
}

$summaryPath = Join-Path (Split-Path -Parent $manifestPath) "format-summary.json"
[ordered]@{
    schemaVersion = 1
    batchCount = @($data.batches).Count
    formulaCount = [int]$data.selectedFormulaCount
    reports = $reports
} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $summaryPath -Encoding UTF8
Write-Host "MathType formatted $($data.selectedFormulaCount) formulas; report=$summaryPath"
