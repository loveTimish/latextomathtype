param(
    [string]$Corpus = "docs\reference\mathtype\latex-coverage-corpus.json",
    [string]$ReferenceOverrides = "docs\reference\mathtype\desktop-reference-overrides.json",
    [string]$FormatLimitations = "docs\reference\mathtype\word-format-limitations.json",
    [string]$OutDir = "target\mathtype-standard",
    [int]$BatchSize = 25,
    [int]$ShardIndex = 0,
    [int]$ShardCount = 1,
    [int]$FormulaIndex = -1,
    [string]$ExpectedWordVersion = "16.0.20228.20124",
    [string]$ExpectedMathTypeVersion = "7.11.1.462",
    [switch]$Resume,
    [switch]$SkipVersionCheck,
    [switch]$Visible
)

$ErrorActionPreference = "Stop"
if ($BatchSize -lt 1) { throw "BatchSize must be >= 1" }
if ($ShardCount -lt 1 -or $ShardIndex -lt 0 -or $ShardIndex -ge $ShardCount) {
    throw "ShardIndex must be in [0, ShardCount)"
}

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$corpusPath = [IO.Path]::GetFullPath((Join-Path $repoRoot $Corpus))
$overridePath = [IO.Path]::GetFullPath((Join-Path $repoRoot $ReferenceOverrides))
$formatLimitationsPath = [IO.Path]::GetFullPath((Join-Path $repoRoot $FormatLimitations))
$outputRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot $OutDir))
$generator = Join-Path $PSScriptRoot "generate_word_mathtype_tex_reference.ps1"
New-Item -ItemType Directory -Force -Path $outputRoot | Out-Null
$preferenceDir = Join-Path $outputRoot "preferences"
$preferenceDocx = Join-Path $preferenceDir "default-12pt.docx"
$preferenceReport = Join-Path $preferenceDir "default-12pt.source-report.json"
if (-not ($Resume -and (Test-Path -LiteralPath $preferenceDocx) -and
        (Test-Path -LiteralPath $preferenceReport))) {
    & $generator -OutDocx $preferenceDocx -OutSourceReport $preferenceReport `
        -Formula @("A") -Visible:$Visible | Out-Host
}

function Get-WordEnvironment {
    $word = New-Object -ComObject Word.Application
    try {
        $exe = Join-Path $word.Path "WINWORD.EXE"
        $fileVersion = [Diagnostics.FileVersionInfo]::GetVersionInfo($exe).ProductVersion
        return [pscustomobject]@{ version = $fileVersion; path = $exe }
    } finally {
        $word.Quit() | Out-Null
        [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}

function Get-MathTypeEnvironment {
    $candidates = @(
        "${env:ProgramFiles(x86)}\MathType\MathType.exe",
        "$env:ProgramFiles\MathType\MathType.exe",
        "${env:ProgramFiles(x86)}\MathType 7\MathType.exe",
        "$env:ProgramFiles\MathType 7\MathType.exe"
    ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }
    if (-not $candidates) {
        $registry = Get-ChildItem HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall,
            HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall -ErrorAction SilentlyContinue |
            Get-ItemProperty -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -like "MathType*" } |
            Select-Object -First 1
        if ($registry -and $registry.InstallLocation) {
            $candidate = Join-Path $registry.InstallLocation "MathType.exe"
            if (Test-Path -LiteralPath $candidate) { $candidates = @($candidate) }
        }
    }
    if (-not $candidates) { throw "MathType.exe was not found" }
    $exe = @($candidates)[0]
    $version = [Diagnostics.FileVersionInfo]::GetVersionInfo($exe).ProductVersion
    return [pscustomobject]@{ version = $version; path = $exe }
}

$wordEnvironment = Get-WordEnvironment
$mathTypeEnvironment = Get-MathTypeEnvironment
if (-not $SkipVersionCheck) {
    if ($wordEnvironment.version -ne $ExpectedWordVersion) {
        throw "Word version mismatch: expected $ExpectedWordVersion, actual $($wordEnvironment.version)"
    }
    if ($mathTypeEnvironment.version -ne $ExpectedMathTypeVersion) {
        throw "MathType version mismatch: expected $ExpectedMathTypeVersion, actual $($mathTypeEnvironment.version)"
    }
}

$corpusData = Get-Content -LiteralPath $corpusPath -Raw -Encoding UTF8 | ConvertFrom-Json
$overrideData = Get-Content -LiteralPath $overridePath -Raw -Encoding UTF8 | ConvertFrom-Json
$formatLimitationsData = Get-Content -LiteralPath $formatLimitationsPath -Raw -Encoding UTF8 | ConvertFrom-Json
$referenceOverrideMap = @{}
foreach ($override in $overrideData.overrides) {
    $referenceOverrideMap[[string]$override.formula] = $override
}
$formatLimitationMap = @{}
foreach ($limitation in $formatLimitationsData.limitations) {
    $formatLimitationMap[[string]$limitation.formula] = $limitation
}
$seen = @{}
$formulas = New-Object System.Collections.Generic.List[object]
$globalIndex = 0
foreach ($entry in $corpusData.entries) {
    foreach ($example in $entry.examples) {
        $formula = [string]$example
        if ($seen.ContainsKey($formula)) { continue }
        $seen[$formula] = $true
        if (($globalIndex % $ShardCount) -eq $ShardIndex -and ($FormulaIndex -lt 0 -or $globalIndex -eq $FormulaIndex)) {
            $override = $referenceOverrideMap[$formula]
            $referenceMode = if ($override) { [string]$override.referenceMode } else { "tex-toggle" }
            $referenceFormula = if ($override -and $override.referenceFormula) {
                [string]$override.referenceFormula
            } else {
                $formula
            }
            $formatLimitation = $formatLimitationMap[$formula]
            $formulas.Add([pscustomobject]@{
                globalIndex = $globalIndex
                formula = $formula
                firstCommand = [string]$entry.command
                referenceMode = $referenceMode
                referenceFormula = $referenceFormula
                referenceReason = if ($override) { [string]$override.reason } else { "" }
                wordFormatMode = if ($formatLimitation) { [string]$formatLimitation.mode } else { "format" }
                wordFormatReason = if ($formatLimitation) { [string]$formatLimitation.reason } else { "" }
            })
        }
        $globalIndex++
    }
}
if ($FormulaIndex -ge 0 -and $formulas.Count -ne 1) {
    throw "FormulaIndex $FormulaIndex is not part of shard $ShardIndex/$ShardCount or does not exist"
}

$batchGroups = @()
$currentGroup = @()
$currentGroupKey = $null
foreach ($formulaItem in $formulas) {
    $itemIsNative = [string]$formulaItem.referenceMode -eq "native-template"
    $itemGroupKey = "$itemIsNative|$([string]$formulaItem.wordFormatMode)"
    if ($currentGroup.Count -gt 0 -and
            ($currentGroupKey -ne $itemGroupKey -or $currentGroup.Count -eq $BatchSize)) {
        $batchGroups += ,@($currentGroup)
        $currentGroup = @()
        $currentGroupKey = $null
    }
    if ($currentGroup.Count -eq 0) {
        $currentGroupKey = $itemGroupKey
    }
    $currentGroup += $formulaItem
}
if ($currentGroup.Count -gt 0) { $batchGroups += ,@($currentGroup) }

$batches = New-Object System.Collections.Generic.List[object]
for ($batchIndex = 0; $batchIndex -lt $batchGroups.Count; $batchIndex++) {
    $items = @($batchGroups[$batchIndex])
    $batchName = "batch-{0:D4}" -f $batchIndex
    $batchDir = Join-Path $outputRoot $batchName
    $referenceDocx = Join-Path $batchDir "standard.docx"
    $sourceReport = Join-Path $batchDir "standard.source-report.json"
    $formulaManifest = Join-Path $batchDir "formulas.json"
    New-Item -ItemType Directory -Force -Path $batchDir | Out-Null
    ConvertTo-Json -InputObject @($items) -Depth 5 |
        Set-Content -LiteralPath $formulaManifest -Encoding UTF8

    $nativeItems = @($items | Where-Object { $_.referenceMode -eq "native-template" })
    $requiresFormattedGeneratedReference = $nativeItems.Count -gt 0
    $requiresWordFormat = @($items | Where-Object { $_.wordFormatMode -eq "format" }).Count -gt 0
    if ($requiresFormattedGeneratedReference -and $nativeItems.Count -ne $items.Count) {
        throw "native-template and TeXToggle formulas must not share a batch: $batchName"
    }
    if ($requiresWordFormat -and @($items | Where-Object { $_.wordFormatMode -ne "format" }).Count -gt 0) {
        throw "Word-format and preserve formulas must not share a batch: $batchName"
    }
    if (-not $requiresFormattedGeneratedReference -and
            -not ($Resume -and (Test-Path -LiteralPath $referenceDocx) -and (Test-Path -LiteralPath $sourceReport))) {
        $formulaTexts = @($items | ForEach-Object { $_.referenceFormula })
        $arguments = @{
            OutDocx = $referenceDocx
            OutSourceReport = $sourceReport
            Formula = $formulaTexts
            Visible = $Visible
        }
        & $generator @arguments | Out-Host
    }
    $batches.Add([pscustomobject]@{
        name = $batchName
        formulaCount = $items.Count
        formulas = $formulaManifest
        standardDocx = $referenceDocx
        sourceReport = $sourceReport
        requiresFormattedGeneratedReference = $requiresFormattedGeneratedReference
        requiresWordFormat = $requiresWordFormat
    })
}

$manifest = [ordered]@{
    schemaVersion = 1
    sourceUrl = [string]$corpusData.sourceUrl
    sourceSha256 = [string]$corpusData.sourceSha256
    referenceOverrides = $overridePath
    referenceOverridesSchemaVersion = [int]$overrideData.schemaVersion
    formatLimitations = $formatLimitationsPath
    formatLimitationsSchemaVersion = [int]$formatLimitationsData.schemaVersion
    retrievedAt = [string]$corpusData.retrievedAt
    officialEntryCount = [int]$corpusData.entryCount
    uniqueFormulaCount = $globalIndex
    selectedFormulaCount = $formulas.Count
    shard = [ordered]@{ index = $ShardIndex; count = $ShardCount }
    batchSize = $BatchSize
    environment = [ordered]@{
        word = $wordEnvironment
        mathType = $mathTypeEnvironment
        defaultPointSize = 12
    }
    formatPreferenceDocx = $preferenceDocx
    formatPreferenceSourceReport = $preferenceReport
    batches = $batches
}
$manifestPath = Join-Path $outputRoot "manifest.json"
$manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
Write-Host "manifest=$manifestPath"
Write-Host "selectedFormulaCount=$($formulas.Count)"
