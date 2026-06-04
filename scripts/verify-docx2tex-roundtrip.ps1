param(
    [string]$DocxPath = "target\rebuild\reorganized-paper-styled.docx",
    [string]$Docx2TexRoot = "J:\docx2tex",
    [string]$OutDir = "target\docx2tex-roundtrip-local-script",
    [string]$RequestJsonPath = "",
    [int]$MinimumTimesCount = 1,
    [int]$MinimumCdotsCount = 0,
    [int]$MinimumFractionCount = 1,
    [string[]]$ExpectedFragments = @()
)

$ErrorActionPreference = "Stop"

$resolvedDocx = Resolve-Path -LiteralPath $DocxPath
$resolvedDocx2Tex = Resolve-Path -LiteralPath $Docx2TexRoot
$d2t = Join-Path $resolvedDocx2Tex "d2t.bat"
$conf = Join-Path $resolvedDocx2Tex "conf\conf.xml"

if (-not (Test-Path -LiteralPath $d2t)) {
    throw "Missing docx2tex runner: $d2t"
}
if (-not (Test-Path -LiteralPath $conf)) {
    throw "Missing docx2tex config: $conf"
}

function Count-Literal([string]$Text, [string]$Needle) {
    if ([string]::IsNullOrEmpty($Needle)) {
        return 0
    }
    return [regex]::Matches($Text, [regex]::Escape($Needle)).Count
}

function Normalize-LatexFragment([string]$Value) {
    if ($null -eq $Value) {
        return ""
    }
    return ($Value -replace "\s+", "").Replace("\left", "").Replace("\right", "")
}

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
& $d2t $resolvedDocx $conf $OutDir
if ($LASTEXITCODE -ne 0) {
    throw "docx2tex failed with exit code $LASTEXITCODE"
}

$texName = [IO.Path]::GetFileNameWithoutExtension($resolvedDocx) + ".tex"
$texPath = Join-Path $OutDir $texName
if (-not (Test-Path -LiteralPath $texPath)) {
    throw "docx2tex did not create expected TeX file: $texPath"
}

$tex = Get-Content -LiteralPath $texPath -Raw
foreach ($needle in @("includegraphics", "times", "textbf")) {
    if (-not $tex.Contains($needle)) {
        throw "TeX output missing expected content: $needle"
    }
}
if ($tex -match "fatal-error|Template xsl:initial-template|I/O error|ERROR") {
    throw "TeX output contains error marker"
}

$badChars = @([string][char]0xFFFD, [string][char]0x25A1)
foreach ($badChar in $badChars) {
    if ($tex.Contains($badChar)) {
        $codePoint = "U+{0:X4}" -f [int][char]$badChar
        throw "TeX output contains malformed replacement character: $codePoint"
    }
}

$timesCount = Count-Literal $tex "\times"
$cdotsCount = Count-Literal $tex "\cdots"
$fractionCount = Count-Literal $tex "\frac"
if ($timesCount -lt $MinimumTimesCount) {
    throw "TeX output has too few \times tokens: actual=$timesCount minimum=$MinimumTimesCount"
}
if ($cdotsCount -lt $MinimumCdotsCount) {
    throw "TeX output has too few \cdots tokens: actual=$cdotsCount minimum=$MinimumCdotsCount"
}
if ($fractionCount -lt $MinimumFractionCount) {
    throw "TeX output has too few \frac tokens: actual=$fractionCount minimum=$MinimumFractionCount"
}

$normalizedTex = Normalize-LatexFragment $tex
foreach ($fragment in $ExpectedFragments) {
    $normalizedFragment = Normalize-LatexFragment $fragment
    if (-not $normalizedTex.Contains($normalizedFragment)) {
        throw "TeX output missing expected LaTeX fragment: $fragment"
    }
}

if ($RequestJsonPath -and $RequestJsonPath.Trim().Length -gt 0) {
    $resolvedRequestJson = Resolve-Path -LiteralPath $RequestJsonPath
    $requestJson = Get-Content -LiteralPath $resolvedRequestJson -Raw
    $requestTimesCount = Count-Literal $requestJson "\times"
    $requestCdotsCount = Count-Literal $requestJson "\cdots"
    $requestFractionCount = Count-Literal $requestJson "\frac"
    if ($requestTimesCount -gt 0 -and $timesCount -lt $requestTimesCount) {
        throw "TeX output lost \times tokens versus request JSON: actual=$timesCount request=$requestTimesCount"
    }
    if ($requestCdotsCount -gt 0 -and $cdotsCount -lt $requestCdotsCount) {
        throw "TeX output lost \cdots tokens versus request JSON: actual=$cdotsCount request=$requestCdotsCount"
    }
    if ($requestFractionCount -gt 0 -and $fractionCount -lt ($requestFractionCount - 5)) {
        throw "TeX output lost too many \frac tokens versus request JSON: actual=$fractionCount request=$requestFractionCount tolerance=5"
    }
}

Get-Item -LiteralPath $texPath | Select-Object FullName,Length,LastWriteTime
Write-Output "tokens: \times=$timesCount \cdots=$cdotsCount \frac=$fractionCount"
Write-Output "docx2tex roundtrip ok"
