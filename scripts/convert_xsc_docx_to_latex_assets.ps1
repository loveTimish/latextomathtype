param(
    [string]$SourceDir = "E:\新加卷\新建文件夹\xsc资料",
    [string]$Docx2TexDir = "J:\latextomathtype\analysis\tools\docx2tex-1.10-release\docx2tex",
    [string]$OutRoot = "J:\latextomathtype\analysis\xsc-latex-fixed",
    [string]$WorkRoot = "J:\latextomathtype\analysis\xsc-docx-ascii-work",
    [int]$Start = 1,
    [int]$End = 0,
    [int]$Limit = 0,
    [switch]$Recurse,
    [switch]$Force
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $SourceDir)) {
    throw "missing source dir: $SourceDir"
}
if (-not (Test-Path -LiteralPath $Docx2TexDir)) {
    throw "missing docx2tex dir: $Docx2TexDir"
}

$d2t = Join-Path $Docx2TexDir "d2t.bat"
$conf = Join-Path $Docx2TexDir "conf\conf.xml"
if (-not (Test-Path -LiteralPath $d2t)) {
    throw "missing d2t.bat: $d2t"
}
if (-not (Test-Path -LiteralPath $conf)) {
    throw "missing conf.xml: $conf"
}

$searchParams = @{
    LiteralPath = $SourceDir
    File = $true
    Filter = "*.docx"
}
if ($Recurse) {
    $searchParams.Recurse = $true
}

$files = @(Get-ChildItem @searchParams |
    Where-Object { $_.Name -notlike "~$*" } |
    Sort-Object FullName)

if ($Limit -gt 0) {
    $files = @($files | Select-Object -First $Limit)
}
if ($End -le 0 -or $End -gt $files.Count) {
    $End = $files.Count
}
if ($Start -lt 1 -or $Start -gt $End) {
    throw "invalid range: $Start..$End for file count $($files.Count)"
}

New-Item -ItemType Directory -Force -Path $OutRoot | Out-Null
New-Item -ItemType Directory -Force -Path $WorkRoot | Out-Null
$manifestPath = Join-Path $OutRoot "manifest.json"
$failurePath = Join-Path $OutRoot "failures.json"
$manifest = @()
$failures = @()

for ($i = $Start; $i -le $End; $i++) {
    $file = $files[$i - 1]
    $outDir = Join-Path $OutRoot "$i"
    $asciiDocx = Join-Path $WorkRoot "$i.docx"
    $texPath = Join-Path $outDir "$i.tex"
    $stableTexPath = Join-Path $outDir "$i.tex"
    $logPath = Join-Path $outDir "$i.convert.log"
    New-Item -ItemType Directory -Force -Path $outDir | Out-Null

    if ((-not $Force) -and (Test-Path -LiteralPath $stableTexPath)) {
        Write-Output "skip existing ${i}: $($file.Name)"
    } else {
        Write-Output "convert ${i}/${End}: $($file.FullName)"
        Copy-Item -LiteralPath $file.FullName -Destination $asciiDocx -Force
        $outDirForD2t = $outDir
        if (-not $outDirForD2t.EndsWith([IO.Path]::DirectorySeparatorChar)) {
            $outDirForD2t += [IO.Path]::DirectorySeparatorChar
        }
        & $d2t $asciiDocx $conf $outDirForD2t *> $logPath
        if ($LASTEXITCODE -ne 0) {
            $failures += [pscustomobject]@{
                index = $i
                sourcePath = $file.FullName
                sourceName = $file.Name
                logPath = $logPath
                exitCode = $LASTEXITCODE
            }
            continue
        }
        if (-not (Test-Path -LiteralPath $texPath)) {
            $candidate = Get-ChildItem -LiteralPath $outDir -File -Filter "*.tex" |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1
            if ($candidate) {
                $texPath = $candidate.FullName
            }
        }
        if (-not (Test-Path -LiteralPath $texPath)) {
            $failures += [pscustomobject]@{
                index = $i
                sourcePath = $file.FullName
                sourceName = $file.Name
                logPath = $logPath
                exitCode = 0
                error = "tex output not found"
            }
            continue
        }
    }

    $manifest += [pscustomobject]@{
        index = $i
        sourcePath = $file.FullName
        sourceName = $file.Name
        latexPath = $stableTexPath
        outDir = $outDir
        length = $file.Length
        lastWriteTime = $file.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
    }
    $manifest | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
    $failures | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $failurePath -Encoding UTF8
}

Write-Output "converted=$($manifest.Count)"
Write-Output "failed=$($failures.Count)"
Write-Output "manifest=$manifestPath"
Write-Output "failures=$failurePath"
