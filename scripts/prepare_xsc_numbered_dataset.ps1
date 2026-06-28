param(
    [string]$SourceDir = "E:\新加卷\新建文件夹\xsc资料",
    [string]$OutDir = "J:\latextomathtype\analysis\xsc-numbered-dataset",
    [int]$Limit = 0,
    [switch]$Recurse
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $SourceDir)) {
    throw "missing source dir: $SourceDir"
}

$searchParams = @{
    LiteralPath = $SourceDir
    File = $true
    Filter = "*.docx"
}
if ($Recurse) {
    $searchParams.Recurse = $true
}

$files = Get-ChildItem @searchParams |
    Where-Object { $_.Name -notlike "~$*" } |
    Sort-Object FullName

if ($Limit -gt 0) {
    $files = $files | Select-Object -First $Limit
}

if (-not $files -or $files.Count -eq 0) {
    throw "no docx files found under: $SourceDir"
}

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$manifest = @()
$index = 1
foreach ($file in $files) {
    $target = Join-Path $OutDir "$index.docx"
    Copy-Item -LiteralPath $file.FullName -Destination $target -Force
    $manifest += [pscustomobject]@{
        index = $index
        numberedPath = $target
        sourcePath = $file.FullName
        sourceName = $file.Name
        length = $file.Length
        lastWriteTime = $file.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
    }
    $index++
}

$manifestPath = Join-Path $OutDir "manifest.json"
$manifest | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
Write-Output "numbered=$($manifest.Count)"
Write-Output "outDir=$OutDir"
Write-Output "manifest=$manifestPath"
