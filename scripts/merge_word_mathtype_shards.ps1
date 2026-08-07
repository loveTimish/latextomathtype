param(
    [Parameter(Mandatory = $true)]
    [string[]]$InputDocx,
    [Parameter(Mandatory = $true)]
    [string]$OutputDocx,
    [ValidateRange(1, 1000000)]
    [int]$ExpectedObjectCount
)

$ErrorActionPreference = "Stop"

function Resolve-RequiredPath([string]$Path) {
    $resolved = if ([IO.Path]::IsPathRooted($Path)) {
        [IO.Path]::GetFullPath($Path)
    } else {
        [IO.Path]::GetFullPath((Join-Path (Get-Location) $Path))
    }
    if (-not (Test-Path -LiteralPath $resolved)) { throw "file not found: $resolved" }
    return $resolved
}

function Resolve-OutputPath([string]$Path) {
    if ([IO.Path]::IsPathRooted($Path)) { return [IO.Path]::GetFullPath($Path) }
    return [IO.Path]::GetFullPath((Join-Path (Get-Location) $Path))
}

function Assert-MathTypeObjects($Document, [int]$Expected, [string]$Label) {
    $actual = $Document.InlineShapes.Count
    if ($actual -ne $Expected) {
        throw "$Label contains $actual inline shapes; expected $Expected"
    }
    for ($index = 1; $index -le $actual; $index++) {
        $progId = ""
        try { $progId = [string]$Document.InlineShapes.Item($index).OLEFormat.ProgID } catch {}
        if ($progId -ne "Equation.DSMT4") {
            throw "$Label inline shape $index is not Equation.DSMT4: $progId"
        }
    }
}

$inputs = @($InputDocx | ForEach-Object { Resolve-RequiredPath $_ })
$output = Resolve-OutputPath $OutputDocx
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $output) | Out-Null
$temporary = Join-Path (Split-Path -Parent $output) (([IO.Path]::GetFileNameWithoutExtension($output)) + ".merge-" + [guid]::NewGuid() + ".docx")

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$destination = $null
try {
    $destination = $word.Documents.Add()
    foreach ($input in $inputs) {
        $source = $word.Documents.Open([string]$input, $false, $true)
        try {
            Assert-MathTypeObjects $source $source.InlineShapes.Count $input
            $insertAt = $destination.Content.End - 1
            $range = $destination.Range($insertAt, $insertAt)
            $range.InsertFile([string]$input)
        } finally {
            $source.Close($false) | Out-Null
            [Runtime.InteropServices.Marshal]::ReleaseComObject($source) | Out-Null
        }
    }
    Assert-MathTypeObjects $destination $ExpectedObjectCount "merged document"
    $destination.SaveAs2([string]$temporary, 16)
    $destination.Close($false) | Out-Null
    [Runtime.InteropServices.Marshal]::ReleaseComObject($destination) | Out-Null
    $destination = $null

    $verification = $word.Documents.Open([string]$temporary, $false, $true)
    try {
        Assert-MathTypeObjects $verification $ExpectedObjectCount "saved merged document"
    } finally {
        $verification.Close($false) | Out-Null
        [Runtime.InteropServices.Marshal]::ReleaseComObject($verification) | Out-Null
    }
    Move-Item -LiteralPath $temporary -Destination $output -Force
    [pscustomobject]@{ OutputDocx = $output; ObjectCount = $ExpectedObjectCount; ShardCount = $inputs.Count }
} finally {
    if ($destination) {
        $destination.Close($false) | Out-Null
        [Runtime.InteropServices.Marshal]::ReleaseComObject($destination) | Out-Null
    }
    $word.Quit() | Out-Null
    [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force }
}
