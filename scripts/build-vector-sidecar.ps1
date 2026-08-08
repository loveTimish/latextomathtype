param(
    [ValidateSet("windows-x64", "linux-x64", "all")]
    [string]$Platform = "all",
    [string]$OutDir = "target/vector-sidecar-dist"
)

$ErrorActionPreference = "Stop"
$nodeVersion = "24.9.0"
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$bundleHasher = [Security.Cryptography.IncrementalHash]::CreateHash(
    [Security.Cryptography.HashAlgorithmName]::SHA256)
try {
    $bundleInputs = @(
        (Join-Path $repoRoot "tools\mathjax\render_mathjax_svg.cjs"),
        (Join-Path $repoRoot "tools\mathjax\mathtype_fit.cjs"),
        (Join-Path $repoRoot "package-lock.json")
    )
    for ($index = 0; $index -lt $bundleInputs.Count; $index++) {
        $bundleHasher.AppendData([IO.File]::ReadAllBytes($bundleInputs[$index]))
        if ($index -lt $bundleInputs.Count - 1) {
            $bundleHasher.AppendData([byte[]]@(0))
        }
    }
    $bundleHash = -join ($bundleHasher.GetHashAndReset() | ForEach-Object { $_.ToString("x2") })
} finally {
    $bundleHasher.Dispose()
}
$outputRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot $OutDir))
$targetRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot "target"))
if (-not $outputRoot.StartsWith($targetRoot + [IO.Path]::DirectorySeparatorChar,
        [StringComparison]::OrdinalIgnoreCase)) {
    throw "OutDir must resolve below the repository target directory: $outputRoot"
}
New-Item -ItemType Directory -Force -Path $outputRoot | Out-Null

$platforms = if ($Platform -eq "all") { @("windows-x64", "linux-x64") } else { @($Platform) }
$specs = @{
    "windows-x64" = @{
        Archive = "node-v$nodeVersion-win-x64.zip"
        Sha256 = "6873514c3e6a012917cc6f95ce48a6289253370d025f1b69db290d70feebfa6e"
        NodeDir = "node-v$nodeVersion-win-x64"
        Format = "zip"
    }
    "linux-x64" = @{
        Archive = "node-v$nodeVersion-linux-x64.tar.xz"
        Sha256 = "f52ec50e959d72d5c680d9731420b2661cd2a8070e94c7369b6ddfcd8b7278be"
        NodeDir = "node-v$nodeVersion-linux-x64"
        Format = "tar.xz"
    }
}

foreach ($name in $platforms) {
    $spec = $specs[$name]
    $download = Join-Path $outputRoot $spec.Archive
    if (-not (Test-Path -LiteralPath $download)) {
        Invoke-WebRequest -Uri "https://nodejs.org/dist/v$nodeVersion/$($spec.Archive)" `
            -OutFile $download -UseBasicParsing
    }
    $actual = (Get-FileHash -LiteralPath $download -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($actual -ne $spec.Sha256) {
        throw "Node archive hash mismatch for $($spec.Archive): $actual"
    }

    $stage = Join-Path $outputRoot "stage-$name"
    if (Test-Path -LiteralPath $stage) {
        Remove-Item -LiteralPath $stage -Recurse -Force
    }
    New-Item -ItemType Directory -Force -Path $stage | Out-Null
    $extract = Join-Path $stage "extract"
    New-Item -ItemType Directory -Force -Path $extract | Out-Null
    if ($spec.Format -eq "zip") {
        Expand-Archive -LiteralPath $download -DestinationPath $extract
    } else {
        & tar -xf $download -C $extract
        $linuxNode = Join-Path $extract "$($spec.NodeDir)\bin\node"
        if ($LASTEXITCODE -ne 0 -and -not (Test-Path -LiteralPath $linuxNode)) {
            throw "tar extraction failed for $download"
        }
    }

    $package = Join-Path $stage "vector-sidecar"
    New-Item -ItemType Directory -Force -Path $package | Out-Null
    Move-Item -LiteralPath (Join-Path $extract $spec.NodeDir) -Destination (Join-Path $package "node")
    New-Item -ItemType Directory -Force -Path (Join-Path $package "tools\mathjax") | Out-Null
    Copy-Item -LiteralPath (Join-Path $repoRoot "tools\mathjax\render_mathjax_svg.cjs") `
        -Destination (Join-Path $package "tools\mathjax\render_mathjax_svg.cjs")
    Copy-Item -LiteralPath (Join-Path $repoRoot "tools\mathjax\mathtype_fit.cjs") `
        -Destination (Join-Path $package "tools\mathjax\mathtype_fit.cjs")
    Copy-Item -LiteralPath (Join-Path $repoRoot "package.json") -Destination $package
    Copy-Item -LiteralPath (Join-Path $repoRoot "package-lock.json") -Destination $package
    & npm ci --omit=dev --ignore-scripts --prefix $package
    if ($LASTEXITCODE -ne 0) { throw "npm ci failed while staging $name" }

    $artifact = Join-Path $outputRoot "vector-sidecar-$name"
    if ($name -eq "windows-x64") {
        $artifact += ".zip"
        if (Test-Path -LiteralPath $artifact) { Remove-Item -LiteralPath $artifact -Force }
        Compress-Archive -Path $package -DestinationPath $artifact -CompressionLevel Optimal
    } else {
        $artifact += ".tar.gz"
        if (Test-Path -LiteralPath $artifact) { Remove-Item -LiteralPath $artifact -Force }
        & python (Join-Path $repoRoot "scripts\package_linux_sidecar.py") $package $artifact
        if ($LASTEXITCODE -ne 0) { throw "Linux sidecar packaging failed for $artifact" }
    }
    $artifactHash = (Get-FileHash -LiteralPath $artifact -Algorithm SHA256).Hash.ToLowerInvariant()
    [PSCustomObject]@{
        platform = $name
        nodeVersion = "v$nodeVersion"
        mathJaxVersion = "3.2.2"
        bundleHash = $bundleHash
        artifact = $artifact
        sha256 = $artifactHash
    } | ConvertTo-Json | Set-Content -LiteralPath "$artifact.json" -Encoding utf8
    Write-Host "Built $artifact ($artifactHash)"
}
