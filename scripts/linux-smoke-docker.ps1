param(
    [string]$Image = "latextomathtype:smoke"
)

$ErrorActionPreference = "Stop"

docker run --rm `
    --entrypoint bash `
    -v "${PWD}:/workspace" `
    -w /workspace `
    $Image `
    -lc "java -version && sh scripts/linux-smoke.sh"
