$EdgeCandidates = @(
    "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
    "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
    "$env:LOCALAPPDATA\Microsoft\Edge\Application\msedge.exe"
)

$EdgePath = $EdgeCandidates |
    Where-Object { Test-Path $_ } |
    Select-Object -First 1

if (-not $EdgePath) {
    Write-Host "[ERROR] Microsoft Edge executable was not found."
    exit 1
}

$ProfilePath = Join-Path $PSScriptRoot ".tiktok_edge_profile"

Write-Host ""
Write-Host "TikTok Edge starting..."
Write-Host "Edge    : $EdgePath"
Write-Host "Profile : $ProfilePath"
Write-Host "CDP     : http://127.0.0.1:9222"
Write-Host ""

& $EdgePath `
    --remote-debugging-port=9222 `
    --user-data-dir="$ProfilePath" `
    "https://www.tiktok.com/"