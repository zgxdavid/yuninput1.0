param(
    [Parameter(Mandatory = $true)]
    [string]$Version,
    [string]$MsiSource = ''
)

$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
$workspaceRoot = Split-Path -Parent $projectRoot

if ([string]::IsNullOrWhiteSpace($MsiSource)) {
    $candidates = Get-ChildItem -Path $projectRoot -Filter 'Yuninput*.msi' -File |
        Sort-Object LastWriteTime -Descending
    if ($candidates.Count -eq 0) {
        throw "No MSI found under project root: $projectRoot"
    }
    $MsiSource = $candidates[0].FullName
}

if (!(Test-Path $MsiSource)) {
    throw "MSI source not found: $MsiSource"
}

$targetName = "Yuninput$Version.msi"
$targetPath = Join-Path $projectRoot $targetName

Copy-Item -Force $MsiSource $targetPath

Write-Host "Prepared patch release artifact: $targetPath"
Write-Host "Source MSI: $MsiSource"
Write-Host ""
Write-Host "Suggested release steps:" 
Write-Host "  git add ."
Write-Host ('  git commit -m "release: v' + $Version + '"')
Write-Host "  git tag v$Version"
Write-Host "  git push origin main --tags"
