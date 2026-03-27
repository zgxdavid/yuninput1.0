param(
    [string]$DllPath
)

$ErrorActionPreference = "Stop"

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

if ([string]::IsNullOrWhiteSpace($DllPath)) {
    $projectRoot = Split-Path -Parent $PSScriptRoot
    $DllPath = Join-Path $projectRoot "build\Release\yuninput.dll"
}

if (-not (Test-Path -Path $DllPath -PathType Leaf)) {
    throw "DLL not found: $DllPath"
}

$DllPath = (Resolve-Path -Path $DllPath).Path

Write-Host "Unregistering IME DLL: $DllPath"
& regsvr32.exe /u /s "$DllPath"
$code = if ($null -eq $LASTEXITCODE) { 0 } else { [int]$LASTEXITCODE }
if ($code -ne 0) {
    throw "regsvr32 /u failed with exit code: $code"
}

Write-Host "Unregister done."
