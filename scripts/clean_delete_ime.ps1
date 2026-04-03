param(
    [string]$InstallRoot = "$env:LOCALAPPDATA\\yuninput",
    [switch]$SkipElevation
)

$ErrorActionPreference = 'Stop'

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

$tipGuid = '{6DE9AB40-3BA8-4B77-8D8F-233966E1C102}'
$profileGuid = '{47DE2FB1-F5E4-4CF8-AB2F-8F7A761731B2}'
$langHex = '0x00000804'

$uninstallScript = Join-Path $PSScriptRoot 'uninstall_clean.ps1'
if (-not (Test-Path $uninstallScript)) {
    throw "Missing script: $uninstallScript"
}

$uninstallArgs = @{
    InstallRoot = $InstallRoot
    SkipElevation = $SkipElevation.IsPresent
}

Write-Host 'Running uninstall_clean.ps1...'
& $uninstallScript @uninstallArgs

$hkcuTip = "Registry::HKEY_CURRENT_USER\\Software\\Microsoft\\CTF\\TIP\\$tipGuid\\LanguageProfile\\$langHex\\$profileGuid"
$hkcuClsid = "Registry::HKEY_CURRENT_USER\\Software\\Classes\\CLSID\\$tipGuid"
$installDirExists = Test-Path $InstallRoot
$hkcuTipExists = Test-Path $hkcuTip
$hkcuClsidExists = Test-Path $hkcuClsid

Write-Host ''
Write-Host '=== Clean Delete Verification ==='
Write-Host "Install root exists: $installDirExists"
Write-Host "HKCU TIP exists: $hkcuTipExists"
Write-Host "HKCU CLSID exists: $hkcuClsidExists"

if ($installDirExists) {
    $leftovers = Get-ChildItem -Path $InstallRoot -Recurse -Force -ErrorAction SilentlyContinue
    if ($leftovers.Count -gt 0) {
        Write-Host ''
        Write-Host 'Locked leftovers (often safe if no longer registered):'
        $leftovers | Select-Object FullName | Format-Table -AutoSize
    }
}

if (-not $hkcuTipExists -and -not $hkcuClsidExists) {
    Write-Host ''
    Write-Host 'Current-user IME registration cleaned successfully.'
}
else {
    Write-Host ''
    Write-Warning 'Some current-user registration keys still exist. You can rerun this script once more.'
}
