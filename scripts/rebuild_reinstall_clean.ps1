param(
    [string]$InstallRoot = "$env:LOCALAPPDATA\\yuninput",
    [switch]$SkipElevation,
    [switch]$NoBuild
)

$ErrorActionPreference = 'Stop'

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

$cleanScript = Join-Path $PSScriptRoot 'clean_delete_ime.ps1'
$installScript = Join-Path $PSScriptRoot 'install_enable.ps1'
$inspectScript = Join-Path $PSScriptRoot 'inspect_ime_state.ps1'

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]::new($identity)
$isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$inVsCodeHost = -not [string]::IsNullOrWhiteSpace($env:VSCODE_PID)

$effectiveSkipElevation = $SkipElevation.IsPresent
if (-not $isAdmin -and -not $effectiveSkipElevation -and $inVsCodeHost) {
    $effectiveSkipElevation = $true
    Write-Warning 'Detected non-admin VS Code terminal. Auto-enabling -SkipElevation to avoid hidden UAC prompt blocking.'
}

foreach ($path in @($cleanScript, $installScript, $inspectScript)) {
    if (-not (Test-Path $path)) {
        throw "Missing required script: $path"
    }
}

function Stop-InstallationInterferingProcesses {
    $processNames = @('ctfmon', 'TextInputHost', 'yuninput_config')
    foreach ($name in $processNames) {
        Get-Process -Name $name -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
}

function Invoke-Step {
    param(
        [string]$Title,
        [scriptblock]$Action
    )

    Write-Host ''
    Write-Host ("=== {0} ===" -f $Title)
    & $Action
}

function Get-ComparablePath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ''
    }

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $collapsed = [System.Text.RegularExpressions.Regex]::Replace($fullPath, '\\+', '\\')
    return $collapsed.TrimEnd('\\')
}

Stop-InstallationInterferingProcesses

$cleanArgs = @{
    InstallRoot = $InstallRoot
    SkipElevation = $effectiveSkipElevation
}

Invoke-Step -Title 'Clean Delete' -Action {
    & $cleanScript @cleanArgs
}

$installArgs = @{
    InstallRoot = $InstallRoot
    NonInteractive = $false
}

if (-not $NoBuild) {
    $installArgs.Build = $true
}
if ($effectiveSkipElevation) {
    $installArgs.SkipElevation = $true
}

Invoke-Step -Title 'Build and Install' -Action {
    & $installScript @installArgs
}

Invoke-Step -Title 'Post-install Verify' -Action {
    & $inspectScript
}

$reportPath = Join-Path $env:LOCALAPPDATA 'yuninput\ime_state_report.txt'
if (-not (Test-Path $reportPath)) {
    throw "Verify report missing: $reportPath"
}

$reportLines = Get-Content -Path $reportPath -ErrorAction Stop

$profileExists = $false
$inprocValue = ''
foreach ($line in $reportLines) {
    if ($line -match '^HKCU profile exists:\s*(True|False)') {
        $profileExists = [bool]::Parse($Matches[1])
    }
    if ($line -match '^HKCU InprocServer32:\s*(.*)$') {
        $inprocValue = $Matches[1].Trim()
    }
}

$normalizedRoot = Get-ComparablePath -Path $InstallRoot
$expectedBinPrefix = ($normalizedRoot + '\\bin\\').ToLowerInvariant()
$normalizedInproc = (Get-ComparablePath -Path $inprocValue).ToLowerInvariant()
$fileName = [System.IO.Path]::GetFileName($inprocValue)
$fileNameValid = $fileName -match '^yuninput(_\d{8}_\d{6})?\.dll$'

$inprocLooksValid =
    -not [string]::IsNullOrWhiteSpace($inprocValue) -and
    $normalizedInproc.StartsWith($expectedBinPrefix) -and
    $fileNameValid

Write-Host ''
Write-Host '=== Final Verdict ==='
Write-Host "HKCU Profile exists: $profileExists"
Write-Host "HKCU InprocServer32: $inprocValue"
Write-Host "HKCU Inproc path valid: $inprocLooksValid"

if (-not $profileExists -or -not $inprocLooksValid) {
    throw 'Reinstall verification failed: HKCU registration is not in expected state.'
}

Write-Host ''
Write-Host 'Rebuild + clean delete + reinstall completed successfully.'
