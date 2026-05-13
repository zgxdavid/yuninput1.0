$ErrorActionPreference = "Stop"

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

$vsInstaller = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vs_installer.exe"
$vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"

if (-not (Test-Path $vsInstaller)) {
    throw "vs_installer.exe not found: $vsInstaller"
}
if (-not (Test-Path $vswhere)) {
    throw "vswhere.exe not found: $vswhere"
}

function Invoke-InstallerModify {
    param(
        [Parameter(Mandatory = $true)][string]$InstallPath,
        [Parameter(Mandatory = $true)][string]$Workload
    )

    Write-Host "Modifying Visual Studio instance: $InstallPath"
    & $vsInstaller modify --installPath "$InstallPath" --add $Workload --includeRecommended --includeOptional --passive --norestart
    if ($LASTEXITCODE -ne 0) {
        throw "VS installer modify failed with exit code: $LASTEXITCODE"
    }
}

function Get-Vs2022Instances {
    $json = & $vswhere -all -products * -version "[17.0,18.0)" -format json 2>$null
    if ([string]::IsNullOrWhiteSpace($json)) {
        return @()
    }

    $items = $json | ConvertFrom-Json
    if ($items -isnot [System.Array]) {
        $items = @($items)
    }

    return @(
        $items |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.installationPath) } |
            ForEach-Object {
                [PSCustomObject]@{
                    InstallPath = ([string]$_.installationPath).Trim()
                    IsComplete = [bool]$_.isComplete
                    IsLaunchable = [bool]$_.isLaunchable
                }
            } |
            Group-Object InstallPath |
            ForEach-Object { $_.Group[0] }
    )
}

function Get-VcvarsPath {
    param(
        [Parameter(Mandatory = $true)][string]$InstallPath
    )

    return Join-Path $InstallPath "VC\Auxiliary\Build\vcvars64.bat"
}

$instances = Get-Vs2022Instances
$readyInstance = $instances |
    Where-Object {
        (Test-Path (Get-VcvarsPath -InstallPath $_.InstallPath)) -and
        $_.IsComplete -and
        $_.IsLaunchable
    } |
    Select-Object -First 1

if ($readyInstance) {
    $vcvars = Get-VcvarsPath -InstallPath $readyInstance.InstallPath
    Write-Host "VC toolchain is already ready: $vcvars"
    Write-Host "You can now run:"
    Write-Host "  scripts/build_release.ps1 -Clean"
    exit 0
}

$vcvarsOnlyInstance = $instances | Where-Object { Test-Path (Get-VcvarsPath -InstallPath $_.InstallPath) } | Select-Object -First 1
if ($vcvarsOnlyInstance) {
    Write-Warning "Detected VC toolchain files, but Visual Studio instance is not complete/launchable. Will trigger installer repair."
}

$installPaths = $instances | Select-Object -ExpandProperty InstallPath
$communityPath = $installPaths | Where-Object { $_ -notmatch '\\BuildTools$' } | Select-Object -First 1

$targetInstallPath = $null
$workload = $null

if ($communityPath -and (Test-Path $communityPath)) {
    $targetInstallPath = $communityPath
    $workload = "Microsoft.VisualStudio.Workload.NativeDesktop"
}
else {
    $buildToolsPath = & $vswhere -all -latest -products * -version "[17.0,18.0)" -requires Microsoft.VisualStudio.Workload.VCTools -property installationPath 2>$null
    $buildToolsPath = if ([string]::IsNullOrWhiteSpace($buildToolsPath)) { $null } else { $buildToolsPath.Trim() }

    if ($buildToolsPath -and (Test-Path $buildToolsPath)) {
        $targetInstallPath = $buildToolsPath
        $workload = "Microsoft.VisualStudio.Workload.VCTools"
    }
}

if ([string]::IsNullOrWhiteSpace($targetInstallPath)) {
    throw "No VS2022 instance found to repair."
}

Invoke-InstallerModify -InstallPath $targetInstallPath -Workload $workload

$vcvars = Get-VcvarsPath -InstallPath $targetInstallPath
if (-not (Test-Path $vcvars)) {
    throw "Repair completed but vcvars64.bat still missing: $vcvars"
}

Write-Host "VC toolchain is ready: $vcvars"
Write-Host "You can now run:"
Write-Host "  scripts/build_release.ps1 -Clean"
