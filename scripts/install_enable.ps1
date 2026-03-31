param(
    [switch]$Build,
    [string]$InstallRoot = "$env:LOCALAPPDATA\\yuninput",
    [string]$LogPath = "",
    [switch]$SkipElevation,
    [switch]$NonInteractive
)

$ErrorActionPreference = 'Stop'

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

$defaultInstallRoot = "$env:LOCALAPPDATA\\yuninput"
if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogPath = Join-Path $defaultInstallRoot 'install.log'
}

function Write-InstallLog {
    param([string]$Message)

    try {
        $logDir = Split-Path -Parent $LogPath
        if (-not [string]::IsNullOrWhiteSpace($logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }

        $line = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
        Add-Content -Path $LogPath -Value $line -Encoding utf8
    }
    catch {
        # Logging must not block installation.
    }
}

# Tolerate typo usage like: - Build
if (-not $Build -and $null -ne $args -and $args.Count -gt 0) {
    foreach ($arg in $args) {
        if ($arg -match '^(Build|-Build|/Build)$') {
            $Build = $true
            Write-Warning "Detected malformed switch usage. Interpreting as -Build."
            break
        }
    }
}

if ([string]::IsNullOrWhiteSpace($InstallRoot) -or $InstallRoot -eq '-') {
    Write-Warning "Invalid InstallRoot '$InstallRoot'. Falling back to default: $defaultInstallRoot"
    $InstallRoot = $defaultInstallRoot
}

if ($InstallRoot.StartsWith('-')) {
    throw "Invalid InstallRoot '$InstallRoot'. Did you mean to use '-Build' (without a space)?"
}

$InstallRoot = $InstallRoot.Trim()
$InstallRoot = $InstallRoot.Trim('"')
if ($InstallRoot.Length -gt 3) {
    # MSI directory properties often end with a trailing slash; normalize to avoid quoting edge cases.
    $InstallRoot = $InstallRoot.TrimEnd('\\')
}

$InstallRoot = [Environment]::ExpandEnvironmentVariables($InstallRoot)

try {
    $InstallRoot = [System.IO.Path]::GetFullPath($InstallRoot)
}
catch {
    throw "Invalid InstallRoot '$InstallRoot': $($_.Exception.Message)"
}

if ($InstallRoot.Length -gt 3) {
    $InstallRoot = $InstallRoot.TrimEnd('\\')
}

if (-not [System.IO.Path]::IsPathRooted($InstallRoot)) {
    throw "InstallRoot must be an absolute path. Current value: $InstallRoot"
}

Write-InstallLog "install_enable.ps1 start. Build=$Build InstallRoot=$InstallRoot"
Write-InstallLog "install_enable.ps1 mode. SkipElevation=$SkipElevation NonInteractive=$NonInteractive"

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]::new($identity)
$isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin -and -not $SkipElevation) {
    $quotedScript = '"' + $PSCommandPath + '"'
    $quotedInstallRoot = '"' + $InstallRoot + '"'
    $quotedLogPath = '"' + $LogPath + '"'

    $argumentString = "-NoProfile -ExecutionPolicy Bypass -File $quotedScript"

    if ($Build) {
        $argumentString += ' -Build'
    }

    if ($InstallRoot -ne $defaultInstallRoot) {
        $argumentString += " -InstallRoot $quotedInstallRoot"
    }

    if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
        $argumentString += " -LogPath $quotedLogPath"
    }

    $argumentString += ' -SkipElevation'

    Write-Host 'Requesting administrator permission...'
    Write-InstallLog 'Requesting elevation via UAC.'
    Write-InstallLog "Elevation command: powershell.exe $argumentString"
    try {
        $proc = Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList $argumentString -Wait -PassThru
        Write-InstallLog "Elevated process exit code: $($proc.ExitCode)"
        if ($proc.ExitCode -ne 0) {
            Write-Error "Elevated installer failed with exit code: $($proc.ExitCode). Log: $LogPath"
            if (Test-Path $LogPath) {
                Write-Host ''
                Write-Host 'Last install log lines:'
                Get-Content $LogPath -Tail 40
            }
        }
        exit $proc.ExitCode
    }
    catch {
        $nativeCode = 0
        if ($_.Exception -is [System.ComponentModel.Win32Exception]) {
            $nativeCode = $_.Exception.NativeErrorCode
        }

        if ($nativeCode -eq 1223) {
            Write-Error "UAC prompt was canceled by user. Installation aborted."
            Write-InstallLog 'Elevation canceled by user (ERROR_CANCELLED=1223).'
        }
        else {
            Write-Error "Elevation failed: $($_.Exception.Message)"
            Write-InstallLog "Elevation failed: $($_.Exception.Message)"
        }

        if (Test-Path $LogPath) {
            Write-Host ''
            Write-Host "Install log: $LogPath"
        }
        exit 1
    }
}
elseif (-not $isAdmin -and $SkipElevation) {
    Write-InstallLog 'SkipElevation specified without admin rights. Continuing without UAC.'
    Write-Warning 'Running without admin rights. Registration may fail. Prefer running this script from an Administrator PowerShell.'
}

$projectRoot = Split-Path -Parent $PSScriptRoot
$buildScript = Join-Path $PSScriptRoot 'build_release.ps1'
$configBuildScript = Join-Path $PSScriptRoot 'build_config_app.ps1'
$registerScript = Join-Path $PSScriptRoot 'register_ime.ps1'
$unregisterScript = Join-Path $PSScriptRoot 'unregister_ime.ps1'

$srcDll = Join-Path $projectRoot 'build\Release\yuninput.dll'
$fallbackSrcDll = Join-Path $projectRoot 'bin\yuninput.dll'
$srcDataDir = Join-Path $projectRoot 'data'
$srcConfigSource = Join-Path $projectRoot 'tools\YuninputConfig.cs'
$srcConfigExe = Join-Path $projectRoot 'tools\yuninput_config.exe'
$fallbackSrcConfigExe = Join-Path $projectRoot 'yuninput_config.exe'

$binDir = Join-Path $InstallRoot 'bin'
$dataDir = Join-Path $InstallRoot 'data'
$dllPath = Join-Path $binDir 'yuninput.dll'
$settingsRoot = Join-Path $env:LOCALAPPDATA 'yuninput'
$settingsPath = Join-Path $settingsRoot 'settings.json'

function Stop-InputProcesses {
    $processNames = @('ctfmon', 'TextInputHost')
    foreach ($name in $processNames) {
        Get-Process -Name $name -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
}

function Copy-WithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Source,
        [Parameter(Mandatory = $true)]
        [string]$Destination,
        [int]$RetryCount = 8,
        [int]$RetryDelayMs = 500
    )

    for ($attempt = 1; $attempt -le $RetryCount; $attempt++) {
        try {
            Copy-Item $Source $Destination -Force
            return
        }
        catch {
            if ($attempt -eq $RetryCount) {
                throw
            }

            Start-Sleep -Milliseconds $RetryDelayMs
        }
    }
}

function Test-SameResolvedPath {
    param(
        [string]$Left,
        [string]$Right
    )

    if ([string]::IsNullOrWhiteSpace($Left) -or [string]::IsNullOrWhiteSpace($Right)) {
        return $false
    }

    if (-not (Test-Path $Left) -or -not (Test-Path $Right)) {
        return $false
    }

    try {
        $leftResolved = (Resolve-Path -Path $Left).Path
        $rightResolved = (Resolve-Path -Path $Right).Path
        return [string]::Equals($leftResolved, $rightResolved, [System.StringComparison]::OrdinalIgnoreCase)
    }
    catch {
        return $false
    }
}

function New-StagedDllPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory
    )

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    return (Join-Path $Directory ("yuninput_" + $stamp + ".dll"))
}

Write-Host "Install root: $InstallRoot"
Write-Host "Install log: $LogPath"
Write-InstallLog 'Running in elevated context.'

try {
    if ($Build) {
        Write-InstallLog 'Building release binary.'
        & $buildScript
    }

    if (Test-Path $configBuildScript) {
        $shouldBuildConfig = $Build
        if (-not $shouldBuildConfig) {
            if (-not (Test-Path $srcConfigExe)) {
                $shouldBuildConfig = $true
            }
            elseif (Test-Path $srcConfigSource) {
                $configExeTime = (Get-Item $srcConfigExe).LastWriteTimeUtc
                $configSrcTime = (Get-Item $srcConfigSource).LastWriteTimeUtc
                if ($configSrcTime -gt $configExeTime) {
                    $shouldBuildConfig = $true
                }
            }
        }

        if ($shouldBuildConfig) {
            Write-InstallLog 'Building config app.'
            & $configBuildScript
        }
    }

    if (-not (Test-Path $srcDll) -and (Test-Path $fallbackSrcDll)) {
        $srcDll = $fallbackSrcDll
    }

    if (-not (Test-Path $srcConfigExe) -and (Test-Path $fallbackSrcConfigExe)) {
        $srcConfigExe = $fallbackSrcConfigExe
    }

    if (-not (Test-Path $srcDll)) {
        throw "Build output not found: $srcDll"
    }

    New-Item -ItemType Directory -Path $InstallRoot -Force | Out-Null
    New-Item -ItemType Directory -Path $binDir -Force | Out-Null
    New-Item -ItemType Directory -Path $dataDir -Force | Out-Null

    # Release previous COM registration and process locks before replacing DLL.
    if (Test-Path $dllPath) {
        try {
            Write-InstallLog "Unregistering previous DLL: $dllPath"
            & $unregisterScript -DllPath $dllPath
        }
        catch {
            Write-Warning "Unregister old DLL failed: $($_.Exception.Message)"
            Write-InstallLog "Warning: unregister failed: $($_.Exception.Message)"
        }
    }

    Stop-InputProcesses
    $registeredDllPath = $dllPath
    if (-not (Test-SameResolvedPath -Left $srcDll -Right $dllPath)) {
        try {
            Copy-WithRetry -Source $srcDll -Destination $dllPath
            Write-InstallLog "Copied DLL to: $dllPath"
        }
        catch {
            $registeredDllPath = New-StagedDllPath -Directory $binDir
            Write-Warning "Primary DLL is locked. Falling back to staged DLL: $registeredDllPath"
            Write-InstallLog "Primary DLL locked. Fallback DLL: $registeredDllPath"
            Copy-Item $srcDll $registeredDllPath -Force
        }
    }
    else {
        Write-InstallLog "Using in-place DLL payload: $dllPath"
    }

    if (Test-Path $srcDataDir) {
        Get-ChildItem -Path $srcDataDir -Filter '*.dict' -File | ForEach-Object {
            if ($_.Name -ieq 'yuninput_user.dict') {
                return
            }

            $destination = Join-Path $dataDir $_.Name
            if (-not (Test-SameResolvedPath -Left $_.FullName -Right $destination)) {
                Copy-Item $_.FullName $destination -Force
            }
        }
    }
    if (Test-Path $srcConfigExe) {
        $installedConfigExe = Join-Path $InstallRoot 'yuninput_config.exe'
        if (-not (Test-SameResolvedPath -Left $srcConfigExe -Right $installedConfigExe)) {
            Copy-Item $srcConfigExe $installedConfigExe -Force
        }
    }

    if (-not (Test-Path $settingsPath)) {
        New-Item -ItemType Directory -Path $settingsRoot -Force | Out-Null
        @"
{
  "chinese_mode": true,
  "full_shape": false,
    "chinese_punctuation": true,
    "smart_symbol_pairs": true,
    "auto_commit_unique_exact": true,
    "auto_commit_min_code_length": 4,
    "empty_candidate_beep": true,
    "tab_navigation": true,
    "enter_exact_priority": true,
    "context_association_enabled": true,
    "context_association_max_entries": 6000,
    "candidate_page_size": 6,
        "dictionary_profile": "zhengma-large",
  "toggle_hotkey": "F9"
}
"@ | Set-Content -Encoding utf8 $settingsPath
    }

    Write-InstallLog "Registering DLL: $registeredDllPath"
    try {
        if ($NonInteractive) {
            & $registerScript -DllPath $registeredDllPath -MachineOnly -LogPath (Join-Path $env:ProgramData 'Yuninput\register_machine.log')
        }
        else {
            & $registerScript -DllPath $registeredDllPath
        }
    }
    catch {
        $registerMessage = $_.Exception.Message
        if ($NonInteractive) {
            Write-InstallLog "Warning: DLL registration failed in NonInteractive mode: $registerMessage"
            Write-InstallLog 'Warning: continuing installation without registered IME. User may need manual register_ime.ps1.'
        }
        else {
            throw
        }
    }

    ${tipGuid} = '{6DE9AB40-3BA8-4B77-8D8F-233966E1C102}'
    try {
        $inproc = & reg.exe query "HKLM\SOFTWARE\Classes\CLSID\$tipGuid\InprocServer32" /ve 2>$null
        if ($LASTEXITCODE -eq 0 -and $null -ne $inproc) {
            Write-InstallLog "HKLM InprocServer32: $($inproc -join ' ')"
        }
        else {
            Write-InstallLog "Warning: HKLM InprocServer32 key missing after register (exit=$LASTEXITCODE)."
        }
    }
    catch {
        Write-InstallLog "Warning: HKLM InprocServer32 query failed: $($_.Exception.Message)"
    }

    if (-not $NonInteractive -and (Test-Path (Join-Path $InstallRoot 'yuninput_config.exe'))) {
        try {
            $shell = New-Object -ComObject WScript.Shell
            $shortcutPath = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Yuninput Config.lnk'
            $shortcut = $shell.CreateShortcut($shortcutPath)
            $shortcut.TargetPath = Join-Path $InstallRoot 'yuninput_config.exe'
            $shortcut.WorkingDirectory = $InstallRoot
            $shortcut.Save()
            Start-Process (Join-Path $InstallRoot 'yuninput_config.exe')
        }
        catch {
            Write-InstallLog "Warning: post-install config launch failed: $($_.Exception.Message)"
        }
    }

    Write-InstallLog 'Install complete.'

    if (-not $NonInteractive) {
        Write-Host ''
        Write-Host 'Install complete.'
        Write-Host "Install path: $InstallRoot"
        Write-Host 'Next: Windows Settings -> Language & region -> Chinese (Simplified) -> Language options -> Add keyboard -> yuninput'
        Write-Host ''
        Write-Host 'Opening relevant settings pages...'
        try {
            Start-Process 'ms-settings:regionlanguage'
            Start-Process 'ms-settings:typing'
            Start-Process 'ctfmon.exe'
        }
        catch {
            Write-InstallLog "Warning: opening settings pages failed: $($_.Exception.Message)"
        }
    }
}
catch {
    $message = $_.Exception.Message
    Write-InstallLog "Install failed: $message"
    Write-Error "Install failed: $message"
    Write-Host "Install log: $LogPath"
    if (Test-Path $LogPath) {
        Write-Host ''
        Write-Host 'Last install log lines:'
        Get-Content $LogPath -Tail 60
    }
    exit 1
}
