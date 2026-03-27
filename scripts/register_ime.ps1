param(
    [string]$DllPath,
    [switch]$SkipElevation
)

$ErrorActionPreference = "Stop"

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

$logRoot = Join-Path $env:LOCALAPPDATA 'yuninput'
New-Item -Path $logRoot -ItemType Directory -Force | Out-Null
$logPath = Join-Path $logRoot 'register_script.log'

function Write-RegisterLog {
    param([string]$Message)
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content -Path $logPath -Encoding UTF8 -Value "[$ts] $Message"
}

Write-RegisterLog 'register_ime.ps1 start'

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]::new($identity)
$isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin -and -not $SkipElevation) {
    $quotedScript = '"' + $PSCommandPath + '"'
    $argumentString = "-NoProfile -ExecutionPolicy Bypass -File $quotedScript -SkipElevation"
    if (-not [string]::IsNullOrWhiteSpace($DllPath)) {
        $argumentString += " -DllPath `"$DllPath`""
    }

    Write-RegisterLog "elevation args: $argumentString"
    Write-Host 'Requesting administrator permission...'
    try {
        $proc = Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList $argumentString -Wait -PassThru
        Write-RegisterLog "elevation child exit code: $($proc.ExitCode)"
        exit $proc.ExitCode
    }
    catch {
        Write-RegisterLog "elevation failed: $($_.Exception.Message)"
        Write-Error "Elevation failed or was canceled: $($_.Exception.Message)"
        exit 1
    }
}

Write-RegisterLog "running elevated=$isAdmin skipElevation=$SkipElevation"

if ([string]::IsNullOrWhiteSpace($DllPath)) {
    $projectRoot = Split-Path -Parent $PSScriptRoot
    $DllPath = Join-Path $projectRoot "build\Release\yuninput.dll"
}

if (-not (Test-Path -Path $DllPath -PathType Leaf)) {
    throw "DLL not found: $DllPath"
}

$DllPath = (Resolve-Path -Path $DllPath).Path
Write-RegisterLog "resolved dll path: $DllPath"

Write-Host "Registering IME DLL: $DllPath"
$regsvr32 = Join-Path $env:SystemRoot 'System32\regsvr32.exe'
$process = Start-Process -FilePath $regsvr32 -ArgumentList '/s', $DllPath -Wait -PassThru
$code = [int]$process.ExitCode
Write-RegisterLog "regsvr32 exit code: $code"
if ($code -ne 0) {
    throw "regsvr32 registration failed with exit code: $code for DLL '$DllPath'. Confirm the installer is running as administrator."
}

${tipGuid} = '{6DE9AB40-3BA8-4B77-8D8F-233966E1C102}'
${profileGuid} = '{47DE2FB1-F5E4-4CF8-AB2F-8F7A761731B2}'
${langHex} = '0x00000804'
$profileDescription = [string]([char]0x5300) + [char]0x7801 + [char]0x8F93 + [char]0x5165 + [char]0x6CD5

$hkcuProfileKey = "HKCU\Software\Microsoft\CTF\TIP\$tipGuid\LanguageProfile\$langHex\$profileGuid"
$hklmProfileKey = "HKLM\Software\Microsoft\CTF\TIP\$tipGuid\LanguageProfile\$langHex\$profileGuid"
$hkcuProfilePath = "Registry::HKEY_CURRENT_USER\Software\Microsoft\CTF\TIP\$tipGuid\LanguageProfile\$langHex\$profileGuid"
$hklmProfilePath = "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\CTF\TIP\$tipGuid\LanguageProfile\$langHex\$profileGuid"
$hkcuClsidPath = "Registry::HKEY_CURRENT_USER\Software\Classes\CLSID\$tipGuid\InprocServer32"

function Set-ProfileValues {
    param(
        [string]$Path,
        [string]$Description,
        [string]$IconFile
    )

    New-Item -Path $Path -Force | Out-Null
    New-ItemProperty -Path $Path -Name Enable -PropertyType DWord -Value 1 -Force | Out-Null
    New-ItemProperty -Path $Path -Name Description -PropertyType String -Value $Description -Force | Out-Null
    New-ItemProperty -Path $Path -Name IconFile -PropertyType String -Value $IconFile -Force | Out-Null
    New-ItemProperty -Path $Path -Name IconIndex -PropertyType DWord -Value 0 -Force | Out-Null
}

Set-ProfileValues -Path $hkcuProfilePath -Description $profileDescription -IconFile $DllPath
Set-ProfileValues -Path $hklmProfilePath -Description $profileDescription -IconFile $DllPath

# Keep per-user CLSID override aligned with the active DLL so TSF loads the latest staged build.
New-Item -Path $hkcuClsidPath -Force | Out-Null
New-ItemProperty -Path $hkcuClsidPath -Name '(default)' -PropertyType String -Value $DllPath -Force | Out-Null
New-ItemProperty -Path $hkcuClsidPath -Name 'ThreadingModel' -PropertyType String -Value 'Apartment' -Force | Out-Null

$hkcuDump = Get-ItemProperty -Path $hkcuProfilePath
$hklmDump = Get-ItemProperty -Path $hklmProfilePath
$hkcuClsidDump = Get-ItemProperty -Path $hkcuClsidPath
Write-RegisterLog "HKCU profile after register: Description=$($hkcuDump.Description) IconFile=$($hkcuDump.IconFile) IconIndex=$($hkcuDump.IconIndex) Enable=$($hkcuDump.Enable)"
Write-RegisterLog "HKLM profile after register: Description=$($hklmDump.Description) IconFile=$($hklmDump.IconFile) IconIndex=$($hklmDump.IconIndex) Enable=$($hklmDump.Enable)"
Write-RegisterLog "HKCU CLSID after register: InprocServer32=$($hkcuClsidDump.'(default)') ThreadingModel=$($hkcuClsidDump.ThreadingModel)"

Write-RegisterLog 'register_ime.ps1 success'

Write-Host "Registration done. Open Windows language settings and enable yuninput keyboard."
