param(
    [string]$DllPath,
    [switch]$SkipElevation,
    [switch]$SkipComRegistration,
    [switch]$CurrentUserOnly,
    [switch]$MachineOnly,
    [string]$LogPath
)

$ErrorActionPreference = "Stop"

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $logRoot = Join-Path $env:LOCALAPPDATA 'yuninput'
    New-Item -Path $logRoot -ItemType Directory -Force | Out-Null
    $LogPath = Join-Path $logRoot 'register_script.log'
}
else {
    $logDir = Split-Path -Parent $LogPath
    if (-not [string]::IsNullOrWhiteSpace($logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }
}

function Write-RegisterLog {
    param([string]$Message)
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content -Path $LogPath -Encoding UTF8 -Value "[$ts] $Message"
}

Write-RegisterLog 'register_ime.ps1 start'
Write-RegisterLog "mode skipComRegistration=$SkipComRegistration currentUserOnly=$CurrentUserOnly machineOnly=$MachineOnly"

if ($CurrentUserOnly -and $MachineOnly) {
    throw 'CurrentUserOnly and MachineOnly cannot both be specified.'
}

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

$dllDir = Split-Path -Parent $DllPath
$iconFilePath = Join-Path $dllDir 'icon_yun.ico'
$profileIconPath = $DllPath
if (Test-Path -Path $iconFilePath -PathType Leaf) {
    $profileIconPath = (Resolve-Path -Path $iconFilePath).Path
}
Write-RegisterLog "resolved profile icon path: $profileIconPath"

${tipGuid} = '{6DE9AB40-3BA8-4B77-8D8F-233966E1C102}'
${profileGuid} = '{47DE2FB1-F5E4-4CF8-AB2F-8F7A761731B2}'
${langHex} = '0x00000804'
$keyboardCategoryGuid = '{34745C63-B2F0-4784-8B67-5E12C8701A31}'
$tipToken = '0804:' + $tipGuid + $profileGuid
$textServiceDescription = 'yuninput Text Service'
$profileDescription = [string]([char]0x5300) + [char]0x7801 + [char]0x8F93 + [char]0x5165 + [char]0x6CD5

$hklmTipRootPath = "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\CTF\TIP\$tipGuid"
$hkcuProfilePath = "Registry::HKEY_CURRENT_USER\Software\Microsoft\CTF\TIP\$tipGuid\LanguageProfile\$langHex\$profileGuid"
$hklmProfilePath = "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\CTF\TIP\$tipGuid\LanguageProfile\$langHex\$profileGuid"
$hklmClsidRootPath = "Registry::HKEY_LOCAL_MACHINE\Software\Classes\CLSID\$tipGuid"
$hklmClsidPath = "$hklmClsidRootPath\InprocServer32"
$hkcuClsidPath = "Registry::HKEY_CURRENT_USER\Software\Classes\CLSID\$tipGuid\InprocServer32"
$hklmCategoryByCategoryPath = "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\CTF\TIP\$tipGuid\Category\Category\$keyboardCategoryGuid\$tipGuid"
$hklmCategoryByItemPath = "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\CTF\TIP\$tipGuid\Category\Item\$tipGuid\$keyboardCategoryGuid"

if (-not $SkipComRegistration) {
    Write-Host "Registering IME DLL: $DllPath"
    $regsvr32 = Join-Path $env:SystemRoot 'System32\regsvr32.exe'
    $process = Start-Process -FilePath $regsvr32 -ArgumentList '/s', $DllPath -Wait -PassThru
    $code = [int]$process.ExitCode
    Write-RegisterLog "regsvr32 exit code: $code"
    if ($code -ne 0) {
        Write-RegisterLog 'Warning: regsvr32 failed; continuing with registry fallback registration.'
    }
}
else {
    Write-RegisterLog 'Skipping COM registration by request.'
}

function Set-RegistryStringValue {
    param(
        [string]$Path,
        [string]$Name,
        [string]$Value
    )

    New-Item -Path $Path -Force | Out-Null
    if ($Name -eq '(default)') {
        New-ItemProperty -Path $Path -Name '(default)' -PropertyType String -Value $Value -Force | Out-Null
        return
    }

    New-ItemProperty -Path $Path -Name $Name -PropertyType String -Value $Value -Force | Out-Null
}

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

function Ensure-MachineRegistrationFallback {
    param([string]$DllPath)

    Set-RegistryStringValue -Path $hklmClsidRootPath -Name '(default)' -Value $textServiceDescription
    Set-RegistryStringValue -Path $hklmClsidPath -Name '(default)' -Value $DllPath
    Set-RegistryStringValue -Path $hklmClsidPath -Name 'ThreadingModel' -Value 'Apartment'

    New-Item -Path $hklmTipRootPath -Force | Out-Null
    New-ItemProperty -Path $hklmTipRootPath -Name Enable -PropertyType String -Value '1' -Force | Out-Null
    New-Item -Path $hklmCategoryByCategoryPath -Force | Out-Null
    New-Item -Path $hklmCategoryByItemPath -Force | Out-Null
}

function Write-OptionalRegistryDump {
    param(
        [string]$Path,
        [string]$Label,
        [string[]]$PropertyNames
    )

    if (-not (Test-Path -Path $Path)) {
        Write-RegisterLog "Warning: $Label missing at $Path"
        return
    }

    $dump = Get-ItemProperty -Path $Path
    $parts = @()
    foreach ($propertyName in $PropertyNames) {
        if ($propertyName -eq '(default)') {
            $parts += "InprocServer32=$($dump.'(default)')"
        }
        else {
            $parts += "$propertyName=$($dump.$propertyName)"
        }
    }

    Write-RegisterLog ($Label + ': ' + ($parts -join ' '))
}

function Refresh-CurrentUserInputStack {
    try {
        Start-Process -FilePath 'ctfmon.exe' -WindowStyle Hidden | Out-Null
        Write-RegisterLog 'Started ctfmon.exe to refresh current user input stack.'
    }
    catch {
        Write-RegisterLog "Warning: failed to start ctfmon.exe: $($_.Exception.Message)"
    }
}

function Update-CurrentUserLanguageList {
    param([string]$TipToken)

    try {
        $languageList = Get-WinUserLanguageList
    }
    catch {
        Write-RegisterLog "Warning: Get-WinUserLanguageList failed: $($_.Exception.Message)"
        return $false
    }

    $targetLanguage = $null
    foreach ($language in $languageList) {
        if ($language.LanguageTag -ieq 'zh-Hans-CN' -or $language.LanguageTag -ieq 'zh-CN') {
            $targetLanguage = $language
            break
        }
    }

    if ($null -eq $targetLanguage) {
        foreach ($language in $languageList) {
            if ($language.LanguageTag -like 'zh*') {
                $targetLanguage = $language
                break
            }
        }
    }

    if ($null -eq $targetLanguage) {
        try {
            $newLanguageList = New-WinUserLanguageList 'zh-Hans-CN'
            foreach ($language in $languageList) {
                [void]$newLanguageList.Add($language)
            }
            $languageList = $newLanguageList
            $targetLanguage = $languageList[0]
            Write-RegisterLog 'Added zh-Hans-CN to current user language list.'
        }
        catch {
            Write-RegisterLog "Warning: New-WinUserLanguageList failed: $($_.Exception.Message)"
            return $false
        }
    }

    if ($null -eq $targetLanguage.InputMethodTips) {
        Write-RegisterLog 'Warning: target language has null InputMethodTips.'
        return $false
    }

    $existingTips = @()
    foreach ($tip in $targetLanguage.InputMethodTips) {
        $tipText = [string]$tip
        if ([string]::Equals($tipText, $TipToken, [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }
        $existingTips += $tipText
    }

    if ($existingTips.Count -ne $targetLanguage.InputMethodTips.Count) {
        Write-RegisterLog "Current user language list contains TIP already; forcing re-add to refresh cache: $TipToken"
        $targetLanguage.InputMethodTips.Clear()
        foreach ($tip in $existingTips) {
            [void]$targetLanguage.InputMethodTips.Add($tip)
        }
    }

    [void]$targetLanguage.InputMethodTips.Add($TipToken)
    try {
        Set-WinUserLanguageList -LanguageList $languageList -Force
        Write-RegisterLog "Applied TIP to current user language list: $TipToken"
        return $true
    }
    catch {
        Write-RegisterLog "Warning: Set-WinUserLanguageList failed: $($_.Exception.Message)"
        return $false
    }
}

if (-not $CurrentUserOnly) {
    Ensure-MachineRegistrationFallback -DllPath $DllPath
    Write-RegisterLog 'Machine registration fallback applied.'
}

if (-not $CurrentUserOnly) {
    Set-ProfileValues -Path $hklmProfilePath -Description $profileDescription -IconFile $profileIconPath
}

if (-not $MachineOnly) {
    Set-ProfileValues -Path $hkcuProfilePath -Description $profileDescription -IconFile $profileIconPath
}

# Keep per-user CLSID override aligned with the active DLL so TSF loads the latest staged build.
if (-not $MachineOnly) {
    New-Item -Path $hkcuClsidPath -Force | Out-Null
    New-ItemProperty -Path $hkcuClsidPath -Name '(default)' -PropertyType String -Value $DllPath -Force | Out-Null
    New-ItemProperty -Path $hkcuClsidPath -Name 'ThreadingModel' -PropertyType String -Value 'Apartment' -Force | Out-Null
    $languageListUpdated = Update-CurrentUserLanguageList -TipToken $tipToken
    Write-RegisterLog "Current user language list updated=$languageListUpdated"
    Refresh-CurrentUserInputStack
}

if (-not $MachineOnly) {
    Write-OptionalRegistryDump -Path $hkcuProfilePath -Label 'HKCU profile after register' -PropertyNames @('Description', 'IconFile', 'IconIndex', 'Enable')
    Write-OptionalRegistryDump -Path $hkcuClsidPath -Label 'HKCU CLSID after register' -PropertyNames @('(default)', 'ThreadingModel')
}
if (-not $CurrentUserOnly) {
    Write-OptionalRegistryDump -Path $hklmProfilePath -Label 'HKLM profile after register' -PropertyNames @('Description', 'IconFile', 'IconIndex', 'Enable')
    Write-OptionalRegistryDump -Path $hklmClsidPath -Label 'HKLM CLSID after register' -PropertyNames @('(default)', 'ThreadingModel')
}

Write-RegisterLog 'register_ime.ps1 success'

Write-Host "Registration done. Open Windows language settings and enable yuninput keyboard."
