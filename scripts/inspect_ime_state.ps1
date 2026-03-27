param(
    [switch]$RepairCurrentUserProfile,
    [switch]$RestartCtfmon,
    [string]$ReportPath = ""
)

$ErrorActionPreference = "Stop"

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

$tipGuid = '{6DE9AB40-3BA8-4B77-8D8F-233966E1C102}'
$profileGuid = '{47DE2FB1-F5E4-4CF8-AB2F-8F7A761731B2}'
$langHex = '0x00000804'
$defaultReportPath = Join-Path $env:LOCALAPPDATA 'yuninput\ime_state_report.txt'

if ([string]::IsNullOrWhiteSpace($ReportPath)) {
    $ReportPath = $defaultReportPath
}

function New-ParentDirectory {
    param([string]$Path)

    $parent = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
}

function Write-ReportLine {
    param([string]$Text)

    Write-Host $Text
    Add-Content -Path $ReportPath -Value $Text -Encoding utf8
}

function Write-Section {
    param([string]$Title)

    Write-ReportLine ""
    Write-ReportLine ("=== {0} ===" -f $Title)
}

function Get-RegQueryText {
    param([string]$KeyPath)

    $stdoutPath = [System.IO.Path]::GetTempFileName()
    $stderrPath = [System.IO.Path]::GetTempFileName()
    try {
        $argumentString = ('query "{0}" /s' -f $KeyPath)
        $proc = Start-Process -FilePath 'reg.exe' -ArgumentList $argumentString -Wait -PassThru -NoNewWindow -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
        $exitCode = if ($null -eq $proc) { 1 } else { [int]$proc.ExitCode }
        if ($exitCode -ne 0) {
            return @("MISSING: $KeyPath")
        }

        $raw = Get-Content -Path $stdoutPath -ErrorAction SilentlyContinue
        if ($raw -is [System.Array]) {
            return $raw
        }

        return @([string]$raw)
    }
    finally {
        if (Test-Path $stdoutPath) {
            Remove-Item $stdoutPath -Force -ErrorAction SilentlyContinue
        }
        if (Test-Path $stderrPath) {
            Remove-Item $stderrPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-RegistryValue {
    param(
        [string]$Path,
        [string]$Name,
        $Default = $null
    )

    try {
        $item = Get-ItemProperty -Path $Path -ErrorAction Stop
        if ($Name -eq '(Default)') {
            return $item.'(default)'
        }
        return $item.$Name
    }
    catch {
        return $Default
    }
}

function Format-ObjectLines {
    param($InputObject)

    if ($null -eq $InputObject) {
        return @('NONE')
    }

    $text = $InputObject | Format-List * | Out-String
    return ($text -split "`r?`n") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
}

function Repair-CurrentUserProfile {
    $hkcuProfilePath = "Registry::HKEY_CURRENT_USER\Software\Microsoft\CTF\TIP\$tipGuid\LanguageProfile\$langHex\$profileGuid"
    $hklmProfilePath = "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\CTF\TIP\$tipGuid\LanguageProfile\$langHex\$profileGuid"
    $hkcuClsidPath = "Registry::HKEY_CURRENT_USER\Software\Classes\CLSID\$tipGuid\InprocServer32"
    $hklmClsidPath = "Registry::HKEY_LOCAL_MACHINE\Software\Classes\CLSID\$tipGuid\InprocServer32"

    if (-not (Test-Path $hklmProfilePath)) {
        Write-ReportLine 'Repair skipped: HKLM profile is missing.'
        return $false
    }

    $description = Get-RegistryValue -Path $hklmProfilePath -Name 'Description' -Default ''
    $iconFile = Get-RegistryValue -Path $hklmProfilePath -Name 'IconFile' -Default ''
    $iconIndex = Get-RegistryValue -Path $hklmProfilePath -Name 'IconIndex' -Default 0
    $enable = Get-RegistryValue -Path $hklmProfilePath -Name 'Enable' -Default 1
    $inproc = Get-RegistryValue -Path $hklmClsidPath -Name '(Default)' -Default ''

    if ([string]::IsNullOrWhiteSpace($description)) {
        Write-ReportLine 'Repair skipped: HKLM Description is empty.'
        return $false
    }

    New-Item -Path $hkcuProfilePath -Force | Out-Null
    New-ItemProperty -Path $hkcuProfilePath -Name 'Enable' -PropertyType DWord -Value ([int]$enable) -Force | Out-Null
    New-ItemProperty -Path $hkcuProfilePath -Name 'Description' -PropertyType String -Value ([string]$description) -Force | Out-Null
    if (-not [string]::IsNullOrWhiteSpace($iconFile)) {
        New-ItemProperty -Path $hkcuProfilePath -Name 'IconFile' -PropertyType String -Value ([string]$iconFile) -Force | Out-Null
    }
    New-ItemProperty -Path $hkcuProfilePath -Name 'IconIndex' -PropertyType DWord -Value ([int]$iconIndex) -Force | Out-Null

    if (-not [string]::IsNullOrWhiteSpace($inproc)) {
        New-Item -Path $hkcuClsidPath -Force | Out-Null
        New-ItemProperty -Path $hkcuClsidPath -Name '(Default)' -Value ([string]$inproc) -Force | Out-Null
        New-ItemProperty -Path $hkcuClsidPath -Name 'ThreadingModel' -PropertyType String -Value 'Apartment' -Force | Out-Null
    }

    Write-ReportLine 'Repair completed: mirrored HKLM profile metadata to HKCU.'
    return $true
}

New-ParentDirectory -Path $ReportPath
Set-Content -Path $ReportPath -Value @(
    ('Timestamp: {0}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')),
    ('User: {0}' -f $env:USERNAME),
    ('Computer: {0}' -f $env:COMPUTERNAME),
    ('RepairCurrentUserProfile: {0}' -f $RepairCurrentUserProfile),
    ('RestartCtfmon: {0}' -f $RestartCtfmon)
) -Encoding utf8

Write-Section 'IME Identity'
Write-ReportLine ("TIP GUID: {0}" -f $tipGuid)
Write-ReportLine ("Profile GUID: {0}" -f $profileGuid)
Write-ReportLine ("Lang ID: {0}" -f $langHex)

if ($RepairCurrentUserProfile) {
    Write-Section 'Repair'
    $repairApplied = Repair-CurrentUserProfile
    if ($repairApplied -and $RestartCtfmon) {
        Get-Process -Name ctfmon -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Process ctfmon.exe | Out-Null
        Write-ReportLine 'ctfmon.exe restarted.'
    }
}

Write-Section 'Canonical Profile Values'
$hkcuProfilePath = "Registry::HKEY_CURRENT_USER\Software\Microsoft\CTF\TIP\$tipGuid\LanguageProfile\$langHex\$profileGuid"
$hklmProfilePath = "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\CTF\TIP\$tipGuid\LanguageProfile\$langHex\$profileGuid"
$hkcuClsidPath = "Registry::HKEY_CURRENT_USER\Software\Classes\CLSID\$tipGuid\InprocServer32"
$hklmClsidPath = "Registry::HKEY_LOCAL_MACHINE\Software\Classes\CLSID\$tipGuid\InprocServer32"

$canonicalRows = @(
    @{ Label = 'HKCU Description'; Value = (Get-RegistryValue -Path $hkcuProfilePath -Name 'Description' -Default '<missing>') },
    @{ Label = 'HKLM Description'; Value = (Get-RegistryValue -Path $hklmProfilePath -Name 'Description' -Default '<missing>') },
    @{ Label = 'HKCU IconFile'; Value = (Get-RegistryValue -Path $hkcuProfilePath -Name 'IconFile' -Default '<missing>') },
    @{ Label = 'HKLM IconFile'; Value = (Get-RegistryValue -Path $hklmProfilePath -Name 'IconFile' -Default '<missing>') },
    @{ Label = 'HKCU IconIndex'; Value = (Get-RegistryValue -Path $hkcuProfilePath -Name 'IconIndex' -Default '<missing>') },
    @{ Label = 'HKLM IconIndex'; Value = (Get-RegistryValue -Path $hklmProfilePath -Name 'IconIndex' -Default '<missing>') },
    @{ Label = 'HKCU Enable'; Value = (Get-RegistryValue -Path $hkcuProfilePath -Name 'Enable' -Default '<missing>') },
    @{ Label = 'HKLM Enable'; Value = (Get-RegistryValue -Path $hklmProfilePath -Name 'Enable' -Default '<missing>') },
    @{ Label = 'HKCU InprocServer32'; Value = (Get-RegistryValue -Path $hkcuClsidPath -Name '(Default)' -Default '<missing>') },
    @{ Label = 'HKLM InprocServer32'; Value = (Get-RegistryValue -Path $hklmClsidPath -Name '(Default)' -Default '<missing>') },
    @{ Label = 'HKCU ThreadingModel'; Value = (Get-RegistryValue -Path $hkcuClsidPath -Name 'ThreadingModel' -Default '<missing>') },
    @{ Label = 'HKLM ThreadingModel'; Value = (Get-RegistryValue -Path $hklmClsidPath -Name 'ThreadingModel' -Default '<missing>') }
)

foreach ($row in $canonicalRows) {
    Write-ReportLine (("{0}: {1}" -f $row.Label, $row.Value))
}

$registryTargets = @(
    @{ Title = 'HKCU TIP Profile'; Path = "HKCU\Software\Microsoft\CTF\TIP\$tipGuid\LanguageProfile\$langHex\$profileGuid" },
    @{ Title = 'HKLM TIP Profile'; Path = "HKLM\Software\Microsoft\CTF\TIP\$tipGuid\LanguageProfile\$langHex\$profileGuid" },
    @{ Title = 'HKCU CLSID'; Path = "HKCU\Software\Classes\CLSID\$tipGuid" },
    @{ Title = 'HKLM CLSID'; Path = "HKLM\Software\Classes\CLSID\$tipGuid" },
    @{ Title = 'CTF LangBar'; Path = 'HKCU\Software\Microsoft\CTF\LangBar' },
    @{ Title = 'Keyboard Preload'; Path = 'HKCU\Keyboard Layout\Preload' },
    @{ Title = 'Keyboard Substitutes'; Path = 'HKCU\Keyboard Layout\Substitutes' }
)

foreach ($target in $registryTargets) {
    Write-Section $target.Title
    foreach ($line in Get-RegQueryText -KeyPath $target.Path) {
        Write-ReportLine ([string]$line)
    }
}

Write-Section 'WinUserLanguageList'
try {
    $languageList = Get-WinUserLanguageList
    foreach ($language in $languageList) {
        foreach ($line in (Format-ObjectLines -InputObject $language)) {
            Write-ReportLine $line
        }
        Write-ReportLine '---'
    }
}
catch {
    Write-ReportLine ("FAILED: Get-WinUserLanguageList - {0}" -f $_.Exception.Message)
}

Write-Section 'Local Files'
$localPaths = @(
    (Join-Path $env:LOCALAPPDATA 'yuninput\bin\yuninput.dll'),
    (Join-Path $env:LOCALAPPDATA 'yuninput\runtime.log'),
    (Join-Path $env:LOCALAPPDATA 'yuninput\register.log'),
    (Join-Path $env:LOCALAPPDATA 'yuninput\register_script.log'),
    (Join-Path $env:LOCALAPPDATA 'yuninput\install.log')
)

foreach ($path in $localPaths) {
    $exists = Test-Path $path
    Write-ReportLine (("{0} :: exists={1}" -f $path, $exists))
}

Write-Section 'Summary'
$hkcuProfileExists = Test-Path "Registry::HKEY_CURRENT_USER\Software\Microsoft\CTF\TIP\$tipGuid\LanguageProfile\$langHex\$profileGuid"
$hklmProfileExists = Test-Path "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\CTF\TIP\$tipGuid\LanguageProfile\$langHex\$profileGuid"
Write-ReportLine (("HKCU profile exists: {0}" -f $hkcuProfileExists))
Write-ReportLine (("HKLM profile exists: {0}" -f $hklmProfileExists))
Write-ReportLine (("Report written to: {0}" -f $ReportPath))