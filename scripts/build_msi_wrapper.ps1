param(
    [string]$Version = '1.3.3',
    [string]$OutputName = 'Yuninput1.3.3.msi',
    [switch]$SkipDictionaryGeneration
)

$ErrorActionPreference = "Stop"

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

$projectRoot = Split-Path -Parent $PSScriptRoot
$buildScript = Join-Path $PSScriptRoot 'build_release.ps1'
$configBuildScript = Join-Path $PSScriptRoot 'build_config_app.ps1'
$generateUserDictScript = Join-Path $PSScriptRoot 'generate_user_dict.ps1'
$wxsPath = Join-Path $projectRoot 'tools\msi\YuninputSetupWrapper.wxs'
$licenseRtf = Join-Path $projectRoot 'tools\msi\license.rtf'
$outMsi = Join-Path $projectRoot $OutputName

function Invoke-TimedStep {
    param(
        [string]$Name,
        [scriptblock]$Action
    )

    Write-Host ""
    Write-Host ("=== {0} ===" -f $Name)
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    & $Action
    $stopwatch.Stop()
    Write-Host ("[{0}] done in {1:n1}s" -f $Name, $stopwatch.Elapsed.TotalSeconds)
}

$requiredFiles = @(
    (Join-Path $projectRoot 'build\Release\yuninput.dll'),
    (Join-Path $projectRoot 'build\Release\yuninput_user_dict_builder.exe'),
    (Join-Path $projectRoot 'assets\icon_yun.ico'),
    (Join-Path $projectRoot 'YuninputManual.zh-CN.md'),
    (Join-Path $projectRoot 'scripts\install_enable.ps1'),
    (Join-Path $projectRoot 'scripts\register_ime.ps1'),
    (Join-Path $projectRoot 'scripts\unregister_ime.ps1'),
    (Join-Path $projectRoot 'data\yuninput_basic.dict'),
    (Join-Path $projectRoot 'data\yuninput_user.dict'),
    (Join-Path $projectRoot 'data\yuninput_user.dict.rules'),
    (Join-Path $projectRoot 'data\yuninput_user-extend.dict'),
    (Join-Path $projectRoot 'data\zhengma-all.dict'),
    (Join-Path $projectRoot 'data\zhengma-all.dict.rules'),
    (Join-Path $projectRoot 'data\zhengma-single.dict'),
    (Join-Path $projectRoot 'data\zhengma.dict'),
    (Join-Path $projectRoot 'data\zhengma.dict.rules'),
    (Join-Path $projectRoot 'data\zhengma-large.dict'),
    (Join-Path $projectRoot 'data\zhengma-large.dict.rules'),
    (Join-Path $projectRoot 'data\zhengma-pinyin.dict')
)

if (-not (Test-Path (Join-Path $projectRoot 'build\Release\yuninput.dll'))) {
    Invoke-TimedStep -Name 'Build Release DLL' -Action {
        & $buildScript
    }
}

if (-not (Test-Path (Join-Path $projectRoot 'tools\yuninput_config.exe'))) {
    Invoke-TimedStep -Name 'Build Config Tool' -Action {
        & $configBuildScript
    }
}

if ($SkipDictionaryGeneration) {
    Write-Host 'Skipping user dictionary generation and using existing dictionary payloads.'
}
else {
    Invoke-TimedStep -Name 'Generate User Dictionaries' -Action {
        & $generateUserDictScript
    }
}

foreach ($requiredFile in $requiredFiles) {
    if (-not (Test-Path $requiredFile)) {
        throw "Missing required MSI payload file: $requiredFile"
    }
}

if (-not (Test-Path $licenseRtf)) {
    throw "Missing MSI license file: $licenseRtf"
}

$wixCmd = Get-Command wix -ErrorAction SilentlyContinue
$wixPath = $null
if ($null -ne $wixCmd) {
    $wixPath = $wixCmd.Source
}

if ([string]::IsNullOrWhiteSpace($wixPath)) {
    $fallbackWix = @(
        "$env:ProgramFiles\WiX Toolset v6.0\bin\wix.exe",
        "$env:ProgramFiles\WiX Toolset v6\bin\wix.exe",
        "$env:LOCALAPPDATA\Microsoft\WinGet\Links\wix.exe"
    )
    foreach ($candidate in $fallbackWix) {
        if (Test-Path $candidate) {
            $wixPath = $candidate
            break
        }
    }
}

if ([string]::IsNullOrWhiteSpace($wixPath) -or -not (Test-Path $wixPath)) {
    throw "wix command not found. Install WiX CLI first (winget install WiXToolset.WiXCLI)."
}

foreach ($extension in @('WixToolset.UI.wixext', 'WixToolset.Util.wixext')) {
    & $wixPath extension add --global $extension | Out-Null
}

if (Test-Path $outMsi) {
    Remove-Item -Force $outMsi
}

Invoke-TimedStep -Name 'Build MSI With WiX' -Action {
    & $wixPath build $wxsPath -arch x64 -culture zh-CN -ext WixToolset.UI.wixext -ext WixToolset.Util.wixext -d ProductVersion="$Version" -d ProductSourceDir="$projectRoot" -bindvariable WixUILicenseRtf="$licenseRtf" -out $outMsi
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to build MSI"
    }
}

Get-Item $outMsi | Select-Object FullName,Length,LastWriteTime | Format-List
