$ErrorActionPreference = 'Stop'

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

$projectRoot = Split-Path -Parent $PSScriptRoot
$builderExe = Join-Path $projectRoot 'build\Release\yuninput_user_dict_builder.exe'
$userDict = Join-Path $projectRoot 'data\yuninput_user.dict'
$extendDict = Join-Path $projectRoot 'data\yuninput_user-extend.dict'
$allDict = Join-Path $projectRoot 'data\zhengma-all.dict'
$singleDict = Join-Path $projectRoot 'data\zhengma-single.dict'
$sourceDicts = @(
    (Join-Path $projectRoot 'data\zhengma-large.dict'),
    (Join-Path $projectRoot 'data\zhengma.dict')
)

function Get-StagedLocalExecutablePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath
    )

    $item = Get-Item $SourcePath
    $stageRoot = Join-Path $env:LOCALAPPDATA 'yuninput\staging'
    New-Item -ItemType Directory -Path $stageRoot -Force | Out-Null

    $stamp = '{0:x}-{1:x}' -f $item.Length, $item.LastWriteTimeUtc.Ticks
    $stagedPath = Join-Path $stageRoot ("yuninput_user_dict_builder_{0}.exe" -f $stamp)
    if (-not (Test-Path $stagedPath)) {
        Copy-Item $SourcePath $stagedPath -Force
        Unblock-File -Path $stagedPath -ErrorAction SilentlyContinue
    }

    return $stagedPath
}

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

if (-not (Test-Path $builderExe)) {
    throw "Missing builder executable: $builderExe"
}

$builderExe = Get-StagedLocalExecutablePath -SourcePath $builderExe

foreach ($path in @($userDict) + $sourceDicts) {
    if (-not (Test-Path $path)) {
        throw "Missing dictionary file: $path"
    }
}

Invoke-TimedStep -Name 'Merge Base Dictionaries' -Action {
    & $builderExe merge $allDict $singleDict $sourceDicts
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to generate merged dictionaries, exit code: $LASTEXITCODE"
    }
}

Invoke-TimedStep -Name 'Extend User Dictionary Variants' -Action {
    & $builderExe extend $extendDict $userDict $singleDict $allDict $userDict
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to generate extend dictionary, exit code: $LASTEXITCODE"
    }
}