param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$DictRelativePath = 'data/zhengma-large.dict',
    [int]$TopN = 12,
    [switch]$UpdateBaseline,
    [string]$BaselineRelativePath = 'diagnostics/sort_regression_baseline.tsv',
    [string]$LatestRelativePath = 'diagnostics/sort_regression_latest.tsv',
    [string[]]$Codes = @(
        'a','aa','ab','b','ba','c',
        'j','jj','jjj','jjjd','jm',
        'ni','nih','nihk',
        'z','zh','zhi'
    ),
    [switch]$AllowProbeErrors
)

$ErrorActionPreference = 'Stop'

# Force a stable UTF-8 console code page so native probe output is decoded consistently.
chcp 65001 | Out-Null
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

function Resolve-RepoPath([string]$basePath, [string]$relativePath) {
    return [System.IO.Path]::GetFullPath((Join-Path $basePath $relativePath))
}

function Test-IsRemotePath([string]$path) {
    if ([string]::IsNullOrWhiteSpace($path)) {
        return $false
    }

    if ($path.StartsWith('\\')) {
        return $true
    }

    try {
        $root = [System.IO.Path]::GetPathRoot($path)
        if ([string]::IsNullOrWhiteSpace($root) -or $root.Length -lt 2) {
            return $false
        }

        $drive = $root.Substring(0, 1) + ':\\'
        $driveType = [System.IO.DriveInfo]::new($drive).DriveType
        return $driveType -eq [System.IO.DriveType]::Network
    }
    catch {
        return $false
    }
}

function Copy-ToLocalStage([string]$sourcePath, [string]$stageRoot) {
    if ([string]::IsNullOrWhiteSpace($sourcePath) -or -not (Test-Path $sourcePath)) {
        return $sourcePath
    }

    if (-not (Test-IsRemotePath $sourcePath)) {
        return $sourcePath
    }

    $targetPath = Join-Path $stageRoot ([System.IO.Path]::GetFileName($sourcePath))
    if (-not (Test-Path $stageRoot)) {
        New-Item -ItemType Directory -Path $stageRoot -Force | Out-Null
    }

    Copy-Item -Path $sourcePath -Destination $targetPath -Force
    Unblock-File -Path $targetPath -ErrorAction SilentlyContinue
    return $targetPath
}

function Invoke-SortProbe([string]$probePath, [string]$dictPath, [string]$code, [int]$topN) {
    $rawOutput = @(& $probePath $dictPath $code $topN 2>&1)
    $exitCode = $LASTEXITCODE
    $lines = @($rawOutput | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.ToString().Trim() })

    return [PSCustomObject]@{
        ExitCode = $exitCode
        Lines = $lines
    }
}

$repo = [System.IO.Path]::GetFullPath($RepoRoot)
$probePath = Resolve-RepoPath $repo 'build/Release/yuninput_sort_probe.exe'
$dictPath = Resolve-RepoPath $repo $DictRelativePath
$baselinePath = Resolve-RepoPath $repo $BaselineRelativePath
$latestPath = Resolve-RepoPath $repo $LatestRelativePath

if (-not (Test-Path $probePath)) {
    throw "Probe not found: $probePath"
}
if (-not (Test-Path $dictPath)) {
    throw "Dictionary not found: $dictPath"
}

$stageRoot = Join-Path ([System.IO.Path]::GetTempPath()) 'yuninput_probe_stage'
$probePath = Copy-ToLocalStage $probePath (Join-Path $stageRoot 'bin')
$dictPath = Copy-ToLocalStage $dictPath (Join-Path $stageRoot 'data')

$baselineDir = Split-Path $baselinePath -Parent
$latestDir = Split-Path $latestPath -Parent
if (-not (Test-Path $baselineDir)) { New-Item -ItemType Directory -Path $baselineDir | Out-Null }
if (-not (Test-Path $latestDir)) { New-Item -ItemType Directory -Path $latestDir | Out-Null }

$rows = New-Object System.Collections.Generic.List[string]
$rows.Add('input_code`trank`tentry_code`ttext')
$probeErrors = New-Object System.Collections.Generic.List[object]

foreach ($code in $Codes) {
    $probeResult = Invoke-SortProbe $probePath $dictPath $code $TopN
    if ($probeResult.ExitCode -ne 0) {
        $probeErrors.Add([PSCustomObject]@{
            Code = $code
            ExitCode = $probeResult.ExitCode
            Output = ($probeResult.Lines -join ' | ')
        }) | Out-Null
        continue
    }

    $probeLines = $probeResult.Lines
    if ($null -eq $probeLines -or $probeLines.Count -eq 0) {
        $rows.Add("$code`t0`t<none>`t<none>")
        continue
    }

    $rank = 0
    foreach ($line in $probeLines) {
        $rank += 1
        $parts = $line -split "`t", 2
        $entryCode = ''
        $entryText = ''
        if ($parts.Count -ge 1) { $entryCode = $parts[0] }
        if ($parts.Count -ge 2) { $entryText = $parts[1] }

        if ([string]::IsNullOrWhiteSpace($entryCode)) { $entryCode = '<empty>' }
        if ([string]::IsNullOrWhiteSpace($entryText)) { $entryText = '<empty>' }
        $rows.Add("$code`t$rank`t$entryCode`t$entryText")
    }
}

Set-Content -Path $latestPath -Value $rows -Encoding UTF8
Write-Host "Latest snapshot written: $latestPath"
Write-Host ("Probe errors: {0}" -f $probeErrors.Count)

if ($probeErrors.Count -gt 0 -and -not $AllowProbeErrors) {
    Write-Host ("Sort regression aborted: {0} probe execution errors." -f $probeErrors.Count)
    $previewProbeErrors = $probeErrors | Select-Object -First 30
    foreach ($probeError in $previewProbeErrors) {
        Write-Host ("PROBE-ERROR code={0} exit={1} output={2}" -f $probeError.Code, $probeError.ExitCode, $probeError.Output)
    }
    if ($probeErrors.Count -gt $previewProbeErrors.Count) {
        Write-Host ("... truncated, remaining probe errors: {0}" -f ($probeErrors.Count - $previewProbeErrors.Count))
    }
    exit 3
}

if ($UpdateBaseline) {
    Set-Content -Path $baselinePath -Value $rows -Encoding UTF8
    Write-Host "Baseline updated: $baselinePath"
    exit 0
}

if (-not (Test-Path $baselinePath)) {
    Write-Host "Baseline missing: $baselinePath"
    Write-Host "Run with -UpdateBaseline to create baseline."
    exit 2
}

$baselineRows = Get-Content -Path $baselinePath -Encoding UTF8
$diffs = Compare-Object -ReferenceObject $baselineRows -DifferenceObject $rows -SyncWindow 0

if ($null -eq $diffs -or $diffs.Count -eq 0) {
    Write-Host 'Sort regression check passed: no differences against baseline.'
    exit 0
}

Write-Host ("Sort regression check failed: {0} differing lines." -f $diffs.Count)
$preview = $diffs | Select-Object -First 30
foreach ($d in $preview) {
    Write-Host ("{0} {1}" -f $d.SideIndicator, $d.InputObject)
}
if ($diffs.Count -gt $preview.Count) {
    Write-Host ("... truncated, remaining diff lines: {0}" -f ($diffs.Count - $preview.Count))
}
exit 1
