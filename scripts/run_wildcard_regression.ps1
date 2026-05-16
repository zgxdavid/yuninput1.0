param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$DictRelativePath = 'data/zhengma-all.dict',
    [int]$TopN = 12,
    [switch]$AllowProbeErrors,
    [string]$OutputRelativePath = 'diagnostics/wildcard_regression_latest.tsv',
    [string[]]$AllowedPatterns = @('a0cd', 'abc0', 'a0c0'),
    [string[]]$RejectedPatterns = @('ab0d', '0bcd', 'ab00', 'a00d', '0b0d', 'abcd', 'abc')
)

$ErrorActionPreference = 'Stop'

chcp 65001 | Out-Null
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

function Resolve-RepoPath([string]$basePath, [string]$relativePath) {
    return [System.IO.Path]::GetFullPath((Join-Path $basePath $relativePath))
}

$expectedExpandedCountByPattern = @{
    'a0cd' = 26
    'abc0' = 26
    'a0c0' = 676
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

function Invoke-WildcardProbe([string]$probePath, [string]$dictPath, [string]$pattern, [int]$topN) {
    $rawOutput = @(& $probePath $dictPath $pattern $topN 2>&1)
    $exitCode = $LASTEXITCODE
    $lines = @($rawOutput | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.ToString().Trim() })

    $summaryLine = $lines | Where-Object { $_.StartsWith('SUMMARY' + "`t") } | Select-Object -Last 1
    $summary = $null
    if (-not [string]::IsNullOrWhiteSpace($summaryLine)) {
        $parts = $summaryLine -split "`t"
        if ($parts.Count -ge 6) {
            $summary = [PSCustomObject]@{
                Pattern = $parts[1]
                Status = $parts[2]
                ExpandedCount = [int]$parts[3]
                MatchedCodes = [int]$parts[4]
                TotalCandidates = [int]$parts[5]
            }
        }
    }

    return [PSCustomObject]@{
        ExitCode = $exitCode
        Lines = $lines
        Summary = $summary
    }
}

$repo = [System.IO.Path]::GetFullPath($RepoRoot)
$probePath = Resolve-RepoPath $repo 'build/Release/yuninput_wildcard_probe.exe'
$dictPath = Resolve-RepoPath $repo $DictRelativePath
$outputPath = Resolve-RepoPath $repo $OutputRelativePath

if (-not (Test-Path $probePath)) {
    throw "Probe not found: $probePath"
}
if (-not (Test-Path $dictPath)) {
    throw "Dictionary not found: $dictPath"
}

$stageRoot = Join-Path ([System.IO.Path]::GetTempPath()) 'yuninput_probe_stage'
$probePath = Copy-ToLocalStage $probePath (Join-Path $stageRoot 'bin')
$dictPath = Copy-ToLocalStage $dictPath (Join-Path $stageRoot 'data')

$outputDir = Split-Path $outputPath -Parent
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

$rows = New-Object System.Collections.Generic.List[string]
$rows.Add('pattern`tstatus`texit_code`texpanded_count`tmatched_codes`ttotal_candidates`tnote')
$probeErrors = New-Object System.Collections.Generic.List[object]
$failures = New-Object System.Collections.Generic.List[string]

foreach ($pattern in $AllowedPatterns) {
    $probeResult = Invoke-WildcardProbe $probePath $dictPath $pattern $TopN
    if ($null -eq $probeResult.Summary) {
        $rows.Add("$pattern`tprobe-error`t$($probeResult.ExitCode)`t0`t0`t0`tmissing-summary")
        $probeErrors.Add([PSCustomObject]@{ Pattern = $pattern; ExitCode = $probeResult.ExitCode; Output = ($probeResult.Lines -join ' | '); Reason = 'missing-summary' }) | Out-Null
        continue
    }

    $summary = $probeResult.Summary
    $expectedExpandedCount = $expectedExpandedCountByPattern[$pattern]
    $isValid =
        $probeResult.ExitCode -eq 0 -and
        $summary.Status -eq 'VALID' -and
        ($expectedExpandedCount -eq $null -or $summary.ExpandedCount -eq $expectedExpandedCount) -and
        $summary.MatchedCodes -gt 0

    $status = if ($isValid) { 'allowed' } else { 'allowed-failed' }
    $rows.Add("$pattern`t$status`t$($probeResult.ExitCode)`t$($summary.ExpandedCount)`t$($summary.MatchedCodes)`t$($summary.TotalCandidates)`tallowed-pattern")
    if (-not $isValid) {
        $failures.Add("allowed pattern failed: $pattern (status=$($summary.Status), exit=$($probeResult.ExitCode), expanded=$($summary.ExpandedCount), matched=$($summary.MatchedCodes))") | Out-Null
    }
}

foreach ($pattern in $RejectedPatterns) {
    $probeResult = Invoke-WildcardProbe $probePath $dictPath $pattern $TopN
    if ($null -eq $probeResult.Summary) {
        $rows.Add("$pattern`tprobe-error`t$($probeResult.ExitCode)`t0`t0`t0`tmissing-summary")
        $probeErrors.Add([PSCustomObject]@{ Pattern = $pattern; ExitCode = $probeResult.ExitCode; Output = ($probeResult.Lines -join ' | '); Reason = 'missing-summary' }) | Out-Null
        continue
    }

    $summary = $probeResult.Summary
    $isRejected = $summary.Status -eq 'INVALID'
    $status = if ($isRejected) { 'rejected' } else { 'unexpected-accept' }
    $rows.Add("$pattern`t$status`t$($probeResult.ExitCode)`t$($summary.ExpandedCount)`t$($summary.MatchedCodes)`t$($summary.TotalCandidates)`trejected-pattern")
    if (-not $isRejected) {
        $failures.Add("rejected pattern unexpectedly accepted: $pattern (status=$($summary.Status), exit=$($probeResult.ExitCode))") | Out-Null
    }
}

Set-Content -Path $outputPath -Value $rows -Encoding UTF8
Write-Host "Wildcard regression snapshot written: $outputPath"
Write-Host "Probe executable: yuninput_wildcard_probe.exe"
Write-Host "Wildcard legality baseline: abc0 is LEGAL (wildcard at 4th position)."
Write-Host "Wildcard legality baseline: ab0d is INVALID (wildcard at 3rd position)."
Write-Host ("Probe errors: {0}" -f $probeErrors.Count)

if ($probeErrors.Count -gt 0 -and -not $AllowProbeErrors) {
    Write-Host ("Wildcard regression aborted: {0} probe execution errors." -f $probeErrors.Count)
    foreach ($probeError in ($probeErrors | Select-Object -First 20)) {
        Write-Host ("PROBE-ERROR pattern={0} exit={1} output={2}" -f $probeError.Pattern, $probeError.ExitCode, $probeError.Output)
    }
    exit 2
}

if ($failures.Count -eq 0) {
    Write-Host 'Wildcard regression check passed.'
    exit 0
}

Write-Host ("Wildcard regression check failed: {0} issues." -f $failures.Count)
foreach ($failure in $failures) {
    Write-Host ("FAIL {0}" -f $failure)
}
exit 1