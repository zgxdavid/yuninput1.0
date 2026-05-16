param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$DictRelativePath = 'data/zhengma-all.dict',
    [int]$Repetitions = 20,
    [switch]$LowNoise,
    [int]$WarmupRepetitions = 3,
    [int]$TrimFirstSamples = 1,
    [int]$InsertCount = 1000,
    [string]$OutputTsvRelativePath = 'diagnostics/perf_sampling_latest.tsv',
    [string]$OutputReportRelativePath = 'diagnostics/perf_sampling_report_latest.md'
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

    if (-not (Test-Path $stageRoot)) {
        New-Item -ItemType Directory -Path $stageRoot -Force | Out-Null
    }

    $targetPath = Join-Path $stageRoot ([System.IO.Path]::GetFileName($sourcePath))
    Copy-Item -Path $sourcePath -Destination $targetPath -Force
    Unblock-File -Path $targetPath -ErrorAction SilentlyContinue
    return $targetPath
}

function Get-Percentile([double[]]$values, [double]$p) {
    if ($null -eq $values -or $values.Count -eq 0) {
        return 0.0
    }

    $sorted = @($values | Sort-Object)
    if ($sorted.Count -eq 1) {
        return [double]$sorted[0]
    }

    $rank = ($p / 100.0) * ($sorted.Count - 1)
    $lower = [int][Math]::Floor($rank)
    $upper = [int][Math]::Ceiling($rank)
    if ($lower -eq $upper) {
        return [double]$sorted[$lower]
    }

    $fraction = $rank - $lower
    return ([double]$sorted[$lower]) + (([double]$sorted[$upper] - [double]$sorted[$lower]) * $fraction)
}

function Build-SampleStats([double[]]$samples) {
    if ($null -eq $samples -or $samples.Count -eq 0) {
        return [PSCustomObject]@{
            MinMs = 0.0
            MaxMs = 0.0
            AvgMs = 0.0
            MedianMs = 0.0
            P95Ms = 0.0
            StdDevMs = 0.0
            Count = 0
        }
    }

    $minMs = ($samples | Measure-Object -Minimum).Minimum
    $maxMs = ($samples | Measure-Object -Maximum).Maximum
    $avgMs = ($samples | Measure-Object -Average).Average
    $medianMs = Get-Percentile $samples 50
    $p95Ms = Get-Percentile $samples 95
    $variance = 0.0
    foreach ($v in $samples) {
        $delta = [double]$v - [double]$avgMs
        $variance += $delta * $delta
    }
    $stdDev = [Math]::Sqrt($variance / [double]$samples.Count)

    return [PSCustomObject]@{
        MinMs = [double]$minMs
        MaxMs = [double]$maxMs
        AvgMs = [double]$avgMs
        MedianMs = [double]$medianMs
        P95Ms = [double]$p95Ms
        StdDevMs = [double]$stdDev
        Count = [int]$samples.Count
    }
}

function Invoke-ProbeWithTiming([string]$exePath, [string[]]$probeArgs, [int]$repetitions) {
    $totalWarmup = if ($LowNoise) { [Math]::Max(0, $WarmupRepetitions) } else { 0 }
    $totalRuns = $repetitions + $totalWarmup
    $allMs = New-Object System.Collections.Generic.List[double]
    $lastLines = @()
    for ($i = 0; $i -lt $totalRuns; $i++) {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $raw = @(& $exePath @probeArgs 2>&1)
        $exitCode = $LASTEXITCODE
        $sw.Stop()
        if ($exitCode -ne 0) {
            $joined = ($raw | ForEach-Object { $_.ToString() }) -join ' | '
            throw "Probe failed: $exePath args=$($probeArgs -join ' ') exit=$exitCode output=$joined"
        }
        if ($i -ge $totalWarmup) {
            $allMs.Add([double]$sw.Elapsed.TotalMilliseconds)
        }
        $lastLines = @($raw | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.ToString().Trim() })
    }

    $effectiveSamples = @([double[]]$allMs.ToArray())
    if ($LowNoise) {
        $dropCount = [Math]::Min([Math]::Max(0, $TrimFirstSamples), [Math]::Max(0, $effectiveSamples.Count - 1))
        if ($dropCount -gt 0) {
            $effectiveSamples = @($effectiveSamples[$dropCount..($effectiveSamples.Count - 1)])
        }
    }
    $stats = Build-SampleStats $effectiveSamples
    $rawStats = Build-SampleStats @([double[]]$allMs.ToArray())

    return [PSCustomObject]@{
        MinMs = [double]$stats.MinMs
        MaxMs = [double]$stats.MaxMs
        AvgMs = [double]$stats.AvgMs
        MedianMs = [double]$stats.MedianMs
        P95Ms = [double]$stats.P95Ms
        StdDevMs = [double]$stats.StdDevMs
        Count = [int]$stats.Count
        RawAvgMs = [double]$rawStats.AvgMs
        RawCount = [int]$rawStats.Count
        WarmupCount = [int]$totalWarmup
        LastLines = $lastLines
    }
}

function Get-WildcardSummaryFromLines([string[]]$lines) {
    $summaryLine = $lines | Where-Object { $_.StartsWith('SUMMARY' + "`t") } | Select-Object -Last 1
    if ([string]::IsNullOrWhiteSpace($summaryLine)) {
        throw 'Missing SUMMARY line in wildcard probe output.'
    }

    $parts = $summaryLine -split "`t"
    if ($parts.Count -lt 6) {
        throw "Malformed SUMMARY line: $summaryLine"
    }

    return [PSCustomObject]@{
        Pattern = $parts[1]
        Status = $parts[2]
        ExpandedCount = [int]$parts[3]
        MatchedCodes = [int]$parts[4]
        TotalCandidates = [int]$parts[5]
    }
}

function Get-InsertSummaryFromLines([string[]]$lines) {
    $summaryLine = $lines | Where-Object { $_.StartsWith('SUMMARY' + "`t") } | Select-Object -Last 1
    if ([string]::IsNullOrWhiteSpace($summaryLine)) {
        throw 'Missing SUMMARY line in index insert probe output.'
    }

    $parts = $summaryLine -split "`t"
    if ($parts.Count -lt 5) {
        throw "Malformed SUMMARY line: $summaryLine"
    }

    return [PSCustomObject]@{
        InsertCount = [int]$parts[1]
        AddedCount = [int]$parts[2]
        ElapsedMs = [double]$parts[3]
        AvgUs = [double]$parts[4]
    }
}

function Get-DictEntryEstimate([string]$dictPath) {
    $count = 0
    foreach ($line in Get-Content -Path $dictPath -Encoding UTF8) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        if ($line.StartsWith('#')) { continue }

        $parts = $line.Trim() -split '\s+'
        if ($parts.Count -lt 2) { continue }

        $first = $parts[0]
        $second = $parts[1]
        $firstIsCode = $first -match '^[A-Za-z]+$'
        $secondIsCode = $second -match '^[A-Za-z]+$'
        if ($firstIsCode -xor $secondIsCode) {
            $count++
        }
    }
    return $count
}

$repo = [System.IO.Path]::GetFullPath($RepoRoot)
$dictPath = Resolve-RepoPath $repo $DictRelativePath
$wildcardProbePath = Resolve-RepoPath $repo 'build/Release/yuninput_wildcard_probe.exe'
$sortProbePath = Resolve-RepoPath $repo 'build/Release/yuninput_sort_probe.exe'
$insertProbePath = Resolve-RepoPath $repo 'build/Release/yuninput_index_insert_probe.exe'
$outputTsvPath = Resolve-RepoPath $repo $OutputTsvRelativePath
$outputReportPath = Resolve-RepoPath $repo $OutputReportRelativePath

foreach ($requiredPath in @($dictPath, $wildcardProbePath, $sortProbePath, $insertProbePath)) {
    if (-not (Test-Path $requiredPath)) {
        throw "Required file not found: $requiredPath"
    }
}

$stageRoot = Join-Path ([System.IO.Path]::GetTempPath()) 'yuninput_perf_stage'
$wildcardProbePath = Copy-ToLocalStage $wildcardProbePath (Join-Path $stageRoot 'bin')
$sortProbePath = Copy-ToLocalStage $sortProbePath (Join-Path $stageRoot 'bin')
$insertProbePath = Copy-ToLocalStage $insertProbePath (Join-Path $stageRoot 'bin')
$dictPath = Copy-ToLocalStage $dictPath (Join-Path $stageRoot 'data')

$rows = New-Object System.Collections.Generic.List[string]
$rows.Add('group`tcase_id`tmetric`tvalue`tnote')

$modeText = if ($LowNoise) {
    "low-noise (warmup=$WarmupRepetitions, trim_first=$TrimFirstSamples)"
} else {
    'normal'
}
$rows.Add("meta`trun	mode	$modeText	sampling mode")
$rows.Add("meta`trun	repetitions	$Repetitions	effective repetitions before trim")

$patterns = @('a0cd', 'abc0', 'a0c0')
$wildcardFastPerCodeLimit = 4
$wildcardFullPerCodeLimit = 8
$wildcardFastExpandedCodeLimit = 96
$wildcardFastTotalQueryBudget = 160
$wildcardFullTotalQueryBudget = 640

$wildcardResults = New-Object System.Collections.Generic.List[object]

foreach ($pattern in $patterns) {
    $fastTiming = Invoke-ProbeWithTiming $wildcardProbePath @($dictPath, $pattern, $wildcardFastPerCodeLimit) $Repetitions
    $fastSummary = Get-WildcardSummaryFromLines $fastTiming.LastLines

    $fullTiming = Invoke-ProbeWithTiming $wildcardProbePath @($dictPath, $pattern, $wildcardFullPerCodeLimit) $Repetitions
    $fullSummary = Get-WildcardSummaryFromLines $fullTiming.LastLines

    if ($fastSummary.Status -ne 'VALID' -or $fullSummary.Status -ne 'VALID') {
        throw "Wildcard pattern is not valid in probe: $pattern"
    }

    $expanded = [int]$fastSummary.ExpandedCount
    $oldFastQueries = [Math]::Min($expanded, $wildcardFastExpandedCodeLimit)
    $oldFastRowsBudget = $oldFastQueries * $wildcardFastPerCodeLimit
    $newFastRowsBudget = [Math]::Min($oldFastRowsBudget, $wildcardFastTotalQueryBudget)

    $oldFullRowsBudget = $expanded * $wildcardFullPerCodeLimit
    $newFullRowsBudget = [Math]::Min($oldFullRowsBudget, $wildcardFullTotalQueryBudget)

    $rows.Add("wildcard`t$pattern`texpanded_codes`t$expanded`tfrom probe summary")
    $rows.Add("wildcard`t$pattern`tfast_avg_ms`t$([Math]::Round($fastTiming.AvgMs, 3))	N=$Repetitions")
    $rows.Add("wildcard`t$pattern`tfast_median_ms`t$([Math]::Round($fastTiming.MedianMs, 3))	effective_n=$($fastTiming.Count)")
    $rows.Add("wildcard`t$pattern`tfast_p95_ms`t$([Math]::Round($fastTiming.P95Ms, 3))	effective_n=$($fastTiming.Count)")
    $rows.Add("wildcard`t$pattern`tfast_stddev_ms`t$([Math]::Round($fastTiming.StdDevMs, 3))	effective_n=$($fastTiming.Count)")
    $rows.Add("wildcard`t$pattern`tfull_avg_ms`t$([Math]::Round($fullTiming.AvgMs, 3))	N=$Repetitions")
    $rows.Add("wildcard`t$pattern`tfull_median_ms`t$([Math]::Round($fullTiming.MedianMs, 3))	effective_n=$($fullTiming.Count)")
    $rows.Add("wildcard`t$pattern`tfull_p95_ms`t$([Math]::Round($fullTiming.P95Ms, 3))	effective_n=$($fullTiming.Count)")
    $rows.Add("wildcard`t$pattern`tfull_stddev_ms`t$([Math]::Round($fullTiming.StdDevMs, 3))	effective_n=$($fullTiming.Count)")
    $rows.Add("wildcard`t$pattern`told_fast_rows_budget`t$oldFastRowsBudget	pre-change estimate")
    $rows.Add("wildcard`t$pattern`tnew_fast_rows_budget`t$newFastRowsBudget	post-change cap")
    $rows.Add("wildcard`t$pattern`told_full_rows_budget`t$oldFullRowsBudget	pre-change estimate")
    $rows.Add("wildcard`t$pattern`tnew_full_rows_budget`t$newFullRowsBudget	post-change cap")

    $wildcardResults.Add([PSCustomObject]@{
        Pattern = $pattern
        Expanded = $expanded
        FastAvgMs = [double]$fastTiming.AvgMs
        FullAvgMs = [double]$fullTiming.AvgMs
        OldFastRowsBudget = $oldFastRowsBudget
        NewFastRowsBudget = $newFastRowsBudget
        OldFullRowsBudget = $oldFullRowsBudget
        NewFullRowsBudget = $newFullRowsBudget
    }) | Out-Null
}

$sortCodes = @('a','aa','ab','b','ba','c','j','jj','jjj','jjjd','jm','ni','nih','nihk','z','zh','zhi')
$sortMs = New-Object System.Collections.Generic.List[double]
foreach ($code in $sortCodes) {
    $timing = Invoke-ProbeWithTiming $sortProbePath @($dictPath, $code, 12) $Repetitions
    $sortMs.Add([double]$timing.AvgMs)
    $rows.Add("sort`t$code`tavg_ms`t$([Math]::Round($timing.AvgMs, 3))	N=$Repetitions")
    $rows.Add("sort`t$code`tmedian_ms`t$([Math]::Round($timing.MedianMs, 3))	effective_n=$($timing.Count)")
    $rows.Add("sort`t$code`tp95_ms`t$([Math]::Round($timing.P95Ms, 3))	effective_n=$($timing.Count)")
}
$sortAvgOfAvg = ($sortMs | Measure-Object -Average).Average
$rows.Add("sort	all	avg_of_avg_ms	$([Math]::Round($sortAvgOfAvg, 3))	across selected codes")

$insertTiming = Invoke-ProbeWithTiming $insertProbePath @($dictPath, $InsertCount) ([Math]::Max(3, [Math]::Min(8, [int]([Math]::Ceiling($Repetitions / 4.0)))))
$insertSummary = Get-InsertSummaryFromLines $insertTiming.LastLines
$rows.Add("index_insert	current	probe_total_ms	$([Math]::Round($insertSummary.ElapsedMs, 3))	insert_count=$InsertCount")
$rows.Add("index_insert	current	probe_avg_us	$([Math]::Round($insertSummary.AvgUs, 3))	insert_count=$InsertCount")
$rows.Add("index_insert	current	probe_run_avg_ms	$([Math]::Round($insertTiming.AvgMs, 3))	repeated process-level timing")

$dictEntryEstimate = Get-DictEntryEstimate $dictPath
$oldPerInsertScanEstimate = [Math]::Max(1, $dictEntryEstimate)
$newPerInsertScanUpperBound = 26 + (26 * 26)
$scanReductionRatio = [double]$oldPerInsertScanEstimate / [double]$newPerInsertScanUpperBound
$rows.Add("index_insert	estimate	old_prefix_scan_units	$oldPerInsertScanEstimate	pre-change RebuildPrefixRanges per insert")
$rows.Add("index_insert	estimate	new_prefix_scan_units	$newPerInsertScanUpperBound	post-change bounded by prefix keyspace")
$rows.Add("index_insert	estimate	scan_unit_reduction_ratio	$([Math]::Round($scanReductionRatio, 3))	old/new")

$outputDir = Split-Path $outputTsvPath -Parent
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}
Set-Content -Path $outputTsvPath -Value $rows -Encoding UTF8

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add('# Performance Sampling Snapshot')
$lines.Add('')
$lines.Add("- GeneratedAt: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$lines.Add("- Dict: $DictRelativePath")
$lines.Add("- Mode: $modeText")
$lines.Add("- Repetitions: $Repetitions")
$lines.Add("- InsertCount: $InsertCount")
$lines.Add('')
$lines.Add('## Wildcard')
foreach ($item in $wildcardResults) {
    $fastRatio = if ($item.OldFastRowsBudget -gt 0) { [Math]::Round([double]$item.NewFastRowsBudget / [double]$item.OldFastRowsBudget, 3) } else { 1.0 }
    $fullRatio = if ($item.OldFullRowsBudget -gt 0) { [Math]::Round([double]$item.NewFullRowsBudget / [double]$item.OldFullRowsBudget, 3) } else { 1.0 }
    $lines.Add("- pattern=$($item.Pattern) expanded=$($item.Expanded) fast_avg_ms=$([Math]::Round($item.FastAvgMs,3)) full_avg_ms=$([Math]::Round($item.FullAvgMs,3))")
    $lines.Add("  - budget_fast old=$($item.OldFastRowsBudget) new=$($item.NewFastRowsBudget) ratio=$fastRatio")
    $lines.Add("  - budget_full old=$($item.OldFullRowsBudget) new=$($item.NewFullRowsBudget) ratio=$fullRatio")
}
$lines.Add('')
$lines.Add('## Sort')
$lines.Add("- avg_of_avg_ms=$([Math]::Round($sortAvgOfAvg,3)) across $($sortCodes.Count) codes")
$lines.Add('')
$lines.Add('## Index Insert')
$lines.Add("- measured probe_avg_us=$([Math]::Round($insertSummary.AvgUs,3)) for insert_count=$InsertCount")
$lines.Add("- estimated old_prefix_scan_units=$oldPerInsertScanEstimate")
$lines.Add("- estimated new_prefix_scan_units=$newPerInsertScanUpperBound")
$lines.Add("- estimated old/new ratio=$([Math]::Round($scanReductionRatio,3))")
$lines.Add('')
$lines.Add('## Notes')
$lines.Add('- wildcard budgets are compared as pre-change estimate vs post-change hard caps.')
$lines.Add('- index insert old/new is complexity-unit estimate because old build is unavailable in this workspace state.')

Set-Content -Path $outputReportPath -Value $lines -Encoding UTF8

Write-Host "Perf sampling TSV written: $outputTsvPath"
Write-Host "Perf sampling report written: $outputReportPath"
