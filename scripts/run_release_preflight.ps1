param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),
    [switch]$SkipBuild,
    [switch]$SkipPerfSampling,
    [switch]$LowNoisePerfSampling,
    [int]$PerfRepetitions = 20,
    [int]$PerfWarmupRepetitions = 3,
    [int]$PerfTrimFirstSamples = 1
)

$ErrorActionPreference = 'Stop'

$repo = [System.IO.Path]::GetFullPath($RepoRoot)
$buildScript = Join-Path $repo 'scripts/build_release.ps1'
$wildcardScript = Join-Path $repo 'scripts/run_wildcard_regression.ps1'
$textServiceScript = Join-Path $repo 'scripts/run_textservice_behavior_regression.ps1'
$perfScript = Join-Path $repo 'scripts/run_perf_sampling.ps1'

foreach ($required in @($wildcardScript, $textServiceScript, $perfScript, $buildScript)) {
    if (-not (Test-Path $required)) {
        throw "Required script not found: $required"
    }
}

Push-Location $repo
try {
    if (-not $SkipBuild) {
        Write-Host '[preflight] step 1/4: build_release.ps1'
        & powershell -NoProfile -ExecutionPolicy Bypass -File $buildScript
        if ($LASTEXITCODE -ne 0) {
            throw "Build failed in preflight, exit=$LASTEXITCODE"
        }
    } else {
        Write-Host '[preflight] step 1/4: build skipped by -SkipBuild'
    }

    Write-Host '[preflight] step 2/4: wildcard regression'
    & powershell -NoProfile -ExecutionPolicy Bypass -File $wildcardScript
    if ($LASTEXITCODE -ne 0) {
        throw "Wildcard regression failed, exit=$LASTEXITCODE"
    }

    Write-Host '[preflight] step 3/4: textservice behavior regression'
    & powershell -NoProfile -ExecutionPolicy Bypass -File $textServiceScript
    if ($LASTEXITCODE -ne 0) {
        throw "TextService behavior regression failed, exit=$LASTEXITCODE"
    }

    if ($SkipPerfSampling) {
        Write-Host '[preflight] step 4/4: perf sampling skipped by -SkipPerfSampling'
    } else {
        Write-Host '[preflight] step 4/4: perf sampling'
        $perfArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $perfScript, '-Repetitions', $PerfRepetitions)
        if ($LowNoisePerfSampling) {
            $perfArgs += @('-LowNoise', '-WarmupRepetitions', $PerfWarmupRepetitions, '-TrimFirstSamples', $PerfTrimFirstSamples)
        }

        & powershell @perfArgs
        if ($LASTEXITCODE -ne 0) {
            throw "Perf sampling failed, exit=$LASTEXITCODE"
        }
    }

    Write-Host '[preflight] all checks passed.'
}
finally {
    Pop-Location
}
